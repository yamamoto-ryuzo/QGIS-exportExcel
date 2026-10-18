# -*- coding: utf-8 -*-
"""geo_excel を AI/外部制御から利用するためのブリッジ。

plugin.bridge 属性にこのオブジェクトをぶら下げると、qgis_live_driver の
bridge_call コマンド(ファイルJSONL / REST /api/cmd / MCP qgis_bridge_call)
から ``qgis.utils.plugins['geo_excel'].bridge.<method>(*args)`` が
プロセス内で直接呼ばれる。

各メソッドは JSON 文字列を返す(bridge_call 側で構造化される)。
ドメインエラー(スクリプト不明・前提未充足等)は ``{"ok": false, "error": ...}``
を返し、予期しない例外は呼び出し元(driver の {"error": traceback})に委ねる。

run() は GUI スレッドで同期的に実行され、Excel起動・帳票生成の時間分だけ
ブロックする(メニュー実行と同じ性質)。スクリプト内の modal QMessageBox
(上書き確認/続行確認/完了通知)は実行中だけ自動応答に差し替える:
confirm=False なら確認系は Cancel(=中断)として安全側に倒す。
また iface.messageBar().pushCritical の呼出は実行中だけ収集して
messages として返す(メッセージバー自体にも表示される)。

例(MCP): qgis_bridge_call pluginId="geo_excel" method="run" args=["list"]
"""

import inspect
import json
import os

from qgis.PyQt.QtWidgets import QMessageBox
from qgis.core import QgsExpressionContextUtils, QgsProject, QgsVectorLayer

# レイヤ変数の既知キー(scripts側の定義と対応)
_SCRIPT_DEFS = [
    {
        "id": "list",
        "file": "exportList.py",
        "title": "一覧帳票を出力 (アクティブレイヤ)",
        "needs_selection": False,
        "vars": ["xlsout_list_template_path", "xlsout_list_output_path_fixed",
                 "xlsout_list_output_path_variable"],
        "description": "アクティブレイヤ全地物を一覧帳票として出力する",
    },
    {
        "id": "single",
        "file": "exportSingle.py",
        "title": "単票を出力 (選択地物)",
        "needs_selection": True,
        "vars": ["xlsout_template_path", "xlsout_output_path_fixed",
                 "xlsout_output_path_variable", "dpi(任意,既定200)"],
        "description": "選択地物ごとに単票を出力する",
    },
    {
        "id": "bulk",
        "file": "exportSingleBulk.py",
        "title": "単票を一括出力 (選択地物または全地物)",
        "needs_selection": False,
        "vars": ["xlsout_template_path", "xlsout_output_path_fixed",
                 "xlsout_output_path_variable", "dpi(任意,既定200)"],
        "description": "選択地物(2件以上)または全地物の単票を一括出力する",
    },
]


class _DialogAutoReply:
    """exec中だけ QMessageBox の modal メソッドを自動応答に差し替える。

    - question → confirm=True なら Ok(上書き/続行)、False なら Cancel(中断)
    - information/critical/warning → メッセージを収集して即 Ok
    """

    _NAMES = ("information", "question", "warning", "critical")

    def __init__(self, confirm):
        self.confirm = bool(confirm)
        self.messages = []
        self._orig = {}

    def _capture(self, title, text):
        self.messages.append({"title": str(title), "text": str(text)})

    def __enter__(self):
        def _ok(*a):
            self._capture(a[1] if len(a) > 1 else "", a[2] if len(a) > 2 else "")
            return QMessageBox.StandardButton.Ok

        def _question(*a):
            self._capture(a[1] if len(a) > 1 else "", a[2] if len(a) > 2 else "")
            return QMessageBox.StandardButton.Ok if self.confirm else QMessageBox.StandardButton.Cancel

        for name in self._NAMES:
            self._orig[name] = getattr(QMessageBox, name)
            setattr(QMessageBox, name,
                    staticmethod(_question if name == "question" else _ok))
        return self

    def __exit__(self, *exc):
        for name, fn in self._orig.items():
            setattr(QMessageBox, name, fn)
        return False


class GeoExcelBridge:
    """``plugins['geo_excel'].bridge`` にぶら下げるAI向けAPI。"""

    NAME = "Excel帳票出力 (geo_excel)"
    SUMMARY = ("ベクターレイヤの属性情報をExcel帳票として出力する。"
               "アクティブレイヤのレイヤ変数(xlsout_*)でテンプレートと"
               "出力先を指定し、一覧帳票・単票・単票一括の3種を実行できる。"
               "実行には win32com(pywin32) と Excel が必要。")
    WHEN_TO_USE = [
        "アクティブレイヤの全地物をExcel一覧帳票に出力したい",
        "選択した地物の属性をExcel単票テンプレートへ差し込み出力したい",
        "選択地物または全地物の単票を一括出力したい",
        "帳票出力用レイヤ変数(xlsout_*)の設定状況を確認・変更したい",
    ]
    METHOD_DOCS = {
        "run": {
            "description":
                "帳票スクリプトを実行する。scriptは list|single|bulk|"
                "ファイル名|index。singleは選択地物または feature_id 指定。"
                "confirm=True で上書き/続行確認にOk応答(既定False=中断)。"
                "Excel起動・生成中はGUIスレッドをブロックする。",
            "example": {"method": "run", "args": ["list", None, True]},
        },
        "status": {
            "description":
                "アクティブレイヤ・選択地物数・win32com利用可否・"
                "レイヤ変数設定状況を返す。",
            "example": {"method": "status", "args": []},
        },
        "listScripts": {
            "description":
                "実行可能な帳票スクリプト一覧(id/ファイル/要選択/"
                "必要レイヤ変数)を返す。",
            "example": {"method": "listScripts", "args": []},
        },
        "layerVariables": {
            "description":
                "対象レイヤ(既定=アクティブ)の xlsout_*/dpi レイヤ変数を返す。",
            "example": {"method": "layerVariables", "args": []},
        },
        "setLayerVariables": {
            "description":
                "対象レイヤ(既定=アクティブ)にレイヤ変数を設定する。"
                "{変数名: 値}。値Noneで削除。",
            "example": {"method": "setLayerVariables",
                        "args": [{"xlsout_output_path_fixed": "C:/out"}]},
        },
    }

    def __init__(self, plugin):
        self._plugin = plugin

    # ---------- 内部ヘルパ ----------

    @staticmethod
    def _ok(payload=None):
        obj = {"ok": True}
        if payload:
            obj.update(payload)
        return json.dumps(obj, ensure_ascii=False)

    @staticmethod
    def _err(msg):
        return json.dumps({"ok": False, "error": str(msg)},
                          ensure_ascii=False)

    def _layer(self, layer_id=None):
        """対象レイヤを解決する。layer_id省略時はアクティブレイヤ。"""
        if layer_id:
            layer = QgsProject.instance().mapLayer(str(layer_id))
        else:
            layer = self._plugin.iface.activeLayer()
        if not isinstance(layer, QgsVectorLayer):
            return None
        return layer

    @staticmethod
    def _script_entry(script):
        """script指定(id/ファイル名/index)をスクリプト定義に解決する。"""
        if isinstance(script, bool):
            return None
        if isinstance(script, int):
            return (_SCRIPT_DEFS[script]
                    if -len(_SCRIPT_DEFS) <= script < len(_SCRIPT_DEFS)
                    else None)
        s = str(script or "").strip().lower()
        for d in _SCRIPT_DEFS:
            if s in (d["id"], d["file"].lower()):
                return d
        for d in _SCRIPT_DEFS:  # 部分一致フォールバック
            if s and (s in d["file"].lower() or s in d["id"]):
                return d
        return None

    def _method_catalog(self):
        """公開メソッドの仕様一覧(signature+docstring+METHOD_DOCS)。"""
        out = []
        for name in sorted(dir(self)):
            if name.startswith("_") or name == "pluginInfo":
                continue
            fn = getattr(self, name)
            if not callable(fn):
                continue
            doc = inspect.getdoc(fn) or ""
            entry = {
                "name": name,
                "signature": f"{name}{inspect.signature(fn)}",
                "description": doc.split("\n")[0] if doc else "",
            }
            extra = self.METHOD_DOCS.get(name) or {}
            if extra.get("description"):
                entry["description"] = extra["description"]
            if extra.get("example"):
                entry["example"] = extra["example"]
            out.append(entry)
        return out

    def _variables(self, layer):
        scope = QgsExpressionContextUtils.layerScope(layer)
        names = set(scope.variableNames()) if scope is not None else set()
        known = {"dpi"}
        for d in _SCRIPT_DEFS:
            known.update(v for v in d["vars"] if "(" not in v)
        targets = sorted(n for n in names
                         if n.startswith("xlsout") or n == "dpi")
        # 未定義の既知キーも None で出して設定漏れを見やすくする
        out = {}
        for n in sorted(known | set(targets)):
            out[n] = scope.variable(n) if n in names else None
        return out

    # ---------- bridge メソッド(JSON文字列を返す) ----------

    def pluginInfo(self):
        """プラグインの自己記述(id/summary/when_to_use/methods)を返す。"""
        return self._ok({
            "id": "geo_excel",
            "name": self.NAME,
            "summary": self.SUMMARY,
            "when_to_use": list(self.WHEN_TO_USE),
            "invoke": {"type": "bridge_call", "plugin_id": "geo_excel",
                       "method": "<method>", "args": []},
            "methods": self._method_catalog(),
            "state": {
                "scriptCount": len(_SCRIPT_DEFS),
                "win32com": self._win32com_available(),
            },
        })

    def status(self):
        """実行時状態(アクティブレイヤ・選択数・依存環境)を返す。"""
        layer = self._plugin.iface.activeLayer()
        info = {"active_layer": None, "selection": 0,
                "win32com": self._win32com_available()}
        if isinstance(layer, QgsVectorLayer):
            info["active_layer"] = {"id": layer.id(), "name": layer.name()}
            info["selection"] = layer.selectedFeatureCount()
            info["variables"] = self._variables(layer)
        return self._ok(info)

    def listScripts(self):
        """帳票スクリプト一覧を返す。"""
        return self._ok({"scripts": [dict(d, index=i)
                                     for i, d in enumerate(_SCRIPT_DEFS)]})

    def layerVariables(self, layer_id=None):
        """対象レイヤの xlsout_*/dpi レイヤ変数を返す。"""
        layer = self._layer(layer_id)
        if layer is None:
            return self._err("vector layer not found or not active")
        return self._ok({"layer_id": layer.id(), "layer": layer.name(),
                         "variables": self._variables(layer)})

    def setLayerVariables(self, variables, layer_id=None):
        """レイヤ変数を設定する。{変数名: 値}、値Noneで削除。"""
        layer = self._layer(layer_id)
        if layer is None:
            return self._err("vector layer not found or not active")
        if not isinstance(variables, dict):
            return self._err("variables must be an object {name: value}")
        for k, v in variables.items():
            QgsExpressionContextUtils.setLayerVariable(layer, str(k), v)
        return self._ok({"layer_id": layer.id(), "layer": layer.name(),
                         "variables": self._variables(layer)})

    def run(self, script="list", feature_id=None, confirm=False):
        """帳票スクリプトを実行する(結果・メッセージを返す)。"""
        entry = self._script_entry(script)
        if entry is None:
            return self._err(
                f"unknown script: {script} "
                f"(use {[d['id'] for d in _SCRIPT_DEFS]})")
        layer = self._layer()
        if layer is None:
            return self._err("no active vector layer")
        if not self._win32com_available():
            return self._err("win32com (pywin32) is not available")

        if entry["needs_selection"]:
            if feature_id is not None:
                fids = [int(feature_id)]
            else:
                fids = list(layer.selectedFeatureIds())
            if not fids:
                return self._err(
                    "no feature_id and no selected features "
                    f"(layer {layer.name()})")
        else:
            fids = [None]

        results = []
        with _DialogAutoReply(confirm) as dlg:
            msgs = self._capture_messagebar(dlg.messages)
            try:
                for fid in fids:
                    r = self._plugin._exec_script(entry["file"],
                                                  layer.id(), fid)
                    ns = r.get("ns") or {}
                    results.append({
                        "fid": fid,
                        "ok": bool(r.get("ok")),
                        "error": r.get("error"),
                        "executed": bool(ns.get("excel_result")),
                        "output": (ns.get("output_path")
                                   or ns.get("output_path_fixed")),
                    })
            finally:
                msgs()
        return self._ok({
            "script": entry["id"],
            "layer": layer.name(),
            "layer_id": layer.id(),
            "results": results,
            "messages": dlg.messages,
        })

    # ---------- 実行環境 ----------

    @staticmethod
    def _win32com_available():
        try:
            import win32com.client  # noqa: F401
            return True
        except Exception:
            return False

    def _capture_messagebar(self, sink):
        """messageBar().pushCritical を収集に回す差し替えを試みる。

        戻り値は restore 関数。SIPオブジェクトへ属性設定できない環境では
        何もしない(no-op restore)。元のpushCriticalも呼ぶので
        ユーザー表示は維持される。
        """
        try:
            bar = self._plugin.iface.messageBar()
            orig = bar.pushCritical
        except Exception:
            return lambda: None

        def _push(title, text):
            sink.append({"level": "critical",
                         "title": str(title), "text": str(text)})
            try:
                orig(title, text)
            except Exception:
                pass

        try:
            bar.pushCritical = _push
        except Exception:
            return lambda: None

        def _restore():
            try:
                bar.pushCritical = orig
            except Exception:
                pass

        return _restore
