# -*- coding: utf-8 -*-
"""geo_excel プラグイン本体。

upstream(QGIS-exportExcel)のスクリプトは「レイヤアクション」として実行される
前提で書かれており、QGISアクションの置換変数([% @layer_id %] / [% $id %])を
含む。本プラグインはプラグインメニューからの実行時に、これらをアクティブ
レイヤ/選択地物の値で置き換えて exec することで、スクリプトを無改変で
再利用する。

またスクリプト内の `from PyQt5.…` 固定importは、exec直前にソース文字列を
`from qgis.PyQt.…` へ書き換えて吸収する(QGISのQt6ビルドでは PyQt5 のimportが
qgis.utils._import フックで禁止されるため、sys.modules shim では不可)。
ファイル自体は upstream のまま維持する。

実行には win32com(pywin32) と Excel が必要。
"""

import os
import re

from qgis.PyQt.QtCore import QCoreApplication, QSettings, QTranslator
from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.core import QgsVectorLayer

from .bridge import GeoExcelBridge

PLUGIN_DIR = os.path.dirname(__file__)

# PyQt5 固定importを qgis.PyQt(Qt5/Qt6 自動切替)へ書き換える
_PYQT5_RE = re.compile(r"PyQt5(?=[\.\s])")


def _rewrite_source(src):
    """アクション置換変数と Qt5 固有記法を実行用に書き換える。

    - `PyQt5.` / `from PyQt5 import` → `qgis.PyQt`
    - `Qt.WaitCursor` → `Qt.CursorShape.WaitCursor`
      (Qt6で非スコープenumは廃止。Qt5でもスコープ記法は有効)
    """
    src = _PYQT5_RE.sub("qgis.PyQt", src)
    return src.replace("Qt.WaitCursor", "Qt.CursorShape.WaitCursor")


def _active_vector_layer(iface):
    layer = iface.activeLayer()
    if not isinstance(layer, QgsVectorLayer):
        iface.messageBar().pushCritical(
            "geo_excel",
            QCoreApplication.translate(
                "geo_excel", "Please activate a vector layer."),
        )
        return None
    return layer


class plugin(object):
    def __init__(self, iface):
        self.iface = iface
        self.actions = []
        self.menu_name = "&geo_excel"
        self._init_language()
        # qgis_live_driver の bridge_call / describe_plugins 経由で
        # AI・外部制御から使われるAPI (plugins['geo_excel'].bridge)
        self.bridge = GeoExcelBridge(self)

    @staticmethod
    def tr(message):
        return QCoreApplication.translate("geo_excel", message)

    def _init_language(self):
        """QGISの locale/userLocale に合わせて i18n/geo_excel_<lang>.qm を読む。

        ソース文字列は英語。en または未対応言語では翻訳を載せない。
        """
        self.translator = None
        try:
            raw = QSettings().value("locale/userLocale", None)
            lang = (str(raw).split("_")[0].split("-")[0].lower()
                    if raw else "")
            if not lang or lang == "en":
                return
            qm_path = os.path.join(
                PLUGIN_DIR, "i18n", f"geo_excel_{lang}.qm")
            if os.path.exists(qm_path):
                translator = QTranslator()
                if translator.load(qm_path):
                    QCoreApplication.installTranslator(translator)
                    self.translator = translator
        except Exception:
            pass

    # ---- GUI --------------------------------------------------------------
    def initGui(self):
        entries = (
            (self.tr("Export list report (active layer)"),
             "exportList.py", False),
            (self.tr("Export single form (selected feature)"),
             "exportSingle.py", True),
            (self.tr("Bulk export single forms (selected or all features)"),
             "exportSingleBulk.py", False),
        )
        for title, script, needs_fid in entries:
            action = QAction(title, self.iface.mainWindow())
            action.triggered.connect(
                lambda checked=False, s=script, f=needs_fid: self._run(s, f)
            )
            self.iface.addPluginToMenu(self.menu_name, action)
            self.actions.append(action)

    def unload(self):
        for action in self.actions:
            self.iface.removePluginMenu(self.menu_name, action)
        self.actions = []

    # ---- script runner ----------------------------------------------------
    def _run(self, script_name, needs_fid):
        layer = _active_vector_layer(self.iface)
        if layer is None:
            return

        if needs_fid:
            fids = layer.selectedFeatureIds()
            if not fids:
                self.iface.messageBar().pushCritical(
                    "geo_excel", self.tr("Please select one or more features.")
                )
                return
            for fid in fids:
                self._exec_script(script_name, layer.id(), fid)
        else:
            self._exec_script(script_name, layer.id(), None)

    def _exec_script(self, script_name, layer_id, fid):
        """スクリプトをexecして結果を返す。

        戻り値: {"ok": bool, "error": str|None, "ns": dict}。
        ns にはスクリプトのグローバルが残るため、bridgeは
        excel_result/output_path 等の終了状態をここから取得する。
        """
        path = os.path.join(PLUGIN_DIR, script_name)
        with open(path, encoding="utf-8") as f:
            src = f.read()

        # QGISアクションの置換変数を実値で置き換える
        src = src.replace("[% @layer_id %]", layer_id)
        if fid is not None:
            src = src.replace("[% $id %]", str(fid))
        src = _rewrite_source(src)

        ns = {"__name__": f"geo_excel.{script_name}", "__file__": path}
        try:
            code = compile(src, path, "exec")
            exec(code, ns)
        except Exception as e:  # noqa: BLE001 - スクリプト側の失敗をQGISを落とさず通知する
            QMessageBox.critical(
                self.iface.mainWindow(),
                "geo_excel",
                self.tr("Failed to run {0}:\n{1}").format(script_name, e),
            )
            return {"ok": False, "error": str(e), "ns": ns}
        return {"ok": True, "ns": ns}
