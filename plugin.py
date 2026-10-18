# -*- coding: utf-8 -*-
"""export_excel プラグイン本体。

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

from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.core import QgsVectorLayer

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
            "export_excel", "ベクターレイヤをアクティブにしてください。"
        )
        return None
    return layer


class plugin(object):
    def __init__(self, iface):
        self.iface = iface
        self.actions = []
        self.menu_name = "&export_excel"

    # ---- GUI --------------------------------------------------------------
    def initGui(self):
        entries = (
            ("一覧帳票を出力 (アクティブレイヤ)", "exportList.py", False),
            ("単票を出力 (選択地物)", "exportSingle.py", True),
            ("単票を一括出力 (選択地物または全地物)", "exportSingleBulk.py", False),
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
                    "export_excel", "地物を選択してください。"
                )
                return
            for fid in fids:
                self._exec_script(script_name, layer.id(), fid)
        else:
            self._exec_script(script_name, layer.id(), None)

    def _exec_script(self, script_name, layer_id, fid):
        path = os.path.join(PLUGIN_DIR, script_name)
        with open(path, encoding="utf-8") as f:
            src = f.read()

        # QGISアクションの置換変数を実値で置き換える
        src = src.replace("[% @layer_id %]", layer_id)
        if fid is not None:
            src = src.replace("[% $id %]", str(fid))
        src = _rewrite_source(src)

        try:
            code = compile(src, path, "exec")
            exec(code, {"__name__": f"export_excel.{script_name}", "__file__": path})
        except Exception as e:  # noqa: BLE001 - スクリプト側の失敗をQGISを落とさず通知する
            QMessageBox.critical(
                self.iface.mainWindow(),
                "export_excel",
                f"{script_name} の実行に失敗しました:\n{e}",
            )
