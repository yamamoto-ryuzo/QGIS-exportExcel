# CHANGELOG — export_excel

## 1.0.0

- QGIS-exportExcel (https://github.com/yamamoto-ryuzo/QGIS-exportExcel) の
  レイヤアクション用スクリプト群(exportList/exportSingle/exportSingleBulk)を
  無改変で同梱し、プラグインメニューから実行できる簡易ラッパーとして組み込み。
  - `[% @layer_id %]` はアクティブなベクターレイヤ、`[% $id %]` は選択地物で
    実行時に置換する
  - スクリプト内の `PyQt5` 固定importは exec 直前にソースを `qgis.PyQt` へ
    書き換えて吸収(Qt6対応。QGIS4のimportフックがPyQt5を禁止するため
    sys.modules shim は不可)。`Qt.WaitCursor` も `Qt.CursorShape.WaitCursor`
    へ同時に書き換える(Qt6のスコープenum対応)
- 実行には win32com(pywin32) と Excel が必要(upstreamと同じ前提)
