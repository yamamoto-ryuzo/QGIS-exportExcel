# -*- coding: utf-8 -*-
"""export_excel - Excel帳票出力スクリプト群の簡易プラグインラッパー"""


def classFactory(iface):  # pylint: disable=invalid-name
    from .plugin import plugin

    return plugin(iface)
