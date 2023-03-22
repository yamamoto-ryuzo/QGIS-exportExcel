# -*- coding: utf-8 -*-
"""
/***************************************************************************
 汎用（単票・単体）調書作成機能
   この機能はベクターレイヤの地物スコープのアクションとして使用することを想定しています。
   アクションフィールドとして以下を使用します。
     [% @layer_id %] -- レイヤID
     [% $id %] -- 地物ID
   
    レイヤ変数に以下の定義をしていおく必要があります。
      xlsout_template_path -- Excelテンプレートブックファイルパス
      xlsout_output_path_fixed -- 帳票出力先ディレクトリパス（必ず存在する必要があります）
      xlsout_output_path_variable -- 出力するファイル名となる地物の属性名
 ***************************************************************************/
"""

import sys
import os
import win32com.client
import tempfile
import datetime
import re

from qgis.core import *
from qgis.gui import *
from qgis.utils import *

from PyQt5.QtCore import Qt, QDate, QTime, QDateTime
from PyQt5.QtWidgets import QAction, QMessageBox, QApplication

# mmへ単位変換する際に使用
unit_for_mm = {
    QgsUnitTypes.DistanceMeters: 1000,
    QgsUnitTypes.DistanceKilometers: 1000000,
    QgsUnitTypes.DistanceFeet: 304.8,
    QgsUnitTypes.DistanceNauticalMiles: 1852000,
    QgsUnitTypes.DistanceYards: 914.4,
    QgsUnitTypes.DistanceMiles: 1609344,
    QgsUnitTypes.DistanceCentimeters: 10,
    QgsUnitTypes.DistanceMillimeters: 1
}

# レイヤ変数定義
TEMPLATE_PATH = "xlsout_template_path"
OUTPUT_PATH_FIXED = "xlsout_output_path_fixed"
OUTPUT_PATH_VARIABLE = "xlsout_output_path_variable"

DEFAULT_CRS = 30167 # 7系

def get_layer_variable(layer_context_name: str, layer: QgsVectorLayer):
    """
    レイヤ変数を取得
    
    Parameters
      layer_context_name -- レイヤ変数名
      layer -- レイヤ

    Returns
      定義されている文字列。
      未定義の場合、メッセージバーにエラーを表示し、空白を返却する。
    """
    if not QgsExpressionContextScope.hasVariable(QgsExpressionContextUtils.layerScope(layer), layer_context_name):
        iface.messageBar().pushCritical("ERROR", f"{layer_context_name} is not defined!")
        return ""

    return QgsExpressionContextScope.variable(QgsExpressionContextUtils.layerScope(layer), layer_context_name)

def create_expression(text: str):
    """
    引数の文字列から新しい式を作成するオブジェクトを作成

    Parameters
      text -- 文字列

    Returns
        引数の文字列から新しい式を作成するオブジェクト。式を含んでいない文字列か、不正な式の場合はNoneを返却する。
    """

    while True:
        exp = QgsExpression(text)

        while True:
            if len(exp.referencedColumns()) > 0:
                # 参照しているフィールドがある
                break

            if len(exp.referencedFunctions()) > 0:
                # 使用している関数がある
                break
        
            if len(exp.referencedVariables()) > 0:
                # 使用している変数がある
                break
            
            # ここまで到達したら単一の文字列なのでNoneを返却する
            return None
        
        if exp.hasEvalError():
            iface.messageBar().pushCritical("ERROR", f"{text} has eval error:{exp.evalErrorString()}")
            return None

        if exp.hasParserError():
            iface.messageBar().pushCritical("ERROR", f"{text} has parser error:{exp.parserErrorString()}")
            return None

        return exp

def variable_based_on_feature(text: str, feature: QgsFeature):
    """
    地物の属性値を使用して文字列を作成

    Parameters
      text -- 文字列
      feature -- 地物

    Returns
      引数の文字列が式を含んでいる場合、地物の属性値等を使用して新しい文字列を作成する。
      式を含んでいない場合は引数の文字列をそのまま返却する。
    """
    exp = create_expression(text)
    if exp is None:
        return text
        
    context = QgsExpressionContext()
    context.setFeature(feature)
    return exp.evaluate(context)

def get_layer_variable_on_feature(layer_context_name: str, layer: QgsVectorLayer, feature: QgsFeature, noerror: bool =True):
    """
    レイヤ変数を取得し、指定地物の属性値から文字列を作成する

    Parameters
      layer_context_name -- レイヤ変数名
      layer -- レイヤ
      feature -- 地物
      noerror -- エラー発生時にエラーメッセージをメッセージバーに表示する場合はTrueを、表示しない場合はFalseを設定

    Returns
      作成した文字列を返却する。
      レイヤ変数が未定義の場合、あるいは式が不正だった場合はNoneを返却。
    """
    ctx = get_layer_variable(layer_context_name, layer)
    if not ctx:
        if noerror==False:
            iface.messageBar().pushCritical("ERROR", f"レイヤ変数 {layer_context_name} が未設定です。")
        return None

    evaluated = variable_based_on_feature(ctx, feature)
    if not evaluated:
        if noerror==False:
            iface.messageBar().pushCritical("ERROR", f"レイヤ変数 {layer_context_name} が正しくありません。")
        return None
    
    return evaluated

def contains_text_in_first_cell(ws, text: str):
    # 最初のセルは結合セルか
    if ws.Cells(1).MergeCells:
        first_cell = ws.Cells(1).MergeArea.Cells(1)
    else:
        first_cell =  ws.Cells(1)

    first_cell_value = first_cell.Value
    if first_cell_value is None:
        return None

    first_cell_text = first_cell.Text
    if text in first_cell_text:
        address = first_cell.Address
        return first_cell
    
    return None

def append_attach(ws, cell, attach_list: list):
    cell_text = cell.Text
    attach_info = cell_text.split("::")

    if len(attach_info) > 1:
        text_info = {
            "sheet": ws.Name,
            "address": cell.Address,
        }
        if feature.fieldNameIndex(attach_info[1]) >= 0:
            # 2023.1.19 QDateエラー修正 start
            # text_info["text"] = feature.attribute(attach_info[1])
            _val = feature.attribute(attach_info[1])
            # Excelでエラーに型は文字列に変換する（対象の型：QDate,QTime,QDateTime）
            if isinstance(_val, QDate):
                text_info["text"] = _val.toString("yyyy/MM/dd")
            elif isinstance(_val, QTime):
                text_info["text"] = _val.toString() # フォーマットは不要
            elif isinstance(_val, QDateTime):
                text_info["text"] = _val.toString("yyyy/MM/dd hh:mm:ss")
            else:
                text_info["text"] = _val
            # 2023.1.19 QDateエラー修正 end
        else:
            text_info["text"] = ""
        attach_list.append(text_info)

def find_attach(wb, feature: QgsFeature, attach_list: list):
    """
    指定ブックの全てのシート内から「##Attach::（地物属性名）」文字列を検索し、その情報をリストに登録。
    この処理では地物から属性値をリストに登録するのみでセル値は置換しない。

    Parameters
      wb -- ワークブック
      feature -- 属性値を取得する地物
      list -- 置換情報リスト

    Returns
      なし
    """

    find_text = "##Attach::"

    # 対象ブックのすべてのシートで検索を行う
    for ws in wb.Worksheets:

        first_cell = contains_text_in_first_cell(ws, find_text)
        if first_cell:
            # 左上端セルに対象文字列が見つかった場合findで検出できないのでここでリストに格納する
            append_attach(ws, first_cell, attach_list)

        found_cell = ws.Cells.Find(What=find_text, LookIn=-4163)   # xlValues
        # 見つからなければ次のシートへ
        if not found_cell:
            continue
        # そのシートで見つかった最初のセルのアドレスを退避しておく
        first_cell = found_cell.Address

        while True:
            # 見つかったセルの情報をリストに格納する
            append_attach(ws, found_cell, attach_list)

            # 次のセルを検索
            found_cell = ws.Cells.FindNext(found_cell)
            if not found_cell:
                # 見つからなければ処理を抜ける
                break
            if found_cell.Address == first_cell:
                # 初めのセルに戻ったら処理を抜ける
                break

def append_attach_fit_image(ws, cell, attach_list: list, image_dict: dict, tmpdir, default_scale: int):
    attach_info = cell.Text.split("::")

    map_theme = attach_info[1] if len(attach_info) > 1 else ""
    map_scale = attach_info[2] if len(attach_info) > 2 else default_scale

    if cell.MergeCells:
        # 結合セル
        target_cell = cell.MergeArea
    else:
        # 非結合セル
        target_cell = cell

    # セルの幅、高さを取得する
    # 単位はポイント
    cell_width = target_cell.Width
    cell_height = target_cell.Height

    map_scale_text = '{:0>10}'.format(map_scale)
    cell_width_text = '{:0>5}'.format(cell_width)
    cell_height_text = '{:0>5}'.format(cell_height)
    image_key = f'{map_theme}_{map_scale_text}_{cell_width_text}_{cell_height_text}'

    # 画像ファイル辞書に既存していないか確認
    if image_key in image_dict:
        # 既存していれば辞書からファイルパスを取得
        image_filepath = image_dict.get(image_key)
    else:
        # なければ画像を作成する
        # 一時ファイル名を作成
        dt_now = datetime.datetime.now()
        now_text = dt_now.strftime('%H%M%S%f')
        image_filepath = os.path.normcase(os.path.join(tmpdir.name, f'{image_key}_{now_text}.png'))
            # 画像ファイル辞書に追加
        image_dict[image_key] = image_filepath

    # 画像は後で作成し、貼り付け位置だけ抽出する
    image_info = {
        "type": "image",
        "sheet": ws.Name,
        "address": target_cell.Address,
        "left": target_cell.Left,
        "top": target_cell.Top,
        "width": cell_width,
        "height": cell_height,
        "filepath": image_filepath,
    }

    attach_list.append(image_info)

def find_attach_fit_image(wb, attach_list: list, tmpdir, default_scale: float, dpi: int):
    """
    指定ブックの全てのシート内から「##AttachFitImage::（テーマ名）::（縮尺）」文字列を検索し、
    内容とセルに合わせて地図画像を一時ファイルとして作成する。
    画像情報と貼り付け先の情報をリストに登録する。
    この処理では画像貼り付けは行わない。

    Parameters
      wb -- ワークブック
      attach_list -- 地図画像貼り付け情報リスト
      tmpdir -- 地図画像を作成する一時フォルダ既定縮尺
      default_scale -- 既定縮尺
      dpi -- DPI

    Returns
      なし
    """

    # 画像作成が重複しないように画像ファイル辞書を作成する
    image_dict = {}
    find_text = "##AttachFitImage::"

    # 対象ブックのすべてのシートで検索を行う
    for ws in wb.Worksheets:
        first_cell = contains_text_in_first_cell(ws, find_text)
        if first_cell:
            # 左上端セルに対象文字列が見つかった場合findで検出できないのでここでリストに格納する
            append_attach_fit_image(ws, first_cell, attach_list, image_dict, tmpdir, default_scale)

        found_cell = ws.Cells.Find(What=find_text, LookIn=-4163)   # xlValues
        # 見つからなければ次のシートへ
        if not found_cell:
            continue
        # そのシートで見つかった最初のセルのアドレスを退避しておく
        first_cell = found_cell.Address

        while True:
            append_attach_fit_image(ws, found_cell, attach_list, image_dict, tmpdir, default_scale)

            # 次のセルを検索
            found_cell = ws.Cells.FindNext(found_cell)
            if not found_cell:
                # 見つからなければ処理を抜ける
                break
            if found_cell.Address == first_cell:
                # 初めのセルに戻ったら処理を抜ける
                break

    # 一括して地図画像ファイルを作成する
    create_map_images(image_dict, dpi)

def create_map_images(image_dict: dict, dpi: int):
    """
    地図画像ファイルを一時フォルダに作成する。

    Parameters
      image_dict -- 地図画像情報辞書。
                    キーは「（テーマ名）_（縮尺（10桁のゼロ埋め））_（幅（5桁のゼロ埋め））_（高さ（5桁のゼロ埋め））」。
                    値はファイルパス
      dpi -- DPI

    Returns
      なし
    """

    # 引数の辞書をテーマ順、縮尺順にソートする
    key_sorted = sorted(image_dict)

    map_canvas = iface.mapCanvas()
    # 現在の縮尺を退避する
    saved_scale = map_canvas.scale()
    # 現在の状態で仮のテーマを作成する
    temp_theme_name = add_temp_theme_from_current_state()

    # テーマを切り替えながら地図画像を作成する
    current_scale = 0
    
    for key in key_sorted:
        key_split = key.split("_")
        theme_name = key_split[0]
        map_scale = float(key_split[1])
        map_width = float(key_split[2])
        map_height = float(key_split[3])

        # テーマを変更する
        if change_theme(theme_name):
            current_theme = theme_name
        else:
            current_theme = ""
        
        if current_scale != map_scale:
            # 縮尺を変更する
            map_canvas.zoomScale(map_scale)
            current_scale = map_scale
        
        create_map_image(image_dict.get(key), map_scale, map_width, map_height, dpi, current_theme)

    # 状態を復元し仮のテーマを削除する
    restore_from_temp_theme(temp_theme_name, saved_scale)


def create_map_image(filepath: str, map_scale: float, pt_width: float, pt_height: float, dpi: int, theme_name:str):
    """
    地図画像を作成する

    Parameters
      filepath -- 画像ファイルパス
      map_scale -- 縮尺
      pt_width -- 画像の幅
      pt_height -- 画像の高さ
      dpi -- DPI

    Returns
      なし
    """

    # マップキャンバス
    map_canvas = iface.mapCanvas()
    # 現在の中心座標
    center = map_canvas.center()

    # 現在の縮尺
    current_scale = map_canvas.scale()

    # 引数の幅と高さはポイント単位なのでピクセルに変換する
    inch_width = pt_width / 72
    inch_height = pt_height / 72
    px_width = inch_width * dpi
    px_height = inch_height * dpi

    adjuster = unit_for_mm.get(map_canvas.mapUnits(), 1)

    # インチからmmへ変換
    extent_width = inch_width * 25.4 / adjuster
    extent_height = inch_height * 25.4 / adjuster

    extent_width *= map_scale
    extent_height *= map_scale

    extent_width /= 2
    extent_height /= 2

    # 2023.1.27 測地系の場合のエラー修正 start
    before_crs = iface.mapCanvas().mapSettings().destinationCrs().authid().replace("EPSG:", "")
    if map_canvas.mapUnits() == QgsUnitTypes.DistanceDegrees:
        xy_flg = center.x() > center.y()
        if xy_flg:
            center_geometry = QgsGeometry.fromPointXY(QgsPointXY(center.x(), center.y()))
        else:
            center_geometry = QgsGeometry.fromPointXY(QgsPointXY(center.y(), center.x())) # xyは逆

        coordwgs84 = QgsCoordinateReferenceSystem(int(before_crs))
        coordweb = QgsCoordinateReferenceSystem(DEFAULT_CRS)
        trans =  QgsCoordinateTransform()
        trans.setSourceCrs(coordwgs84)
        trans.setDestinationCrs(coordweb)
        center_geometry.transform(trans)
        center = center_geometry.asPoint()

        # 中心座標から変換した幅と高さの範囲を産出する
        extent = QgsRectangle(
            center.x() - extent_width/1000,
            center.y() - extent_height/1000,
            center.x() + extent_width/1000,
            center.y() + extent_height/1000
        )

        rect_geom = QgsGeometry.fromRect(extent)

        trans.setSourceCrs(coordweb)
        trans.setDestinationCrs(coordwgs84)
        rect_geom.transform(trans)

        rect_polygon = rect_geom.asPolygon()[0]
        max_x = rect_polygon[0].x()
        min_x = rect_polygon[0].x()
        max_y = rect_polygon[0].y()
        min_y = rect_polygon[0].y()

        for point in rect_polygon:
            x = point.x()
            y = point.y()

            if x > max_x:
                max_x = x
            if y > max_y:
                max_y = y
            if x < min_x:
                min_x = x
            if y < min_y:
                min_y = y

        if xy_flg:
            extent = QgsRectangle(min_x, min_y, max_x, max_y)
        else:
            # xyは逆にして生成
            extent = QgsRectangle(min_y, min_x, max_y, max_x)

    # 2023.1.27 測地系の場合のエラー修正 end
    else:
        # 中心座標から変換した幅と高さの範囲を産出する
        extent = QgsRectangle(
            center.x() - extent_width,
            center.y() - extent_height,
            center.x() + extent_width,
            center.y() + extent_height
        )

    # レンダリング設定
    # 2023.1.27 テーマのレイヤ取得 修正 start
    project = QgsProject.instance()
    map_col = project.mapThemeCollection()
    layers = map_col.mapThemeVisibleLayers(theme_name)
    if layers == None or len(layers) == 0:
        layers = map_canvas.layers()
    # 2023.1.27 テーマのレイヤ取得 修正 end

    ms = QgsMapSettings()
    ms.setLayers(layers)
    ms.setOutputDpi(dpi)
    ms.setBackgroundColor(map_canvas.canvasColor())
    ms.setOutputSize(QtCore.QSize(int(px_width), int(px_height))) # 2023.1.27 WARNING回避
    ms.setExtent(extent)

    render = QgsMapRendererParallelJob(ms)

    def finished():
        img = render.renderedImage()
        img.save(filepath, "png")

    render.finished.connect(finished)

    render.start()
    render.waitForFinished()

def change_theme(theme_name: str):
    """
    テーマ変更

      Parameters
       theme_name -- テーマ名

      Returns
        テーマ変更が正常に完了したらTrue、失敗したらFalseを返却する
    """
    project = QgsProject.instance()
    col = project.mapThemeCollection()

    if not col.hasMapTheme(theme_name):
        # 現プロジェクトに対象の名前のテーマが存在しない場合は実行前の状態で行う
        return False

    # 現プロジェクトのテーマコレクションから対象のテーマで地図表示を変更する
    root = project.layerTreeRoot()
    model = iface.layerTreeView().layerTreeModel()
    col.applyTheme(theme_name, root, model)

    # テーマ変更正常終了を返却する
    return True

def add_temp_theme_from_current_state():
    """
    現在の状態で仮のテーマを作成する

      Returns
        仮のテーマ名
    """

    project = QgsProject.instance()
    col = project.mapThemeCollection()
    
    # 現在の状態から仮のテーマを作成する
    root = project.layerTreeRoot()
    model = iface.layerTreeView().layerTreeModel()
    current_state = col.createThemeFromCurrentState(root, model)
    
    # 仮のテーマの名前を作成
    dt_now = datetime.datetime.now()
    now_text = dt_now.strftime('%H%M%S%f')
    theme_name = f"temptheme_{now_text}"
    col.insert(theme_name, current_state)

    return theme_name

def restore_from_temp_theme(temp_theme_name: str, map_scale: float):
    """
    指定の仮のテーマを適用後、削除する

      Parameters
        temp_theme_name -- 仮のテーマ名
        map_scale -- 縮尺
    """

    # 仮のテーマで復元する
    change_theme(temp_theme_name)
    # 仮のテーマを削除する
    QgsProject.instance().mapThemeCollection().removeMapTheme(temp_theme_name)
    # 縮尺を復元する
    iface.mapCanvas().zoomScale(map_scale)

def replace_attach(wb, attach_list: list):
    """
    指定ブックの全シートの該当セルの値をリストの内容で置換する

      Parameters
        wb -- ワークブック
        attach_list -- 置換情報リスト

      Returns
        True:正常, False:エラーがあり強制終了
    """
    
    # 2023.1.19 QDateエラー修正 start
    ret_flg = True

    current_sheet = ""
    for info in attach_list:
        if current_sheet != info.get("sheet"):
            current_sheet = info.get("sheet")
            ws = wb.Worksheets(current_sheet)
        
        _address = info.get("address")
        _str = info.get("text")

        try:
            # セルに置換した文字列を代入
            if _str:
                ws.Range(_address).Value = _str
            else:
                ws.Range(_address).ClearContents()

        except Exception as e:
            iface.messageBar().pushCritical("ERROR", f"セル【{_address}】に【{str(_str)}】は出力できませんでした。")
            ret_flg = False
            break
   
    return ret_flg
    # 2023.1.19 QDateエラー修正 end

def insert_images(wb, attach_image_list: list):
    """
    指定ブックの全シートの該当セルにリストの内容に従い画像を挿入する

      Parameters
        wb -- ワークブック
        attach_image_list -- list[dict] 挿入する画像の情報辞書のリスト
         辞書の内容は以下の通り
           sheet -- シート名
           filepath -- 画像ファイルパス
           width -- 画像の幅（ポイント）
           height -- 画像の高さ（ポイント）
           address -- セルのアドレス
    """

    current_sheet = ""
    for info in attach_image_list:

        if current_sheet != info.get("sheet"):
            current_sheet = info.get("sheet")
            ws = wb.Worksheets(current_sheet)
        
        # 作成した画像を挿入
        filepath = info.get("filepath")
        width = info.get("width")
        height = info.get("height")
        if os.path.exists(filepath):
            ws.Range(info.get("address")).ClearContents()
            # 指定の縮尺、サイズで作成した地図画像を挿入
            ws.Shapes.AddPicture(filepath, False, True, info.get("left"), info.get("top"), width, height)

def check_save_folder(path: str):
    """
    保存ファイルパスのディレクトリが存在しなければ再帰的に作成する

      Parameters
        path -- 保存ファイルパス
    """

    # ディレクトリ部分だけ抜き出す
    target = os.path.dirname(path)

    # 再帰的にディレクトリを作成
    os.makedirs(target, exist_ok=True)

def output_single_report(excel_app, template_path: str, feature: QgsFeature, output_path: str, dpi: int):
    """
    対象地物の情報で単票形式のExcelブックを作成、保存する

      Parameters
        excel_app -- Excelアプリケーション
        template_path -- テンプレートファイルパス
        feature -- 地物
        output_path -- 保存ファイルパス
        dpi -- DPI

      Returns
        正常終了の場合はTrueを、それ以外はFalseを返却する
    """

    success = True

    attach_list = []
    attach_image_list = []

    # 現在の縮尺を既定縮尺にする
    default_scale = iface.mapCanvas().scale()

    # 画像を作成しておく一時フォルダを作成する
    tmpdir = tempfile.TemporaryDirectory()

    try:
        # Excelで指定テンプレートで新規ブックを作成する
        wb = excel_app.Workbooks.Add(template_path)
        wb.Activate

        # 全シートから##Attach::を見つける
        find_attach(wb, feature, attach_list)

        # 全シートから##AttachFitImage::を見つける
        find_attach_fit_image(wb, attach_image_list, tmpdir, default_scale, dpi)

        # 全シートの##Attach::を置換
        ret = replace_attach(wb, attach_list)
        if(ret == False):
            return False

        # 全シートの##AttachFitImage::に画像を挿入
        insert_images(wb, attach_image_list)

        if overwrite == False:
            # フォルダが存在していなければ作成する
            check_save_folder(output_path)

        excel_app.DisplayAlerts = False
        # ブックを保存する
        wb.SaveAs(output_path, wb.FileFormat)
        excel_app.DisplayAlerts = True

    except Exception as e:
        iface.messageBar().pushCritical("ERROR", f"エラー発生: {e}")
        success = False
        
    finally:
        # 使い終わった画像ファイルを一時フォルダごと開放する
        tmpdir.cleanup()

    return success

def check_same_name_opend(excel_app, book_name: str):
    """
    同名のブックが現在開かれていないか確認する

      Parameters
        excel_app -- Excelアプリケーション
        book_name -- ブック名

      Returns
        既に開かれていた場合エラーをメッセージバーに表示してFalseを返却する。それ以外はTrueを返却する
    """

    # 同名のブックが開いていないか確認する
    for wb in excel_app.Workbooks:
        if wb.Name == book_name:
            iface.messageBar().pushCritical("ERROR", f"同じ名前のブック{book_name} \n が既に開かれています。")
            return False
    
    return True

excel_result = False
while True:
    # レイヤーを取得
    layer = QgsProject.instance().mapLayer('[% @layer_id %]')
    # 該当地物を取得
    fid = int('[% $id %]')
    feature = layer.getFeature(fid)

    # 現在の位置と縮尺を退避する
    map_canvas = iface.mapCanvas()
    save_center = map_canvas.center()
    save_scale = map_canvas.scale()

    # レイヤ変数からテンプレートパス、出力パス、dpiを取得作成する
    dpi_raw = get_layer_variable("dpi", layer)
    if dpi_raw:
        # dpiがレイヤ変数に正しく設定されていたら数値変換する
        if not dpi_raw.isdecimal():
            iface.messageBar().pushCritical("ERROR", f"レイヤ変数 dpi が正しくありません。")
            break
        dpi = int(dpi_raw)
    else:
        dpi = 200
        QgsMessageLog.logMessage(f"dpi={dpi}として処理します。")

    # テンプレートファイルパス取得
    template_path = get_layer_variable_on_feature(TEMPLATE_PATH, layer, feature)
    if not template_path:
        break
    
    # 出力ファイルディレクトリ（ユーザー指定部）パス取得
    output_path_fixed = get_layer_variable_on_feature(OUTPUT_PATH_FIXED, layer, feature)
    if not output_path_fixed:
        iface.messageBar().pushCritical("ERROR", f"出力先フォルダが設定されていません。")
        break
    # 出力ファイルディレクトリ（可変部）パス取得
    output_path_variable = get_layer_variable_on_feature(OUTPUT_PATH_VARIABLE, layer, feature)
    if not output_path_variable:
        iface.messageBar().pushCritical("ERROR", f"出力先フォルダ(可変部)が設定されていません。")
        break
    
    # テンプレートファイルの存在チェック
    if not os.path.isabs(template_path):
        template_path = os.path.join(os.path.dirname(QgsProject.instance().fileName()), template_path)
    else:
        # 絶対パスの場合、先頭が\\\\では認識しないため、\\に変換する
        template_path = template_path.replace("\\\\\\\\", "\\\\")

    if os.path.exists(template_path) == False:
        iface.messageBar().pushCritical("ERROR", f"テンプレートファイル\n {template_path} \nが見つかりません。")
        break

    # 出力先フォルダ（ユーザー指定部）の存在チェック
    if not os.path.isabs(output_path_fixed):
        # 相対パス
        projectpath  = QgsProject.instance().fileName(); 
        projectdirpath = os.path.dirname(str(projectpath)) ; 
        output_path_fixed = os.path.abspath(os.path.join(projectdirpath, output_path_fixed))
    else:
        # 絶対パスの場合、先頭が\\\\では認識しないため、\\に変換する
        output_path_fixed = output_path_fixed.replace("\\\\\\\\", "\\\\")
    
    if not os.path.exists(output_path_fixed):
        iface.messageBar().pushCritical("ERROR", f"出力先フォルダ\n {output_path_fixed} \nは存在していません。")
        break

    # 出力ファイルパス作成
    output_path = os.path.join(output_path_fixed, output_path_variable)

    # 出力先ファイルの存在チェック
    overwrite = False
    if os.path.exists(output_path):
        msgbox_return = QMessageBox.question(iface.mainWindow(), "確認", f"{output_path} \nは既に存在しています。\n上書きしますか。", QMessageBox.Ok, QMessageBox.Cancel)
        if msgbox_return == QMessageBox.Cancel:
            break
        overwrite = True

    with OverrideCursor(Qt.WaitCursor):
        # 中央に対象地物を表示する
        map_canvas.panToFeatureIds(layer, [feature.id()])
        map_canvas.refresh()

        # Excelを開く
        try:
            excel_app = win32com.client.GetObject(Class="Excel.Application")
        except:
            excel_app = win32com.client.Dispatch("Excel.Application")

        # 同名のブックが開いていないか確認
        book_name = os.path.basename(output_path)
        if check_same_name_opend(excel_app, book_name):
            # 更新停止
            excel_app.ScreenUpdating = False

            # 単票出力
            excel_result = output_single_report(excel_app, template_path, feature, output_path, dpi)

            # 更新再開
            excel_app.ScreenUpdating = True

        excel_app.Visible = True
        
        if excel_app.Workbooks.Count == 0:
            excel_app.Quit()

    # 永久ループを抜ける
    break

# 縮尺と位置を元に戻す
flg_refresh = False
if map_canvas.center().x() != save_center.x() or map_canvas.center().y() != save_center.y():
    map_canvas.setCenter(save_center)
    flg_refresh = True
if map_canvas.scale() != save_scale:
    map_canvas.zoomScale(save_scale)
    flg_refresh = True
if flg_refresh:
    map_canvas.refresh()
    
if excel_result:
    QMessageBox.information(iface.mainWindow(), "情報", f"{output_path}を出力しました。")