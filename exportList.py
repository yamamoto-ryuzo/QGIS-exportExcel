# -*- coding: utf-8 -*-
"""
/***************************************************************************
 汎用（一覧）調書作成機能
   この機能はベクターレイヤのレイヤスコープのアクションとして使用することを想定しています。
   アクションフィールドとして以下を使用します。
     [% @layer_id %] -- レイヤID
   
    レイヤ変数に以下の定義をしていおく必要があります。
      xlsout_list_template_path -- Excelテンプレートブックファイルパス
      xlsout_list_output_path_fixed -- 帳票出力先ディレクトリパス（必ず存在する必要があります）
      xlsout_list_output_path_variable -- xlsout_output_path_fixedで指定したディレクトリの下に作成する動的作成するディレクトリ名。
                                     ・QGISで使用できる関数を指定可能
                                     ・文字列はシングルクォーテーションで括る
                                     ・対象地物のフィールド名はダブルクォーテーションで括る
                                     ・パス区切り文字に\を使用する際は\\とする
 ***************************************************************************/

"""
import sys
import os
import win32com.client
import datetime
import re

from qgis.core import *
from qgis.gui import *
from qgis.utils import *

from PyQt5.QtCore import Qt, QDate, QTime, QDateTime
from PyQt5.QtWidgets import QAction, QMessageBox, QApplication

# レイヤ変数定義
TEMPLATE_PATH = "xlsout_list_template_path"
OUTPUT_PATH_FIXED = "xlsout_list_output_path_fixed"
OUTPUT_PATH_VARIABLE = "xlsout_list_output_path_variable"

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

def get_layer_variable_evaluated(layer_context_name: str, layer: QgsVectorLayer):
    """
    指定の名前のレイヤ変数を取得し、式なら評価後の文字列を返却する。
    式でない場合は定義された文字列をそのまま返却する。

      Parameters
        layer_context_name -- レイヤ変数名
        layer -- レイヤ

      Returns
        式なら評価後の文字列を返却し、それ以外はそれ以外は定義された文字列をそのまま返却する
    """

    ctx = get_layer_variable(layer_context_name, layer)
    if not ctx:
        return ""
    exp = create_expression(ctx)
    if exp is None or exp.evaluate() is None:
        # 単一の文字列を返却する
        return ctx
    
    return exp.evaluate()

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

def append_list_insert(ws, cell, list_insert_list, extent):

    merged = False
    if cell.MergeCells:
        target_cell = cell.MergeArea
        cell_text = cell.Cells(1).Text
        merged = True
    else:
        target_cell = cell
        cell_text = cell.Text

    attach_info = cell_text.split("::")

    if len(attach_info) > 1:

        row =  target_cell.Row if merged == False else target_cell.Cells(1).Row
        column = target_cell.Column if merged == False else target_cell.Cells(1).Column
        row_span = target_cell.Rows.Count
        column_span = target_cell.Columns.Count

        if extent["min_row"] == 0 or row < extent["min_row"]:
            extent["min_row"] = row 

        if extent["min_column"] == 0 or column < extent["min_column"]:
            extent["min_column"] = column

        if extent["max_row"] < (row + row_span - 1):
            extent["max_row"] = row + row_span - 1

        if extent["max_column"] < (column + column_span - 1):
            extent["max_column"] = column + column_span - 1

        list_info = {
            "sheet": ws.Name,
            "name": attach_info[1],
            "row": row,
            "column": column,
            "row_span": row_span,
            "column_span": column_span,
        }

        list_insert_list.append(list_info)

def find_list_insert(ws, list_insert_list: list, extent: dict):
    """
    シート内の##ListInsert::を含むセルから一覧の行として繰り返す情報をリストにする。

      Parameters
        ws -- ワークシート
        list_insert_list: 一覧の行として繰り返す情報リスト
        extent: 一覧の行として繰り返す範囲を辞書として設定する
                min_column -- 最小の列番号
                min_row -- 最小の行番号
                max_column -- 最大の列番号
                max_row -- 最大の行番号
    
      Returns
        なし
    """
    
    # 検索対象文字列
    find_text = "##ListInsert::"
    
    # ##ListInsesrt::範囲を決める変数
    extent["min_column"] = 0
    extent["min_row"] = 0
    extent["max_column"] = 0
    extent["max_row"] = 0

    first_cell = contains_text_in_first_cell(ws, find_text)
    if first_cell:
        # 左上端セルに対象文字列が見つかった場合findで検出できないのでここでリストに格納する
        append_list_insert(ws, first_cell, list_insert_list, extent)

    found_cell = ws.Cells.Find(What=find_text, LookIn=-4163)   # xlValues

    # 見つからなければ処理終了
    if not found_cell:
        return

    # シート内で見つかった最初のセルのアドレスを退避しておく
    first_cell = found_cell.Address

    while True:
        append_list_insert(ws, found_cell, list_insert_list, extent)

        # 次のセルを検索
        found_cell = ws.Cells.FindNext(found_cell)
        if not found_cell:
            # 見つからなければ処理を抜ける
            break
        if found_cell.Address == first_cell:
            # 初めのセルに戻ったら処理を抜ける
            break

def insert_list_values(ws, layer: QgsVectorLayer, list_insert_list: list, extent: dict):
    """
    一覧の行として繰り返す情報リストを元に、レイヤの全地物の情報をシートに出力する.

      Parameters
        ws -- ワークシート
        layer -- レイヤ
        list_insert_list -- 一覧の行として繰り返す情報リスト
        extent -- 一覧の行として繰り返す範囲示す辞書
                min_column -- 最小の列番号
                min_row -- 最小の行番号
                max_column -- 最大の列番号
                max_row -- 最大の行番号

      Returns
        なし
    """

    if len(list_insert_list) <= 0:
        return

    data_count = layer.featureCount()
    
    # 繰り返し範囲を決定する
    min_row = extent["min_row"]
    min_column = extent["min_column"]
    max_row = extent["max_row"]
    max_column = extent["max_column"]

    row_span = max_row - min_row + 1
    
    # 繰り返し範囲の書式と数式を行単位でコピー（行高もコピーするため）
    ws.Range(ws.Cells(min_row, min_column), ws.Cells(max_row, max_column)).EntireRow.Copy()

    # 貼り付け
    start_row = max_row + 1
    start_col = min_column
    end_row = max_row + ((data_count - 1) * row_span)
    end_column = max_column
    ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_column)).EntireRow.Select()
    excel_app.Selection.PasteSpecial(Paste=-4104)       # 全てペースト

    excel_app.CutCopyMode = False

    # データ設定範囲に設定されている値を二次元配列（タプル）を取得する
    data_range = ws.Range(ws.Cells(min_row, min_column), ws.Cells(min_row + (data_count * row_span), max_column))

    # データ設定範囲を既に設定されているデータごと二次元配列に格納する
    # Valueを取得するとタプルになるので行列ともにリストに変換する
    cells_tuple = data_range.Value
    cells_list = []
    for row_tuple in cells_tuple:
        cells_list.append(list(row_tuple))

    request = QgsFeatureRequest().setFlags(QgsFeatureRequest.NoGeometry)

    # fidで並び変える
    fid_list = []
    for feature in layer.getFeatures(request):
        fid_list.append(feature.id())
    fid_list.sort()

    # 地物
    for idx, fid in enumerate(fid_list):
        feature = layer.getFeature(fid)
        for target in list_insert_list:
            column = target["column"] - start_col
            row = target["row"] - min_row + (idx * row_span)

            if feature.fieldNameIndex(target["name"]) < 0:
                cells_list[row][column] = None
                continue

            attr_value = feature.attribute(target["name"])

            if attr_value:
                # 2023.1.19 QDateエラー修正 start
                # cells_list[row][column] = str(feature.attribute(target["name"]))

                # Excelでエラーに型は文字列に変換する（対象の型：QDate,QTime,QDateTime）
                if isinstance(attr_value, QDate):
                    cells_list[row][column] = attr_value.toString("yyyy/MM/dd")
                elif isinstance(attr_value, QTime):
                    cells_list[row][column] = attr_value.toString() # フォーマットは不要
                elif isinstance(attr_value, QDateTime):
                    cells_list[row][column] = attr_value.toString("yyyy/MM/dd hh:mm:ss")
                else:
                    cells_list[row][column] = str(attr_value)
                # 2023.1.19 QDateエラー修正 end
            else:
                cells_list[row][column] = None

    cells_list2 = []
    for row_list in cells_list:
        cells_list2.append(tuple(row_list))

    data_range.Value =  tuple(cells_list2)

    ws.Cells(1).Select()

def insert_list(wb, layer: QgsVectorLayer):
    """
    ワークブックの全シートに一覧を作成する
    
      Parameters
        wb -- ワークブック
        layer -- レイヤ

      Returns
        なし
    """

    for ws in wb.Worksheets:
        ws.Select()
        # 対象ブックのすべてのシートで ##ListInsesrt:: を探して一覧を作成する
        list_insert_list = []
        extent = {}

        # このシートの##ListInsesrt::を探す
        find_list_insert(ws, list_insert_list, extent)
        if len(list_insert_list) <= 0:
            # 見つからなければ次のシートへ
            continue
        
        # ##ListInsesrt::の設定に沿って一覧を出力する
        insert_list_values(ws, layer, list_insert_list, extent)

    wb.Worksheets(1).Select()


def check_save_folder(path):
    """
    保存ファイルパスのディレクトリが存在しなければ再帰的に作成する

      Parameters
        path -- 保存ファイルパス
    """

    # ディレクトリ部分だけ抜き出す
    target = os.path.dirname(path)

    # 再帰的にディレクトリを作成
    os.makedirs(target, exist_ok=True)

def output_list_report(excel_app, layer: QgsVectorLayer, template_path: str, output_path: str):
    """
    一覧出力

    Parameters
      excel_app -- Excelアプリケーション
      layer -- レイヤ
      template_path -- テンプレートファイルパス
      output_path -- 出力ファイルパス

    Returns
      一覧出力に成功したらTrueを、失敗したらFalseを返却する。
    """

    success = True

    attach_list = []

    try:
        # Excelで指定テンプレートで新規ブックを作成する
        wb = excel_app.Workbooks.Add(template_path)
        wb.Activate

        # 全シートから##ListInsert::を見つけて一覧を作成する
        insert_list(wb, layer)

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

    # レイヤー変数からテンプレートパス、出力パスを取得作成する
    # テンプレートパス
    template_path = get_layer_variable_evaluated(TEMPLATE_PATH, layer)
    if not template_path:
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

    # 出力パス(ユーザー指定部)
    output_path_fixed = get_layer_variable_evaluated(OUTPUT_PATH_FIXED, layer)
    if not output_path_fixed:
        iface.messageBar().pushCritical("ERROR", f"出力先フォルダが設定されていません。")
        break

    # 出力パス(可変部)
    output_path_variable = get_layer_variable_evaluated(OUTPUT_PATH_VARIABLE, layer)
    if not output_path_variable:
        iface.messageBar().pushCritical("ERROR", f"出力先フォルダ(可変部)が設定されていません。")
        break
    
    # 出力先フォルダの存在チェック
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

    output_path = os.path.join(output_path_fixed, output_path_variable)

    # 出力先ファイルの存在チェック
    overwrite = False
    if os.path.exists(output_path):
        msgbox_return = QMessageBox.question(iface.mainWindow(), "確認", f"{output_path} \nは既に存在しています。\n上書きしますか。", QMessageBox.Ok, QMessageBox.Cancel)
        if msgbox_return == QMessageBox.Cancel:
            break
        overwrite = True
        
    with OverrideCursor(Qt.WaitCursor):
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

            # 一覧作成
            excel_result = output_list_report(excel_app, layer, template_path, output_path)

            # 更新再開
            excel_app.ScreenUpdating = True

        excel_app.Visible = True
        
        if excel_app.Workbooks.Count == 0:
            excel_app.Quit()

    # 永久ループを抜ける
    break

if excel_result:
    QMessageBox.information(iface.mainWindow(), "情報", f"{output_path}を出力しました。")