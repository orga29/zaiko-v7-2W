import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Border, Side
import datetime
# import tkinter as tk # --- 削除 ---
# from tkinter import filedialog, messagebox, ttk # --- 削除 ---
import os
from copy import copy
# from tkcalendar import DateEntry # --- 削除 ---

import streamlit as st  # --- 追加 ---
import io               # --- 追加 ---


def find_sheet(workbook, target_name):
    """シート名をスペース無視で照合して取得"""
    for name in workbook.sheetnames:
        if name.strip() == target_name.strip():
            return workbook[name]
    return None


def create_categorized_inventory_excel(
    input_file_buffer,  # 変更: ファイルパス -> ファイルバッファ
    target_date_str: str
): # 変更: 戻り値の型が変わる (エラー時はstr, 成功時はTuple)
    """
    在庫集計ファイルから「箱もの」「こもの」を分類し、
    既存シート「在庫表（箱）」「在庫表（こもの）」に転記。
    書式は保持しつつ、空白行や残データを確実に除外。
    """

    input_sheet_name = "在庫集計表"  # 固定シート名

    try:
        # --- 日付解析 ---
        try:
            target_date = datetime.datetime.strptime(target_date_str, '%Y-%m-%d').date()
        except ValueError:
            return "エラー: 日付形式が無効です。YYYY-MM-DD形式で入力してください。"

        # --- 曜日→列マッピング（日曜=M） ---
        ordered_days_jp = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日']
        ordered_columns_letters = ['M', 'R', 'W', 'AB', 'AG', 'AL']
        python_weekday_to_jp_day = {
            0: '月曜日', 1: '火曜日', 2: '水曜日',
            3: '木曜日', 4: '金曜日', 5: '土曜日', 6: '日曜日'
        }

        base_day_jp = python_weekday_to_jp_day[target_date.weekday()]
        if base_day_jp == "土曜日":
            return "エラー: 土曜日は集計対象外です。"

        col_letter = ordered_columns_letters[ordered_days_jp.index(base_day_jp)]

        # --- 入力ファイル読込 (データ取得用) ---
        # 変更: ファイルパスからバッファを読み込む
        wb_input = openpyxl.load_workbook(input_file_buffer, data_only=True)
        if input_sheet_name not in wb_input.sheetnames:
            return f"エラー: シート『{input_sheet_name}』が見つかりません。"
        ws_input = wb_input[input_sheet_name]

        header_row = 7
        exclusion_keywords = ["配達料", "運賃", "カステラ", "十勝の息吹", "有機納豆", "ひきわり", "豆腐", "丸大豆"]
        exclusion_toichi = "東一"

        boxed, smalls = [], []

        # --- データ抽出 ---
        for r in range(header_row + 1, ws_input.max_row + 1):
            code = ws_input.cell(row=r, column=column_index_