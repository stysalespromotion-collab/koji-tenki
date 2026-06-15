"""
転記ロジックモジュール
見積決定.xlsx → 出来高.xlsx / 工事完了.xlsx
"""

import shutil
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def _get_val(ws, coord):
    """セルの値を取得（結合セルも考慮）"""
    return ws[coord].value


def _set_val(ws, coord, value):
    """セルに値をセット（結合セルの先頭に書き込む）"""
    if value is None:
        return
    ws[coord] = value


def _copy_rows(ws_src, ws_dst, src_start, src_end, dst_start, dst_end, col_map, skip_keywords=None):
    """
    行範囲をコピー。空行・合計行はスキップ。
    col_map: {src_col_num: dst_col_num}
    skip_keywords: この文字列を含むセルがある行はスキップ
    """
    skip_keywords = skip_keywords or ['合計', '合　計']
    dst_row = dst_start
    for src_row in range(src_start, src_end + 1):
        if dst_row > dst_end:
            break
        # 合計行スキップ
        is_skip = any(
            skip_kw in str(ws_src.cell(src_row, c).value or '')
            for c in col_map
            for skip_kw in skip_keywords
        )
        if is_skip:
            continue
        # データがあるか確認
        has_data = any(
            ws_src.cell(src_row, c).value not in (None, 0, '')
            for c in col_map
        )
        if not has_data:
            continue
        for src_col, dst_col in col_map.items():
            val = ws_src.cell(src_row, src_col).value
            if val is not None and val != '':
                ws_dst.cell(dst_row, dst_col).value = val
        dst_row += 1


def transfer_dekidaka(mitsumori_bytes, kaime):
    """
    出来高書式への転記
    mitsumori_bytes: 見積決定.xlsxのバイト列
    kaime: 何回目か（int）
    戻り値: 出来高.xlsxのバイト列
    """
    wb_src = load_workbook(io.BytesIO(mitsumori_bytes), data_only=True)
    ws_src = wb_src['見積決定']

    with open('templates/出来高.xlsx', 'rb') as f:
        tmpl_bytes = f.read()
    wb_dst = load_workbook(io.BytesIO(tmpl_bytes))
    ws_dst = wb_dst['出来高']

    # ヘッダー不要（出来高はヘッダー転記しない）

    # 回目を記入
    ws_dst['J10'] = f'{kaime}'
    ws_dst['J11'] = f'{kaime}'

    # 売上・請求金額
    _set_val(ws_dst, 'F9', _get_val(ws_src, 'H16'))   # 契約金額（売上合計）

    # 下請支払（見積決定 行20-38 → 出来高 行17-20 / 4行分）
    # 見積決定の下請工事店行（合計行除く）
    _copy_rows(
        ws_src, ws_dst,
        src_start=20, src_end=38,
        dst_start=17, dst_end=20,
        col_map={
            2: 2,   # B: 店名
            8: 8,   # H: 支払合計金額
        }
    )

    # 材料明細（見積決定 行42-56 → 出来高 行23-37）
    _copy_rows(
        ws_src, ws_dst,
        src_start=42, src_end=56,
        dst_start=23, dst_end=37,
        col_map={
            2: 2,   # B: 材料名
            4: 4,   # D: 数量
            5: 5,   # E: 単位
            6: 6,   # F: 仕入単価
            7: 7,   # G: 販売単価
            8: 8,   # H: 仕入金額
            9: 9,   # I: 販売金額
            11: 11, # K: 販売先
        }
    )

    out = io.BytesIO()
    wb_dst.save(out)
    out.seek(0)
    return out.read()


def transfer_kanryo(mitsumori_bytes, dekidaka_bytes_list, kaime):
    """
    工事完了書式への転記（最終明細）
    mitsumori_bytes: 見積決定.xlsxのバイト列
    dekidaka_bytes_list: 過去の出来高.xlsxのバイト列リスト（下請累計用）
    kaime: 今回が何回目か（int）
    戻り値: 工事完了.xlsxのバイト列
    """
    wb_src = load_workbook(io.BytesIO(mitsumori_bytes), data_only=True)
    ws_src = wb_src['見積決定']

    with open('templates/工事完了.xlsx', 'rb') as f:
        tmpl_bytes = f.read()
    wb_dst = load_workbook(io.BytesIO(tmpl_bytes))
    ws_dst = wb_dst['１回完了']

    # ヘッダー情報
    _set_val(ws_dst, 'A2',  _get_val(ws_src, 'B2'))   # No.
    _set_val(ws_dst, 'C3',  _get_val(ws_src, 'C3'))   # 工事種類
    _set_val(ws_dst, 'C5',  _get_val(ws_src, 'C5'))   # 工事名称
    _set_val(ws_dst, 'C6',  _get_val(ws_src, 'C6'))   # 工事場所
    _set_val(ws_dst, 'C7',  _get_val(ws_src, 'C7'))   # 売上先
    _set_val(ws_dst, 'I5',  _get_val(ws_src, 'I5'))   # 元請
    _set_val(ws_dst, 'I6',  _get_val(ws_src, 'I6'))   # 契約No
    _set_val(ws_dst, 'I7',  _get_val(ws_src, 'I7'))   # 施工日
    _set_val(ws_dst, 'K1',  _get_val(ws_src, 'K1'))   # 所属
    _set_val(ws_dst, 'K2',  _get_val(ws_src, 'K2'))   # 見積担当者
    _set_val(ws_dst, 'K3',  _get_val(ws_src, 'K3'))   # 施工担当者

    # 売上・請求金額
    _set_val(ws_dst, 'F8', _get_val(ws_src, 'H16'))   # 契約金額

    # 下請支払 累計（全出来高ファイルの合計）
    # 店名ごとに支払合計を集計
    shita_totals = {}  # {店名: 累計金額}
    shita_order = []

    for db in dekidaka_bytes_list:
        wb_d = load_workbook(io.BytesIO(db), data_only=True)
        ws_d = wb_d['出来高']
        for r in range(17, 21):  # 出来高の下請支払は行17-20
            name = ws_d.cell(r, 2).value
            amount = ws_d.cell(r, 8).value
            if name and str(name).strip():
                name = str(name).strip()
                if name not in shita_totals:
                    shita_totals[name] = 0
                    shita_order.append(name)
                shita_totals[name] += (amount or 0)

    # 工事完了の下請支払欄（行17-22）に累計を書き込む
    dst_row = 17
    for name in shita_order:
        if dst_row > 22:
            break
        ws_dst.cell(dst_row, 2).value = name
        ws_dst.cell(dst_row, 8).value = shita_totals[name]
        dst_row += 1

    # 材料明細（見積決定から全部）
    _copy_rows(
        ws_src, ws_dst,
        src_start=42, src_end=56,
        dst_start=25, dst_end=39,
        col_map={
            2: 2,   # B: 材料名
            4: 4,   # D: 数量
            5: 5,   # E: 単位
            6: 6,   # F: 仕入単価
            7: 7,   # G: 販売単価
            8: 8,   # H: 仕入金額
            9: 9,   # I: 販売金額
            11: 11, # K: 販売先
        }
    )

    # 塗厚管理（見積決定からそのまま）
    for src_r, dst_r in [(64, 54), (65, 55), (66, 56)]:
        for col in [10, 11]:  # J, K
            val = ws_src.cell(src_r, col).value
            if val is not None:
                ws_dst.cell(dst_r, col).value = val

    out = io.BytesIO()
    wb_dst.save(out)
    out.seek(0)
    return out.read()
