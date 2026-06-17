"""
転記ロジックモジュール
見積決定（.xlsx / .xls）→ 出来高.xlsx / 工事完了.xlsx
"""

import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def _load_workbook_any(file_bytes, filename):
    fname = filename.lower()
    if fname.endswith('.xls') and not fname.endswith('.xlsx'):
        import xlrd
        from openpyxl import Workbook as OpenpyxlWorkbook
        xls_wb = xlrd.open_workbook(file_contents=file_bytes)
        wb = OpenpyxlWorkbook()
        wb.remove(wb.active)
        for sheet_name in xls_wb.sheet_names():
            xls_ws = xls_wb.sheet_by_name(sheet_name)
            ws = wb.create_sheet(title=sheet_name)
            for row in range(xls_ws.nrows):
                for col in range(xls_ws.ncols):
                    cell = xls_ws.cell(row, col)
                    ws.cell(row + 1, col + 1, cell.value)
        return wb
    else:
        return load_workbook(io.BytesIO(file_bytes), data_only=True)


def _get_val(ws, coord):
    return ws[coord].value


def _set_val(ws, coord, value):
    if value is None:
        return
    ws[coord] = value


def _copy_rows(ws_src, ws_dst, src_start, src_end, dst_start, dst_end, col_map, skip_keywords=None):
    skip_keywords = skip_keywords or ['合計', '合　計']
    dst_row = dst_start
    for src_row in range(src_start, src_end + 1):
        if dst_row > dst_end:
            break
        is_skip = any(
            skip_kw in str(ws_src.cell(src_row, c).value or '')
            for c in col_map
            for skip_kw in skip_keywords
        )
        if is_skip:
            continue
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


def transfer_dekidaka(mitsumori_bytes, mitsumori_name, kaime):
    wb_src = _load_workbook_any(mitsumori_bytes, mitsumori_name)
    ws_src = wb_src['見積決定']

    with open('templates/出来高.xlsx', 'rb') as f:
        tmpl_bytes = f.read()
    wb_dst = load_workbook(io.BytesIO(tmpl_bytes))
    ws_dst = wb_dst['出来高']

    ws_dst['J10'] = f'{kaime}'
    ws_dst['J11'] = f'{kaime}'

    _set_val(ws_dst, 'F9', _get_val(ws_src, 'H16'))

    _copy_rows(
        ws_src, ws_dst,
        src_start=20, src_end=38,
        dst_start=17, dst_end=20,
        col_map={2: 2, 8: 8}
    )

    _copy_rows(
        ws_src, ws_dst,
        src_start=42, src_end=56,
        dst_start=23, dst_end=37,
        col_map={2: 2, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 11: 11}
    )

    out = io.BytesIO()
    wb_dst.save(out)
    out.seek(0)
    return out.read()


def transfer_kanryo(mitsumori_bytes, mitsumori_name, dekidaka_list, kaime):
    wb_src = _load_workbook_any(mitsumori_bytes, mitsumori_name)
    ws_src = wb_src['見積決定']

    with open('templates/工事完了.xlsx', 'rb') as f:
        tmpl_bytes = f.read()
    wb_dst = load_workbook(io.BytesIO(tmpl_bytes))
    ws_dst = wb_dst['１回完了']

    _set_val(ws_dst, 'A2', _get_val(ws_src, 'B2'))
    _set_val(ws_dst, 'C3', _get_val(ws_src, 'C3'))
    _set_val(ws_dst, 'C5', _get_val(ws_src, 'C5'))
    _set_val(ws_dst, 'C6', _get_val(ws_src, 'C6'))
    _set_val(ws_dst, 'C7', _get_val(ws_src, 'C7'))
    _set_val(ws_dst, 'I5', _get_val(ws_src, 'I5'))
    _set_val(ws_dst, 'I6', _get_val(ws_src, 'I6'))
    _set_val(ws_dst, 'I7', _get_val(ws_src, 'I7'))
    _set_val(ws_dst, 'K1', _get_val(ws_src, 'K1'))
    _set_val(ws_dst, 'K2', _get_val(ws_src, 'K2'))
    _set_val(ws_dst, 'K3', _get_val(ws_src, 'K3'))

    _set_val(ws_dst, 'F8', _get_val(ws_src, 'H16'))

    shita_totals = {}
    shita_order = []
    for db, dname in dekidaka_list:
        wb_d = _load_workbook_any(db, dname)
        ws_d = wb_d['出来高']
        for r in range(17, 21):
            name = ws_d.cell(r, 2).value
            amount = ws_d.cell(r, 8).value
            if name and str(name).strip():
                name = str(name).strip()
                if name not in shita_totals:
                    shita_totals[name] = 0
                    shita_order.append(name)
                shita_totals[name] += (amount or 0)

    dst_row = 17
    for name in shita_order:
        if dst_row > 22:
            break
        ws_dst.cell(dst_row, 2).value = name
        ws_dst.cell(dst_row, 8).value = shita_totals[name]
        dst_row += 1

    _copy_rows(
        ws_src, ws_dst,
        src_start=42, src_end=56,
        dst_start=25, dst_end=39,
        col_map={2: 2, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 11: 11}
    )

    for src_r, dst_r in [(64, 54), (65, 55), (66, 56)]:
        for col in [10, 11]:
            val = ws_src.cell(src_r, col).value
            if val is not None:
                ws_dst.cell(dst_r, col).value = val

    out = io.BytesIO()
    wb_dst.save(out)
    out.seek(0)
    return out.read()
