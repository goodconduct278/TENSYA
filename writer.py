"""
転記先（見積書）Excelシートへの書き込み処理。

1明細 = 3行ブロック：
  行N   : 仕様1 / 数量 / 単位 / 単価   （名称は空）
  行N+1 : 名称1 / 仕様2
  行N+2 : 名称2 / 仕様3

数量・単位・単価はブロック1行目に書く（金額式が1行目を参照するため）。
金額列（IF式）には書き込まない。
"""

from openpyxl.utils import column_index_from_string, get_column_letter
from logger import Logger


def _col(letter: str) -> int:
    return column_index_from_string(letter.upper())


def _safe_set(ws, row: int, col_letter: str, value) -> None:
    """結合セル対応のセル書き込み。結合範囲なら左上セルに書く。"""
    col = _col(col_letter)
    for merge in ws.merged_cells.ranges:
        if (merge.min_row <= row <= merge.max_row
                and merge.min_col <= col <= merge.max_col):
            ws.cell(merge.min_row, merge.min_col).value = value
            return
    ws.cell(row, col).value = value


def find_total_row(ws, settings: dict) -> int:
    """合計行キーワードをA列で検索して行番号を返す。見つからなければ0。"""
    keyword = settings['sum_keyword']
    for (cell,) in ws.iter_rows(min_col=1, max_col=1):
        if cell.value is not None and keyword in str(cell.value):
            return cell.row
    return 0


def clear_destination(ws, settings: dict) -> None:
    """転記先の名称・仕様・数量・単位・単価を値のみクリア（書式・数式は保持）"""
    total_row = find_total_row(ws, settings)
    start_row = settings['dst_start_row']

    if total_row > 0:
        end_row = total_row - 1
    else:
        end_row = 0
        for row in ws.iter_rows(min_row=start_row):
            for cell in row:
                if cell.value is not None:
                    end_row = cell.row
        if end_row == 0:
            return

    for r in range(start_row, end_row + 1):
        _safe_set(ws, r, settings['dst_col_name'],  None)
        _safe_set(ws, r, settings['dst_col_spec'],  None)
        _safe_set(ws, r, settings['dst_col_qty'],   None)
        _safe_set(ws, r, settings['dst_col_unit'],  None)
        _safe_set(ws, r, settings['dst_col_price'], None)


def ensure_rows(ws, settings: dict, needed_blocks: int, logger: Logger) -> int:
    """
    必要なブロック数に対して行数が不足している場合、3行単位で行を挿入する。
    戻り値：追加した行数（合計行が無い場合は0）
    """
    total_row = find_total_row(ws, settings)
    if total_row == 0:
        logger.log("合計行なし：行追加処理をスキップします")
        return 0

    start_row = settings['dst_start_row']
    current_blocks = (total_row - start_row) // 3
    short_blocks = needed_blocks - current_blocks

    if short_blocks <= 0:
        return 0

    add_rows = short_blocks * 3

    for _ in range(short_blocks):
        total_row = find_total_row(ws, settings)
        ws.insert_rows(total_row, amount=3)
        # insert_rows は total_row の直前に挿入するため、
        # 合計行は total_row + 3 に移動する（find_total_row で再取得）

    # 挿入された行の値をクリア（insert_rows は空行を挿入するが念のため）
    new_total_row = find_total_row(ws, settings)
    for r in range(new_total_row - add_rows, new_total_row):
        _safe_set(ws, r, settings['dst_col_name'],  None)
        _safe_set(ws, r, settings['dst_col_spec'],  None)
        _safe_set(ws, r, settings['dst_col_qty'],   None)
        _safe_set(ws, r, settings['dst_col_unit'],  None)
        _safe_set(ws, r, settings['dst_col_price'], None)

    return add_rows


def write_to_destination(ws, settings: dict, blocks: list[dict]) -> None:
    """明細ブロックを3行単位で転記先に書き込む"""
    write_row = settings['dst_start_row']

    for block in blocks:
        # 行1：仕様1 / 数量 / 単位 / 単価（名称は空のまま）
        _safe_set(ws, write_row,     settings['dst_col_spec'],  block['spec1'])
        _safe_set(ws, write_row,     settings['dst_col_qty'],   block['qty'])
        _safe_set(ws, write_row,     settings['dst_col_unit'],  block['unit'])
        _safe_set(ws, write_row,     settings['dst_col_price'], block['price'])
        # 行2：名称1 / 仕様2
        _safe_set(ws, write_row + 1, settings['dst_col_name'],  block['name1'])
        _safe_set(ws, write_row + 1, settings['dst_col_spec'],  block['spec2'])
        # 行3：名称2 / 仕様3
        _safe_set(ws, write_row + 2, settings['dst_col_name'],  block['name2'])
        _safe_set(ws, write_row + 2, settings['dst_col_spec'],  block['spec3'])

        write_row += 3


def rebuild_sum_formula(ws, settings: dict, logger: Logger) -> str:
    """
    金額列のSUM式を完全再生成して合計行に書き込む。
    開始行から3行おきにブロック1行目のセルを収集して =SUM(...) を組み立てる。
    """
    total_row = find_total_row(ws, settings)
    if total_row == 0:
        logger.log("合計行なし：SUM式更新をスキップします")
        return "(SUM式更新スキップ)"

    start_row  = settings['dst_start_row']
    col_letter = settings['dst_col_amount'].upper()
    col_idx    = _col(settings['dst_col_amount'])

    cell_refs = [
        f"{col_letter}{r}"
        for r in range(start_row, total_row, 3)
    ]
    formula = f"=SUM({','.join(cell_refs)})"

    # 合計行の金額列セルに書き込む（結合セル対応）
    for merge in ws.merged_cells.ranges:
        if (merge.min_row <= total_row <= merge.max_row
                and merge.min_col <= col_idx <= merge.max_col):
            ws.cell(merge.min_row, merge.min_col).value = formula
            return formula

    ws.cell(total_row, col_idx).value = formula
    return formula
