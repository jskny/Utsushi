"""エッジケース検証用の帳票テンプレート(.xlsx)を生成する。

`tools/Utsushi.SampleGenerator`(C#)が作る `samples/reports/` の合成サンプルは、Utsushi の対応機能を
狙って配置した「お手本どおりの」テンプレートである。こちらは、帳票の担当者が普段の Excel 操作で作りがちな
テンプレート(既定の游ゴシック・テーマの色・列全体の書式・数式・条件付き書式・入力規則・コメント・
グラフ・非表示の行列・「すべての列を1ページに印刷」など)を再現し、差し込み値の極端な値と組み合わせて
ライブラリの挙動を確かめるためのものである。

各帳票で何を確かめるかは `samples/edge-cases/README.md` を参照。

実行方法(リポジトリのルートで):

    pip install -r tools/edge-case-samples/requirements.txt
    python3 tools/edge-case-samples/generate.py samples/edge-cases

openpyxl は Excel と違い数式の計算結果(キャッシュ値)を保存しないため、保存後に
シートの XML へ計算結果を書き込み、Excel で保存したブックと同じ状態にしている(`_patch_cached_values`)。
"""

import datetime
import io
import os
import re
import sys
import zipfile

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Color, Font, GradientFill, PatternFill, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break

# 日本語版 Excel 2016 以降の新規ブックの標準フォント
YU_GOTHIC = "游ゴシック"
THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FIXED_ZIP_TIME = (2026, 4, 1, 0, 0, 0)


def new_workbook():
    """日本語版 Excel の新規ブックと同じく、標準フォントが游ゴシック 11pt のブックを作る。"""
    wb = Workbook()
    default = Font(name=YU_GOTHIC, size=11, family=3, charset=128, scheme="minor", color=Color(theme=1))
    wb._fonts[0] = default  # noqa: SLF001 - openpyxl に標準フォントを変える公開 API が無い
    wb._named_styles["Normal"].font = default  # noqa: SLF001
    # 游ゴシック 11pt のブックを Excel で保存すると、標準の行の高さは 18.75pt になる
    wb.active.sheet_format.defaultRowHeight = 18.75
    return wb


def new_sheet(wb, title):
    ws = wb.create_sheet(title)
    ws.sheet_format.defaultRowHeight = 18.75
    return ws


def font(size=11, bold=False, name=YU_GOTHIC, color=None, **kwargs):
    return Font(name=name, size=size, bold=bold, family=3, charset=128, color=color, **kwargs)


def put(ws, ref, value=None, *, f=None, align=None, border=None, fill=None, fmt=None):
    cell = ws[ref]
    if value is not None:
        cell.value = value
    cell.font = f or font()
    if align is not None:
        cell.alignment = align
    if border is not None:
        cell.border = border
    if fill is not None:
        cell.fill = fill
    if fmt is not None:
        cell.number_format = fmt
    return cell


def box_range(ws, rng, f=None, align=None, fill=None):
    """結合セルの外周に罫線を引くため、範囲内の全セルに同じ書式を設定する(Excel の操作と同じ結果になる)。"""
    for row in ws[rng]:
        for cell in row:
            cell.font = f or font()
            cell.border = BOX
            if align is not None:
                cell.alignment = align
            if fill is not None:
                cell.fill = fill


def setup_page(ws, *, area, orientation="portrait", top=0.75, bottom=0.75, left=0.7, right=0.7):
    ws.print_area = area
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = orientation
    ws.page_margins.top = top
    ws.page_margins.bottom = bottom
    ws.page_margins.left = left
    ws.page_margins.right = right
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3


def save(wb, path, cached_values):
    """ブックを保存し、数式セルへ計算結果を書き込む。"""
    buffer = io.BytesIO()
    wb.save(buffer)
    _patch_cached_values(buffer.getvalue(), path, cached_values)


def _patch_cached_values(data, path, cached_values):
    """cached_values: {シート名: {セル番地: 計算結果(数値・文字列・None)}}。None は計算結果を書かない(未計算のまま)。"""
    src = zipfile.ZipFile(io.BytesIO(data))
    sheet_paths = _sheet_paths(src)
    entries = []
    for item in src.infolist():
        content = src.read(item.filename)
        for sheet_name, values in cached_values.items():
            if sheet_paths.get(sheet_name) != item.filename:
                continue
            xml = content.decode("utf-8")
            for ref, value in values.items():
                xml = _patch_cell(xml, ref, value)
            content = xml.encode("utf-8")
        entries.append((item.filename, content))
    _write_entries(path, entries)


def _write_entries(path, entries):
    """再生成しても同じバイト列になるよう、ZIP の日時と文書プロパティの作成・更新日時を固定して書き出す。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, content in entries:
            if name == "docProps/core.xml":
                content = re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*", rb"\g<1>2026-04-01T00:00:00Z", content)
            info = zipfile.ZipInfo(name, date_time=FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            dst.writestr(info, content)


def _sheet_paths(src):
    workbook = src.read("xl/workbook.xml").decode("utf-8")
    rels = src.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    targets.update({k: v for v, k in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    result = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', workbook):
        target = targets[rid]
        result[name] = target.lstrip("/") if target.startswith("/xl/") else "xl/" + target
    return result


def _patch_cell(xml, ref, value):
    pattern = re.compile(r'<c r="%s"([^>]*)><f>(.*?)</f><v ?/>(?:</v>)?</c>' % re.escape(ref))
    match = pattern.search(xml)
    if match is None:
        raise ValueError("数式セルが見つからない: " + ref)
    attrs, formula = match.group(1), match.group(2)
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    if value is None:
        replacement = '<c r="%s"%s><f>%s</f></c>' % (ref, attrs, formula)
    elif isinstance(value, str):
        escaped = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        replacement = '<c r="%s"%s t="str"><f>%s</f><v>%s</v></c>' % (ref, attrs, formula, escaped)
    else:
        number = str(value) if isinstance(value, int) else repr(float(value))
        replacement = '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (ref, attrs, formula, number)
    return xml[: match.start()] + replacement + xml[match.end():]


def png_logo(width, height, rgb):
    from PIL import Image as PilImage, ImageDraw

    image = PilImage.new("RGB", (width, height), rgb)
    draw = ImageDraw.Draw(image)
    draw.rectangle([4, 4, width - 5, height - 5], outline=(255, 255, 255), width=3)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def jpeg_photo(width, height):
    from PIL import Image as PilImage

    image = PilImage.new("RGB", (width, height))
    pixels = image.load()
    for y in range(height):
        for x in range(width):
            pixels[x, y] = (x * 255 // width, y * 255 // height, 160)
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# 1. 送付状: 差し込み値の極端な値(長い宛名・複数行の住所・外字・空の値・改行やタブを含む値)
# ---------------------------------------------------------------------------
def generate_cover_letter(root):
    wb = new_workbook()
    ws = wb.active
    ws.title = "送付状"

    widths = {"A": 6, "B": 13, "C": 13, "D": 13, "E": 5, "F": 4, "G": 10, "H": 15}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 30
    # 表題の下の帯: 行番号をクリックして行全体を塗りつぶした行(セルは作らない。行の書式として保存される)
    ws.row_dimensions[2].height = 6
    ws.row_dimensions[2].fill = PatternFill("solid", fgColor=Color(theme=4, tint=0.3999755851924192))
    ws.merge_cells("A1:H1")
    put(ws, "A1", "書類送付のご案内", f=font(18, True, name="ＭＳ 明朝"),
        align=Alignment(horizontal="center", vertical="center"))

    # 宛先ブロック。郵便番号(はみ出し)・住所(3行ぶんの高さの折り返し。縦位置は Excel 既定の「下」)・
    # 会社名(縮小して全体を表示)・部署(セル幅で切る)・担当者名(はみ出し。右隣の「様」と重なりうる)
    put(ws, "A3", "〒000-0000")
    ws.merge_cells("A4:D6")
    put(ws, "A4", "(住所)", align=Alignment(wrap_text=True))
    ws.merge_cells("A7:D7")
    put(ws, "A7", "(会社名)", f=font(12, True), align=Alignment(shrink_to_fit=True, vertical="center"))
    ws.row_dimensions[7].height = 20
    put(ws, "A8", "(部署)", align=Alignment(horizontal="left"))
    ws.merge_cells("A8:B8")
    put(ws, "C8", "(担当者名)", f=font(12))
    put(ws, "D8", "様", f=font(12), align=Alignment(horizontal="right"))

    # 発行情報(右上)
    put(ws, "G3", "発送日", align=Alignment(horizontal="right"))
    put(ws, "H3", datetime.date(2026, 4, 1), fmt='yyyy"年"m"月"d"日"', align=Alignment(horizontal="right"))
    put(ws, "G4", "文書番号", align=Alignment(horizontal="right"))
    put(ws, "H4", 1, fmt="0000", align=Alignment(horizontal="right"))

    # 差出人(右側、4行の折り返し。縦位置は上)
    ws.merge_cells("F5:H8")
    put(ws, "F5", "(差出人)", f=font(10), align=Alignment(wrap_text=True, vertical="top"))

    # 本文(固定文言)
    ws.merge_cells("A10:H12")
    put(ws, "A10",
        "拝啓 時下ますますご清栄のこととお慶び申し上げます。平素は格別のお引き立てを賜り、厚く御礼申し上げます。"
        "下記の書類をお送りいたしますので、ご査収くださいますようお願い申し上げます。",
        align=Alignment(wrap_text=True, vertical="top"))

    # 明細表(品名は折り返しなしの通常セル)
    ws.row_dimensions[14].height = 20
    headers = [("A14", "No."), ("B14", "書類名"), ("G14", "部数"), ("H14", "備考")]
    ws.merge_cells("B14:F14")
    box_range(ws, "A14:H14", f=font(11, True), align=Alignment(horizontal="center", vertical="center"),
              fill=PatternFill("solid", fgColor=Color(theme=4, tint=0.7999816888943144)))
    for ref, text in headers:
        ws[ref].value = text
    for i in range(5):
        row = 15 + i
        ws.merge_cells(f"B{row}:F{row}")
        box_range(ws, f"A{row}:H{row}", align=Alignment(vertical="center"))
        ws[f"A{row}"].value = i + 1
        ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"G{row}"].number_format = '#,##0"部"'
        ws[f"G{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"H{row}"].alignment = Alignment(vertical="center", shrink_to_fit=True)

    # 金額: 横位置を「標準」のままにしたセル(差し込んだ文字列は左に寄る)と、右揃えにしたセル
    put(ws, "F21", "金額(標準)", f=font(10), align=Alignment(horizontal="right"))
    ws.merge_cells("G21:H21")
    box_range(ws, "G21:H21")
    put(ws, "F22", "金額(右揃え)", f=font(10), align=Alignment(horizontal="right"))
    ws.merge_cells("G22:H22")
    box_range(ws, "G22:H22", align=Alignment(horizontal="right"))

    # 備考(5行ぶんの折り返し、上揃え)
    put(ws, "A24", "備考", f=font(11, True))
    ws.merge_cells("A25:H29")
    box_range(ws, "A25:H29", align=Alignment(wrap_text=True, vertical="top"))

    # 書式の作り方による違いを確かめるセル
    #  B31: 英字フォント(Calibri)のセル。日本語の値を差し込むと日本語の部分は同梱フォントで描かれる
    #  B32: テンプレートに存在しないセル(一度も入力・書式設定をしていない)
    #  C33: 列全体に書式(太字・赤)を設定した列のセル(セル自体は未設定)
    #  (A列はラベル)
    put(ws, "A30", "以下は書式の作り方を変えた差し込みセル", f=font(9, color="808080"))
    put(ws, "B31", "", f=Font(name="Calibri", size=11))
    ws["A31"].value = "英字"
    ws["A31"].font = font(9)
    ws["A32"].value = "未設定"
    ws["A32"].font = font(9)
    ws["A33"].value = "列書式"
    ws["A33"].font = font(9)
    ws.column_dimensions["C"].font = font(11, True, color="C00000")

    ws.oddFooter.center.text = "&P / &N"
    setup_page(ws, area="A1:H34")
    save(wb, os.path.join(root, "edge-cover-letter", "template.xlsx"), {})


# ---------------------------------------------------------------------------
# 2. 売上一覧: 「すべての列を1ページに印刷」・横向き・非表示の行と列・高さ0の行・タイトル行・多ページ
# ---------------------------------------------------------------------------
def generate_sales_list(root):
    wb = new_workbook()
    ws = wb.active
    ws.title = "売上一覧"

    columns = [("A", 5, "No."), ("B", 13, "売上日"), ("C", 30, "得意先"), ("D", 24, "品目"), ("E", 8, "数量"),
               ("F", 11, "単価"), ("G", 13, "金額"), ("H", 11, "原価(社外秘)"), ("I", 9, "粗利率"),
               ("J", 22, "担当者メモ"), ("K", 12, "入金予定日")]
    for col, width, _ in columns:
        ws.column_dimensions[col].width = width
    # 社外秘の原価列を非表示にする(印刷されてはいけない)
    ws.column_dimensions["H"].hidden = True

    ws.row_dimensions[1].height = 26
    ws.merge_cells("A1:K1")
    put(ws, "A1", "売上一覧表", f=font(16, True), align=Alignment(horizontal="center", vertical="center"))
    put(ws, "A2", "(対象期間)", f=font(11, True))
    put(ws, "J2", "(部署名)", align=Alignment(horizontal="right"))
    ws.merge_cells("J2:K2")

    header_fill = PatternFill("solid", fgColor=Color(theme=8, tint=0.5999938962981048))
    for col, _, text in columns:
        put(ws, f"{col}4", text, f=font(11, True), border=BOX, fill=header_fill,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[4].height = 30

    customers = ["株式会社サンプル商事", "有限会社テスト工業", "合同会社見本ソリューションズ東日本支社",
                 "サンプル物産株式会社", "株式会社ﾃｽﾄｶﾅ"]
    items = ["コピー用紙A4(500枚×5冊)", "トナーカートリッジ", "保守サービス(年間)", "ノートPC", "USBメモリ 64GB"]
    cached = {}
    first = 5
    count = 150
    total_amount = 0
    total_cost = 0
    for i in range(count):
        row = first + i
        quantity = (i % 9) + 1
        price = [480, 12800, 98000, 158000, 1980][i % 5]
        amount = quantity * price
        cost = round(amount * 0.72)
        total_amount += amount
        total_cost += cost
        date = datetime.date(2026, 4, 1) + datetime.timedelta(days=i // 5)
        put(ws, f"A{row}", i + 1, border=BOX, align=Alignment(horizontal="center"))
        put(ws, f"B{row}", date, border=BOX, fmt="yyyy/m/d(aaa)")
        put(ws, f"C{row}", customers[i % 5], border=BOX)
        put(ws, f"D{row}", items[i % 5], border=BOX)
        put(ws, f"E{row}", quantity, border=BOX, fmt="#,##0")
        put(ws, f"F{row}", price, border=BOX, fmt='"¥"#,##0')
        put(ws, f"G{row}", f"=E{row}*F{row}", border=BOX, fmt='"¥"#,##0;[Red]"¥"-#,##0')
        cached[f"G{row}"] = amount
        put(ws, f"H{row}", cost, border=BOX, fmt="#,##0")
        put(ws, f"I{row}", f"=IF(G{row}=0,\"\",(G{row}-H{row})/G{row})", border=BOX, fmt="0.0%")
        cached[f"I{row}"] = (amount - cost) / amount
        put(ws, f"J{row}", "要確認" if i % 17 == 0 else None, border=BOX)
        put(ws, f"K{row}", date + datetime.timedelta(days=30), border=BOX, fmt="[$-ja-JP]ggge\"年\"m\"月\"d\"日\"")
        if i % 25 == 24:
            # 作業用に一時的に隠した行(印刷されてはいけない)
            ws.row_dimensions[row].hidden = True

    # 高さ0の行(非表示と同じく印刷しない)。openpyxl は高さ0を書き出さないため、保存後に書き込む
    zero_height_row = first + 10
    ws.row_dimensions[zero_height_row].height = 0

    total_row = first + count
    ws.merge_cells(f"A{total_row}:F{total_row}")
    box_range(ws, f"A{total_row}:K{total_row}", f=font(11, True),
              fill=PatternFill("solid", fgColor=Color(theme=0, tint=-0.1499984740745262)))
    ws[f"A{total_row}"].value = "合計"
    ws[f"A{total_row}"].alignment = Alignment(horizontal="center")
    ws[f"G{total_row}"].value = f"=SUBTOTAL(9,G{first}:G{total_row - 1})"
    ws[f"G{total_row}"].number_format = '"¥"#,##0'
    cached[f"G{total_row}"] = total_amount
    ws[f"H{total_row}"].value = f"=SUM(H{first}:H{total_row - 1})"
    ws[f"H{total_row}"].number_format = "#,##0"
    cached[f"H{total_row}"] = total_cost

    # 手動の改ページ(「すべての列を1ページに印刷」と併用すると Excel と同じく無視される)
    ws.row_breaks.append(Break(id=first + 49))

    setup_page(ws, area=f"A1:K{total_row}", orientation="landscape", top=0.6, bottom=0.6, left=0.4, right=0.4)
    ws.print_title_rows = "4:4"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.oddHeader.right.text = "&A"
    ws.oddFooter.center.text = "&P / &N ページ"
    ws.oddFooter.right.text = "印刷日 &D"
    ws.freeze_panes = "A5"
    path = os.path.join(root, "edge-sales-list", "template.xlsx")
    save(wb, path, {"売上一覧": cached})
    _rewrite_sheet(path, "売上一覧", lambda xml: _replace_once(
        xml, '<row r="%d" customHeight="1">' % zero_height_row, '<row r="%d" ht="0" customHeight="1">' % zero_height_row))


# ---------------------------------------------------------------------------
# 3. 工程表: 印刷タイトル行・列、行と列の両方向の改ページ、手動改ページ、ページの方向、ページ中央、
#    改ページをまたぐ結合セル、先頭ページ番号、先頭ページのみ別指定のヘッダー/フッター
# ---------------------------------------------------------------------------
def generate_schedule(root):
    wb = new_workbook()
    ws = wb.active
    ws.title = "工程表"

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 10
    day_columns = 40
    for i in range(day_columns):
        ws.column_dimensions[_col(4 + i)].width = 4.5

    ws.merge_cells("A1:C1")
    put(ws, "A1", "工程表", f=font(16, True), align=Alignment(vertical="center"))
    ws.row_dimensions[1].height = 26
    put(ws, "A2", "(案件名)", f=font(11, True))
    put(ws, "A3", "(作成者)", f=font(10))

    for ref, text in (("A4", "No."), ("B4", "作業項目"), ("C4", "担当")):
        put(ws, ref, text, f=font(10, True), border=BOX, align=Alignment(horizontal="center", vertical="center"))
    start = datetime.date(2026, 4, 1)
    for i in range(day_columns):
        put(ws, f"{_col(4 + i)}4", start + datetime.timedelta(days=i), f=font(8), border=BOX, fmt="m/d",
            align=Alignment(horizontal="center", vertical="center", text_rotation=90))
    ws.row_dimensions[4].height = 36

    bar = PatternFill("solid", fgColor="5B9BD5")
    rows = 70
    for i in range(rows):
        row = 5 + i
        put(ws, f"A{row}", i + 1, f=font(10), border=BOX, align=Alignment(horizontal="center"))
        put(ws, f"B{row}", f"作業項目 {i + 1:03d}", f=font(10), border=BOX)
        put(ws, f"C{row}", ["佐藤", "鈴木", "髙橋", "渡邊"][i % 4], f=font(10), border=BOX,
            align=Alignment(horizontal="center"))
        begin = (i * 3) % (day_columns - 6)
        for j in range(day_columns):
            put(ws, f"{_col(4 + j)}{row}", None, f=font(8), border=BOX,
                fill=bar if begin <= j < begin + 6 else None)

    # 改ページをまたぐ位置に、縦に結合したセル(工程の区切り)を置く
    ws.merge_cells("B26:B30")
    ws["B26"].value = "改ページをまたぐ結合セル(5行)"
    ws["B26"].alignment = Alignment(vertical="center", wrap_text=True)

    # 手動改ページ: 行は30行目の手前、列は25列目(Y)の手前
    ws.row_breaks.append(Break(id=29))
    ws.col_breaks.append(Break(id=24))

    last_row = 5 + rows - 1
    setup_page(ws, area=f"A1:{_col(3 + day_columns)}{last_row}", orientation="landscape")
    ws.print_title_rows = "4:4"
    ws.print_title_cols = "A:C"
    ws.page_setup.scale = 90
    ws.page_setup.pageOrder = "overThenDown"
    ws.page_setup.firstPageNumber = 5
    ws.page_setup.useFirstPageNumber = True
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.HeaderFooter.differentFirst = True
    ws.firstHeader.center.text = "&\"游ゴシック,太字\"&14工程表(表紙ページ)"
    ws.firstFooter.center.text = "- &P -"
    ws.oddHeader.left.text = "&A"
    ws.oddFooter.center.text = "&P / &N"
    save(wb, os.path.join(root, "edge-schedule", "template.xlsx"), {})


def _col(index):
    name = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


# ---------------------------------------------------------------------------
# 4. 注文書: よく使われる Excel の機能(テーマの色、数式、条件付き書式、入力規則、コメント、リッチテキスト、
#    縦書き、均等割り付け、網かけ・グラデーション、各種罫線、各種表示形式、画像、グラフ、複数シート)
# ---------------------------------------------------------------------------
def generate_order_form(root):
    wb = new_workbook()
    cover = wb.active
    cover.title = "表紙"
    put(cover, "A1", "この表紙シートは印刷対象ではない")

    ws = new_sheet(wb, "注文書")
    master = new_sheet(wb, "マスタ")
    master.sheet_state = "hidden"
    for i, (code, name, price) in enumerate([("P-001", "コピー用紙", 480), ("P-002", "トナー", 12800)], start=1):
        master[f"A{i}"], master[f"B{i}"], master[f"C{i}"] = code, name, price

    widths = {"A": 4, "B": 12, "C": 22, "D": 8, "E": 6, "F": 12, "G": 14, "H": 3}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # 表題: テーマの色(濃い青)の塗りつぶしに白文字
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A1:G1")
    put(ws, "A1", "注 文 書", f=Font(name=YU_GOTHIC, size=20, bold=True, color=Color(theme=0), family=3),
        fill=PatternFill("solid", fgColor=Color(theme=4, tint=-0.249977111117893)),
        align=Alignment(horizontal="center", vertical="center"))

    # 宛先: 文字列の表示形式 @" 御中"(差し込んだ値に表示形式は効かないことの確認)
    ws.merge_cells("A3:D3")
    box_range(ws, "A3:D3", f=font(14, True))
    for c in ws["A3:D3"][0]:
        c.border = Border(bottom=Side(style="double", color="000000"))
    put(ws, "A3", "(宛先)", f=font(14, True), fmt='@" 御中"', border=Border(bottom=Side(style="double")))

    # 発注日(和暦)・注文番号
    put(ws, "F3", "発注日", align=Alignment(horizontal="distributed", indent=1))
    put(ws, "G3", datetime.date(2026, 4, 1), fmt='[$-ja-JP]ggge"年"m"月"d"日"', align=Alignment(horizontal="right"))
    put(ws, "F4", "注文番号", align=Alignment(horizontal="distributed", indent=1))
    put(ws, "G4", "(注文番号)", align=Alignment(horizontal="right"))

    # 会社ロゴ(PNG)と、納品場所の写真(JPEG、twoCellAnchor 相当にセルへ合わせて伸縮)
    logo = Image(png_logo(160, 48, (31, 78, 140)))
    logo.width, logo.height = 120, 36
    ws.add_image(logo, "F6")

    # 支払条件: 入力規則(プルダウン)のセル
    put(ws, "A5", "支払条件", f=font(10, True))
    ws.merge_cells("B5:C5")
    box_range(ws, "B5:C5")
    dv = DataValidation(type="list", formula1='"現金,銀行振込,手形"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("B5")
    ws["B5"].value = "銀行振込"

    # 納期: リッチテキスト(一部だけ赤字・太字)
    put(ws, "A6", "納期", f=font(10, True))
    ws.merge_cells("B6:D6")
    ws["B6"].value = CellRichText(
        "2026年4月30日 ",
        TextBlock(InlineFont(rFont=YU_GOTHIC, b=True, color="FFC00000", sz=11), "(厳守)"))
    ws["B6"].font = font()

    # 明細表
    header_fill = PatternFill("solid", fgColor=Color(theme=4, tint=0.7999816888943144))
    for col, text in zip("ABCDEFG", ["No", "品番", "品名", "数量", "単位", "単価", "金額"]):
        put(ws, f"{col}9", text, f=font(10, True), border=BOX, fill=header_fill,
            align=Alignment(horizontal="center", vertical="center"))
    items = [("P-001", "コピー用紙 A4", 10, "箱", 4800), ("P-002", "トナーカートリッジ(黒)", 2, "本", 12800),
             ("P-003", "値引き", 1, "式", -3000)]
    cached = {}
    for i in range(8):
        row = 10 + i
        for col in "ABCDEFG":
            put(ws, f"{col}{row}", None, border=Border(left=THIN, right=THIN, top=Side(style="hair"),
                                                     bottom=Side(style="hair")))
        ws[f"A{row}"].alignment = Alignment(horizontal="center")
        ws[f"D{row}"].number_format = "#,##0"
        ws[f"F{row}"].number_format = '#,##0_ ;[Red]\\-#,##0\\ '
        ws[f"G{row}"].number_format = '#,##0_ ;[Red]\\-#,##0\\ '
        ws[f"G{row}"].value = f'=IF(D{row}="","",D{row}*F{row})'
        if i < len(items):
            code, name, qty, unit, price = items[i]
            ws[f"A{row}"].value = i + 1
            ws[f"B{row}"].value = code
            ws[f"C{row}"].value = name
            ws[f"D{row}"].value = qty
            ws[f"E{row}"].value = unit
            ws[f"F{row}"].value = price
            cached[f"G{row}"] = qty * price
        else:
            cached[f"G{row}"] = ""
    # 明細の最終行の下だけ二重線
    for col in "ABCDEFG":
        ws[f"{col}17"].border = Border(left=THIN, right=THIN, top=Side(style="hair"), bottom=Side(style="double"))

    subtotal = sum(q * p for _, _, q, _, p in items)
    tax = int(subtotal * 0.1)
    for row, label, formula, value in ((18, "小計", "=SUM(G10:G17)", subtotal),
                                       (19, "消費税(10%)", "=ROUNDDOWN(G18*0.1,0)", tax),
                                       (20, "合計", "=G18+G19", subtotal + tax)):
        ws.merge_cells(f"E{row}:F{row}")
        box_range(ws, f"E{row}:G{row}", f=font(10, row == 20))
        ws[f"E{row}"].value = label
        ws[f"E{row}"].alignment = Alignment(horizontal="center")
        ws[f"G{row}"].value = formula
        ws[f"G{row}"].number_format = '"¥"#,##0;[Red]"¥"-#,##0'
        cached[f"G{row}"] = value
    # 金額がマイナスなら赤(条件付き書式。Utsushi は反映しない)
    ws.conditional_formatting.add("G10:G20", CellIsRule(operator="lessThan", formula=["0"],
                                                       fill=PatternFill("solid", bgColor="FFC7CE")))

    # 差し込んだ宛先を数式で参照するセル(差し込んでも再計算されない)
    put(ws, "A22", "宛先の確認:", f=font(9, color="808080"))
    ws.merge_cells("B22:D22")
    put(ws, "B22", '=A3&" 御中"', f=font(9, color="808080"))
    cached["B22"] = "(宛先) 御中"

    # 縦書き・均等割り付け・選択範囲内で中央・網かけ・グラデーション・斜線・各種線種
    put(ws, "A24", "社内記入欄", f=font(9), align=Alignment(text_rotation=255, vertical="center", horizontal="center"),
        border=Border(left=Side(style="medium"), right=THIN, top=Side(style="medium"), bottom=Side(style="medium")))
    ws.merge_cells("A24:A28")
    for r in range(25, 29):
        ws[f"A{r}"].border = Border(left=Side(style="medium"), bottom=Side(style="medium") if r == 28 else None)
    labels = [("B24", "承認者", Alignment(horizontal="distributed", vertical="center")),
              ("B25", "受付日", Alignment(horizontal="centerContinuous")),
              ("B26", "網かけ", None), ("B27", "グラデーション", None), ("B28", "斜線(対象外)", None)]
    for ref, text, align in labels:
        put(ws, ref, text, f=font(9), align=align, border=Border(left=THIN, right=THIN, top=Side(style="dashed"),
                                                               bottom=Side(style="dashed")))
    put(ws, "C25", None, border=Border(top=Side(style="dashed"), bottom=Side(style="dashed")))
    ws["B26"].fill = PatternFill("darkTrellis", fgColor="BFBFBF", bgColor="FFFFFF")
    ws["B27"].fill = GradientFill(stop=("FFFFFF", "9BC2E6"))
    ws["B28"].border = Border(left=THIN, right=THIN, top=Side(style="dashed"), bottom=Side(style="medium"),
                              diagonal=Side(style="thin", color="FF0000"), diagonalUp=True, diagonalDown=True)
    ws.merge_cells("C24:G24")
    put(ws, "C24", "(承認者)", align=Alignment(vertical="center"),
        border=Border(top=Side(style="medium"), right=Side(style="medium"), bottom=Side(style="dotted")))
    for col in "DEFG":
        ws[f"{col}24"].border = Border(top=Side(style="medium"), right=Side(style="medium") if col == "G" else None,
                                       bottom=Side(style="dotted"))
    put(ws, "D26", 0.25, fmt="0%")
    put(ws, "E26", 1.5, fmt="[h]:mm")
    put(ws, "F26", 0.75, fmt="h:mm AM/PM")
    put(ws, "G26", 1234567.891, fmt='#,##0,"千円"')
    put(ws, "D27", 12, fmt="0000")
    put(ws, "E27", -5, fmt="0;\"▲\"0")
    put(ws, "F27", 0.5, fmt="# ?/?")
    put(ws, "G27", 12345678, fmt="0.00E+00")
    put(ws, "D28", datetime.date(2026, 4, 1), fmt="ge.m.d")
    put(ws, "E28", datetime.date(2026, 4, 1), fmt="mmm-yy")
    put(ws, "F28", True)
    put(ws, "G28", "=1/0")
    cached["G28"] = None  # エラー値はキャッシュしない(下で #DIV/0! を書き込む)
    for r in range(25, 29):
        ws[f"G{r}"].border = Border(right=Side(style="medium"), bottom=Side(style="medium") if r == 28 else None)
    for col in "CDEF":
        ws[f"{col}28"].border = Border(bottom=Side(style="medium"))

    # コメント(メモ)
    ws["G4"].comment = Comment("注文番号は基幹システムから採番する", "担当者")

    # ハイパーリンク
    put(ws, "A30", "お問い合わせ: https://example.com/contact", f=font(9, color="0563C1", underline="single"))
    ws["A30"].hyperlink = "https://example.com/contact"

    # 簡単なグラフ(Utsushi は出力しない)
    chart = BarChart()
    chart.title = "品目別金額"
    chart.add_data(Reference(ws, min_col=7, min_row=9, max_row=12), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=3, min_row=10, max_row=12))
    chart.width, chart.height = 8, 4
    ws.add_chart(chart, "C32")

    # 写真(JPEG)
    photo = Image(jpeg_photo(120, 80))
    photo.width, photo.height = 90, 60
    ws.add_image(photo, "A32")

    setup_page(ws, area="A1:G36")
    ws.oddFooter.left.text = "&F"
    ws.oddFooter.right.text = "&P/&N"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    wb.active = 1  # 保存時に開いていたシート = 注文書

    path = os.path.join(root, "edge-order-form", "template.xlsx")
    save(wb, path, {"注文書": cached})
    _write_error_value(path, "注文書", "G28", "#DIV/0!")


def _write_error_value(path, sheet_name, ref, error):
    """Excel が保存するエラー値(t="e")を数式セルへ書き込む。"""
    def rewrite(xml):
        result, count = re.subn(
            r'<c r="%s"([^>]*)><f>(.*?)</f></c>' % re.escape(ref),
            lambda m: '<c r="%s"%s t="e"><f>%s</f><v>%s</v></c>' % (ref, m.group(1), m.group(2), error),
            xml)
        if count != 1:
            raise ValueError("エラー値を書き込む数式セルが見つからない: " + ref)
        return result

    _rewrite_sheet(path, sheet_name, rewrite)


def _replace_once(xml, old, new):
    """置換が空振りして意図しないブックができるのを防ぐため、ちょうど1か所であることを確かめて置き換える。"""
    if xml.count(old) != 1:
        raise ValueError("置換対象がちょうど1か所ではない: " + old)
    return xml.replace(old, new)


def _rewrite_sheet(path, sheet_name, rewrite):
    """保存済みのブックのシート XML を書き換える(openpyxl で表現できない状態を作るため)。"""
    with zipfile.ZipFile(path) as src:
        sheet_path = _sheet_paths(src)[sheet_name]
        entries = [(item.filename, src.read(item.filename)) for item in src.infolist()]
    _write_entries(path, [
        (name, rewrite(content.decode("utf-8")).encode("utf-8") if name == sheet_path else content)
        for name, content in entries])


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else os.path.join("samples", "edge-cases")
    generate_cover_letter(root)
    generate_sales_list(root)
    generate_schedule(root)
    generate_order_form(root)
    print("エッジケースの帳票サンプルを生成しました: " + os.path.abspath(root))


if __name__ == "__main__":
    main()
