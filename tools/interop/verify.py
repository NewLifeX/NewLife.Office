#!/usr/bin/env python3
"""跨库互操作校验：读取本库（NewLife.Office）产出的文档，并用第三方库产出文档供本库读取。

用法：
    python tools/interop/verify.py <interop_dir>

目录约定：
    <interop_dir>/ours    本库生成，交给第三方读取
    <interop_dir>/theirs  第三方生成，交给本库读取

核心格式（xlsx/docx/pptx/pdf）读取失败即返回非 0；其它格式失败会打印警告但仍返回非 0，
由 CI 侧决定是否阻塞。
"""
from __future__ import annotations

import datetime
import os
import sys
import traceback
import zipfile


def _require(name):
    try:
        return __import__(name)
    except Exception:  # pragma: no cover - 缺少依赖时直接报错退出
        raise SystemExit(f"缺少依赖 {name}，请先执行 pip install -r tools/interop/requirements.txt")


def check_xlsx(path: str) -> None:
    openpyxl = _require("openpyxl")
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0][0] == "Id", f"表头不符: {rows[0]}"
    assert rows[1][0] == 1 and rows[1][1] == "甲", f"首行数据不符: {rows[1]}"
    assert abs(float(rows[1][2]) - 1234.5) < 1e-9, f"数值不符: {rows[1][2]}"
    assert rows[1][3] is True, f"布尔值不符: {rows[1][3]}"


def check_xls(path: str) -> None:
    xlrd = _require("xlrd")
    book = xlrd.open_workbook(path)
    names = book.sheet_names()
    # 本库默认会先建一个 Sheet1，数据写在显式命名的 S1 上
    sheet = book.sheet_by_name("S1") if "S1" in names else book.sheet_by_index(len(names) - 1)
    assert sheet.nrows >= 2, f"行数不足: {sheet.nrows}, sheets={names}"
    assert sheet.cell_value(0, 0) == "Id", f"表头不符: {sheet.row_values(0)}"
    assert sheet.cell_value(1, 0) == 1, f"首行数据不符: {sheet.row_values(1)}"


def check_docx(path: str) -> None:
    docx = _require("docx")
    document = docx.Document(path)
    text = "\n".join(p.text for p in document.paragraphs)
    assert "第一段中文内容" in text, f"正文缺失: {text!r}"


def check_pptx(path: str) -> None:
    pptx = _require("pptx")
    presentation = pptx.Presentation(path)
    assert len(presentation.slides) >= 1, "没有幻灯片"
    texts = "\n".join(
        shape.text
        for slide in presentation.slides
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
    )
    assert "幻灯片标题" in texts, f"文本缺失: {texts!r}"


def check_pdf(path: str) -> None:
    pypdf = _require("pypdf")
    reader = pypdf.PdfReader(path)
    assert len(reader.pages) >= 1, "没有页面"
    text = reader.pages[0].extract_text() or ""
    assert "Hello PDF" in text, f"文本缺失: {text!r}"


def check_ods(path: str) -> None:
    from odf import opendocument, table, teletype

    doc = opendocument.load(path)
    cells = doc.getElementsByType(table.Table)
    assert cells, "没有工作表"
    texts = [teletype.extractText(c) for c in doc.getElementsByType(table.TableCell)]
    assert "甲" in texts or "Name" in texts, f"单元格内容不符: {texts[:8]}"


def check_epub(path: str) -> None:
    zf = zipfile.ZipFile(path)
    names = zf.namelist()
    assert "mimetype" in names, "缺少 mimetype"
    assert zf.read("mimetype") == b"application/epub+zip", "mimetype 内容不符"
    assert "META-INF/container.xml" in names, "缺少 container.xml"
    assert any(n.startswith("OEBPS/") and n.endswith(".xhtml") for n in names), "缺少章节 XHTML"


def check_ics(path: str) -> None:
    icalendar = _require("icalendar")
    cal = icalendar.Calendar.from_ical(open(path, "rb").read())
    events = [c for c in cal.walk() if c.name == "VEVENT"]
    assert events, "没有 VEVENT"
    assert str(events[0]["SUMMARY"]) == "会议;含逗号,和分号", f"摘要不符: {events[0]['SUMMARY']}"


def check_vcf(path: str) -> None:
    vobject = _require("vobject")
    card = vobject.readOne(open(path, encoding="utf-8").read())
    assert str(card.fn.value) == "张三", f"FN 不符: {card.fn.value}"
    assert str(card.email.value) == "zhangsan@example.com", f"EMAIL 不符: {card.email.value}"


def check_eml(path: str) -> None:
    import email

    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f)
    assert msg["Subject"] is not None, "缺少 Subject 头"
    assert msg["From"] == "sender@example.com", f"From 不符: {msg['From']}"
    body = msg.get_payload(decode=True) or b""
    if body[:3] == b"\xef\xbb\xbf":
        raise AssertionError("正文不应带 UTF-8 BOM")
    assert b"\r\r\n" not in body, "QP 编码产生了重复回车"


CHECKS = {
    "ours.xlsx": check_xlsx,
    "ours.xls": check_xls,
    "ours.docx": check_docx,
    "ours.pptx": check_pptx,
    "ours.pdf": check_pdf,
    "ours.ods": check_ods,
    "ours.epub": check_epub,
    "ours.ics": check_ics,
    "ours.vcf": check_vcf,
    "ours.eml": check_eml,
}

CORE = {"ours.xlsx", "ours.xls", "ours.docx", "ours.pptx", "ours.pdf"}


def verify_ours(ours_dir: str) -> int:
    failures = 0
    for name, check in CHECKS.items():
        path = os.path.join(ours_dir, name)
        if not os.path.exists(path):
            print(f"[skip] {name} 不存在")
            if name in CORE:
                failures += 1
            continue
        try:
            check(path)
            print(f"[ok]   {name}")
        except Exception as ex:
            print(f"[FAIL] {name}: {type(ex).__name__}: {ex}")
            traceback.print_exc()
            failures += 1
    return failures


def generate_theirs(theirs_dir: str) -> None:
    os.makedirs(theirs_dir, exist_ok=True)
    openpyxl = _require("openpyxl")

    # xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S1"
    ws.append(["Id", "Name", "Amount"])
    ws.append([1, "甲", 1234.5])
    wb.save(os.path.join(theirs_dir, "theirs.xlsx"))

    # docx
    docx = _require("docx")
    document = docx.Document()
    document.add_paragraph("第三方库写入的段落 paragraph from python-docx")
    document.save(os.path.join(theirs_dir, "theirs.docx"))

    # pptx
    pptx = _require("pptx")
    presentation = pptx.Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    box = slide.shapes.add_textbox(0, 0, 4000000, 1000000)
    box.text_frame.text = "第三方幻灯片文本"
    presentation.save(os.path.join(theirs_dir, "theirs.pptx"))

    # pdf（reportlab，使用 ASCII 文本便于本库提取）
    reportlab = _require("reportlab")
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(os.path.join(theirs_dir, "theirs.pdf"))
    c.drawString(72, 720, "ReportLab generated PDF")
    c.showPage()
    c.save()

    # ods
    from odf import opendocument, table, text as odftext

    ods = opendocument.OpenDocumentSpreadsheet()
    t = table.Table(name="S1")
    ods.spreadsheet.addElement(t)
    row = table.TableRow()
    t.addElement(row)
    cell = table.TableCell(valuetype="string")
    cell.addElement(odftext.P(text="甲"))
    row.addElement(cell)
    ods.save(os.path.join(theirs_dir, "theirs.ods"))

    # ics
    icalendar = _require("icalendar")
    cal = icalendar.Calendar()
    cal.add("prodid", "-//interop//EN")
    cal.add("version", "2.0")
    event = icalendar.Event()
    event.add("uid", "theirs-1@interop")
    event.add("summary", "第三方日程")
    event.add("dtstart", datetime.datetime(2024, 5, 6, 8, 0, 0, tzinfo=datetime.timezone.utc))
    cal.add_component(event)
    with open(os.path.join(theirs_dir, "theirs.ics"), "wb") as f:
        f.write(cal.to_ical())

    # vcf
    vobject = _require("vobject")
    card = vobject.vCard()
    card.add("fn").value = "李四"
    card.add("n").value = vobject.vcard.Name(family="李", given="四")
    email = card.add("email")
    email.value = "lisi@example.com"
    email.type_param = "WORK"
    with open(os.path.join(theirs_dir, "theirs.vcf"), "w", encoding="utf-8", newline="") as f:
        f.write(card.serialize())

    # eml
    from email.message import EmailMessage

    msg = EmailMessage()
    msg["From"] = "third@example.com"
    msg["To"] = "someone@example.com"
    msg["Subject"] = "第三方邮件"
    msg.set_content("第三方邮件正文\r\nsecond line")
    with open(os.path.join(theirs_dir, "theirs.eml"), "wb") as f:
        f.write(bytes(msg))

    # epub
    ebooklib = _require("ebooklib")
    from ebooklib import epub

    book = epub.EpubBook()
    book.set_identifier("theirs-epub")
    book.set_title("第三方 EPUB")
    book.set_language("zh")
    chapter = epub.EpubHtml(title="第一章", file_name="chap_01.xhtml", lang="zh")
    chapter.content = "<html><body><p>第三方章节</p></body></html>"
    book.add_item(chapter)
    book.toc = (chapter,)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav", chapter]
    epub.write_epub(os.path.join(theirs_dir, "theirs.epub"), book)


def main() -> int:
    if len(sys.argv) < 2:
        print("用法: verify.py <interop_dir>", file=sys.stderr)
        return 2

    root = sys.argv[1]
    ours = os.path.join(root, "ours")
    theirs = os.path.join(root, "theirs")

    print("== 第三方读取本库文档 ==")
    failures = verify_ours(ours)

    print("== 生成第三方文档 ==")
    try:
        generate_theirs(theirs)
        print(f"[ok]   已生成 {theirs}")
    except Exception as ex:
        print(f"[FAIL] 生成第三方文档失败: {type(ex).__name__}: {ex}")
        traceback.print_exc()
        failures += 1

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
