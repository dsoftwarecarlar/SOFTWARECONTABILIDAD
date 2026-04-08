from __future__ import annotations

import argparse
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


DOCUMENT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""


SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>
"""


STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="es-EC"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="220"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="1F1F1F"/>
      <w:sz w:val="34"/>
      <w:szCs w:val="34"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="160"/>
    </w:pPr>
    <w:rPr>
      <w:color w:val="555555"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="280" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="0E4F7A"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="240" w:after="100"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="0E4F7A"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="200" w:after="80"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="0E4F7A"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BulletList">
    <w:name w:val="Bullet List"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="540" w:hanging="220"/>
      <w:spacing w:after="100"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="NumberList">
    <w:name w:val="Number List"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:ind w:left="540" w:hanging="220"/>
      <w:spacing w:after="100"/>
    </w:pPr>
  </w:style>
</w:styles>
"""


INLINE_TOKEN_RE = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`)")


def make_core_xml(title: str) -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    title_text = escape(title)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{title_text}</dc:title>
  <dc:subject>Manual de usuario</dc:subject>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>
"""


APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>Automotores Carlos Larrea</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>1.0</AppVersion>
</Properties>
"""


def parse_inline(text: str) -> list[dict[str, str | bool]]:
    parts: list[dict[str, str | bool]] = []
    position = 0
    for match in INLINE_TOKEN_RE.finditer(text):
        if match.start() > position:
            parts.append({"text": text[position:match.start()], "bold": False, "code": False})
        token = match.group(0)
        if token.startswith("**") and token.endswith("**"):
            parts.append({"text": token[2:-2], "bold": True, "code": False})
        elif token.startswith("`") and token.endswith("`"):
            parts.append({"text": token[1:-1], "bold": False, "code": True})
        position = match.end()
    if position < len(text):
        parts.append({"text": text[position:], "bold": False, "code": False})
    if not parts:
        parts.append({"text": text, "bold": False, "code": False})
    return parts


def parse_markdown(markdown_text: str) -> list[dict[str, object]]:
    lines = markdown_text.splitlines()
    blocks: list[dict[str, object]] = []
    paragraph_lines: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph_lines
        if not paragraph_lines:
            return
        text = " ".join(part.strip() for part in paragraph_lines if part.strip())
        if text:
            blocks.append({"type": "paragraph", "text": text})
        paragraph_lines = []

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped == "[[PAGE_BREAK]]":
            flush_paragraph()
            blocks.append({"type": "page_break"})
            continue

        if stripped == "":
            flush_paragraph()
            continue

        heading_match = re.match(r"^(#{1,3})\s+(.*)$", stripped)
        if heading_match:
            flush_paragraph()
            level = len(heading_match.group(1))
            blocks.append({"type": f"heading{level}", "text": heading_match.group(2).strip()})
            continue

        if re.match(r"^- ", stripped):
            flush_paragraph()
            blocks.append({"type": "bullet", "text": stripped[2:].strip()})
            continue

        if re.match(r"^\d+\.\s+", stripped):
            flush_paragraph()
            blocks.append({"type": "number", "text": stripped})
            continue

        paragraph_lines.append(stripped)

    flush_paragraph()
    return blocks


def make_run_xml(text: str, *, bold: bool = False, code: bool = False) -> str:
    preserve = ' xml:space="preserve"' if text.startswith(" ") or text.endswith(" ") or "  " in text else ""
    run_props: list[str] = []
    if bold:
        run_props.append("<w:b/>")
    if code:
        run_props.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>')
        run_props.append("<w:color w:val=\"404040\"/>")
    props_xml = f"<w:rPr>{''.join(run_props)}</w:rPr>" if run_props else ""
    return f"<w:r>{props_xml}<w:t{preserve}>{escape(text)}</w:t></w:r>"


def make_paragraph_xml(style: str, runs: list[dict[str, str | bool]]) -> str:
    runs_xml = "".join(
        make_run_xml(
            str(run["text"]),
            bold=bool(run["bold"]),
            code=bool(run["code"]),
        )
        for run in runs
    )
    return (
        "<w:p>"
        f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>"
        f"{runs_xml}"
        "</w:p>"
    )


def build_document_xml(blocks: list[dict[str, object]]) -> str:
    body_parts: list[str] = []
    first_heading_rendered = False
    first_block = True

    for block in blocks:
        block_type = str(block["type"])
        if block_type == "page_break":
            body_parts.append("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>")
            first_block = False
            continue

        text = str(block.get("text", ""))
        runs = parse_inline(text)

        if first_block and block_type == "heading1":
            style = "Title"
            first_heading_rendered = True
        elif not first_heading_rendered and block_type == "paragraph":
            style = "Subtitle"
        elif block_type == "heading1":
            style = "Heading1"
            first_heading_rendered = True
        elif block_type == "heading2":
            style = "Heading2"
        elif block_type == "heading3":
            style = "Heading3"
        elif block_type == "bullet":
            runs = parse_inline(f"- {text}")
            style = "BulletList"
        elif block_type == "number":
            style = "NumberList"
        else:
            style = "Normal"

        body_parts.append(make_paragraph_xml(style, runs))
        first_block = False

    body_parts.append(
        "<w:sectPr>"
        "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
        "</w:sectPr>"
    )

    body_xml = "".join(body_parts)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{body_xml}</w:body>
</w:document>
"""


def extract_title(blocks: list[dict[str, object]], fallback: str) -> str:
    for block in blocks:
        if block["type"] == "heading1":
            return str(block.get("text", fallback))
    return fallback


def generate_docx(markdown_path: Path, output_path: Path) -> None:
    markdown_text = markdown_path.read_text(encoding="utf-8")
    blocks = parse_markdown(markdown_text)
    title = extract_title(blocks, "Manual de Usuario")
    document_xml = build_document_xml(blocks)
    core_xml = make_core_xml(title)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        package.writestr("_rels/.rels", ROOT_RELS_XML)
        package.writestr("docProps/core.xml", core_xml)
        package.writestr("docProps/app.xml", APP_XML)
        package.writestr("word/document.xml", document_xml)
        package.writestr("word/styles.xml", STYLES_XML)
        package.writestr("word/settings.xml", SETTINGS_XML)
        package.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera un archivo DOCX simple a partir de un Markdown controlado.")
    parser.add_argument("source", type=Path, help="Ruta del archivo Markdown fuente.")
    parser.add_argument("output", type=Path, help="Ruta del archivo DOCX de salida.")
    args = parser.parse_args()
    generate_docx(args.source, args.output)


if __name__ == "__main__":
    main()
