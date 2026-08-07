from __future__ import annotations

import io
import os
import tempfile
from contextlib import suppress
from collections import Counter
from pathlib import Path
from statistics import median

import pandas as pd
import pdfplumber
import streamlit as st
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches, Pt


OUTPUT_FORMATS = {
    "Word (.docx)": {
        "extension": "docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
    "Excel (.xlsx)": {
        "extension": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    "PowerPoint (.pptx)": {
        "extension": "pptx",
        "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    },
}


def set_page_style() -> None:
    st.set_page_config(page_title="PDF変換アプリ", page_icon="📄", layout="centered")
    st.markdown(
        """
        <style>
            .stApp {
                background: linear-gradient(180deg, #f7f9fc 0%, #eef3f9 100%);
            }
            .block-container {
                max-width: 860px;
                padding-top: 2rem;
                padding-bottom: 2rem;
            }
            .hero {
                background: white;
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 20px;
                padding: 1.5rem 1.6rem;
                box-shadow: 0 16px 40px rgba(15, 23, 42, 0.08);
                margin-bottom: 1.25rem;
            }
            .hero h1 {
                margin: 0;
                font-size: 1.8rem;
            }
            .hero p {
                margin: 0.45rem 0 0;
                color: #475569;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def make_safe_sheet_name(name: str, fallback: str) -> str:
    sanitized = "".join(char if char not in r'[]:*?/\\' else "_" for char in name).strip()
    sanitized = sanitized[:31]
    return sanitized or fallback


def normalize_table(table: list[list[str | None]]) -> pd.DataFrame:
    if not table:
        return pd.DataFrame()

    max_columns = max((len(row) for row in table), default=0)
    normalized_rows: list[list[str | None]] = []
    for row in table:
        normalized_rows.append(list(row) + [None] * (max_columns - len(row)))

    header = normalized_rows[0]
    body = normalized_rows[1:]
    if any(value is None for value in header):
        header = [f"column_{index + 1}" for index in range(max_columns)]

    return pd.DataFrame(body, columns=[str(value).strip() if value is not None else f"column_{index + 1}" for index, value in enumerate(header)])


def group_pdf_chars_into_lines(chars: list[dict], tolerance: float = 2.5) -> list[list[dict]]:
    if not chars:
        return []

    ordered_chars = sorted(chars, key=lambda char: (float(char.get("top", 0.0)), float(char.get("x0", 0.0))))
    grouped_lines: list[list[dict]] = []
    current_line: list[dict] = []
    current_top: float | None = None

    for char in ordered_chars:
        char_top = float(char.get("top", 0.0))
        if not current_line:
            current_line = [char]
            current_top = char_top
            continue

        if current_top is not None and abs(char_top - current_top) <= tolerance:
            current_line.append(char)
            current_top = (current_top + char_top) / 2
            continue

        grouped_lines.append(current_line)
        current_line = [char]
        current_top = char_top

    if current_line:
        grouped_lines.append(current_line)

    return grouped_lines


def create_textbox_for_line(slide, line_chars: list[dict], page_width: float, page_height: float, slide_width_in: float, slide_height_in: float) -> None:
    if not line_chars:
        return

    ordered_chars = sorted(line_chars, key=lambda char: float(char.get("x0", 0.0)))
    text = "".join(str(char.get("text", "")) for char in ordered_chars).rstrip()
    if not text.strip():
        return

    left_pt = min(float(char.get("x0", 0.0)) for char in ordered_chars)
    top_pt = min(float(char.get("top", 0.0)) for char in ordered_chars)
    right_pt = max(float(char.get("x1", 0.0)) for char in ordered_chars)
    bottom_pt = max(float(char.get("bottom", 0.0)) for char in ordered_chars)
    font_sizes = [float(char.get("size", 0.0)) for char in ordered_chars if char.get("size")]
    font_size_pt = max(median(font_sizes), 1.0) if font_sizes else max(bottom_pt - top_pt, 10.0)

    scale_x = slide_width_in / (page_width / 72.0)
    scale_y = slide_height_in / (page_height / 72.0)

    left = Inches((left_pt / 72.0) * scale_x)
    top = Inches((top_pt / 72.0) * scale_y)
    width = Inches(max(((right_pt - left_pt) / 72.0) * scale_x, font_size_pt / 72.0 * scale_x * 0.6))
    height = Inches(max(((bottom_pt - top_pt) / 72.0) * scale_y, (font_size_pt / 72.0) * scale_y * 1.3))

    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.clear()
    text_frame.word_wrap = False
    text_frame.margin_left = Pt(0)
    text_frame.margin_right = Pt(0)
    text_frame.margin_top = Pt(0)
    text_frame.margin_bottom = Pt(0)

    paragraph = text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(font_size_pt)

    font_names = [str(char.get("fontname")) for char in ordered_chars if char.get("fontname")]
    if font_names:
        run.font.name = Counter(font_names).most_common(1)[0][0]


def convert_pdf_to_docx(pdf_path: str) -> bytes:
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = os.path.join(temp_dir, "converted.docx")
        converter = Converter(pdf_path)
        try:
            converter.convert(output_path, start=0, end=None)
        finally:
            converter.close()

        with open(output_path, "rb") as output_file:
            return output_file.read()


def convert_pdf_to_xlsx(pdf_path: str) -> bytes:
    output_buffer = io.BytesIO()

    with pdfplumber.open(pdf_path) as pdf:
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            sheet_written = False
            for page_index, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables() or []
                if not tables:
                    continue

                for table_index, table in enumerate(tables, start=1):
                    dataframe = normalize_table(table)
                    if dataframe.empty:
                        continue

                    sheet_name = make_safe_sheet_name(
                        f"page{page_index}_table{table_index}",
                        fallback=f"page{page_index}",
                    )
                    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheet_written = True

            if not sheet_written:
                info_frame = pd.DataFrame(
                    {
                        "message": [
                            "このPDFからは表を抽出できませんでした。",
                            "テキスト主体のPDFや罫線の弱い表では抽出精度が下がることがあります。",
                        ]
                    }
                )
                info_frame.to_excel(writer, sheet_name="result", index=False)

    output_buffer.seek(0)
    return output_buffer.read()


def convert_pdf_to_pptx(pdf_path: str) -> bytes:
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages
        if not pages:
            raise ValueError("PDFからスライドを生成できませんでした。")

        max_page_width = max(float(page.width) for page in pages)
        max_page_height = max(float(page.height) for page in pages)

        presentation = Presentation()
        presentation.slide_width = Inches(max_page_width / 72.0)
        presentation.slide_height = Inches(max_page_height / 72.0)

        slide_width_in = max_page_width / 72.0
        slide_height_in = max_page_height / 72.0
        blank_layout = presentation.slide_layouts[6]

        for page in pages:
            slide = presentation.slides.add_slide(blank_layout)
            chars = page.chars or []

            for line_chars in group_pdf_chars_into_lines(chars):
                create_textbox_for_line(
                    slide,
                    line_chars,
                    page_width=float(page.width),
                    page_height=float(page.height),
                    slide_width_in=slide_width_in,
                    slide_height_in=slide_height_in,
                )

    output_buffer = io.BytesIO()
    presentation.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.read()


def cleanup_path(file_path: str | None) -> None:
    if not file_path:
        return

    with suppress(FileNotFoundError, PermissionError, OSError):
        Path(file_path).unlink()


def write_upload_to_temp_file(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile) -> str:
    suffix = Path(uploaded_file.name).suffix or ".pdf"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        return temp_file.name


def run_conversion(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile, output_label: str) -> tuple[bytes, str, str]:
    pdf_path = write_upload_to_temp_file(uploaded_file)
    try:
        if output_label == "Word (.docx)":
            output_bytes = convert_pdf_to_docx(pdf_path)
        elif output_label == "Excel (.xlsx)":
            output_bytes = convert_pdf_to_xlsx(pdf_path)
        elif output_label == "PowerPoint (.pptx)":
            output_bytes = convert_pdf_to_pptx(pdf_path)
        else:
            raise ValueError("未対応の変換形式です。")

        meta = OUTPUT_FORMATS[output_label]
        output_name = f"{Path(uploaded_file.name).stem}.{meta['extension']}"
        return output_bytes, output_name, meta["mime"]
    finally:
        cleanup_path(pdf_path)


def main() -> None:
    set_page_style()

    st.markdown(
        """
        <div class="hero">
            <h1>PDF 変換アプリ</h1>
            <p>PDF を Word / Excel / PowerPoint に変換して、すぐにダウンロードできます。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("PDFファイルをアップロード", type=["pdf"])
    output_label = st.selectbox("変換先の形式", list(OUTPUT_FORMATS.keys()))

    convert_clicked = st.button("変換開始", use_container_width=True, disabled=uploaded_file is None)

    if convert_clicked and uploaded_file is not None:
        try:
            with st.spinner("変換中です。ファイルサイズやページ数によって少し時間がかかる場合があります。"):
                output_bytes, output_name, mime_type = run_conversion(uploaded_file, output_label)

            st.session_state["converted_file"] = {
                "bytes": output_bytes,
                "name": output_name,
                "mime": mime_type,
            }
            st.success("変換が完了しました。")
        except Exception as error:
            st.session_state.pop("converted_file", None)
            st.error(f"変換に失敗しました: {error}")

    converted_file = st.session_state.get("converted_file")
    if converted_file:
        st.download_button(
            label="変換ファイルをダウンロード",
            data=converted_file["bytes"],
            file_name=converted_file["name"],
            mime=converted_file["mime"],
            use_container_width=True,
        )


if __name__ == "__main__":
    main()