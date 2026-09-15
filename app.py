#!/usr/bin/env python3
"""Streamlit front end for ppt2word.

Users upload one or more PDF/PPTX/PPTM files.  PDFs are mandatory because they
supply page images.  When a matching PowerPoint file is present its structured
text extraction is preferred; otherwise the app falls back to Poppler text
extraction from the PDF.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from ppt2word import (
    DEFAULT_OUTPUT_NAME,
    build_minutes_from_sources,
    pair_source_files,
)

APP_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = APP_DIR / "templates"
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


def list_templates() -> list[Path]:
    return sorted(
        (
            path
            for path in TEMPLATES_DIR.iterdir()
            if path.is_file()
            and path.suffix.casefold() == ".docx"
            and not path.name.startswith("~$")
        ),
        key=lambda path: path.name.casefold(),
    )


def safe_uploaded_name(name: str) -> str:
    safe_name = Path(name).name
    if not safe_name or safe_name in {".", ".."}:
        raise ValueError("ファイル名が不正です。")
    return safe_name


def write_uploaded_files(uploaded_files, destination: Path) -> list[Path]:
    destination.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    seen_names: set[str] = set()

    for uploaded_file in uploaded_files:
        name = safe_uploaded_name(uploaded_file.name)
        key = name.casefold()
        if key in seen_names:
            raise ValueError(f"同名ファイルが複数アップロードされています: {name}")
        seen_names.add(key)

        path = destination / name
        path.write_bytes(uploaded_file.getvalue())
        written.append(path)

    return written


def save_template_to_library(uploaded_file) -> Path:
    """Persist an uploaded template without overwriting an existing template."""
    name = safe_uploaded_name(uploaded_file.name)
    if Path(name).suffix.casefold() != ".docx":
        raise ValueError("テンプレートはDOCX形式で指定してください。")

    data = uploaded_file.getvalue()
    candidate = TEMPLATES_DIR / name
    if candidate.exists():
        if candidate.read_bytes() == data:
            return candidate
        stem = candidate.stem
        suffix = candidate.suffix
        index = 2
        while True:
            alternate = TEMPLATES_DIR / f"{stem}_{index}{suffix}"
            if not alternate.exists():
                candidate = alternate
                break
            index += 1

    candidate.write_bytes(data)
    return candidate


def normalized_output_name(raw_name: str) -> str:
    name = Path(raw_name.strip()).name
    if not name:
        raise ValueError("出力ファイル名を指定してください。")
    if Path(name).suffix.casefold() != ".docx":
        name = f"{name}.docx"
    return name


def clear_result() -> None:
    st.session_state.pop("result_bytes", None)
    st.session_state.pop("result_name", None)


def main() -> None:
    st.set_page_config(page_title="PowerPoint / PDF to meeting-minutes DOCX")
    st.title("統合議事録ベース生成")
    st.caption(
        "PDFをページ画像に使用し、同名PPTX/PPTMがあれば文字列抽出を優先します。"
        "PPTX/PPTMがないPDFはPDF本文抽出へフォールバックします。"
    )

    templates = list_templates()
    if templates:
        template_mode = st.radio(
            "テンプレート",
            ("保存済みテンプレート", "新規アップロード"),
            horizontal=True,
        )
    else:
        template_mode = "新規アップロード"
        st.info(
            f"保存済みテンプレートがありません。{TEMPLATES_DIR.name}/ にDOCXを置くか、"
            "下からアップロードしてください。"
        )

    selected_template: Path | None = None
    uploaded_template = None
    register_template = False

    if template_mode == "保存済みテンプレート":
        selected_name = st.selectbox(
            "使用するテンプレート",
            [path.name for path in templates],
        )
        selected_template = next(path for path in templates if path.name == selected_name)
    else:
        uploaded_template = st.file_uploader(
            "新規テンプレート (DOCX)",
            type=["docx"],
            accept_multiple_files=False,
            key="template_upload",
        )
        register_template = st.checkbox(
            "変換成功後、このテンプレートを templates に保存する",
            value=False,
        )

    uploaded_sources = st.file_uploader(
        "変換対象 (PDF / PPTX / PPTM、複数選択可)",
        type=["pdf", "pptx", "pptm"],
        accept_multiple_files=True,
        key="source_uploads",
        help=(
            "各資料はPDFが必須です。同じベース名のPPTX/PPTMが同時にある場合は、"
            "PowerPointから文字列を抽出します。"
        ),
    )

    output_name_raw = st.text_input("出力ファイル名", value=DEFAULT_OUTPUT_NAME)

    with st.expander("詳細設定"):
        dpi = st.number_input(
            "PDF画像化 DPI",
            min_value=72,
            max_value=1200,
            value=300,
            step=25,
        )
        jpeg_quality = st.slider(
            "中間JPEG品質",
            min_value=1,
            max_value=95,
            value=85,
        )

    if st.button("変換実行", type="primary"):
        clear_result()
        try:
            if not uploaded_sources:
                raise ValueError("変換対象ファイルをアップロードしてください。")
            if selected_template is None and uploaded_template is None:
                raise ValueError("DOCXテンプレートを選択またはアップロードしてください。")

            output_name = normalized_output_name(output_name_raw)

            # All uploaded source files, uploaded templates, generated images and
            # the generated DOCX live under this directory.  The DOCX is copied
            # into memory before the context closes, so normal and exceptional
            # exits both remove the job directory.
            with tempfile.TemporaryDirectory(prefix="ppt2word_app_") as job_directory:
                job_path = Path(job_directory)
                input_paths = write_uploaded_files(
                    uploaded_sources,
                    job_path / "inputs",
                )
                sources = pair_source_files(input_paths)

                pdf_only = [source.pdf_path.name for source in sources if source.uses_pdf_text]
                if pdf_only:
                    st.warning(
                        "PPTX/PPTMがないためPDF文字列抽出にフォールバックします: "
                        + ", ".join(pdf_only)
                    )

                if selected_template is not None:
                    template_path = selected_template
                else:
                    assert uploaded_template is not None
                    template_path = job_path / "uploaded_template.docx"
                    template_path.write_bytes(uploaded_template.getvalue())

                output_path = job_path / output_name
                with st.spinner("変換中..."):
                    build_minutes_from_sources(
                        sources,
                        output_path,
                        template_path=template_path,
                        dpi=int(dpi),
                        jpeg_quality=int(jpeg_quality),
                        keep_images=False,
                        verbose=False,
                    )
                result_bytes = output_path.read_bytes()

            # At this point the complete temporary job directory has already
            # been removed.  Only the generated DOCX bytes remain in session RAM.
            st.session_state["result_bytes"] = result_bytes
            st.session_state["result_name"] = output_name

            if uploaded_template is not None and register_template:
                try:
                    saved_template = save_template_to_library(uploaded_template)
                    st.success(
                        f"変換完了。テンプレートを保存しました: {saved_template.name}"
                    )
                except OSError as exc:
                    st.warning(
                        "変換は完了しましたが、templates へのテンプレート保存に失敗しました: "
                        f"{exc}"
                    )
            else:
                st.success("変換完了。一時ファイルは削除済みです。")

        except (FileNotFoundError, ValueError, OSError) as exc:
            st.error(str(exc))
        except Exception as exc:
            # Avoid showing uploaded content or an application traceback in the UI.
            st.error(f"変換に失敗しました: {exc}")

    result_bytes = st.session_state.get("result_bytes")
    result_name = st.session_state.get("result_name")
    if result_bytes is not None and result_name is not None:
        st.download_button(
            "DOCXをダウンロード",
            data=result_bytes,
            file_name=result_name,
            mime=DOCX_MIME,
            type="primary",
        )
        if st.button("生成結果をメモリから破棄"):
            clear_result()
            st.rerun()


if __name__ == "__main__":
    main()
