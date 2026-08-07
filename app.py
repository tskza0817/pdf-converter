from __future__ import annotations

import os
from pathlib import Path

import streamlit as st
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.io.cloud_asset import CloudAsset
from adobe.pdfservices.operation.io.stream_asset import StreamAsset
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult


OUTPUT_FORMATS = {
    "Word (.docx)": {
        "extension": "docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "target_format": ExportPDFTargetFormat.DOCX,
    },
    "Excel (.xlsx)": {
        "extension": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "target_format": ExportPDFTargetFormat.XLSX,
    },
    "PowerPoint (.pptx)": {
        "extension": "pptx",
        "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "target_format": ExportPDFTargetFormat.PPTX,
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
                max-width: 900px;
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


def read_config_value(key: str) -> str:
    env_value = os.getenv(key, "")
    if env_value:
        return env_value

    try:
        secret_value = st.secrets[key]
    except Exception:
        return ""

    return str(secret_value)


def load_credentials() -> tuple[str, str]:
    client_id = read_config_value("PDF_SERVICES_CLIENT_ID").strip()
    client_secret = read_config_value("PDF_SERVICES_CLIENT_SECRET").strip()
    return client_id, client_secret


def create_pdf_services(client_id: str, client_secret: str) -> PDFServices:
    credentials = ServicePrincipalCredentials(client_id=client_id, client_secret=client_secret)
    return PDFServices(credentials=credentials)


def export_pdf_with_adobe(uploaded_pdf: bytes, output_label: str, client_id: str, client_secret: str) -> bytes:
    pdf_services = create_pdf_services(client_id, client_secret)
    input_asset = pdf_services.upload(input_stream=uploaded_pdf, mime_type=PDFServicesMediaType.PDF)

    export_pdf_params = ExportPDFParams(target_format=OUTPUT_FORMATS[output_label]["target_format"])
    export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)

    location = pdf_services.submit(export_pdf_job)
    pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
    result_asset: CloudAsset = pdf_services_response.get_result().get_asset()
    stream_asset: StreamAsset = pdf_services.get_content(result_asset)
    return stream_asset.get_input_stream()


def run_conversion(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile, output_label: str, client_id: str, client_secret: str) -> tuple[bytes, str, str]:
    output_bytes = export_pdf_with_adobe(uploaded_file.getvalue(), output_label, client_id, client_secret)
    meta = OUTPUT_FORMATS[output_label]
    output_name = f"{Path(uploaded_file.name).stem}.{meta['extension']}"
    return output_bytes, output_name, meta["mime"]


def main() -> None:
    set_page_style()
    client_id, client_secret = load_credentials()

    st.markdown(
        """
        <div class="hero">
            <h1>PDF 変換アプリ</h1>
            <p>Adobe PDF Services API を使って、PDF を Word / Excel / PowerPoint の編集可能形式へ変換します。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("PDFファイルをアップロード", type=["pdf"])
    output_label = st.selectbox("変換先の形式", list(OUTPUT_FORMATS.keys()))

    credentials_ready = bool(client_id and client_secret)
    if not credentials_ready:
        st.warning("管理者にお問い合わせください（Secrets未設定）")

    convert_clicked = st.button("変換開始", use_container_width=True, disabled=uploaded_file is None or not credentials_ready)

    if convert_clicked and uploaded_file is not None and credentials_ready:
        try:
            with st.spinner("Adobe のエンジンで変換中です。ページ数やPDFの内容によって時間がかかる場合があります。"):
                output_bytes, output_name, mime_type = run_conversion(uploaded_file, output_label, client_id, client_secret)

            st.session_state["converted_file"] = {
                "bytes": output_bytes,
                "name": output_name,
                "mime": mime_type,
            }
            st.success("変換が完了しました。")
        except (ServiceApiException, ServiceUsageException, SdkException) as error:
            st.session_state.pop("converted_file", None)
            st.error(f"Adobe PDF Services で変換に失敗しました: {error}")
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