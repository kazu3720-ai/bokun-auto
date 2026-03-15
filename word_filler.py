"""
Word テンプレート差し込み処理
docxtpl (Jinja2 ベース) を使って .docx テンプレートにデータを流し込みます。

テンプレート内のプレースホルダー一覧:
  {{ full_name }}        → 氏名（姓＋名）
  {{ last_name }}        → 姓
  {{ first_name }}       → 名
  {{ address }}          → 住所
  {{ phone }}            → 電話番号
  {{ date_of_birth }}    → 生年月日
  {{ confirmation_code }}→ 確認番号
  {{ booking_date }}     → 予約日
  {{ email }}            → メールアドレス
  {{ today }}            → 書類作成日（実行当日）
"""

import io
from datetime import datetime
from docxtpl import DocxTemplate


def fill_template(template_path: str, context: dict) -> bytes:
    """
    Word テンプレートにコンテキストデータを差し込み、
    バイト列として返します（Streamlit のダウンロードに使用）。

    Args:
        template_path: .docx テンプレートファイルのパス
        context: プレースホルダー名をキーとした辞書

    Returns:
        差し込み済み .docx のバイト列
    """
    doc = DocxTemplate(template_path)

    # 今日の日付を自動付与
    context.setdefault("today", datetime.now().strftime("%Y年%m月%d日"))

    context.pop("_raw_customer", None)
    doc.render(context)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


def fill_template_bytes(template_bytes: bytes, context: dict) -> bytes:
    """
    テンプレートをファイルパスではなくバイト列で受け取るバージョン。
    Streamlit の UploadedFile から直接使用できます。

    Args:
        template_bytes: .docx テンプレートのバイト列
        context: プレースホルダー名をキーとした辞書

    Returns:
        差し込み済み .docx のバイト列
    """
    template_buffer = io.BytesIO(template_bytes)
    doc = DocxTemplate(template_buffer)

    context.setdefault("today", datetime.now().strftime("%Y年%m月%d日"))
    context.pop("_raw_customer", None)

    doc.render(context)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.read()
