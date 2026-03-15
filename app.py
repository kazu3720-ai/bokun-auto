"""
Bokun 予約データ → Word 差し込みツール

起動方法:
  streamlit run app.py
"""

import streamlit as st
import os
from pathlib import Path
from datetime import date, timedelta
from dotenv import load_dotenv

from bokun_client import BokunClient, BokunAPIError
from word_filler import fill_template_bytes

load_dotenv()

SAVED_TEMPLATE_PATH = Path(__file__).parent / "saved_template.docx"
DEFAULT_TEMPLATE_PATH = Path(__file__).parent / "default_template.docx"


def _get_secret(key: str, default: str = "") -> str:
    """Streamlit Cloud の st.secrets → .env の順で値を取得します"""
    try:
        return st.secrets.get(key, os.getenv(key, default))
    except Exception:
        return os.getenv(key, default)

# ─────────────────────────────────────────────
# ページ設定
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Bokun 予約 → Word 差し込みツール",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Bokun 予約データ → Word 差し込みツール")
st.caption("予約情報を取得して Word 書類を自動作成します")


# ─────────────────────────────────────────────
# テンプレートのバイト列を取得（保存済み or 新規アップロード）
# ─────────────────────────────────────────────
def get_template_bytes(uploaded_file) -> bytes | None:
    """
    テンプレートのバイト列を返します。優先順位：
    1. 新規アップロード（保存して以降も使用）
    2. 保存済みテンプレート（saved_template.docx）
    3. GitHubに含めたデフォルトテンプレート（default_template.docx）
    """
    if uploaded_file is not None:
        data = uploaded_file.read()
        try:
            SAVED_TEMPLATE_PATH.write_bytes(data)
        except Exception:
            pass  # クラウド環境でファイル書き込みできない場合はスキップ
        return data
    if SAVED_TEMPLATE_PATH.exists():
        return SAVED_TEMPLATE_PATH.read_bytes()
    if DEFAULT_TEMPLATE_PATH.exists():
        return DEFAULT_TEMPLATE_PATH.read_bytes()
    return None


# ─────────────────────────────────────────────
# 共通: Word ダウンロードセクション
# ─────────────────────────────────────────────
def render_download_section(info: dict, template_bytes: bytes | None, key_suffix: str):
    """差し込み済み Word のダウンロードボタンを表示します"""
    if template_bytes is None:
        st.warning(
            "サイドバーから Word テンプレート (.docx) をアップロードしてください。\n\n"
            "テンプレートがない場合は `python create_template.py` を実行してください。"
        )
        return

    try:
        output_bytes = fill_template_bytes(template_bytes, info.copy())
        conf = info.get("confirmation_code", "document")
        name = info.get("full_name", "").replace(" ", "_")
        filename = f"予約書_{conf}_{name}.docx"

        st.download_button(
            label="⬇️ 差し込み済み Word をダウンロード",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key=f"download_{key_suffix}",
        )
    except Exception as e:
        st.error(
            f"Word 生成エラー: {e}\n\n"
            "テンプレートの {{ プレースホルダー }} が正しいか確認してください。"
        )


# ─────────────────────────────────────────────
# サイドバー：API 設定 & テンプレートアップロード
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ API 設定")
    access_key = st.text_input(
        "アクセスキー",
        value=_get_secret("BOKUN_ACCESS_KEY"),
        type="password",
    )
    secret_key = st.text_input(
        "シークレットキー",
        value=_get_secret("BOKUN_SECRET_KEY"),
        type="password",
    )

    st.divider()
    st.header("📂 Word テンプレート")

    if SAVED_TEMPLATE_PATH.exists():
        st.success("保存済みテンプレートを使用中")
        if st.button("🗑️ テンプレートを削除して差し替える"):
            SAVED_TEMPLATE_PATH.unlink()
            st.rerun()
        uploaded_template = None
    elif DEFAULT_TEMPLATE_PATH.exists():
        st.success("デフォルトテンプレートを使用中")
        uploaded_template = st.file_uploader(
            "別のテンプレートに差し替える（任意）",
            type=["docx"],
        )
    else:
        uploaded_template = st.file_uploader(
            "テンプレート (.docx) をアップロード",
            type=["docx"],
            help="{{ full_name }} などのプレースホルダーを含む .docx ファイル",
        )
        st.caption("一度アップロードすると次回起動時も自動で読み込まれます。")

template_bytes = get_template_bytes(uploaded_template)


# ─────────────────────────────────────────────
# メインエリア：タブ切り替え
# ─────────────────────────────────────────────
tab_search, tab_manual = st.tabs(
    ["🔍 予約を検索して差し込み", "✏️ 手動入力で差し込み"]
)


# ── タブ1: API から予約を検索 ──────────────────
with tab_search:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("確認番号で検索")
        confirmation_code = st.text_input(
            "予約確認番号",
            placeholder="例: OKI-12345",
        )
        search_by_code = st.button("🔎 この番号で取得", use_container_width=True)

    with col2:
        st.subheader("日付範囲で検索")
        d_from = st.date_input("開始日", value=date.today())
        d_to = st.date_input("終了日", value=date.today() + timedelta(days=7))
        date_type = st.radio(
            "日付の種類",
            options=["start", "creation"],
            format_func=lambda x: "🚴 ツアー利用日" if x == "start" else "📋 予約作成日",
            horizontal=True,
            help="ツアー利用日：実際に体験する日　予約作成日：お客様が予約した日",
        )
        product_filter = st.text_input(
            "プラン名で絞り込み（任意・部分一致）",
            value="E-BIKE",
            help="入力したキーワードを含む予約だけ表示します。空欄で全件取得。",
        )
        search_by_date = st.button("📅 一覧を取得", use_container_width=True)

    st.divider()

    # ── 確認番号で取得 ──
    if search_by_code:
        if not access_key or not secret_key:
            st.error("サイドバーでアクセスキーとシークレットキーを入力してください。")
        elif not confirmation_code.strip():
            st.warning("確認番号を入力してください。")
        else:
            with st.spinner("予約データを取得中..."):
                try:
                    client = BokunClient(access_key, secret_key)
                    bookings = client.search_bookings_by_confirmation(confirmation_code.strip())
                    if not bookings:
                        st.warning("該当する予約が見つかりませんでした。確認番号を再確認してください。")
                    else:
                        st.session_state["fetched_bookings"] = bookings
                        st.session_state["selected_idx"] = 0
                        st.success(f"{len(bookings)} 件の予約データを取得しました。")
                except BokunAPIError as e:
                    st.error(f"取得エラー: HTTP {e.status_code}")
                    st.code(e.response_body, language="json")
                except Exception as e:
                    st.error(f"取得エラー: {e}")

    # ── 日付範囲で取得 ──
    if search_by_date:
        if not access_key or not secret_key:
            st.error("サイドバーでアクセスキーとシークレットキーを入力してください。")
        else:
            with st.spinner("予約一覧を取得中..."):
                try:
                    client = BokunClient(access_key, secret_key)
                    bookings = client.search_bookings_by_date(
                        date_from=d_from.strftime("%Y-%m-%d"),
                        date_to=d_to.strftime("%Y-%m-%d"),
                        product_keyword=product_filter,
                        date_type=date_type,
                    )
                    st.session_state["fetched_bookings"] = bookings
                    st.session_state["selected_idx"] = 0
                    st.success(f"{len(bookings)} 件の予約が見つかりました。")
                except BokunAPIError as e:
                    st.error(f"取得エラー: HTTP {e.status_code}")
                    st.code(e.response_body, language="json")
                except Exception as e:
                    st.error(f"取得エラー: {e}")

    # ── 取得した予約の表示・選択 ──
    if "fetched_bookings" in st.session_state:
        bookings = st.session_state["fetched_bookings"]

        if len(bookings) > 1:
            # インデックス番号付きラベルでキー重複を防ぐ
            labels = []
            for i, b in enumerate(bookings):
                code = b.get("confirmationCode", "不明")
                customer = b.get("customer") or {}
                name = (
                    f"{customer.get('lastName', '')} {customer.get('firstName', '')}".strip()
                    or "名前不明"
                )
                start = b.get("startDate", "")
                labels.append(f"[{i+1:02d}] {code} — {name}")

            selected_idx = st.selectbox(
                f"差し込む予約を選択してください（全{len(bookings)}件）",
                range(len(labels)),
                format_func=lambda i: labels[i],
                key="booking_selector",
            )
            selected_booking = bookings[selected_idx]
        else:
            selected_booking = bookings[0] if bookings else None

        if selected_booking:
            client = BokunClient(access_key, secret_key)
            conf_code = selected_booking.get("confirmationCode", "")
            with st.spinner("詳細データ（生年月日など）を取得中..."):
                try:
                    full_booking = client.get_full_booking(conf_code)
                    selected_booking = full_booking
                except Exception:
                    pass

            info = client.extract_customer_info(selected_booking)

            st.subheader("取得した予約情報")
            col_a, col_b = st.columns(2)
            with col_a:
                st.metric("氏名", info["full_name"] or "—")
                st.metric("電話番号", info["phone"] or "—")
                st.metric("生年月日", info["date_of_birth"] or "—")
                st.metric("利用日", info["start_date"] or "—")
            with col_b:
                st.metric("確認番号", info["confirmation_code"] or "—")
                st.metric("予約日", info["booking_date"] or "—")
                st.metric("プラン名", info["product_title"] or "—")
                st.metric("メール", info["email"] or "—")
            st.write("**住所:**", info["address"] or "—")

            if not info["date_of_birth"]:
                st.warning("生年月日が取得できませんでした。Bokun の予約データに登録がない可能性があります。")

            with st.expander("生データを確認する（デバッグ用）"):
                st.write("**顧客情報:**")
                st.json(info.get("_raw_customer", {}))
                st.write("**activityBookings（利用日の確認）:**")
                st.json(selected_booking.get("activityBookings", []))

            st.divider()
            render_download_section(info, template_bytes, key_suffix="search")


# ── タブ2: 手動入力 ────────────────────────────
with tab_manual:
    st.subheader("顧客情報を手動で入力")
    st.caption("API を使わず直接入力して Word を生成することもできます。")

    col1, col2 = st.columns(2)
    with col1:
        m_last = st.text_input("姓（漢字）", placeholder="例: 山田", key="m_last")
        m_first = st.text_input("名（漢字）", placeholder="例: 太郎", key="m_first")
        m_phone = st.text_input("電話番号", placeholder="例: 090-1234-5678", key="m_phone")
        m_dob = st.text_input("生年月日", placeholder="例: 1990年01月15日", key="m_dob")
    with col2:
        m_address = st.text_area(
            "住所", placeholder="例: 東京都渋谷区1-2-3", height=100, key="m_addr"
        )
        m_email = st.text_input(
            "メールアドレス", placeholder="例: yamada@example.com", key="m_email"
        )
        m_conf = st.text_input("確認番号（任意）", placeholder="例: OKI-12345", key="m_conf")

    st.divider()
    if st.button("📄 Word を生成してダウンロード", use_container_width=True, key="gen_manual"):
        manual_info = {
            "full_name": f"{m_last} {m_first}".strip(),
            "last_name": m_last,
            "first_name": m_first,
            "address": m_address,
            "phone": m_phone,
            "date_of_birth": m_dob,
            "email": m_email,
            "confirmation_code": m_conf,
            "booking_date": "",
            "start_date": "",
            "product_title": "",
        }
        render_download_section(manual_info, template_bytes, key_suffix="manual")
