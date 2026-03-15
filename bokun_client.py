"""
Bokun API クライアント

認証仕様（公式ドキュメントより）:
  ヘッダー  : X-Bokun-Date, X-Bokun-AccessKey, X-Bokun-Signature
  日付形式  : "YYYY-MM-DD HH:MM:SS" (UTC)
  署名アルゴリズム: HMAC-SHA1 → Base64エンコード
  署名メッセージ : {date}{access_key}{METHOD}{path}  ← 改行なし
"""

import hashlib
import hmac
import base64
import re
import requests
from datetime import datetime, timezone
import os
from dotenv import load_dotenv

load_dotenv()


class BokunAPIError(Exception):
    """Bokun API のエラー詳細を保持する例外クラス"""
    def __init__(self, status_code: int, url: str, response_body: str):
        self.status_code = status_code
        self.url = url
        self.response_body = response_body
        super().__init__(f"HTTP {status_code}: {url}")


class BokunClient:
    def __init__(self, access_key: str = None, secret_key: str = None):
        self.access_key = access_key or os.getenv("BOKUN_ACCESS_KEY", "")
        self.secret_key = secret_key or os.getenv("BOKUN_SECRET_KEY", "")
        self.base_url = os.getenv("BOKUN_BASE_URL", "https://api.bokun.io").rstrip("/")

    def _make_date_str(self) -> str:
        """UTC時刻を Bokun が要求する形式で返します: 'YYYY-MM-DD HH:MM:SS'"""
        return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    def _make_signature(self, date_str: str, method: str, path: str) -> str:
        """
        HMAC-SHA1 署名を生成します。
        メッセージ: {date}{access_key}{METHOD}{path} (改行なし)
        """
        message = f"{date_str}{self.access_key}{method.upper()}{path}"
        secret_bytes = self.secret_key.encode("utf-8")
        message_bytes = message.encode("utf-8")
        signature = hmac.new(secret_bytes, message_bytes, hashlib.sha1).digest()
        return base64.b64encode(signature).decode("utf-8")

    def _get_headers(self, method: str, path: str) -> dict:
        """認証ヘッダーを生成します"""
        date_str = self._make_date_str()
        signature = self._make_signature(date_str, method, path)
        return {
            "X-Bokun-Date": date_str,
            "X-Bokun-AccessKey": self.access_key,
            "X-Bokun-Signature": signature,
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json",
        }

    def _request(self, method: str, path: str, params: dict = None, json: dict = None):
        """API リクエストを実行します"""
        full_path = path
        if params:
            query = "&".join(f"{k}={v}" for k, v in params.items())
            full_path = f"{path}?{query}"

        url = self.base_url + full_path
        headers = self._get_headers(method, full_path)
        resp = requests.request(
            method, url, headers=headers, json=json, timeout=30
        )
        if not resp.ok:
            raise BokunAPIError(
                status_code=resp.status_code,
                url=url,
                response_body=resp.text,
            )
        return resp.json()

    def get_full_booking(self, confirmation_code: str) -> dict:
        """
        確認番号でフル予約データを取得します。
        生年月日を含む完全な顧客情報が得られます。
        GET /booking.json/booking/{confirmationCode}
        """
        path = f"/booking.json/booking/{confirmation_code}"
        return self._request("GET", path)

    def search_bookings_by_confirmation(self, confirmation_code: str) -> list:
        """確認番号（例: ICT-8582）で予約を検索します。"""
        path = "/booking.json/product-booking-search"
        body = {
            "confirmationCode": confirmation_code.strip(),
            "pageSize": 5,
            "page": 0,
        }
        result = self._request("POST", path, json=body)
        return _extract_results(result)

    def search_bookings_by_date(
        self,
        date_from: str,
        date_to: str,
        product_keyword: str = "",
        date_type: str = "start",
        page_size: int = 100,
    ) -> list:
        """
        日付範囲で予約一覧を全件取得し、アプリ側でキーワード絞り込みします。
        date_type : "start"=ツアー利用日, "creation"=予約作成日
        product_keyword : プラン名の部分一致（大文字小文字を無視）
        """
        path = "/booking.json/product-booking-search"
        date_range = {
            "from": f"{date_from}T00:00:00",
            "to": f"{date_to}T23:59:59",
            "includeLower": True,
            "includeUpper": True,
        }

        all_raw: list = []
        page = 0

        while True:
            body: dict = {
                "pageSize": page_size,
                "page": page,
            }
            if date_type == "start":
                body["startDateRange"] = date_range
            else:
                body["creationDateRange"] = date_range

            result = self._request("POST", path, json=body)
            page_items = _extract_results_raw(result)
            all_raw.extend(page_items)

            if len(page_items) < page_size:
                break
            page += 1

        # 確認番号で重複除去
        seen: set = set()
        unique: list = []
        for b in all_raw:
            code = b.get("confirmationCode", "")
            key = code if code else id(b)
            if key in seen:
                continue
            seen.add(key)
            unique.append(b)

        # キーワードでアプリ側フィルタリング（部分一致・大文字小文字無視）
        if product_keyword.strip():
            kw = product_keyword.strip().lower()
            unique = [b for b in unique if _booking_matches_keyword(b, kw)]

        return unique

    def extract_customer_info(self, booking: dict) -> dict:
        """
        予約データから差し込み用のフィールドを抽出します。
        Word テンプレートのプレースホルダーと対応しています。
        """
        customer = booking.get("customer") or {}

        first_name = customer.get("firstName", "")
        last_name = customer.get("lastName", "")
        full_name = f"{last_name} {first_name}".strip()

        address = _format_address_jp(customer)

        dob_raw = customer.get("dateOfBirth", "")
        dob = _format_date(dob_raw)
        age = _calc_age(dob_raw)
        dob_with_age = f"{dob}（{age}歳）" if dob and age is not None else dob

        booking_date = _format_date(booking.get("creationDate", ""))
        phone = _format_phone(customer.get("phoneNumber", ""))

        start_date = _extract_start_date(booking)
        product_title = _extract_product_title(booking)

        return {
            "full_name": full_name,
            "last_name": last_name,
            "first_name": first_name,
            "address": address,
            "phone": phone,
            "date_of_birth": dob_with_age,
            "age": str(age) if age is not None else "",
            "confirmation_code": booking.get("confirmationCode", ""),
            "booking_date": booking_date,
            "start_date": start_date,
            "product_title": product_title,
            "email": customer.get("email", ""),
            "_raw_customer": customer,
        }


def _extract_start_date(booking: dict) -> str:
    """
    予約データから利用日時を取り出します。
    複数のフィールドパスを順番に試します。
    """
    candidates = []

    # フル予約: activityBookings リストを探索
    for ab in (booking.get("activityBookings") or []):
        candidates += [
            ab.get("startDate"),
            ab.get("date"),
            (ab.get("activityAvailability") or {}).get("startTime"),
            (ab.get("activityAvailability") or {}).get("date"),
        ]

    # 検索結果レスポンスのフラットフィールド
    candidates += [
        booking.get("startDate"),
        booking.get("startTime"),
        booking.get("date"),
    ]

    for raw in candidates:
        if raw:
            return _format_datetime(raw)

    return ""


def _extract_product_title(booking: dict) -> str:
    """予約データからプラン名（商品名）を取り出します"""
    activity_bookings = booking.get("activityBookings") or []
    if activity_bookings:
        activity = activity_bookings[0].get("activity") or {}
        return activity.get("title", "")

    # 検索結果レスポンス
    product = booking.get("product") or {}
    return product.get("title", "") or booking.get("productTitle", "")


def _extract_results_raw(api_response) -> list:
    """API レスポンスから予約アイテムをそのまま取り出します（重複除去なし）"""
    if isinstance(api_response, list):
        items = []
        for item in api_response:
            items.extend(item.get("results", []) if isinstance(item, dict) else [])
        return items
    if isinstance(api_response, dict):
        return api_response.get("results", [])
    return []


def _extract_results(api_response) -> list:
    """後方互換のためのラッパー（確認番号で重複除去あり）"""
    raw = _extract_results_raw(api_response)
    seen: set = set()
    unique: list = []
    for b in raw:
        code = b.get("confirmationCode", "")
        if code and code in seen:
            continue
        seen.add(code)
        unique.append(b)
    return unique


def _booking_matches_keyword(booking: dict, keyword: str) -> bool:
    """
    予約データのいずれかのフィールドにキーワードが含まれるか確認します。
    大文字小文字を無視した部分一致。
    """
    targets = []

    # 検索結果レスポンスのフラットフィールド
    product = booking.get("product") or {}
    targets.append(product.get("title", ""))
    targets.append(booking.get("productTitle", ""))
    targets.append(booking.get("title", ""))

    # フル予約レスポンスの activityBookings
    for ab in (booking.get("activityBookings") or []):
        activity = ab.get("activity") or {}
        targets.append(activity.get("title", ""))
        targets.append(ab.get("title", ""))

    return any(keyword in str(t).lower() for t in targets if t)


WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]


def _format_datetime(raw) -> str:
    """日付データを「YYYY年MM月DD日（曜）」形式に変換します"""
    if not raw:
        return ""
    if isinstance(raw, (int, float)):
        try:
            dt = datetime.fromtimestamp(raw / 1000, tz=timezone.utc)
            return dt.strftime("%Y年%m月%d日") + f"（{WEEKDAYS_JP[dt.weekday()]}）"
        except Exception:
            return str(raw)
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(str(raw)[: len(fmt)], fmt)
            return dt.strftime("%Y年%m月%d日") + f"（{WEEKDAYS_JP[dt.weekday()]}）"
        except ValueError:
            continue
    return str(raw)


def _calc_age(raw) -> "Optional[int]":
    """生年月日から現在の年齢を計算します"""
    if not raw:
        return None
    try:
        if isinstance(raw, (int, float)):
            dob = datetime.fromtimestamp(raw / 1000, tz=timezone.utc)
        else:
            dob = None
            for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
                try:
                    dob = datetime.strptime(str(raw)[: len(fmt)], fmt)
                    break
                except ValueError:
                    continue
        if dob is None:
            return None
        today = datetime.now()
        age = today.year - dob.year
        if (today.month, today.day) < (dob.month, dob.day):
            age -= 1
        return age
    except Exception:
        return None


def _format_address_jp(customer: dict) -> str:
    """
    Bokun の顧客データから日本語順の住所を生成します。
    〒XXX-XXXX 都道府県＋市区町村・番地
    """
    raw_post = re.sub(r"\D", "", str(customer.get("postCode", "")))
    if len(raw_post) == 7:
        postcode = f"〒{raw_post[:3]}-{raw_post[3:]}"
    elif raw_post:
        postcode = f"〒{raw_post}"
    else:
        postcode = ""

    prefecture = customer.get("place") or ""
    street = customer.get("address") or ""

    parts = [p for p in [postcode, prefecture + street] if p]
    return "　".join(parts)


def _format_phone(raw: str) -> str:
    """電話番号を日本の形式（090-XXXX-XXXX）に整形します"""
    if not raw:
        return ""
    # 数字のみ抽出
    digits = re.sub(r"\D", "", str(raw))
    # 国際番号 +81 → 先頭を 0 に変換
    if digits.startswith("81") and len(digits) >= 11:
        digits = "0" + digits[2:]
    # 桁数でフォーマット分岐
    if len(digits) == 11:
        # 携帯・IP電話: 090-XXXX-XXXX
        return f"{digits[0:3]}-{digits[3:7]}-{digits[7:11]}"
    elif len(digits) == 10:
        # 固定電話: 03-XXXX-XXXX 等
        return f"{digits[0:2]}-{digits[2:6]}-{digits[6:10]}"
    return raw


def _format_date(raw) -> str:
    """日付データ（Unixタイムスタンプ整数 or ISO文字列）を日本語表記に変換します"""
    if not raw:
        return ""
    # Unixタイムスタンプ（ミリ秒）の場合
    if isinstance(raw, (int, float)):
        try:
            dt = datetime.fromtimestamp(raw / 1000, tz=timezone.utc)
            return dt.strftime("%Y年%m月%d日")
        except Exception:
            return str(raw)
    # 文字列の場合
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(str(raw)[: len(fmt)], fmt)
            return dt.strftime("%Y年%m月%d日")
        except ValueError:
            continue
    return str(raw)
