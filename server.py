from __future__ import annotations

import csv
import io
import json
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


HOST = "127.0.0.1"
PORT = 8765
HOLIDAY_CSV_URL = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"


class ActivityLogHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, directory: str | None = None, **kwargs):
        root = directory or str(Path(__file__).resolve().parent)
        super().__init__(*args, directory=root, **kwargs)

    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/holidays":
            self.handle_holidays_api(parsed.query)
            return
        if parsed.path == "/":
            self.path = "/index.html"
        super().do_GET()

    def handle_holidays_api(self, query: str) -> None:
        try:
            years = parse_years(query)
            holidays = load_holidays(years)
            payload = {
                "holidays": holidays,
                "source": "cao",
                "fetchedAt": datetime.now().isoformat(timespec="minutes"),
            }
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except ValueError as error:
            self.send_json_error(HTTPStatus.BAD_REQUEST, str(error))
        except urllib.error.URLError as error:
            self.send_json_error(HTTPStatus.BAD_GATEWAY, f"祝日CSVの取得に失敗しました: {error.reason}")
        except Exception as error:  # pragma: no cover
            self.send_json_error(HTTPStatus.INTERNAL_SERVER_ERROR, f"予期しないエラー: {error}")

    def send_json_error(self, status: HTTPStatus, message: str) -> None:
        body = json.dumps({"error": message}, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def parse_years(query: str) -> list[int]:
    params = urllib.parse.parse_qs(query)
    raw_years = params.get("years", [""])[0]
    years = []
    for part in raw_years.split(","):
        token = part.strip()
        if not token:
            continue
        year = int(token)
        if year < 1948 or year > 2100:
            raise ValueError("years パラメータが不正です。")
        years.append(year)
    if not years:
        years = [datetime.now().year]
    return sorted(set(years))


def load_holidays(years: list[int]) -> list[dict[str, str]]:
    request = urllib.request.Request(
      HOLIDAY_CSV_URL,
      headers={
          "User-Agent": "activity-log-holiday-sync/1.0",
          "Accept": "text/csv,*/*",
      },
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        raw = response.read()

    text = decode_csv_bytes(raw)
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if len(rows) < 2:
        return []

    target_years = {str(year) for year in years}
    holidays = []
    for row in rows[1:]:
        if len(row) < 2:
            continue
        date_value = normalize_csv_date(row[0])
        if not date_value:
            continue
        if date_value[:4] not in target_years:
            continue
        holidays.append({
            "date": date_value,
            "name": row[1].strip(),
        })

    holidays.sort(key=lambda item: item["date"])
    return holidays


def decode_csv_bytes(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "cp932", "shift_jis", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("祝日CSVの文字コードを判定できませんでした。")


def normalize_csv_date(value: str) -> str:
    parts = value.strip().split("/")
    if len(parts) != 3:
        return ""
    try:
        year, month, day = [int(part) for part in parts]
    except ValueError:
        return ""
    return f"{year:04d}-{month:02d}-{day:02d}"


def main() -> None:
    port = PORT
    if len(sys.argv) >= 2:
        port = int(sys.argv[1])

    server = ThreadingHTTPServer((HOST, port), ActivityLogHandler)
    print(f"Activity Log server running on http://{HOST}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
