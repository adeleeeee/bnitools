#!/usr/bin/env python3
"""小房間預約的日期計算（純標準庫）。

**skill 與測試共用同一份邏輯**：SKILL.md 不再內嵌一份手抄的計算程式碼，
一律呼叫本模組，避免測試測的是另一份、skill 實際跑的沒被覆蓋。

三個容易翻車的地方，全部在這裡處理掉：
  1. 時區：用設定的 timezone 取「今天」，不用 date.today()（雲端沙箱是 UTC，會差一天）
  2. 同日多筆：時間戳保留到秒，選真正最新的那筆，與輸入列的順序無關
  3. 日期比對：一律正規化成 date 物件再比，`9/3` 與 `09/03` 視為同一天

用法（給 skill 呼叫）：
    python3 tools/room_dates.py --flow bni-local --input payload.json

payload.json 形狀：
    {
      "bookings": [
        {"timestamp": "2026/7/6 上午 11:08:53", "date": "10/14",
         "who": "032 某某某", "group": "某某分享", "projector": true}
      ],
      "available": ["9/30", "11/18"]
    }

`available` 的元素可以是表單原樣的 `M/D`，也可以是 `YYYY-MM-DD`。
表單選項沒有時間戳可推年份，**今年已過的 M/D 一律回報歧義**（可能是過期殘留、
也可能是明年場次），人工確認後改填完整日期即可。

`bookings` 與 `available` **都必須存在**。抓取失敗時請中止流程，
不要傳空陣列 —— 空陣列代表「確實查到沒有」，不是「查不到」。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

WEDNESDAY = 2  # date.weekday()：週一=0
DEFAULT_LATER_LIMIT = 4

# 試算表時間戳：2026/7/6 上午 11:08:53
_TS_CJK_RE = re.compile(
    r"^(\d{4})/(\d{1,2})/(\d{1,2})\s+(上午|下午)\s*(\d{1,2}):(\d{2}):(\d{2})$"
)
# 24 小時制備援：2026/7/6 11:08:53 或 2026-07-06 11:08:53
_TS_24H_RE = re.compile(
    r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})[ T](\d{1,2}):(\d{2}):(\d{2})$"
)
_MD_RE = re.compile(r"^\s*(\d{1,2})\s*/\s*(\d{1,2})\s*$")
# 已人工確認年份的完整日期
_ISO_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


class AmbiguousDate(ValueError):
    """日期無法無歧義地決定年份，必須人工確認。"""


@dataclass(frozen=True)
class Booking:
    when: date
    timestamp: datetime
    who: str
    group: str
    projector: bool


def today_in(timezone: str) -> date:
    """用指定時區取今天。雲端沙箱是 UTC，絕不能用 date.today()。"""
    return datetime.now(ZoneInfo(timezone)).date()


def parse_sheet_timestamp(text: str) -> datetime:
    """解析試算表時間戳，保留到秒。無法解析一律 fail fast。"""
    raw = text.strip()
    m = _TS_CJK_RE.match(raw)
    if m:
        year, month, day, meridiem, hour, minute, second = m.groups()
        hour = int(hour)
        # 12 小時制先驗範圍再換算：上午 13:00 是壞資料，不能默默當成 13 時
        if not 1 <= hour <= 12:
            raise ValueError(f"12 小時制的時鐘數必須是 1–12，收到 {hour}：{text!r}")
        if meridiem == "下午" and hour != 12:
            hour += 12
        elif meridiem == "上午" and hour == 12:
            hour = 0
        return datetime(int(year), int(month), int(day), hour, int(minute), int(second))
    m = _TS_24H_RE.match(raw)
    if m:
        year, month, day, hour, minute, second = (int(x) for x in m.groups())
        return datetime(year, month, day, hour, minute, second)
    raise ValueError(f"無法解析時間戳：{text!r}")


def parse_md(text: str) -> tuple[int, int]:
    """`9/3`、`09/03`、` 9 / 3 ` 都解析成 (9, 3)。"""
    m = _MD_RE.match(text)
    if not m:
        raise ValueError(f"無法解析日期：{text!r}（預期 M/D）")
    month, day = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise ValueError(f"月份超出範圍：{text!r}")
    if not 1 <= day <= 31:
        raise ValueError(f"日期超出範圍：{text!r}")
    return month, day


def _try_date(year: int, month: int, day: int) -> date | None:
    """該年沒有這個日期（如平年 2/29）就回 None，讓呼叫端試下一個候選年。"""
    try:
        return date(year, month, day)
    except ValueError:
        return None


def resolve_booking_date(md: str, timestamp: datetime) -> date:
    """試算表的預約日期只有 M/D。

    年份＝讓該日期落在「填表當天或之後」且最接近填表日的那一年。
    候選年裡若某一年沒有這個日期（2/29），跳過該年繼續試下一年；
    每一年都不合法（4/31）才報錯。
    """
    month, day = parse_md(md)
    submitted = timestamp.date()
    # 2/29 最多要等 8 年才輪到閏年，候選年給足
    for year in range(submitted.year, submitted.year + 9):
        candidate = _try_date(year, month, day)
        if candidate is None:
            continue
        if candidate >= submitted:
            return candidate
    raise AmbiguousDate(
        f"預約日期 {md} 無法依填表時間 {submitted} 推定年份"
        f"（可能不是合法日期，例如 4/31）"
    )


def resolve_option_date(md: str, today: date) -> date:
    """表單選項的日期只有 M/D，而且**沒有時間戳可當基準**。

    只能以「今天」推年：落在今天或之後就是今年。
    **今年已經過去的一律回報歧義** —— 可能是表單還沒移除的過期選項，
    也可能是明年的場次，光靠「過了幾天」分不出來，不能默默推到下一年。

    要處理這種日期，請在 payload 的 `available` 直接給完整日期（`YYYY-MM-DD`），
    代表已經人工確認過年份。
    """
    month, day = parse_md(md)
    candidate = _try_date(today.year, month, day)
    if candidate is None:
        raise AmbiguousDate(f"可預約日期 {md} 在 {today.year} 年不是合法日期")
    if candidate >= today:
        return candidate
    raise AmbiguousDate(
        f"可預約日期 {md} 在今年（{candidate.isoformat()}）已經過去："
        f"可能是表單未移除的過期選項，也可能是明年的場次，無法自行判斷。"
        f"請人工確認後，在 available 改填完整日期（YYYY-MM-DD）再跑一次。"
    )


def resolve_option_entry(value: str, today: date) -> date:
    """表單選項可以是 `M/D`（依今天推年）或已人工確認的 `YYYY-MM-DD`。"""
    text = value.strip()
    if _ISO_RE.match(text):
        return date.fromisoformat(text)
    return resolve_option_date(text, today)


def this_wednesday(today: date) -> date:
    """今天起（含今天）最近的星期三。"""
    return today + timedelta(days=(WEDNESDAY - today.weekday()) % 7)


def latest_per_date(bookings: list[Booking]) -> tuple[dict[date, Booking], list[date]]:
    """同一天多筆時保留時間戳最新的那筆，與輸入順序無關。"""
    chosen: dict[date, Booking] = {}
    duplicated: set[date] = set()
    for booking in bookings:
        existing = chosen.get(booking.when)
        if existing is None:
            chosen[booking.when] = booking
            continue
        duplicated.add(booking.when)
        if booking.timestamp > existing.timestamp:
            chosen[booking.when] = booking
    return chosen, sorted(duplicated)


def fmt(day: date) -> str:
    """訊息用的顯示格式，不補零。"""
    return f"{day.month}/{day.day}"


def build_summary(
    payload: dict,
    timezone: str,
    today: date | None = None,
    later_limit: int = DEFAULT_LATER_LIMIT,
) -> dict:
    """把原始資料整理成訊息需要的結構。"""
    for key in ("bookings", "available"):
        if key not in payload:
            raise ValueError(
                f"輸入缺少 {key!r}。抓取失敗請中止流程，不要用空陣列頂替 —— "
                "空陣列代表「確實查到沒有」，不是「查不到」。"
            )

    if today is None:
        today = today_in(timezone)

    parsed: list[Booking] = []
    for row in payload["bookings"]:
        raw_date = (row.get("date") or "").strip()
        if not raw_date:
            continue  # 試算表末尾常有沒填日期的殘留列
        timestamp = parse_sheet_timestamp(row["timestamp"])
        when = resolve_booking_date(raw_date, timestamp)
        if when < today:
            continue
        parsed.append(
            Booking(
                when=when,
                timestamp=timestamp,
                who=row["who"].strip(),
                group=row["group"].strip(),
                projector=bool(row["projector"]),
            )
        )

    chosen, duplicated = latest_per_date(parsed)
    wednesday = this_wednesday(today)
    upcoming = sorted(chosen)

    available = [resolve_option_entry(x, today) for x in payload["available"]]
    clashes = sorted(set(available) & set(chosen))

    def as_row(booking: Booking) -> dict:
        return {
            "date": booking.when.isoformat(),
            "display": fmt(booking.when),
            "who": booking.who,
            "group": booking.group,
            "projector": booking.projector,
        }

    this_week = chosen.get(wednesday)
    return {
        "today": today.isoformat(),
        "this_wednesday": wednesday.isoformat(),
        "this_week": as_row(this_week) if this_week else None,
        "later": [as_row(chosen[d]) for d in upcoming if d > wednesday][:later_limit],
        "available": [{"date": d.isoformat(), "display": fmt(d)} for d in sorted(available)],
        "warnings": {
            "clashes": [fmt(d) for d in clashes],
            "duplicated": [fmt(d) for d in duplicated],
        },
    }


def main():
    p = argparse.ArgumentParser(description="小房間預約的日期計算")
    p.add_argument("--flow", required=True, choices=["bni-local", "bni-cloud"])
    p.add_argument("--config", help="改讀指定的設定檔（雲端流程用）")
    p.add_argument("--input", required=True, help="payload JSON 檔路徑")
    p.add_argument("--later-limit", type=int, default=DEFAULT_LATER_LIMIT)
    args = p.parse_args()

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import bni_config

    try:
        cfg = bni_config.load(args.flow, args.config)
    except bni_config.ConfigError as exc:
        sys.exit(f"設定錯誤：{exc}")

    payload = json.loads(Path(args.input).read_text(encoding="utf-8"))
    try:
        summary = build_summary(payload, cfg["timezone"], later_limit=args.later_limit)
    except (AmbiguousDate, ValueError) as exc:
        sys.exit(f"日期處理失敗：{exc}")

    summary["chapter_short"] = cfg["chapter_short"]
    summary["form_url"] = cfg["room_form_responder_url"]
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
