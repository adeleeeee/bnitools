"""room_dates 的邊界測試（純標準庫 unittest，全部使用合成資料）。

直接跑：
    python3 tools/tests/test_room_dates.py
"""
import importlib.util
import sys
import unittest
from datetime import date, datetime
from pathlib import Path

_TOOL = Path(__file__).resolve().parent.parent / "room_dates.py"
_spec = importlib.util.spec_from_file_location("room_dates", _TOOL)
room_dates = importlib.util.module_from_spec(_spec)
# dataclass 需要能從 sys.modules 反查自己的模組，載入前必須先註冊
sys.modules["room_dates"] = room_dates
_spec.loader.exec_module(room_dates)


def booking(ts, d, who="001 王小明", group="示範組", projector=False):
    return {"timestamp": ts, "date": d, "who": who, "group": group, "projector": projector}


class ParseTimestampTest(unittest.TestCase):
    def test_cjk_morning(self):
        self.assertEqual(
            room_dates.parse_sheet_timestamp("2030/7/6 上午 11:08:53"),
            datetime(2030, 7, 6, 11, 8, 53),
        )

    def test_cjk_afternoon(self):
        self.assertEqual(
            room_dates.parse_sheet_timestamp("2030/5/17 下午 1:20:07"),
            datetime(2030, 5, 17, 13, 20, 7),
        )

    def test_hour_out_of_12h_range_rejected(self):
        """12 小時制不該接受 13 或 0，那是壞資料。"""
        for bad in ("2030/7/6 上午 13:00:00", "2030/7/6 下午 13:00:00", "2030/7/6 上午 0:30:00"):
            with self.assertRaises(ValueError, msg=bad):
                room_dates.parse_sheet_timestamp(bad)

    def test_noon_is_twelve(self):
        """下午 12:30 = 12:30，不是 24:30。"""
        self.assertEqual(
            room_dates.parse_sheet_timestamp("2030/1/2 下午 12:30:00"),
            datetime(2030, 1, 2, 12, 30, 0),
        )

    def test_midnight_is_zero(self):
        """上午 12:05 = 00:05。"""
        self.assertEqual(
            room_dates.parse_sheet_timestamp("2030/1/2 上午 12:05:00"),
            datetime(2030, 1, 2, 0, 5, 0),
        )

    def test_24h_fallback(self):
        self.assertEqual(
            room_dates.parse_sheet_timestamp("2030-01-02 13:05:00"),
            datetime(2030, 1, 2, 13, 5, 0),
        )

    def test_garbage_fails_fast(self):
        with self.assertRaises(ValueError):
            room_dates.parse_sheet_timestamp("上週三下午")


class ParseMdTest(unittest.TestCase):
    def test_zero_padded_equals_plain(self):
        self.assertEqual(room_dates.parse_md("9/3"), room_dates.parse_md("09/03"))

    def test_whitespace_tolerated(self):
        self.assertEqual(room_dates.parse_md(" 9 / 3 "), (9, 3))

    def test_out_of_range(self):
        with self.assertRaises(ValueError):
            room_dates.parse_md("13/1")


class ResolveBookingDateTest(unittest.TestCase):
    def test_same_year(self):
        ts = datetime(2030, 8, 3, 10, 0, 0)
        self.assertEqual(room_dates.resolve_booking_date("12/9", ts), date(2030, 12, 9))

    def test_crosses_year_end(self):
        """12 月填表約隔年 1 月。"""
        ts = datetime(2030, 12, 20, 9, 0, 0)
        self.assertEqual(room_dates.resolve_booking_date("1/7", ts), date(2031, 1, 7))

    def test_same_day_as_submission(self):
        ts = datetime(2030, 6, 5, 9, 0, 0)
        self.assertEqual(room_dates.resolve_booking_date("6/5", ts), date(2030, 6, 5))

    def test_zero_padded_same_result(self):
        ts = datetime(2030, 8, 3, 10, 0, 0)
        self.assertEqual(
            room_dates.resolve_booking_date("09/03", ts),
            room_dates.resolve_booking_date("9/3", ts),
        )

    def test_leap_day_skips_non_leap_years(self):
        """2/29 要跳過平年找到下一個閏年，不是直接報錯。"""
        ts = datetime(2027, 12, 1, 9, 0, 0)
        self.assertEqual(room_dates.resolve_booking_date("2/29", ts), date(2028, 2, 29))

    def test_never_valid_date_reports(self):
        """4/31 不存在於任何一年，必須明確報錯，不可被吞掉。"""
        with self.assertRaises(room_dates.AmbiguousDate):
            room_dates.resolve_booking_date("4/31", datetime(2030, 1, 1, 9, 0, 0))


class ResolveOptionDateTest(unittest.TestCase):
    def test_future_is_this_year(self):
        self.assertEqual(
            room_dates.resolve_option_date("9/30", date(2030, 9, 7)), date(2030, 9, 30)
        )

    def test_today_counts_as_available(self):
        self.assertEqual(
            room_dates.resolve_option_date("9/7", date(2030, 9, 7)), date(2030, 9, 7)
        )

    def test_any_past_date_is_ambiguous(self):
        """今年已過的選項一律報歧義：剛過去和過去很久都分不出是殘留還是明年。"""
        for md in ("9/1", "7/29", "1/8"):
            with self.assertRaises(room_dates.AmbiguousDate, msg=md):
                room_dates.resolve_option_date(md, date(2030, 9, 7))

    def test_invalid_calendar_date_reports(self):
        with self.assertRaises(room_dates.AmbiguousDate):
            room_dates.resolve_option_date("2/30", date(2030, 1, 1))

    def test_explicit_iso_date_accepted(self):
        """人工確認過年份後，用完整日期繞過歧義。"""
        self.assertEqual(
            room_dates.resolve_option_entry("2031-07-29", date(2030, 9, 7)),
            date(2031, 7, 29),
        )

    def test_iso_and_md_mix(self):
        self.assertEqual(
            room_dates.resolve_option_entry("9/30", date(2030, 9, 7)), date(2030, 9, 30)
        )


class ThisWednesdayTest(unittest.TestCase):
    def test_monday_gives_this_week(self):
        self.assertEqual(room_dates.this_wednesday(date(2030, 9, 9)), date(2030, 9, 11))

    def test_tuesday_boundary(self):
        """週二 → 隔天週三。"""
        self.assertEqual(room_dates.this_wednesday(date(2030, 9, 10)), date(2030, 9, 11))

    def test_wednesday_is_itself(self):
        self.assertEqual(room_dates.this_wednesday(date(2030, 9, 11)), date(2030, 9, 11))

    def test_thursday_rolls_to_next_week(self):
        self.assertEqual(room_dates.this_wednesday(date(2030, 9, 12)), date(2030, 9, 18))


class TodayInTimezoneTest(unittest.TestCase):
    def test_taipei_is_ahead_of_utc_before_midnight(self):
        """台北午夜前後：UTC 還在前一天時，台北已經是隔天。

        不直接呼叫 today_in（會依賴當下時間），改驗證同一瞬間在兩個時區的日期差異，
        這正是原本用 date.today() 會在雲端沙箱算錯「本週三」的原因。
        """
        from datetime import timezone
        from zoneinfo import ZoneInfo

        instant = datetime(2030, 9, 10, 16, 30, tzinfo=timezone.utc)
        self.assertEqual(instant.astimezone(ZoneInfo("UTC")).date(), date(2030, 9, 10))
        self.assertEqual(instant.astimezone(ZoneInfo("Asia/Taipei")).date(), date(2030, 9, 11))
        # 同一瞬間，UTC 算出的「本週三」比台北早一週
        self.assertEqual(room_dates.this_wednesday(date(2030, 9, 10)), date(2030, 9, 11))
        self.assertEqual(room_dates.this_wednesday(date(2030, 9, 11)), date(2030, 9, 11))


class LatestPerDateTest(unittest.TestCase):
    def _mk(self, ts, who):
        return room_dates.Booking(
            when=date(2030, 9, 11), timestamp=ts, who=who, group="示範組", projector=False
        )

    def test_picks_latest_regardless_of_input_order(self):
        early = self._mk(datetime(2030, 9, 1, 9, 0, 0), "001 王小明")
        late = self._mk(datetime(2030, 9, 1, 17, 30, 0), "002 陳大文")
        forward, dup_a = room_dates.latest_per_date([early, late])
        reverse, dup_b = room_dates.latest_per_date([late, early])
        self.assertEqual(forward[date(2030, 9, 11)].who, "002 陳大文")
        self.assertEqual(reverse[date(2030, 9, 11)].who, "002 陳大文")
        self.assertEqual(dup_a, dup_b, [date(2030, 9, 11)])

    def test_no_duplicate_reported_when_unique(self):
        _, dupes = room_dates.latest_per_date([self._mk(datetime(2030, 9, 1, 9, 0, 0), "001 王小明")])
        self.assertEqual(dupes, [])


class BuildSummaryTest(unittest.TestCase):
    def _payload(self, **over):
        base = {
            "bookings": [
                booking("2030/8/1 上午 10:00:00", "9/11", who="001 王小明", projector=True),
                booking("2030/8/1 上午 10:01:00", "9/18", who="002 陳大文"),
                booking("2030/8/1 上午 10:02:00", "9/25", who="003 李小華"),
                booking("2030/8/1 上午 10:03:00", "10/2", who="004 張三"),
                booking("2030/8/1 上午 10:04:00", "10/9", who="005 李四"),
                booking("2030/8/1 上午 10:05:00", "10/16", who="006 王五"),
                booking("2030/7/1 上午 09:00:00", "8/28", who="007 趙六"),  # 已過去
                booking("2030/8/1 上午 10:06:00", "", who="008 殘留列"),      # 沒填日期
            ],
            "available": ["10/23", "10/30"],
        }
        base.update(over)
        return base

    def _summary(self, payload=None, today=date(2030, 9, 9)):
        return room_dates.build_summary(
            payload or self._payload(), "Asia/Taipei", today=today
        )

    def test_this_week_and_later_limit(self):
        s = self._summary()
        self.assertEqual(s["this_wednesday"], "2030-09-11")
        self.assertEqual(s["this_week"]["who"], "001 王小明")
        self.assertTrue(s["this_week"]["projector"])
        self.assertEqual([r["display"] for r in s["later"]], ["9/18", "9/25", "10/2", "10/9"])

    def test_past_and_blank_rows_dropped(self):
        s = self._summary()
        shown = {r["display"] for r in s["later"]} | {s["this_week"]["display"]}
        self.assertNotIn("8/28", shown)
        self.assertNotIn("008 殘留列", [r["who"] for r in s["later"]])

    def test_no_booking_this_week(self):
        payload = self._payload(bookings=[booking("2030/8/1 上午 10:00:00", "9/18")])
        s = self._summary(payload)
        self.assertIsNone(s["this_week"])

    def test_clash_detected_across_zero_padding(self):
        """試算表寫 09/18、表單寫 9/18，必須偵測為衝突。"""
        payload = self._payload(
            bookings=[booking("2030/8/1 上午 10:00:00", "09/18")],
            available=["9/18"],
        )
        s = self._summary(payload)
        self.assertEqual(s["warnings"]["clashes"], ["9/18"])

    def test_duplicate_reported(self):
        payload = self._payload(
            bookings=[
                booking("2030/8/1 上午 10:00:00", "9/11", who="001 王小明"),
                booking("2030/8/2 下午 3:00:00", "9/11", who="002 陳大文"),
            ]
        )
        s = self._summary(payload)
        self.assertEqual(s["warnings"]["duplicated"], ["9/11"])
        self.assertEqual(s["this_week"]["who"], "002 陳大文")

    def test_missing_key_is_error_not_empty(self):
        payload = self._payload()
        del payload["available"]
        with self.assertRaises(ValueError) as ctx:
            self._summary(payload)
        self.assertIn("available", str(ctx.exception))

    def test_empty_available_is_allowed(self):
        s = self._summary(self._payload(available=[]))
        self.assertEqual(s["available"], [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
