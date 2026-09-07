#!/usr/bin/env python3
"""私有設定的載入與驗證（純標準庫）。

設計原則：**缺欄位或還是佔位符就停止並指出欄位名**，不猜值、不退回硬編碼、
不以空值繼續。設定裡的分會名稱與 Google 檔案 ID 屬私有資料，只存在
未進版控的 config.json；公開 repo 只有 config.example.json 的佔位符版本。

兩個設定來源互斥：
  - 本機：repo 根目錄的 config.json
  - 雲端 routine：把 prompt 內的設定區塊寫成暫存 JSON，用 --config 指過來
同時偵測到兩者 → 停止並要求釐清，不自行挑一個。

直接檢查目前設定：
    python3 tools/bni_config.py --flow bni-local
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

SCHEMA_VERSION = 1

# 各流程的必要欄位。多餘欄位不得要求 —— 例如雲端 bni 走 WebFetch，
# 既不需要 room_form_id，也不需要任何投影片欄位。
REQUIRED_BY_FLOW = {
    "bni-local": [
        "schema_version", "chapter_short", "timezone",
        "room_form_id", "room_form_responder_url", "room_sheet_id",
    ],
    "bni-cloud": [
        "schema_version", "chapter_short", "timezone",
        "room_form_responder_url", "room_sheet_id",
    ],
    "bnippt": [
        "schema_version", "guest_slides_file_id",
    ],
}

REPO_ROOT = Path(__file__).resolve().parent.parent


class ConfigError(Exception):
    """設定缺漏或無效。訊息一律指出欄位名，讓使用者知道要補什麼。"""


def _is_placeholder(value: str) -> bool:
    """config.example.json 的佔位符長成 <說明> 這樣。"""
    stripped = value.strip()
    return stripped.startswith("<") and stripped.endswith(">")


def resolve_config_path(explicit: str | None, repo_root: Path = REPO_ROOT) -> Path:
    """決定要讀哪一份設定，並強制兩來源互斥。"""
    repo_config = repo_root / "config.json"
    if explicit:
        explicit_path = Path(explicit).expanduser().resolve()
        if repo_config.exists() and explicit_path != repo_config.resolve():
            raise ConfigError(
                "偵測到兩個設定來源同時存在，拒絕自行挑選：\n"
                f"  1. repo 設定 {repo_config}\n"
                f"  2. 指定設定 {explicit_path}\n"
                "本機流程請不要傳 --config；雲端流程請在沒有 repo config.json 的環境執行。"
            )
        if not explicit_path.exists():
            raise ConfigError(f"找不到指定的設定檔：{explicit_path}")
        return explicit_path
    if repo_config.exists():
        return repo_config
    raise ConfigError(
        f"找不到設定檔 {repo_config}。\n"
        "請複製 config.example.json 成 config.json 並填入自己的值"
        "（config.json 已被 .gitignore 排除，不會進版控）。"
    )


def validate(raw: dict, flow: str) -> dict:
    """依流程驗證必要欄位，回傳驗證過的設定。"""
    if flow not in REQUIRED_BY_FLOW:
        raise ConfigError(f"未知的流程 {flow!r}，可用：{sorted(REQUIRED_BY_FLOW)}")

    version = raw.get("schema_version")
    if version is None:
        raise ConfigError("設定缺少必要欄位：schema_version")
    if not isinstance(version, int) or isinstance(version, bool):
        raise ConfigError(f"schema_version 必須是整數，目前是 {type(version).__name__}")
    if version != SCHEMA_VERSION:
        raise ConfigError(
            f"不支援的 schema_version={version}（本工具支援 {SCHEMA_VERSION}）。"
            "請對照 config.example.json 更新設定。"
        )

    missing, placeholder, blank = [], [], []
    for field in REQUIRED_BY_FLOW[flow]:
        if field == "schema_version":
            continue
        if field not in raw:
            missing.append(field)
            continue
        value = raw[field]
        if not isinstance(value, str):
            raise ConfigError(f"欄位 {field} 必須是字串，目前是 {type(value).__name__}")
        if not value.strip():
            blank.append(field)
        elif _is_placeholder(value):
            placeholder.append(field)

    problems = []
    if missing:
        problems.append(f"缺少欄位：{', '.join(missing)}")
    if blank:
        problems.append(f"欄位為空白：{', '.join(blank)}")
    if placeholder:
        problems.append(f"欄位還是 config.example.json 的佔位符：{', '.join(placeholder)}")
    if problems:
        raise ConfigError(f"流程 {flow} 的設定不完整 —— " + "；".join(problems))

    if "timezone" in REQUIRED_BY_FLOW[flow]:
        tz = raw["timezone"]
        try:
            ZoneInfo(tz)
        except (ZoneInfoNotFoundError, ValueError) as exc:
            raise ConfigError(f"timezone {tz!r} 不是有效時區：{exc}") from exc

    return raw


def load(flow: str, explicit: str | None = None, repo_root: Path = REPO_ROOT) -> dict:
    path = resolve_config_path(explicit, repo_root)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"設定檔 {path} 不是合法 JSON：{exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigError(f"設定檔 {path} 的最外層必須是物件")
    return validate(raw, flow)


def main():
    p = argparse.ArgumentParser(description="檢查私有設定是否可供指定流程使用")
    p.add_argument("--flow", required=True, choices=sorted(REQUIRED_BY_FLOW))
    p.add_argument("--config", help="改讀指定的設定檔（雲端流程用）")
    args = p.parse_args()
    try:
        cfg = load(args.flow, args.config)
    except ConfigError as exc:
        sys.exit(f"設定錯誤：{exc}")
    # 只印欄位名，不回顯值（值含分會名稱與檔案 ID）
    print(f"✅ 流程 {args.flow} 的設定齊全，已驗證欄位："
          f"{', '.join(REQUIRED_BY_FLOW[args.flow])}")
    print(f"   schema_version={cfg['schema_version']}")


if __name__ == "__main__":
    main()
