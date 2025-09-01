"""Data models and storage helpers."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Dict, List

from .utils import DATA_DIR, BASE_SAVE_PATH


@dataclass
class MonthData:
    """Serialized representation of monthly working data."""

    year: int
    month: int
    days: Dict[int, List[Dict[str, str]]] = field(default_factory=dict)

    @property
    def path(self) -> str:
        os.makedirs(DATA_DIR, exist_ok=True)
        return os.path.join(DATA_DIR, f"{self.year:04d}-{self.month:02d}.json")

    def save(self) -> None:
        days: Dict[str, List[Dict[str, str]]] = {}
        for day, rows in self.days.items():
            row_list: List[Dict[str, str]] = []
            for r in rows:
                row_list.append(
                    {
                        "work": r.get("work", ""),
                        "plan": r.get("plan", ""),
                        "done": r.get("done", ""),
                    }
                )
            if row_list:
                days[str(day)] = row_list
        data = {"year": self.year, "month": self.month, "days": days}
        with open(self.path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, year: int, month: int) -> "MonthData":
        path = os.path.join(DATA_DIR, f"{year:04d}-{month:02d}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            days: Dict[int, List[Dict[str, str]]] = {}
            for k, v in data.get("days", {}).items():
                row_list: List[Dict[str, str]] = []
                for row in v:
                    if isinstance(row, dict):
                        row_list.append(
                            {
                                "work": row.get("work", ""),
                                "plan": row.get("plan", ""),
                                "done": row.get("done", ""),
                            }
                        )
                    elif isinstance(row, list):
                        row_list.append(
                            {
                                "work": row[0] if len(row) > 0 else "",
                                "plan": row[1] if len(row) > 1 else "",
                                "done": row[2] if len(row) > 2 else "",
                            }
                        )
                days[int(k)] = row_list
            return cls(year=data.get("year", year), month=data.get("month", month), days=days)
        return cls(year=year, month=month)


def ensure_year_dirs(year: int) -> str:
    """Ensure required directory layout exists for *year*."""

    base = os.path.join(BASE_SAVE_PATH, str(year))
    for sub in ("stats", "release", "top", "year"):
        os.makedirs(os.path.join(base, sub), exist_ok=True)
    return base


def stats_dir(year: int) -> str:
    return os.path.join(ensure_year_dirs(year), "stats")


def release_dir(year: int) -> str:
    return os.path.join(ensure_year_dirs(year), "release")


def top_dir(year: int) -> str:
    return os.path.join(ensure_year_dirs(year), "top")


def year_dir(year: int) -> str:
    return os.path.join(ensure_year_dirs(year), "year")

