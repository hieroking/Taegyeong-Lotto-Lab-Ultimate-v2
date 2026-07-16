"""
太炅 Lotto Lab Ultimate v0.2

기능
- 역대 로또 Excel 불러오기
- 번호 빈도 / 페어 / 트리플 분석
- 사진 파일 목록 등록
- 번호 직접 입력 및 출현횟수 집계
- 역대 1등·2등 동일 조합 제외
- 조건 기반 추천조합 생성
- 직접 만든 조합 검사
- 추천 결과 Excel 저장
"""

from __future__ import annotations

import math
import re
import sys
import json
import base64
import os
import subprocess
import traceback
import tempfile
import urllib.request
import urllib.parse
import shutil
from datetime import datetime
from collections import Counter, defaultdict
from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from typing import Iterable

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QFont, QColor, QBrush, QImage
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSpinBox, QDialog,
    QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
)

APP_NAME = "太炅 Lotto Lab Ultimate"
VERSION = "7.0.0"



WINDOWS_OCR_PS = '$ErrorActionPreference = "Stop"\n[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n\nfunction Await($AsyncOperation, [Type]$ResultType) {\n    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() |\n        Where-Object {\n            $_.Name -eq "AsTask" -and\n            $_.IsGenericMethod -and\n            $_.GetParameters().Count -eq 1\n        }\n    $method = $methods | Select-Object -First 1\n    if ($null -eq $method) {\n        throw "Windows Runtime AsTask 메서드를 찾지 못했습니다."\n    }\n    $generic = $method.MakeGenericMethod($ResultType)\n    $task = $generic.Invoke($null, @($AsyncOperation))\n    $task.Wait()\n    return $task.Result\n}\n\ntry {\n    Add-Type -AssemblyName System.Runtime.WindowsRuntime\n\n    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]\n    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]\n    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime]\n\n    $imagePath = $env:LOTTO_OCR_IMAGE\n    if ([string]::IsNullOrWhiteSpace($imagePath)) {\n        throw "사진 경로가 전달되지 않았습니다."\n    }\n    if (!(Test-Path -LiteralPath $imagePath)) {\n        throw "사진 파일을 찾을 수 없습니다: $imagePath"\n    }\n\n    $fullPath = [System.IO.Path]::GetFullPath($imagePath)\n    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)) ([Windows.Storage.StorageFile])\n    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])\n    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])\n    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])\n\n    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()\n    if ($null -eq $engine) {\n        throw "Windows OCR 엔진을 만들 수 없습니다. Windows 설정에서 한국어 OCR 언어 기능을 설치하세요."\n    }\n\n    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])\n    $text = $result.Text\n\n    # 휴대폰 캡처에서 자주 섞이는 시간/날짜/페이지 표시 제거\n    $clean = [regex]::Replace($text, \'\\b\\d{1,2}:\\d{2}\\b\', \' \')\n    $clean = [regex]::Replace($clean, \'\\b\\d{4}[./-]\\d{1,2}[./-]\\d{1,2}\\b\', \' \')\n    $clean = [regex]::Replace($clean, \'\\b\\d+\\s*/\\s*\\d+\\b\', \' \')\n\n    $numbers = @()\n    foreach ($m in [regex]::Matches($clean, \'(?<!\\d)\\d{1,2}(?!\\d)\')) {\n        $n = [int]$m.Value\n        if ($n -ge 1 -and $n -le 45) {\n            $numbers += $n\n        }\n    }\n\n    @{ ok = $true; text = $text; numbers = $numbers } |\n        ConvertTo-Json -Compress -Depth 4\n    exit 0\n}\ncatch {\n    @{ ok = $false; error = $_.Exception.Message; numbers = @() } |\n        ConvertTo-Json -Compress -Depth 4\n    exit 1\n}'


@dataclass(frozen=True)
class Draw:
    round_no: int
    numbers: tuple[int, int, int, int, int, int]
    bonus: int | None


def parse_numbers(text: str) -> list[int]:
    values = [int(x) for x in re.findall(r"\d+", text)]
    invalid = [x for x in values if not 1 <= x <= 45]
    if invalid:
        raise ValueError(f"1~45 범위를 벗어난 번호: {invalid}")
    return values


class LottoAnalyzer:
    def __init__(self) -> None:
        self.draws: list[Draw] = []
        self.number_counts: Counter[int] = Counter()
        self.pair_counts: Counter[tuple[int, int]] = Counter()
        self.triple_counts: Counter[tuple[int, int, int]] = Counter()
        self.recent_number_counts: Counter[int] = Counter()
        self.recent_pair_counts: Counter[tuple[int, int]] = Counter()
        self.recent_triple_counts: Counter[tuple[int, int, int]] = Counter()
        self.first_prize: set[tuple[int, ...]] = set()
        self.second_prize: set[tuple[int, ...]] = set()

    def load_excel(self, path: str | Path) -> None:
        xls = pd.ExcelFile(path)
        best: list[Draw] = []

        for sheet in xls.sheet_names:
            try:
                raw = pd.read_excel(path, sheet_name=sheet, header=None)
            except Exception:
                continue
            draws = self._parse_sheet(raw)
            if len(draws) > len(best):
                best = draws

        if not best:
            raise ValueError(
                "회차와 당첨번호 6개를 찾지 못했습니다. "
                "첫 행에 회차·당첨번호·보너스가 있는 파일을 사용하세요."
            )

        self.draws = sorted(best, key=lambda d: d.round_no)
        self._analyze()

    @staticmethod
    def _parse_sheet(df: pd.DataFrame) -> list[Draw]:
        """엑셀 시트에서 회차·당첨번호 6개·보너스를 안전하게 추출합니다."""
        if df.empty or df.shape[1] < 7:
            return []

        header_row = None
        for i in range(min(30, len(df))):
            texts = [str(v).strip().lower() for v in df.iloc[i].tolist()]
            if any("회차" in x or x == "round" for x in texts):
                header_row = i
                break

        if header_row is None:
            return []

        headers = [
            "" if pd.isna(v) else str(v).strip()
            for v in df.iloc[header_row].tolist()
        ]
        body = df.iloc[header_row + 1:].reset_index(drop=True)

        def find_index(predicate) -> int | None:
            for idx, header in enumerate(headers):
                if predicate(header):
                    return idx
            return None

        round_idx = find_index(
            lambda h: "회차" in h.lower() or h.lower() == "round"
        )
        bonus_idx = find_index(
            lambda h: "보너스" in h.lower() or "bonus" in h.lower()
        )

        if round_idx is None:
            return []

        order_words = ("첫번째", "두번째", "세번째", "네번째", "다섯번째", "여섯번째")
        number_indices: list[int] = []
        for word in order_words:
            idx = find_index(lambda h, w=word: w in h)
            if idx is not None:
                number_indices.append(idx)

        if len(number_indices) < 6:
            number_indices = []
            for n in range(1, 7):
                idx = find_index(
                    lambda h, n=n: bool(
                        re.search(rf"(번호|num|ball)\s*{n}$", h, re.I)
                    )
                )
                if idx is not None:
                    number_indices.append(idx)

        if len(number_indices) < 6:
            candidates: list[int] = []
            for idx in range(df.shape[1]):
                if idx in (round_idx, bonus_idx):
                    continue
                series = pd.to_numeric(body.iloc[:, idx], errors="coerce").dropna()
                if len(series) >= 10 and float(series.between(1, 45).mean()) >= 0.85:
                    candidates.append(idx)
            number_indices = candidates[:6]

        if len(number_indices) < 6:
            return []

        draws: list[Draw] = []
        for _, row in body.iterrows():
            try:
                round_no = int(float(row.iloc[round_idx]))
                nums = tuple(
                    sorted(int(float(row.iloc[idx])) for idx in number_indices[:6])
                )
            except (ValueError, TypeError, IndexError):
                continue

            if len(set(nums)) != 6 or not all(1 <= x <= 45 for x in nums):
                continue

            bonus = None
            if bonus_idx is not None:
                try:
                    bonus_value = row.iloc[bonus_idx]
                    if pd.notna(bonus_value):
                        candidate = int(float(bonus_value))
                        if 1 <= candidate <= 45:
                            bonus = candidate
                except (ValueError, TypeError, IndexError):
                    bonus = None

            draws.append(Draw(round_no, nums, bonus))

        return draws

    def _analyze(self) -> None:
        self.number_counts.clear()
        self.pair_counts.clear()
        self.triple_counts.clear()
        self.recent_number_counts.clear()
        self.recent_pair_counts.clear()
        self.recent_triple_counts.clear()
        self.first_prize.clear()
        self.second_prize.clear()

        for draw in self.draws:
            self.number_counts.update(draw.numbers)
            self.pair_counts.update(combinations(draw.numbers, 2))
            self.triple_counts.update(combinations(draw.numbers, 3))
            self.first_prize.add(draw.numbers)

            # 2등 조합 = 본번호 5개 + 보너스번호
            if draw.bonus is not None:
                for five in combinations(draw.numbers, 5):
                    self.second_prize.add(tuple(sorted((*five, draw.bonus))))

        # 최근패턴은 최신 100회를 기준으로 계산
        for draw in self.draws[-100:]:
            self.recent_number_counts.update(draw.numbers)
            self.recent_pair_counts.update(combinations(draw.numbers, 2))
            self.recent_triple_counts.update(combinations(draw.numbers, 3))

    def check_combo(self, combo: tuple[int, ...]) -> dict:
        combo = tuple(sorted(combo))
        same_first = combo in self.first_prize
        same_second = combo in self.second_prize
        matches: list[tuple[int, int]] = []
        s = set(combo)
        for draw in self.draws:
            count = len(s.intersection(draw.numbers))
            if count >= 4:
                matches.append((draw.round_no, count))
        matches.sort(key=lambda x: (-x[1], -x[0]))
        return {"first": same_first, "second": same_second, "matches": matches}



class TKPerformanceEngine:
    """1~1218회 워크포워드 검증으로 선택한 성과 중심 번호·조합 엔진."""

    FEATURE_NAMES = ['최근10회', '최근30회', '최근100회', '최근300회', '전체빈도', '미출현간격', '직전이월수', '2회전재등장', '끝수흐름', '직전번호동반수', '인접연속수', '간격수흐름']
    OPTIMIZED_WEIGHTS = [0.00378073, 0.116455048, 0.104777671, 0.073486328, 0.005267944, 0.221622601, 0.030370023, 0.076836631, 0.257287055, 0.041604862, 0.022926405, 0.045584697]
    OPTIMIZATION_RESULT = {'tested_settings': 30000, 'feature_names': ['최근10회', '최근30회', '최근100회', '최근300회', '전체빈도', '미출현간격', '직전이월수', '2회전재등장', '끝수흐름', '직전번호동반수', '인접연속수', '간격수흐름'], 'weights': {'최근10회': 0.003781, '최근30회': 0.116455, '최근100회': 0.104778, '최근300회': 0.073486, '전체빈도': 0.005268, '미출현간격': 0.221623, '직전이월수': 0.03037, '2회전재등장': 0.076837, '끝수흐름': 0.257287, '직전번호동반수': 0.041605, '인접연속수': 0.022926, '간격수흐름': 0.045585}, 'train': {'average_top15_hits': 2.088, 'three_plus_rate': 0.316, 'four_plus_rate': 0.088, 'max_hits': 5}, 'validation': {'average_top15_hits': 2.0797, 'three_plus_rate': 0.3116, 'four_plus_rate': 0.1014, 'max_hits': 6}, 'holdout': {'round_start': 1081, 'round_end': 1218, 'average_top15_hits': 2.1232, 'three_plus_rate': 0.3188, 'four_plus_rate': 0.1449, 'max_hits': 5, 'random_expected_hits': 2.0}, 'elapsed_seconds': 24.08}

    @staticmethod
    def normalize(values):
        values = list(map(float, values))
        low, high = min(values), max(values)
        if high <= low:
            return [0.0] * len(values)
        return [(value - low) / (high - low) for value in values]

    @classmethod
    def number_scores(cls, draws):
        if len(draws) < 30:
            raise ValueError("성과최적추천은 최소 30회 이상의 데이터가 필요합니다.")
        history = [tuple(draw.numbers) for draw in draws]
        flat = lambda rows: [n for row in rows for n in row]
        feature_columns = []

        for window in (10, 30, 100, 300):
            counts = Counter(flat(history[-window:]))
            feature_columns.append(cls.normalize([counts[n] for n in range(1, 46)]))

        all_counts = Counter(flat(history))
        feature_columns.append(cls.normalize([all_counts[n] for n in range(1, 46)]))

        last_seen = {n: -1 for n in range(1, 46)}
        for index, row in enumerate(history):
            for number in row:
                last_seen[number] = index
        gaps = [len(history) - 1 - last_seen[n] for n in range(1, 46)]
        feature_columns.append(cls.normalize(gaps))

        last = set(history[-1])
        previous = set(history[-2])
        feature_columns.append([1.0 if n in last else 0.0 for n in range(1, 46)])
        feature_columns.append([
            1.0 if n in previous and n not in last else 0.0
            for n in range(1, 46)
        ])

        ending_counts = Counter(n % 10 for n in flat(history[-30:]))
        feature_columns.append(
            cls.normalize([ending_counts[n % 10] for n in range(1, 46)])
        )

        partner = Counter()
        for row in history[-100:]:
            row_set = set(row)
            overlap = len(row_set & last)
            if overlap:
                for number in row_set - last:
                    partner[number] += overlap
        feature_columns.append(cls.normalize([partner[n] for n in range(1, 46)]))

        adjacent = []
        for n in range(1, 46):
            adjacent.append(
                1.0 if n not in last and any(abs(n - x) == 1 for x in last) else 0.0
            )
        feature_columns.append(adjacent)

        gap_counts = Counter()
        for row in history[-30:]:
            for a, b in combinations(sorted(row), 2):
                if 1 <= b - a <= 15:
                    gap_counts[b - a] += 1
        common_gaps = [gap for gap, _ in gap_counts.most_common(5)]
        interval = [0.0] * 45
        for rank, gap in enumerate(common_gaps):
            value = 1.0 - rank * 0.15
            for source in last:
                for candidate in (source - gap, source + gap):
                    if 1 <= candidate <= 45 and candidate not in last:
                        interval[candidate - 1] = max(interval[candidate - 1], value)
        feature_columns.append(interval)

        scores = {}
        details = {}
        for number in range(1, 46):
            contributions = []
            for name, weight, column in zip(
                cls.FEATURE_NAMES, cls.OPTIMIZED_WEIGHTS, feature_columns
            ):
                value = column[number - 1]
                contributions.append((name, value * weight))
            scores[number] = sum(value for _, value in contributions) * 100.0
            details[number] = sorted(
                contributions, key=lambda item: (-item[1], item[0])
            )
        return scores, details

    @classmethod
    def generate(
        cls,
        analyzer,
        count=100,
        fixed_numbers=(),
        excluded_numbers=(),
        candidate_numbers=(),
    ):
        scores, details = cls.number_scores(analyzer.draws)
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        ranked = [
            n for n, _ in sorted(scores.items(), key=lambda item: (-item[1], item[0]))
            if n not in excluded_set
        ]
        pool = []
        for n in list(fixed_set | candidate_set) + ranked:
            if n not in pool and n not in excluded_set:
                pool.append(n)
            if len(pool) >= 20:
                break
        pool = sorted(pool)

        # 번호 선택과 조합 배치를 분리:
        # 후보 TOP20 안에서 조합을 만들고 구조·분산·동반출현을 별도 평가합니다.
        raw = []
        recent_pair = analyzer.recent_pair_counts
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in analyzer.first_prize or combo in analyzer.second_prize:
                continue
            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            total = sum(combo)
            zones = [
                sum(lo <= n <= hi for n in combo)
                for lo, hi in ((1, 10), (11, 20), (21, 30), (31, 40), (41, 45))
            ]
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not 95 <= total <= 185:
                continue
            if max(zones) > 3:
                continue

            number_score = sum(scores[n] for n in combo) / 6.0
            pair_score = sum(
                recent_pair[tuple(sorted(pair))]
                for pair in combinations(combo, 2)
            ) / 15.0
            candidate_bonus = len(combo_set & candidate_set) * 2.5
            consecutive = sum(b - a == 1 for a, b in zip(combo, combo[1:]))
            structure = 100.0
            structure -= abs(odd - 3) * 5.0
            structure -= abs(high - 3) * 4.0
            structure -= abs(total - 140) * 0.12
            structure -= max(0, consecutive - 1) * 8.0
            final_score = number_score * 0.72 + min(100.0, pair_score * 10) * 0.13 + structure * 0.15 + candidate_bonus

            top_reasons = []
            for number in combo:
                strongest = details[number][:2]
                top_reasons.append(
                    f"{number}번: " + ", ".join(name for name, _ in strongest)
                )

            metrics = {
                "performance": final_score,
                "composite": final_score,
                "input": 0.0,
                "pair": min(100.0, pair_score * 10),
                "triple": 0.0,
                "recent": number_score,
                "structure": structure,
                "pattern_votes": 0,
                "strategy": "성과최적엔진",
                "performance_reasons": top_reasons,
                "candidate_hits": len(combo_set & candidate_set),
                "candidate_bonus": candidate_bonus,
                "filter_mode": "성과최적화",
            }
            raw.append((final_score, combo, metrics))

        raw.sort(key=lambda row: (-row[0], row[1]))

        # 지나치게 비슷한 조합을 줄여 실제 추천 묶음의 포착 범위를 확대합니다.
        selected = []
        number_usage = Counter()
        for score, combo, metrics in raw:
            overlap5 = any(len(set(combo) & set(old_combo)) >= 5 for _, old_combo, _ in selected)
            usage_penalty = sum(number_usage[n] for n in combo) * 0.35
            adjusted = score - usage_penalty
            if overlap5 and len(selected) >= 10:
                continue
            selected.append((adjusted, combo, metrics))
            number_usage.update(combo)
            if len(selected) >= count:
                break
        selected.sort(key=lambda row: (-row[0], row[1]))
        return selected

class Recommender:
    """입력빈도·동반수·트리플·최근패턴을 자동 종합해 순위를 계산합니다."""

    CATEGORY_NAMES = {
        "추천조합": "composite",
        "나온횟수": "input",
        "동반수": "pair",
        "트리플": "triple",
        "최근패턴": "recent",
        "통합데이터추천": "mixed",
        "성과최적추천": "performance",
        "특이패턴추천": "pattern",
        "자체추천": "self",
    }

    def __init__(self, analyzer: LottoAnalyzer) -> None:
        self.a = analyzer
        self.max_pair = max(self.a.pair_counts.values(), default=1)
        self.max_triple = max(self.a.triple_counts.values(), default=1)
        self.max_recent_number = max(self.a.recent_number_counts.values(), default=1)
        self.max_recent_pair = max(self.a.recent_pair_counts.values(), default=1)

    @staticmethod
    def consecutive_pairs(combo: tuple[int, ...]) -> int:
        return sum(1 for a, b in zip(combo, combo[1:]) if b - a == 1)

    def pair_details(self, combo: tuple[int, ...], top_n: int = 3):
        details = [
            (pair, self.a.pair_counts[pair])
            for pair in combinations(combo, 2)
        ]
        details.sort(key=lambda item: (-item[1], item[0]))
        return details[:top_n]

    def triple_details(self, combo: tuple[int, ...], top_n: int = 3):
        details = [
            (triple, self.a.triple_counts[triple])
            for triple in combinations(combo, 3)
        ]
        details.sort(key=lambda item: (-item[1], item[0]))
        return details[:top_n]

    @staticmethod
    def confidence_score(score: float, metrics: dict[str, float]) -> float:
        """0~100 추천 신뢰도. 통계점수와 조합 균형을 함께 반영합니다."""
        base = min(100.0, max(0.0, score))
        stability = (
            metrics.get("pair", 0.0) * 0.25
            + metrics.get("triple", 0.0) * 0.20
            + metrics.get("recent", 0.0) * 0.20
            + metrics.get("structure", 0.0) * 0.20
            + metrics.get("input", 0.0) * 0.15
        )
        return round(min(100.0, base * 0.65 + stability * 0.35), 1)

    @staticmethod
    def confidence_grade(confidence: float) -> str:
        if confidence >= 95:
            return "S"
        if confidence >= 90:
            return "A"
        if confidence >= 80:
            return "B"
        if confidence >= 70:
            return "C"
        return "D"

    @staticmethod
    def select_diverse(
        candidates: list[tuple[float, tuple[int, ...], dict[str, float]]],
        count: int,
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """서로 너무 비슷한 조합을 줄여 결과 다양성을 높입니다."""
        selected = []
        selected_sets = []

        # 1차: 기존 조합과 5개 이상 겹치는 후보는 제외
        for row in candidates:
            combo_set = set(row[1])
            if all(len(combo_set & prev) <= 4 for prev in selected_sets):
                selected.append(row)
                selected_sets.append(combo_set)
                if len(selected) >= count:
                    return selected

        # 2차: 부족하면 점수순으로 남은 조합을 보충
        existing = {row[1] for row in selected}
        for row in candidates:
            if row[1] in existing:
                continue
            selected.append(row)
            existing.add(row[1])
            if len(selected) >= count:
                break
        return selected

    @staticmethod
    def _normalize(value: float, maximum: float) -> float:
        if maximum <= 0:
            return 0.0
        return max(0.0, min(100.0, value / maximum * 100.0))

    def metrics(
        self,
        combo: tuple[int, ...],
        source_weights: Counter[int],
    ) -> dict[str, float]:
        max_input = max(source_weights.values(), default=1)
        input_raw = sum(source_weights[n] for n in combo)
        input_score = self._normalize(input_raw, max_input * 6)

        pair_values = [self.a.pair_counts[p] for p in combinations(combo, 2)]
        pair_score = self._normalize(
            sum(sorted(pair_values, reverse=True)[:5]),
            self.max_pair * 5,
        )

        triple_values = [self.a.triple_counts[t] for t in combinations(combo, 3)]
        triple_score = self._normalize(
            sum(sorted(triple_values, reverse=True)[:5]),
            self.max_triple * 5,
        )

        recent_number = sum(self.a.recent_number_counts[n] for n in combo)
        recent_pair = sum(self.a.recent_pair_counts[p] for p in combinations(combo, 2))
        recent_score = (
            self._normalize(recent_number, self.max_recent_number * 6) * 0.55
            + self._normalize(recent_pair, self.max_recent_pair * 15) * 0.45
        )

        odd = sum(n % 2 for n in combo)
        high = sum(n >= 23 for n in combo)
        total = sum(combo)
        structure = 100.0
        structure -= abs(odd - 3) * 12
        structure -= abs(high - 3) * 10
        if total < 100:
            structure -= (100 - total) * 0.8
        elif total > 180:
            structure -= (total - 180) * 0.8
        structure -= max(0, self.consecutive_pairs(combo) - 1) * 15
        structure = max(0.0, min(100.0, structure))

        # 자동 종합 기준: 사용자가 따로 가중치를 조절하지 않아도 됨
        composite = (
            input_score * 0.30
            + pair_score * 0.25
            + triple_score * 0.20
            + recent_score * 0.15
            + structure * 0.10
        )

        return {
            "input": input_score,
            "pair": pair_score,
            "triple": triple_score,
            "recent": recent_score,
            "structure": structure,
            "composite": composite,
        }

    STRATEGY_WEIGHTS = {
        "균형형": {
            "input": 0.25, "pair": 0.20, "triple": 0.15,
            "recent": 0.15, "structure": 0.25,
        },
        "출현형": {
            "input": 0.50, "pair": 0.15, "triple": 0.10,
            "recent": 0.15, "structure": 0.10,
        },
        "동반수형": {
            "input": 0.15, "pair": 0.50, "triple": 0.15,
            "recent": 0.10, "structure": 0.10,
        },
        "트리플형": {
            "input": 0.10, "pair": 0.20, "triple": 0.50,
            "recent": 0.10, "structure": 0.10,
        },
        "최근형": {
            "input": 0.15, "pair": 0.15, "triple": 0.10,
            "recent": 0.50, "structure": 0.10,
        },
        "AI Ultimate": {
            "input": 0.20, "pair": 0.20, "triple": 0.15,
            "recent": 0.20, "structure": 0.25,
        },
    }

    def strategy_score(
        self,
        metrics: dict[str, float],
        strategy: str,
    ) -> float:
        weights = self.STRATEGY_WEIGHTS.get(
            strategy, self.STRATEGY_WEIGHTS["균형형"]
        )
        return sum(metrics.get(key, 0.0) * weight for key, weight in weights.items())

    MIXED_PRESETS = {
        "입력중심형": (0.50, 0.20, 0.20, 0.10),
        "최근중심형": (0.25, 0.40, 0.25, 0.10),
        "균형형": (0.30, 0.25, 0.25, 0.20),
        "장기형": (0.20, 0.15, 0.25, 0.40),
    }

    @staticmethod
    def _window_analyzer(draws: list[Draw], size: int) -> LottoAnalyzer:
        analyzer = LottoAnalyzer()
        analyzer.draws = list(draws[-min(size, len(draws)):])
        analyzer._analyze()
        return analyzer

    @staticmethod
    def _historical_combo_score(
        combo: tuple[int, ...],
        analyzer: LottoAnalyzer,
    ) -> float:
        max_number = max(analyzer.number_counts.values(), default=1)
        max_pair = max(analyzer.pair_counts.values(), default=1)
        max_triple = max(analyzer.triple_counts.values(), default=1)

        number_score = (
            sum(analyzer.number_counts[n] for n in combo)
            / max(1, max_number * 6)
            * 100
        )
        pair_values = sorted(
            (analyzer.pair_counts[p] for p in combinations(combo, 2)),
            reverse=True,
        )
        pair_score = sum(pair_values[:5]) / max(1, max_pair * 5) * 100

        triple_values = sorted(
            (analyzer.triple_counts[t] for t in combinations(combo, 3)),
            reverse=True,
        )
        triple_score = sum(triple_values[:5]) / max(1, max_triple * 5) * 100

        odd = sum(n % 2 for n in combo)
        high = sum(n >= 23 for n in combo)
        total = sum(combo)
        structure = 100.0
        structure -= abs(odd - 3) * 12
        structure -= abs(high - 3) * 10
        if total < 100:
            structure -= (100 - total) * 0.8
        elif total > 180:
            structure -= (total - 180) * 0.8
        structure = max(0.0, min(100.0, structure))

        return (
            number_score * 0.35
            + pair_score * 0.30
            + triple_score * 0.20
            + structure * 0.15
        )

    def _auto_mixed_preset(
        self,
        source_weights: Counter[int],
    ) -> str:
        """최근 20회를 검증구간으로 사용해 네 혼합비율 중 하나를 빠르게 선택합니다."""
        if len(self.a.draws) < 120:
            return "균형형"

        history = self.a.draws[:-20]
        targets = self.a.draws[-20:]
        target_counts = Counter()
        for draw in targets:
            target_counts.update(draw.numbers)

        windows = {
            100: self._window_analyzer(history, 100),
            500: self._window_analyzer(history, 500),
            1000: self._window_analyzer(history, 1000),
        }
        max_input = max(source_weights.values(), default=1)
        scores = {}

        for preset, weights in self.MIXED_PRESETS.items():
            input_w, w100, w500, w1000 = weights
            number_scores = {}

            for number in range(1, 46):
                input_score = source_weights[number] / max_input * 100
                history_scores = []
                for size in (100, 500, 1000):
                    analyzer = windows[size]
                    max_count = max(analyzer.number_counts.values(), default=1)
                    history_scores.append(
                        analyzer.number_counts[number] / max_count * 100
                    )

                number_scores[number] = (
                    input_score * input_w
                    + history_scores[0] * w100
                    + history_scores[1] * w500
                    + history_scores[2] * w1000
                )

            top_numbers = [
                n for n, _ in sorted(
                    number_scores.items(),
                    key=lambda item: (-item[1], item[0]),
                )[:15]
            ]
            scores[preset] = sum(target_counts[n] for n in top_numbers)

        return max(scores, key=scores.get)

    def generate_mixed(
        self,
        source_weights: Counter[int],
        count: int,
        preset: str,
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """입력번호와 최근 100·500·1000회 데이터를 실제 추천점수에 혼합합니다."""
        if len(source_weights) < 6:
            raise ValueError("통합데이터추천은 입력번호가 최소 6개 필요합니다.")

        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if preset == "자동최적형":
            applied_preset = self._auto_mixed_preset(source_weights)
        else:
            applied_preset = preset

        input_w, w100, w500, w1000 = self.MIXED_PRESETS.get(
            applied_preset,
            self.MIXED_PRESETS["균형형"],
        )

        analyzers = {
            100: self._window_analyzer(self.a.draws, 100),
            500: self._window_analyzer(self.a.draws, 500),
            1000: self._window_analyzer(self.a.draws, 1000),
        }

        max_input = max(source_weights.values(), default=1)
        number_blend = {}
        for number in range(1, 46):
            input_score = source_weights[number] / max_input * 100
            window_scores = []
            for size in (100, 500, 1000):
                analyzer = analyzers[size]
                max_count = max(analyzer.number_counts.values(), default=1)
                window_scores.append(
                    analyzer.number_counts[number] / max_count * 100
                )
            number_blend[number] = (
                input_score * input_w
                + window_scores[0] * w100
                + window_scores[1] * w500
                + window_scores[2] * w1000
            )

        ranked = [
            number for number, _ in sorted(
                number_blend.items(),
                key=lambda item: (-item[1], item[0]),
            )
            if number not in excluded_set
        ]

        # 입력번호는 우선 보존하고, 역사 데이터 상위번호로 후보군을 확장합니다.
        pool = list(dict.fromkeys(
            list(source_weights.keys()) + ranked + list(fixed_set | candidate_set)
        ))
        pool = [n for n in pool if n not in excluded_set]

        mandatory = list(fixed_set | candidate_set)
        selected = []
        for number in mandatory + pool:
            if number not in selected:
                selected.append(number)
            if len(selected) >= 18:
                break
        pool = sorted(selected)

        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue
            if not 20 <= sum(combo) <= 300:
                continue

            base_metrics = dict(self.metrics(combo, source_weights))
            input_score = base_metrics["input"]
            score100 = self._historical_combo_score(combo, analyzers[100])
            score500 = self._historical_combo_score(combo, analyzers[500])
            score1000 = self._historical_combo_score(combo, analyzers[1000])

            candidate_hits = len(combo_set & candidate_set)
            candidate_bonus = min(12.0, candidate_hits * 4.0)
            mixed_score = (
                input_score * input_w
                + score100 * w100
                + score500 * w500
                + score1000 * w1000
                + candidate_bonus
            )

            base_metrics.update({
                "mixed": mixed_score,
                "mixed_preset": applied_preset,
                "mixed_input_weight": input_w,
                "mixed_100_weight": w100,
                "mixed_500_weight": w500,
                "mixed_1000_weight": w1000,
                "score100": score100,
                "score500": score500,
                "score1000": score1000,
                "candidate_hits": candidate_hits,
                "candidate_bonus": candidate_bonus,
                "strategy": f"통합-{applied_preset}",
            })
            candidates.append((mixed_score, combo, base_metrics))

        candidates.sort(key=lambda row: (-row[0], row[1]))
        return self.select_diverse(candidates, count)

    PATTERN_NAMES = (
        "이월수", "2회전재등장", "단기강세", "장기미출현복귀",
        "끝수흐름", "연속수후보", "동반수확장", "간격수흐름",
    )

    @staticmethod
    def _normalize_counter(
        values: Counter[int] | dict[int, float],
    ) -> dict[int, float]:
        maximum = max(values.values(), default=0)
        if maximum <= 0:
            return {n: 0.0 for n in range(1, 46)}
        return {
            n: float(values.get(n, 0)) / maximum * 100.0
            for n in range(1, 46)
        }

    @staticmethod
    def _number_gaps(draws: list[Draw]) -> dict[int, int]:
        latest_index = len(draws) - 1
        last_seen = {}
        for index, draw in enumerate(draws):
            for number in draw.numbers:
                last_seen[number] = index
        return {
            number: latest_index - last_seen.get(number, -1)
            for number in range(1, 46)
        }

    @classmethod
    def _pattern_number_scores(
        cls,
        draws: list[Draw],
    ) -> tuple[dict[str, dict[int, float]], dict[int, list[str]]]:
        """각 패턴별 1~45 번호 점수와 번호별 추천 근거를 계산합니다."""
        if len(draws) < 10:
            raise ValueError("특이패턴 분석에는 최소 10회 이상의 데이터가 필요합니다.")

        last = draws[-1].numbers
        previous = draws[-2].numbers
        recent10 = draws[-10:]
        recent30 = draws[-30:]
        recent100 = draws[-100:]

        pattern_scores: dict[str, dict[int, float]] = {
            name: {n: 0.0 for n in range(1, 46)}
            for name in cls.PATTERN_NAMES
        }
        reasons: dict[int, list[str]] = defaultdict(list)

        # 1. 이월수: 역대 이월수 평균과 직전 회차 번호
        rollover_counts = Counter()
        for before, after in zip(draws[:-1], draws[1:]):
            rollover_counts[len(set(before.numbers) & set(after.numbers))] += 1
        total_transitions = max(1, sum(rollover_counts.values()))
        rollover_probability = 1.0 - rollover_counts[0] / total_transitions
        rollover_base = 72.0 + rollover_probability * 25.0
        for number in last:
            pattern_scores["이월수"][number] = rollover_base
            reasons[number].append("직전회차 이월수 후보")

        # 2. 2회 전 재등장: 2회 전에는 있었지만 직전에는 없던 번호
        for number in set(previous) - set(last):
            pattern_scores["2회전재등장"][number] = 82.0
            reasons[number].append("2회 전 번호 재등장 후보")

        # 3. 단기강세: 최근 10회와 30회 빈도를 혼합
        count10 = Counter(n for draw in recent10 for n in draw.numbers)
        count30 = Counter(n for draw in recent30 for n in draw.numbers)
        norm10 = cls._normalize_counter(count10)
        norm30 = cls._normalize_counter(count30)
        for number in range(1, 46):
            score = norm10[number] * 0.65 + norm30[number] * 0.35
            pattern_scores["단기강세"][number] = score
            if score >= 72:
                reasons[number].append("최근 10·30회 강세")

        # 4. 장기 미출현 복귀
        gaps = cls._number_gaps(draws)
        sorted_gaps = sorted(gaps.values())
        q70 = sorted_gaps[int(len(sorted_gaps) * 0.70)]
        max_gap = max(sorted_gaps, default=1)
        for number, gap in gaps.items():
            score = min(100.0, gap / max(1, max_gap) * 100.0)
            pattern_scores["장기미출현복귀"][number] = score
            if gap >= q70:
                reasons[number].append(f"{gap}회 미출현 복귀 후보")

        # 5. 끝수 흐름
        ending10 = Counter(n % 10 for draw in recent10 for n in draw.numbers)
        ending30 = Counter(n % 10 for draw in recent30 for n in draw.numbers)
        max_ending = max(
            (ending10[e] * 0.7 + ending30[e] * 0.3 for e in range(10)),
            default=1,
        )
        for number in range(1, 46):
            ending = number % 10
            raw = ending10[ending] * 0.7 + ending30[ending] * 0.3
            score = raw / max(1, max_ending) * 100.0
            pattern_scores["끝수흐름"][number] = score
            if score >= 78:
                reasons[number].append(f"끝수 {ending} 흐름 강세")

        # 6. 연속수 후보: 직전 번호의 앞뒤 번호
        for number in last:
            for candidate in (number - 1, number + 1):
                if 1 <= candidate <= 45 and candidate not in last:
                    pattern_scores["연속수후보"][candidate] = max(
                        pattern_scores["연속수후보"][candidate],
                        88.0,
                    )
                    reasons[candidate].append(f"{number}번 인접 연속수 후보")

        # 7. 동반수 확장: 직전 번호들과 최근100회에 자주 함께 나온 번호
        recent_pairs = Counter(
            pair for draw in recent100 for pair in combinations(draw.numbers, 2)
        )
        partner_raw = Counter()
        for number in last:
            for candidate in range(1, 46):
                if candidate == number or candidate in last:
                    continue
                pair = tuple(sorted((number, candidate)))
                partner_raw[candidate] += recent_pairs[pair]
        partner_norm = cls._normalize_counter(partner_raw)
        for number, score in partner_norm.items():
            pattern_scores["동반수확장"][number] = score
            if score >= 72:
                reasons[number].append("직전번호 동반수 확장")

        # 8. 간격수 흐름: 최근 당첨 조합에서 자주 나온 번호 간 차이를 직전번호에 적용
        gap_counts = Counter()
        for draw in recent30:
            nums = draw.numbers
            for a, b in combinations(nums, 2):
                gap = b - a
                if 1 <= gap <= 15:
                    gap_counts[gap] += 1
        common_gaps = [gap for gap, _ in gap_counts.most_common(5)]
        for source in last:
            for rank, gap in enumerate(common_gaps):
                score = 92.0 - rank * 8.0
                for candidate in (source - gap, source + gap):
                    if 1 <= candidate <= 45 and candidate not in last:
                        pattern_scores["간격수흐름"][candidate] = max(
                            pattern_scores["간격수흐름"][candidate],
                            score,
                        )
                        if score >= 76:
                            reasons[candidate].append(f"간격 {gap} 흐름 후보")

        return pattern_scores, reasons

    @classmethod
    def pattern_reliability(
        cls,
        draws: list[Draw],
        test_rounds: int = 60,
    ) -> dict[str, float]:
        """최근 과거 회차에서 패턴별 TOP10의 평균 적중도를 계산합니다."""
        reliability = {name: 50.0 for name in cls.PATTERN_NAMES}
        if len(draws) < 80:
            return reliability

        hit_totals = Counter()
        tested = 0
        start = max(20, len(draws) - test_rounds)
        for target_index in range(start, len(draws)):
            history = draws[:target_index]
            target = set(draws[target_index].numbers)
            scores, _ = cls._pattern_number_scores(history)
            for name, number_scores in scores.items():
                top10 = {
                    n for n, _ in sorted(
                        number_scores.items(),
                        key=lambda item: (-item[1], item[0]),
                    )[:10]
                }
                hit_totals[name] += len(top10 & target)
            tested += 1

        if tested:
            for name in cls.PATTERN_NAMES:
                average_hits = hit_totals[name] / tested
                # TOP10 무작위 기대 적중은 약 1.33개. 이를 기준으로 35~100점 환산.
                reliability[name] = round(
                    max(35.0, min(100.0, 50.0 + (average_hits - 1.33) * 28.0)),
                    1,
                )
        return reliability

    def pattern_board(
        self,
    ) -> tuple[list[dict[str, object]], dict[str, float]]:
        scores, reasons = self._pattern_number_scores(self.a.draws)
        reliability = self.pattern_reliability(self.a.draws)

        rows = []
        for number in range(1, 46):
            votes = []
            weighted_score = 0.0
            weight_sum = 0.0
            for name in self.PATTERN_NAMES:
                score = scores[name][number]
                rel = reliability[name]
                weighted_score += score * rel
                weight_sum += rel
                if score >= 68:
                    votes.append(name)

            final_score = weighted_score / max(1.0, weight_sum)
            rows.append({
                "number": number,
                "score": round(final_score, 1),
                "votes": len(votes),
                "patterns": votes,
                "reasons": reasons.get(number, []),
            })

        rows.sort(
            key=lambda row: (
                -int(row["votes"]),
                -float(row["score"]),
                int(row["number"]),
            )
        )
        return rows, reliability

    def pattern_briefing(self) -> str:
        board, reliability = self.pattern_board()
        top = board[:12]
        exclusions = sorted(
            board,
            key=lambda row: (
                int(row["votes"]),
                float(row["score"]),
                int(row["number"]),
            ),
        )[:8]

        latest = self.a.draws[-1]
        recent10 = self.a.draws[-10:]
        average_sum = sum(sum(d.numbers) for d in recent10) / len(recent10)
        average_odd = sum(
            sum(n % 2 for n in d.numbers) for d in recent10
        ) / len(recent10)
        strongest = sorted(
            reliability.items(),
            key=lambda item: (-item[1], item[0]),
        )[:3]

        top_text = " · ".join(
            f"{row['number']}({row['votes']}표)"
            for row in top
        )
        excluded_text = " · ".join(str(row["number"]) for row in exclusions)
        pattern_text = " / ".join(
            f"{name} {score:.1f}점" for name, score in strongest
        )
        return (
            f"이번 주 특이패턴 브리핑\n"
            f"최신 기준 회차: {latest.round_no}회\n"
            f"최근 10회 평균 합계: {average_sum:.1f} / 평균 홀수: {average_odd:.1f}개\n"
            f"검증점수 상위 패턴: {pattern_text}\n\n"
            f"패턴투표 핵심번호 TOP12\n{top_text}\n\n"
            f"AI 제외 검토번호 8개\n{excluded_text}\n\n"
            f"※ 패턴투표는 과거 통계를 이용한 분석이며 당첨을 보장하지 않습니다."
        )

    def historical_similar_draws(
        self,
        combo: tuple[int, ...],
        top_n: int = 5,
    ) -> list[tuple[int, float, tuple[int, ...]]]:
        """합계·홀짝·구간·끝수·연속수 특성이 비슷한 과거 회차를 찾습니다."""
        def signature(numbers: tuple[int, ...]) -> tuple:
            total = sum(numbers)
            odd = sum(n % 2 for n in numbers)
            low = sum(n <= 22 for n in numbers)
            endings = len({n % 10 for n in numbers})
            consecutive = self.consecutive_pairs(numbers)
            zones = (
                sum(1 <= n <= 10 for n in numbers),
                sum(11 <= n <= 20 for n in numbers),
                sum(21 <= n <= 30 for n in numbers),
                sum(31 <= n <= 40 for n in numbers),
                sum(41 <= n <= 45 for n in numbers),
            )
            return total, odd, low, endings, consecutive, zones

        target = signature(combo)
        rows = []
        for draw in self.a.draws:
            sig = signature(draw.numbers)
            distance = (
                abs(target[0] - sig[0]) / 25.0
                + abs(target[1] - sig[1]) * 0.8
                + abs(target[2] - sig[2]) * 0.6
                + abs(target[3] - sig[3]) * 0.4
                + abs(target[4] - sig[4]) * 0.8
                + sum(abs(a - b) for a, b in zip(target[5], sig[5])) * 0.35
            )
            similarity = max(0.0, 100.0 - distance * 12.0)
            rows.append((draw.round_no, round(similarity, 1), draw.numbers))
        rows.sort(key=lambda row: (-row[1], -row[0]))
        return rows[:top_n]

    def generate_pattern(
        self,
        count: int,
        mode: str = "자동종합",
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        board, reliability = self.pattern_board()
        board_map = {int(row["number"]): row for row in board}
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if mode != "자동종합" and mode in self.PATTERN_NAMES:
            pattern_scores, reasons = self._pattern_number_scores(self.a.draws)
            ranked_numbers = sorted(
                range(1, 46),
                key=lambda n: (-pattern_scores[mode][n], n),
            )
        else:
            ranked_numbers = [int(row["number"]) for row in board]
            _, reasons = self._pattern_number_scores(self.a.draws)

        selected = []
        for number in list(fixed_set | candidate_set) + ranked_numbers:
            if number in excluded_set or number in selected:
                continue
            selected.append(number)
            if len(selected) >= 20:
                break
        pool = sorted(selected)

        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_set and not fixed_set.issubset(combo_set):
                continue
            if excluded_set & combo_set:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            total = sum(combo)
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if not 85 <= total <= 200:
                continue

            values = [float(board_map[n]["score"]) for n in combo]
            votes = [int(board_map[n]["votes"]) for n in combo]
            pattern_score = sum(values) / 6.0 + sum(votes) * 1.7

            pair_values = [
                self.a.recent_pair_counts[pair]
                for pair in combinations(combo, 2)
            ]
            pair_bonus = min(12.0, sum(sorted(pair_values, reverse=True)[:4]) / 8.0)
            candidate_bonus = min(12.0, len(combo_set & candidate_set) * 4.0)
            final_score = pattern_score + pair_bonus + candidate_bonus

            metrics = dict(self.metrics(combo, Counter({n: 1 for n in pool})))
            combo_patterns = Counter()
            combo_reasons = []
            for number in combo:
                for pattern in board_map[number]["patterns"]:
                    combo_patterns[pattern] += 1
                combo_reasons.extend(reasons.get(number, []))

            leading_patterns = [
                name for name, _ in combo_patterns.most_common(3)
            ]
            metrics.update({
                "pattern": final_score,
                "pattern_score": pattern_score,
                "pattern_votes": sum(votes),
                "pattern_mode": mode,
                "pattern_names": leading_patterns,
                "pattern_reasons": list(dict.fromkeys(combo_reasons))[:5],
                "pattern_reliability": reliability,
                "strategy": f"특이패턴-{mode}",
            })
            candidates.append((final_score, combo, metrics))

        candidates.sort(key=lambda row: (-row[0], row[1]))
        return self.select_diverse(candidates, count)

    def category_score(self, metrics: dict[str, float], category: str) -> float:
        key = self.CATEGORY_NAMES.get(category, "composite")
        if key == "composite":
            return metrics["composite"]
        if key == "self":
            return metrics.get("self", metrics["composite"])
        # 항목별 순위는 해당 항목 70% + 자동 종합 30%
        return metrics[key] * 0.70 + metrics["composite"] * 0.30

    def self_number_scores(self) -> dict[int, float]:
        """역대 전체 데이터만으로 1~45 번호별 자체 점수를 계산합니다."""
        max_all = max(self.a.number_counts.values(), default=1)
        max_recent = max(self.a.recent_number_counts.values(), default=1)

        # 번호별 동반수 중심성: 다른 번호들과 같이 나온 횟수의 합
        pair_centrality = {}
        for number in range(1, 46):
            total = 0
            for other in range(1, 46):
                if number == other:
                    continue
                pair = tuple(sorted((number, other)))
                total += self.a.pair_counts[pair]
            pair_centrality[number] = total
        max_centrality = max(pair_centrality.values(), default=1)

        # 최신 출현 회차와 지연 정도
        latest_round = self.a.draws[-1].round_no if self.a.draws else 0
        last_seen = {n: 0 for n in range(1, 46)}
        for draw in self.a.draws:
            for n in draw.numbers:
                last_seen[n] = draw.round_no
        max_delay = max((latest_round - last_seen[n] for n in range(1, 46)), default=1)

        scores = {}
        for n in range(1, 46):
            all_score = self.a.number_counts[n] / max_all * 100
            recent_score = self.a.recent_number_counts[n] / max_recent * 100
            central_score = pair_centrality[n] / max_centrality * 100
            delay = latest_round - last_seen[n]
            delay_score = delay / max(1, max_delay) * 100

            scores[n] = (
                all_score * 0.35
                + recent_score * 0.30
                + central_score * 0.20
                + delay_score * 0.15
            )
        return scores

    def generate_self(
        self,
        count: int = 100,
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """사진·직접입력 없이 역대 전체 당첨번호만으로 추천합니다."""
        number_scores = self.self_number_scores()

        # 45개 전체 조합(약 814만 개)은 느리므로 자체 점수 상위 28개를 우선 탐색합니다.
        # 구간 다양성을 위해 각 10번대 구간의 강한 번호도 포함합니다.
        ranked = sorted(number_scores, key=lambda n: (-number_scores[n], n))
        pool = set(ranked[:24])
        for low, high in [(1, 10), (11, 20), (21, 30), (31, 40), (41, 45)]:
            zone = [n for n in ranked if low <= n <= high][:2]
            pool.update(zone)

        # 최대 28개로 제한
        pool = sorted(pool, key=lambda n: (-number_scores[n], n))[:28]
        pool = sorted(pool)

        # metrics()의 나온횟수 점수를 역대 출현빈도로 계산하기 위한 가중치
        historical_weights = Counter({
            n: max(1, self.a.number_counts[n])
            for n in range(1, 46)
        })

        candidates = []
        for combo in combinations(pool, 6):
            total = sum(combo)
            if not 20 <= total <= 300:
                continue

            odd = sum(n % 2 for n in combo)
            high = sum(n >= 23 for n in combo)
            if odd not in (2, 3, 4) or high not in (2, 3, 4):
                continue
            if self.consecutive_pairs(combo) > 2:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = self.metrics(combo, historical_weights)
            own_number_score = sum(number_scores[n] for n in combo) / 6.0

            # 자체추천 최종 점수: 번호 자체점수와 조합 통계의 자동 종합
            self_score = own_number_score * 0.45 + metrics["composite"] * 0.55
            metrics = dict(metrics)
            metrics["self"] = self_score
            candidates.append((self_score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[:count]

    def select_pool(
        self,
        source_weights: Counter[int],
        category: str,
        limit: int = 24,
    ) -> list[int]:
        """입력번호가 많아도 계산이 멈추지 않도록 카테고리별 핵심 번호를 선별합니다."""
        numbers = sorted(source_weights)
        if len(numbers) <= limit:
            return numbers

        max_input = max(source_weights.values(), default=1)

        pair_centrality = {}
        triple_centrality = {}
        for n in numbers:
            pair_centrality[n] = sum(
                self.a.pair_counts[tuple(sorted((n, other)))]
                for other in numbers
                if other != n
            )
            triple_centrality[n] = sum(
                count
                for triple, count in self.a.triple_counts.items()
                if n in triple
            )

        max_pair_c = max(pair_centrality.values(), default=1)
        max_triple_c = max(triple_centrality.values(), default=1)
        max_recent = max(
            (self.a.recent_number_counts[n] for n in numbers),
            default=1,
        )

        scores = {}
        for n in numbers:
            input_score = source_weights[n] / max_input * 100
            pair_score = pair_centrality[n] / max_pair_c * 100
            triple_score = triple_centrality[n] / max_triple_c * 100
            recent_score = self.a.recent_number_counts[n] / max_recent * 100

            if category == "나온횟수":
                score = input_score * 0.75 + pair_score * 0.10 + recent_score * 0.15
            elif category == "동반수":
                score = pair_score * 0.75 + input_score * 0.15 + recent_score * 0.10
            elif category == "트리플":
                score = triple_score * 0.75 + pair_score * 0.15 + input_score * 0.10
            elif category == "최근패턴":
                score = recent_score * 0.75 + pair_score * 0.15 + input_score * 0.10
            else:
                score = (
                    input_score * 0.30
                    + pair_score * 0.25
                    + triple_score * 0.20
                    + recent_score * 0.25
                )
            scores[n] = score

        ranked = sorted(numbers, key=lambda n: (-scores[n], n))

        # 각 번호 구간에서 최소 2개씩 포함해 특정 구간 쏠림을 방지
        selected = set(ranked[: max(14, limit - 10)])
        for low, high in [(1, 10), (11, 20), (21, 30), (31, 40), (41, 45)]:
            zone = [n for n in ranked if low <= n <= high][:2]
            selected.update(zone)

        final = sorted(selected, key=lambda n: (-scores[n], n))[:limit]
        return sorted(final)

    @staticmethod
    def filter_mode(input_count: int) -> tuple[str, str]:
        """입력 번호 개수에 따라 자동 필터 강도를 결정합니다."""
        if input_count <= 14:
            return "기본", "10~14개 입력: 결과 확보를 우선하는 완화 필터"
        if input_count <= 24:
            return "고급", "15~24개 입력: 균형과 통계를 함께 보는 고급 필터"
        return "정밀", "25개 이상 입력: 후보가 많아 더 엄격한 정밀 필터"

    def passes_filter(
        self,
        combo: tuple[int, ...],
        mode: str,
        sum_min: int,
        sum_max: int,
        allow_consecutive: bool,
    ) -> bool:
        total = sum(combo)
        if not sum_min <= total <= sum_max:
            return False

        odd = sum(n % 2 for n in combo)
        high = sum(n >= 23 for n in combo)
        consecutive = self.consecutive_pairs(combo)

        if not allow_consecutive and consecutive > 0:
            return False

        if mode == "기본":
            if odd not in (1, 2, 3, 4, 5):
                return False
            if high not in (0, 1, 2, 3, 4, 5, 6):
                return False
            if consecutive > 4:
                return False
            return True

        if mode == "고급":
            if odd not in (2, 3, 4):
                return False
            if high not in (2, 3, 4):
                return False
            if consecutive > 2:
                return False
            return True

        if odd not in (2, 3, 4):
            return False
        if high not in (2, 3, 4):
            return False
        if consecutive > 1:
            return False
        if total < 85 or total > 195:
            return False
        return True

    def relaxed_fallback(
        self,
        pool: list[int],
        source_weights: Counter[int],
        category: str,
        count: int,
        sum_min: int,
        sum_max: int,
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
        strategy: str = "균형형",
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        """필터 결과가 부족할 때 점수순으로 자동 보충합니다."""
        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_numbers and not set(fixed_numbers).issubset(combo_set):
                continue
            if excluded_numbers and combo_set.intersection(excluded_numbers):
                continue
            total = sum(combo)
            if not sum_min <= total <= sum_max:
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = dict(self.metrics(combo, source_weights))
            candidate_hits = len(combo_set.intersection(candidate_numbers))
            candidate_bonus = min(12.0, candidate_hits * 4.0)
            metrics["candidate_hits"] = candidate_hits
            metrics["candidate_bonus"] = candidate_bonus
            base_score = (
                self.strategy_score(metrics, strategy)
                if category == "추천조합"
                else self.category_score(metrics, category)
            )
            metrics["strategy"] = strategy
            score = base_score + candidate_bonus
            candidates.append((score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[:count]

    def recommendation_reason(
        self,
        metrics: dict[str, float],
        fixed_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
        combo: tuple[int, ...] = (),
    ) -> str:
        """추천 근거를 한 줄로 요약합니다."""
        labels = [
            ("나온횟수 강함", metrics.get("input", 0.0)),
            ("동반수 강함", metrics.get("pair", 0.0)),
            ("트리플 강함", metrics.get("triple", 0.0)),
            ("최근패턴 우수", metrics.get("recent", 0.0)),
            ("조합 균형 우수", metrics.get("structure", 0.0)),
        ]
        labels.sort(key=lambda item: (-item[1], item[0]))
        selected = [name for name, score in labels[:2] if score >= 35]
        if not selected:
            selected = [labels[0][0]]

        if fixed_numbers:
            selected.append("필수번호 " + ",".join(map(str, fixed_numbers)))

        included_candidates = sorted(set(combo) & set(candidate_numbers))
        if included_candidates:
            selected.append(
                "후보번호 포함 " + ",".join(map(str, included_candidates))
            )

        return " / ".join(selected)

    def generate(
        self,
        source_weights: Counter[int],
        count: int,
        sum_min: int,
        sum_max: int,
        allow_consecutive: bool,
        category: str,
        fixed_numbers: tuple[int, ...] = (),
        excluded_numbers: tuple[int, ...] = (),
        candidate_numbers: tuple[int, ...] = (),
        strategy: str = "균형형",
    ) -> list[tuple[float, tuple[int, ...], dict[str, float]]]:
        input_count = len(source_weights)
        mode, _ = self.filter_mode(input_count)

        fixed_numbers = tuple(sorted(set(fixed_numbers)))
        excluded_numbers = tuple(sorted(set(excluded_numbers)))
        candidate_numbers = tuple(sorted(set(candidate_numbers)))

        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if fixed_set & excluded_set:
            raise ValueError("필수번호와 제외번호가 중복됩니다.")
        if fixed_set & candidate_set:
            raise ValueError("필수번호와 후보번호가 중복됩니다.")
        if excluded_set & candidate_set:
            raise ValueError("제외번호와 후보번호가 중복됩니다.")

        missing_fixed = [n for n in fixed_numbers if n not in source_weights]
        if missing_fixed:
            raise ValueError(
                "필수번호는 번호 입력란에도 포함되어야 합니다: "
                + ", ".join(map(str, missing_fixed))
            )

        pool_limit = 24 if input_count >= 25 else input_count
        pool = self.select_pool(source_weights, category, limit=pool_limit)
        pool = sorted((set(pool) | fixed_set | candidate_set) - excluded_set)

        if len(pool) < 6:
            raise ValueError("고유 번호가 최소 6개 필요합니다.")

        candidates = []
        for combo in combinations(pool, 6):
            combo_set = set(combo)
            if fixed_numbers and not set(fixed_numbers).issubset(combo_set):
                continue
            if excluded_numbers and combo_set.intersection(excluded_numbers):
                continue
            if not self.passes_filter(
                combo, mode, sum_min, sum_max, allow_consecutive
            ):
                continue
            if combo in self.a.first_prize or combo in self.a.second_prize:
                continue

            metrics = dict(self.metrics(combo, source_weights))
            metrics["filter_mode"] = mode
            candidate_hits = len(combo_set.intersection(candidate_numbers))
            candidate_bonus = min(12.0, candidate_hits * 4.0)
            metrics["candidate_hits"] = candidate_hits
            metrics["candidate_bonus"] = candidate_bonus
            base_score = (
                self.strategy_score(metrics, strategy)
                if category == "추천조합"
                else self.category_score(metrics, category)
            )
            metrics["strategy"] = strategy
            score = base_score + candidate_bonus
            candidates.append((score, combo, metrics))

        candidates.sort(key=lambda x: (-x[0], x[1]))

        if len(candidates) < count:
            existing = {combo for _, combo, _ in candidates}
            fallback = self.relaxed_fallback(
                pool, source_weights, category, count, sum_min, sum_max,
                fixed_numbers=fixed_numbers,
                excluded_numbers=excluded_numbers,
                candidate_numbers=candidate_numbers,
                strategy=strategy,
            )
            for score, combo, metrics in fallback:
                if combo in existing:
                    continue
                metrics = dict(metrics)
                metrics["filter_mode"] = f"{mode}→자동완화"
                candidates.append((score, combo, metrics))
                existing.add(combo)
                if len(candidates) >= count:
                    break

        candidates.sort(key=lambda x: (-x[0], x[1]))
        return self.select_diverse(candidates, count)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.analyzer = LottoAnalyzer()
        self.photo_paths: list[str] = []
        self.recommendations: list[tuple[float, tuple[int, ...], dict[str, float]]] = []
        self.ocr_cache: dict[tuple[str, int, int], list[int]] = {}
        self.pattern_cache: dict[str, object] = {}
        self.suspend_auto_recommend = False

        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(1320, 850)
        self.setMinimumSize(1100, 700)

        self.stack = QStackedWidget()
        self.pages = [
            self.make_source_page(),
            self.make_recommend_page(),
        ]
        for p in self.pages:
            self.stack.addWidget(p)

        root = QWidget()
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.make_sidebar())
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        self.make_menu()
        self.apply_theme()
        self.statusBar().showMessage("역대 로또 Excel 파일을 불러오세요.")

    def make_sidebar(self) -> QWidget:
        box = QFrame()
        box.setObjectName("sidebar")
        box.setFixedWidth(245)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(18, 22, 18, 18)

        logo = QLabel("太炅")
        logo.setObjectName("logo")
        logo.setAlignment(Qt.AlignCenter)
        sub = QLabel("Lotto Lab Ultimate")
        sub.setObjectName("subtitle")
        sub.setAlignment(Qt.AlignCenter)
        lay.addWidget(logo)
        lay.addWidget(sub)
        lay.addSpacing(20)

        names = [
            "번호 입력", "추천조합", "나온횟수", "동반수", "트리플",
            "최근패턴", "통합데이터추천", "성과최적추천", "특이패턴추천", "자체추천"
        ]
        for i, name in enumerate(names):
            b = QPushButton(name)
            if i == 0:
                b.clicked.connect(lambda checked=False: self.stack.setCurrentIndex(0))
            else:
                category_name = name
                b.clicked.connect(
                    lambda checked=False, cat=category_name: self.show_recommend_category(cat)
                )
            lay.addWidget(b)

        lay.addStretch()
        b = QPushButton("역대 Excel 불러오기")
        b.setObjectName("primary")
        b.clicked.connect(self.open_excel)
        lay.addWidget(b)
        return box

    def make_menu(self) -> None:
        menu = self.menuBar().addMenu("파일")
        open_action = QAction("역대 Excel 불러오기", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_excel)
        menu.addAction(open_action)

        export_action = QAction("추천 결과 Excel 저장", self)
        export_action.triggered.connect(self.export_results)
        menu.addAction(export_action)

    def make_dashboard(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("대시보드")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        self.dashboard_info = QLabel(
            "역대 로또 당첨번호 Excel을 불러오면 분석이 시작됩니다.\n\n"
            "사진을 추가하면 별도 OCR 파일 없이 Windows 내장 OCR로 번호를 자동 인식합니다."
        )
        self.dashboard_info.setObjectName("card")
        self.dashboard_info.setAlignment(Qt.AlignCenter)
        self.dashboard_info.setMinimumHeight(230)
        lay.addWidget(self.dashboard_info)

        self.progress = QProgressBar()
        lay.addWidget(self.progress)
        lay.addStretch()
        return p

    def make_source_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("사진·번호 입력")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        self.excel_status = QLabel(
            "역대 Excel이 아직 등록되지 않았습니다. 왼쪽 아래 '역대 Excel 불러오기'를 누르세요."
        )
        self.excel_status.setObjectName("card")
        self.excel_status.setWordWrap(True)
        lay.addWidget(self.excel_status)

        self.excel_progress = QProgressBar()
        self.excel_progress.setRange(0, 100)
        self.excel_progress.setValue(0)
        lay.addWidget(self.excel_progress)

        grid = QGridLayout()
        left = QFrame()
        left.setObjectName("card")
        ll = QVBoxLayout(left)
        ll.addWidget(QLabel("사진 파일 등록"))
        self.photo_list = QListWidget()
        ll.addWidget(self.photo_list)
        row = QHBoxLayout()
        add = QPushButton("사진 추가·자동 인식")
        add.clicked.connect(self.add_photos)
        delete = QPushButton("선택 삭제")
        delete.clicked.connect(self.delete_photo)
        rerun = QPushButton("선택 사진 다시 인식")
        rerun.clicked.connect(self.rerun_selected_photo_ocr)
        row.addWidget(add)
        row.addWidget(delete)
        ll.addLayout(row)
        ll.addWidget(rerun)

        right = QFrame()
        right.setObjectName("card")
        rl = QVBoxLayout(right)
        rl.addWidget(QLabel(
            "사진 또는 메모에 나온 번호를 그대로 입력하세요.\n"
            "같은 번호가 반복되면 출현횟수 가중치로 반영됩니다."
        ))
        self.source_input = QPlainTextEdit()
        self.source_input.setPlaceholderText(
            "예:\n16 29 42 12 13\n"
            "2 6 8 9 15 18 22 28 30 34 35 37"
        )
        rl.addWidget(self.source_input)

        fixed_label = QLabel(
            "필수번호 — 자체추천을 제외한 모든 추천 조합에 반드시 포함됩니다."
        )
        fixed_label.setWordWrap(True)
        rl.addWidget(fixed_label)

        self.fixed_input = QLineEdit()
        self.fixed_input.setPlaceholderText("예: 3 6 또는 3, 6, 24")
        rl.addWidget(self.fixed_input)

        excluded_label = QLabel(
            "제외번호 — 자체추천을 제외한 모든 추천 조합에서 제거됩니다."
        )
        excluded_label.setWordWrap(True)
        rl.addWidget(excluded_label)

        self.excluded_input = QLineEdit()
        self.excluded_input.setPlaceholderText("예: 18 29")
        rl.addWidget(self.excluded_input)

        candidate_label = QLabel(
            "후보번호 — 포함 시 가산점을 받지만 필수는 아닙니다."
        )
        candidate_label.setWordWrap(True)
        rl.addWidget(candidate_label)

        self.candidate_input = QLineEdit()
        self.candidate_input.setPlaceholderText("예: 7 14 33")
        rl.addWidget(self.candidate_input)

        analyze = QPushButton("입력 번호 집계")
        analyze.setObjectName("primary")
        analyze.clicked.connect(self.update_source_counts)
        rl.addWidget(analyze)
        self.source_summary = QLabel("입력 대기")
        self.source_summary.setWordWrap(True)
        rl.addWidget(self.source_summary)

        grid.addWidget(left, 0, 0)
        grid.addWidget(right, 0, 1)
        lay.addLayout(grid)
        return p

    def make_stats_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)
        title = QLabel("통계 분석")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        self.stats_type = QComboBox()
        self.stats_type.addItems(["번호 빈도", "페어 상위 100", "트리플 상위 100"])
        self.stats_type.currentIndexChanged.connect(self.refresh_stats_table)
        lay.addWidget(self.stats_type)

        self.stats_table = QTableWidget(0, 3)
        self.stats_table.setHorizontalHeaderLabels(["순위", "번호/조합", "출현 횟수"])
        self.stats_table.setEditTriggers(QTableWidget.NoEditTriggers)
        lay.addWidget(self.stats_table)
        return p

    def make_recommend_page(self) -> QWidget:
        p = QWidget()
        lay = QVBoxLayout(p)

        title = QLabel("자동 추천 결과")
        title.setObjectName("pageTitle")
        lay.addWidget(title)

        guide = QLabel(
            "사진 또는 직접 입력한 번호를 기준으로 자동 계산합니다.\n"
            "추천 100조합 · 역대 1등·2등 동일 조합 제외\n"
            "특이패턴추천은 이월수·재등장·강세·미출현·끝수·연속수·동반수·간격수의 번호를 투표식으로 종합합니다.\n"
            "자체추천은 사진이나 직접입력 없이 역대 전체 당첨번호만으로 계산합니다."
        )
        guide.setObjectName("card")
        guide.setWordWrap(True)
        lay.addWidget(guide)

        self.rec_category = QComboBox()
        self.rec_category.addItems(
            [
                "추천조합", "나온횟수", "동반수", "트리플",
                "최근패턴", "통합데이터추천", "성과최적추천", "특이패턴추천", "자체추천"
            ]
        )
        self.rec_category.currentTextChanged.connect(self.generate_recommendations)
        lay.addWidget(self.rec_category)

        strategy_row = QHBoxLayout()
        strategy_label = QLabel("추천 전략")
        self.strategy_combo = QComboBox()
        self.strategy_combo.addItems(
            ["균형형", "출현형", "동반수형", "트리플형", "최근형", "AI Ultimate"]
        )
        self.strategy_combo.currentTextChanged.connect(self.generate_recommendations)
        strategy_row.addWidget(strategy_label)
        strategy_row.addWidget(self.strategy_combo)

        self.strategy_battle_btn = QPushButton("전략 배틀 백테스트")
        self.strategy_battle_btn.clicked.connect(self.run_strategy_battle)
        strategy_row.addWidget(self.strategy_battle_btn)
        lay.addLayout(strategy_row)

        mixed_row = QHBoxLayout()
        mixed_label = QLabel("통합데이터 비율")
        self.mixed_preset_combo = QComboBox()
        self.mixed_preset_combo.addItems(
            ["입력중심형", "최근중심형", "균형형", "장기형", "자동최적형"]
        )
        self.mixed_preset_combo.currentTextChanged.connect(
            self.generate_recommendations
        )
        mixed_row.addWidget(mixed_label)
        mixed_row.addWidget(self.mixed_preset_combo)
        mixed_help = QLabel(
            "입력번호 + 최근 100회 + 500회 + 1000회 데이터를 섞어 새 조합을 만듭니다."
        )
        mixed_help.setWordWrap(True)
        mixed_row.addWidget(mixed_help, 1)
        lay.addLayout(mixed_row)

        pattern_row = QHBoxLayout()
        pattern_label = QLabel("특이패턴")
        self.pattern_mode_combo = QComboBox()
        self.pattern_mode_combo.addItems(
            [
                "자동종합", "이월수", "2회전재등장", "단기강세",
                "장기미출현복귀", "끝수흐름", "연속수후보",
                "동반수확장", "간격수흐름",
            ]
        )
        self.pattern_mode_combo.currentTextChanged.connect(
            self.generate_recommendations
        )
        pattern_row.addWidget(pattern_label)
        pattern_row.addWidget(self.pattern_mode_combo)

        self.pattern_brief_btn = QPushButton("이번 주 패턴 브리핑")
        self.pattern_brief_btn.clicked.connect(self.show_pattern_briefing)
        pattern_row.addWidget(self.pattern_brief_btn)
        self.performance_report_btn = QPushButton("30,000회 최적화 결과")
        self.performance_report_btn.clicked.connect(self.show_performance_report)
        pattern_row.addWidget(self.performance_report_btn)

        self.round_search_input = QLineEdit()
        self.round_search_input.setPlaceholderText("회차/번호 검색: 33 37 40")
        pattern_row.addWidget(self.round_search_input, 1)
        self.round_search_btn = QPushButton("회차 검색")
        self.round_search_btn.clicked.connect(self.search_rounds)
        pattern_row.addWidget(self.round_search_btn)
        lay.addLayout(pattern_row)

        self.rec_status = QLabel(
            "역대 Excel을 불러오면 자체추천이 자동 계산됩니다.\n"
            "추천조합·나온횟수·동반수·트리플·최근패턴은 입력번호 6개 이상이 필요합니다."
        )
        self.rec_status.setWordWrap(True)
        self.rec_status.setObjectName("card")
        lay.addWidget(self.rec_status)

        self.rec_table = QTableWidget(0, 13)
        self.rec_table.setHorizontalHeaderLabels(
            [
                "순위", "추천조합", "추천 이유", "신뢰도", "등급",
                "카테고리 점수", "종합 점수", "나온횟수", "동반수",
                "트리플", "최근패턴", "동반출현 횟수", "합계"
            ]
        )
        self.rec_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.rec_table.cellClicked.connect(self.show_recommend_detail)
        lay.addWidget(self.rec_table)

        self.detail_box = QPlainTextEdit()
        self.detail_box.setReadOnly(True)
        self.detail_box.setPlaceholderText(
            "추천조합을 클릭하면 동반출현·트리플·점수 상세가 표시됩니다."
        )
        self.detail_box.setMaximumHeight(170)
        lay.addWidget(self.detail_box)

        export = QPushButton("현재 순위 Excel 저장")
        export.clicked.connect(self.export_results)
        lay.addWidget(export)
        return p

    def show_recommend_category(self, category: str) -> None:
        self.stack.setCurrentIndex(1)
        index = self.rec_category.findText(category)
        if index >= 0:
            self.rec_category.blockSignals(True)
            self.rec_category.setCurrentIndex(index)
            self.rec_category.blockSignals(False)
        self.generate_recommendations()

    def open_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "역대 로또 당첨번호 Excel 선택",
            "",
            "Excel (*.xlsx *.xls)",
        )
        if not path:
            return

        try:
            self.excel_progress.setValue(10)
            self.excel_status.setText(f"Excel 읽는 중: {Path(path).name}")
            self.statusBar().showMessage("역대 Excel을 읽고 있습니다...")
            QApplication.processEvents()

            self.analyzer.load_excel(path)
            self.pattern_cache.clear()

            self.excel_progress.setValue(75)
            latest = self.analyzer.draws[-1].round_no
            QApplication.processEvents()

            self.excel_status.setText(
                f"Excel 등록 완료: {Path(path).name}\n"
                f"분석 회차 {len(self.analyzer.draws):,}개 · 최신 {latest}회 · "
                f"1등 조합 {len(self.analyzer.first_prize):,}개"
            )
            self.excel_progress.setValue(100)

            # Excel만으로 계산 가능한 자체추천을 즉시 실행
            self.statusBar().showMessage("Excel 분석 완료 — 자체추천 계산 중...")
            self.show_recommend_category("자체추천")

            # 사진/직접입력 번호가 이미 있다면 다른 5개 항목도 사용할 수 있음을 표시
            try:
                counts = self.source_weights()
            except Exception:
                counts = Counter()

            if len(counts) >= 6:
                self.rec_status.setText(
                    "Excel과 입력번호가 모두 준비되었습니다. "
                    "왼쪽 항목을 누르면 각 기준 100조합이 계산됩니다."
                )

        except Exception as exc:
            self.excel_progress.setValue(0)
            self.excel_status.setText("Excel 등록 실패")
            QMessageBox.critical(
                self,
                "불러오기 오류",
                f"{exc}\n\n{traceback.format_exc(limit=3)}",
            )

    def prepare_ocr_image(self, image_path: str) -> tuple[str, str | None]:
        """고해상도 사진을 OCR에 충분한 크기로 축소해 처리시간을 줄입니다."""
        image = QImage(image_path)
        if image.isNull():
            return image_path, None

        max_side = max(image.width(), image.height())
        if max_side <= 1800:
            return image_path, None

        scaled = image.scaled(
            1800,
            1800,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation,
        )
        temp_file = tempfile.NamedTemporaryFile(
            prefix="taegyeong_ocr_",
            suffix=".jpg",
            delete=False,
        )
        temp_path = temp_file.name
        temp_file.close()
        if not scaled.save(temp_path, "JPG", 88):
            try:
                os.unlink(temp_path)
            except OSError:
                pass
            return image_path, None
        return temp_path, temp_path

    def run_windows_ocr(self, image_path: str) -> list[int]:
        """Windows OCR을 축소 이미지와 캐시로 빠르게 호출합니다."""
        if sys.platform != "win32":
            raise RuntimeError("사진 OCR은 Windows 10/11에서만 사용할 수 있습니다.")

        source_path = Path(image_path).resolve()
        stat = source_path.stat()
        cache_key = (str(source_path), stat.st_mtime_ns, stat.st_size)
        if cache_key in self.ocr_cache:
            return list(self.ocr_cache[cache_key])

        prepared_path, temp_path = self.prepare_ocr_image(str(source_path))

        encoded = base64.b64encode(
            WINDOWS_OCR_PS.encode("utf-16le")
        ).decode("ascii")

        env = os.environ.copy()
        env["LOTTO_OCR_IMAGE"] = prepared_path

        command = [
            "powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy", "Bypass",
            "-EncodedCommand", encoded,
        ]

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8-sig",
            errors="replace",
            timeout=120,
            env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        stdout = completed.stdout.strip()
        stderr = completed.stderr.strip()

        # PowerShell이 JSON 앞에 공백/경고를 붙인 경우 마지막 JSON 객체를 찾음
        json_line = ""
        for line in reversed(stdout.splitlines()):
            line = line.strip()
            if line.startswith("{") and line.endswith("}"):
                json_line = line
                break

        if not json_line:
            detail = stderr or stdout or "Windows OCR에서 결과를 받지 못했습니다."
            raise RuntimeError(detail[:1000])

        try:
            payload = json.loads(json_line)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "Windows OCR 결과를 해석하지 못했습니다.\n"
                f"출력: {json_line[:500]}"
            ) from exc

        if not payload.get("ok"):
            raise RuntimeError(payload.get("error", "Windows OCR 처리 실패"))

        numbers = payload.get("numbers") or []
        result = [int(n) for n in numbers if 1 <= int(n) <= 45]
        self.ocr_cache[cache_key] = list(result)
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        return result

    def append_ocr_numbers(self, numbers: list[int]) -> None:
        if not numbers:
            return
        current = self.source_input.toPlainText().rstrip()
        added = " ".join(map(str, numbers))
        self.source_input.blockSignals(True)
        self.source_input.setPlainText((current + "\n" + added).strip())
        self.source_input.blockSignals(False)
        self.update_source_counts()

    def add_photos(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "번호 사진 선택", "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp)"
        )
        if not paths:
            return

        all_numbers: list[int] = []
        failures: list[str] = []
        self.suspend_auto_recommend = True

        for path in paths:
            if path not in self.photo_paths:
                self.photo_paths.append(path)
                self.photo_list.addItem(Path(path).name)

            try:
                self.statusBar().showMessage(f"내장 OCR 인식 중: {Path(path).name}")
                QApplication.processEvents()
                all_numbers.extend(self.run_windows_ocr(path))
            except Exception as exc:
                failures.append(f"{Path(path).name}: {exc}")

        self.suspend_auto_recommend = False
        if all_numbers:
            self.append_ocr_numbers(all_numbers)
            self.statusBar().showMessage(
                f"사진 {len(paths)}장 처리 완료 — 숫자 {len(all_numbers)}개 인식, 추천조합 계산 완료"
            )
        else:
            self.statusBar().showMessage("사진에서 1~45 숫자를 찾지 못했습니다.")

        message = (
            f"사진 {len(paths)}장 처리 완료\n"
            f"인식된 숫자: {len(all_numbers)}개"
        )
        if failures:
            message += "\n\n일부 오류:\n" + "\n".join(failures[:5])
        QMessageBox.information(self, "사진 OCR 결과", message)

    def rerun_selected_photo_ocr(self) -> None:
        row = self.photo_list.currentRow()
        if row < 0:
            QMessageBox.information(self, "사진 선택", "다시 인식할 사진을 선택하세요.")
            return

        path = self.photo_paths[row]
        try:
            numbers = self.run_windows_ocr(path)
            if numbers:
                self.append_ocr_numbers(numbers)
                QMessageBox.information(
                    self,
                    "OCR 완료",
                    f"{Path(path).name}\n숫자 {len(numbers)}개를 입력란에 추가했습니다."
                )
            else:
                QMessageBox.information(
                    self,
                    "OCR 결과",
                    "사진에서 1~45 숫자를 찾지 못했습니다."
                )
        except Exception as exc:
            QMessageBox.warning(self, "OCR 오류", str(exc))

    def delete_photo(self) -> None:
        row = self.photo_list.currentRow()
        if row >= 0:
            self.photo_list.takeItem(row)
            self.photo_paths.pop(row)

    @staticmethod
    def _parse_special_numbers(text: str, label: str, maximum: int = 10) -> tuple[int, ...]:
        numbers = sorted({
            int(x)
            for x in re.findall(r"\d{1,2}", text)
            if 1 <= int(x) <= 45
        })
        if len(numbers) > maximum:
            raise ValueError(f"{label}는 최대 {maximum}개까지 입력할 수 있습니다.")
        return tuple(numbers)

    def fixed_numbers(self) -> tuple[int, ...]:
        text = self.fixed_input.text() if hasattr(self, "fixed_input") else ""
        return self._parse_special_numbers(text, "필수번호", 5)

    def excluded_numbers(self) -> tuple[int, ...]:
        text = self.excluded_input.text() if hasattr(self, "excluded_input") else ""
        return self._parse_special_numbers(text, "제외번호", 10)

    def candidate_numbers(self) -> tuple[int, ...]:
        text = self.candidate_input.text() if hasattr(self, "candidate_input") else ""
        return self._parse_special_numbers(text, "후보번호", 10)

    def validate_special_numbers(
        self,
        source_weights: Counter[int],
        fixed_numbers: tuple[int, ...],
        excluded_numbers: tuple[int, ...],
        candidate_numbers: tuple[int, ...],
    ) -> None:
        fixed_set = set(fixed_numbers)
        excluded_set = set(excluded_numbers)
        candidate_set = set(candidate_numbers)

        if fixed_set & excluded_set:
            raise ValueError(
                "필수번호와 제외번호에 같은 번호가 있습니다: "
                + ", ".join(map(str, sorted(fixed_set & excluded_set)))
            )
        if fixed_set & candidate_set:
            raise ValueError(
                "필수번호와 후보번호에 같은 번호가 있습니다: "
                + ", ".join(map(str, sorted(fixed_set & candidate_set)))
            )
        if excluded_set & candidate_set:
            raise ValueError(
                "제외번호와 후보번호에 같은 번호가 있습니다: "
                + ", ".join(map(str, sorted(excluded_set & candidate_set)))
            )

        missing_fixed = [n for n in fixed_numbers if n not in source_weights]
        if missing_fixed:
            raise ValueError(
                "필수번호는 일반 번호 입력란에도 포함되어야 합니다: "
                + ", ".join(map(str, missing_fixed))
            )

    def source_weights(self) -> Counter[int]:
        # 필수·제외·후보번호 입력란은 여기 합산하지 않습니다.
        # 따라서 일반 입력번호의 나온횟수가 중복 증가하지 않습니다.
        nums = parse_numbers(self.source_input.toPlainText())
        return Counter(nums)

    def update_source_counts(self) -> None:
        try:
            counts = self.source_weights()
            if not counts:
                self.source_summary.setText("입력된 번호가 없습니다.")
                return
            ranked = sorted(counts.items(), key=lambda x: (-x[1], x[0]))
            fixed = self.fixed_numbers()
            excluded = self.excluded_numbers()
            candidate = self.candidate_numbers()

            fixed_text = "없음" if not fixed else ", ".join(map(str, fixed))
            excluded_text = "없음" if not excluded else ", ".join(map(str, excluded))
            candidate_text = "없음" if not candidate else ", ".join(map(str, candidate))

            self.source_summary.setText(
                f"고유 번호 {len(counts)}개 / 전체 입력 {sum(counts.values())}개\n"
                f"필수번호: {fixed_text} / 제외번호: {excluded_text} / 후보번호: {candidate_text}\n"
                + " · ".join(f"{n}번 {c}회" for n, c in ranked)
            )
            if self.analyzer.draws and len(counts) >= 6:
                self.excel_status.setText(
                    self.excel_status.text()
                    + "\n입력번호 준비 완료 — 모든 추천 항목을 사용할 수 있습니다."
                )
                # 여러 사진 처리 중에는 매 사진마다 100조합을 재계산하지 않습니다.
                if not self.suspend_auto_recommend:
                    self.show_recommend_category("추천조합")
            elif not self.analyzer.draws:
                self.source_summary.setText(
                    self.source_summary.text()
                    + "\n역대 Excel을 불러오면 추천 계산이 시작됩니다."
                )
        except Exception as e:
            QMessageBox.warning(self, "번호 입력 오류", str(e))

    def refresh_stats_table(self) -> None:
        idx = self.stats_type.currentIndex()
        if idx == 0:
            items = sorted(
                ((n, self.analyzer.number_counts[n]) for n in range(1, 46)),
                key=lambda x: (-x[1], x[0])
            )
        elif idx == 1:
            items = self.analyzer.pair_counts.most_common(100)
        else:
            items = self.analyzer.triple_counts.most_common(100)

        self.stats_table.setRowCount(len(items))
        for r, (key, count) in enumerate(items, 1):
            text = str(key) if isinstance(key, int) else " · ".join(map(str, key))
            self.stats_table.setItem(r - 1, 0, QTableWidgetItem(str(r)))
            self.stats_table.setItem(r - 1, 1, QTableWidgetItem(text))
            self.stats_table.setItem(r - 1, 2, QTableWidgetItem(str(count)))
        self.stats_table.resizeColumnsToContents()

    def generate_recommendations(self, *_args) -> None:
        if not self.analyzer.draws:
            if hasattr(self, "rec_table"):
                self.rec_table.setRowCount(0)
            if hasattr(self, "rec_status"):
                self.rec_status.setText(
                    "먼저 왼쪽 아래의 '역대 Excel 불러오기'로 당첨번호 파일을 등록하세요."
                )
            self.statusBar().showMessage("역대 Excel이 필요합니다.")
            return
        try:
            recommender = Recommender(self.analyzer)
            category = self.rec_category.currentText()

            if category == "자체추천":
                self.statusBar().showMessage("역대 전체 데이터로 자체추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = recommender.generate_self(100)
                fixed_numbers = ()
                excluded_numbers = ()
                candidate_numbers = ()
                strategy = "자체추천"
            elif category == "성과최적추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.rec_status.setText(
                    "성과최적추천 계산 중 · 30,000개 설정 경쟁에서 선택된 가중치 적용"
                )
                self.statusBar().showMessage("성과최적추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = TKPerformanceEngine.generate(
                    self.analyzer,
                    100,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = "성과최적엔진"
            elif category == "특이패턴추천":
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                source_weights = self.source_weights()
                self.validate_special_numbers(
                    source_weights,
                    fixed_numbers,
                    excluded_numbers,
                    candidate_numbers,
                )
                mode = self.pattern_mode_combo.currentText()
                self.rec_status.setText(
                    f"특이패턴추천 계산 중 · {mode} · 패턴투표와 검증점수를 종합합니다."
                )
                self.statusBar().showMessage("특이패턴 추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = recommender.generate_pattern(
                    100,
                    mode=mode,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = f"특이패턴-{mode}"
            elif category == "통합데이터추천":
                weights = self.source_weights()
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                self.validate_special_numbers(
                    weights,
                    fixed_numbers,
                    excluded_numbers,
                    candidate_numbers,
                )
                if len(weights) < 6:
                    self.rec_table.setRowCount(0)
                    self.rec_status.setText(
                        "통합데이터추천은 입력번호가 최소 6개 필요합니다."
                    )
                    return
                preset = self.mixed_preset_combo.currentText()
                self.rec_status.setText(
                    f"통합데이터추천 계산 중 · 입력번호 + 최근100/500/1000회 · {preset}"
                )
                self.statusBar().showMessage("통합데이터추천 100조합 계산 중...")
                QApplication.processEvents()
                self.recommendations = recommender.generate_mixed(
                    weights,
                    100,
                    preset,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                )
                strategy = f"통합-{preset}"
            else:
                weights = self.source_weights()
                fixed_numbers = self.fixed_numbers()
                excluded_numbers = self.excluded_numbers()
                candidate_numbers = self.candidate_numbers()
                strategy = self.strategy_combo.currentText()
                self.validate_special_numbers(
                    weights,
                    fixed_numbers,
                    excluded_numbers,
                    candidate_numbers,
                )
                if len(weights) < 6:
                    self.rec_table.setRowCount(0)
                    self.rec_status.setText(
                        f"{category}은 사진 또는 직접 입력에서 고유 번호 6개 이상이 필요합니다."
                    )
                    self.statusBar().showMessage(
                        f"{category}: 사진 또는 직접 입력으로 고유 번호 6개 이상을 입력하세요."
                    )
                    return
                filter_name, filter_desc = recommender.filter_mode(len(weights))
                self.rec_status.setText(
                    f"{category} 100조합 계산 중 · {filter_desc} · 전략: {strategy}"
                )
                self.statusBar().showMessage(
                    f"{category} 계산 중 · {filter_name} 필터"
                )
                QApplication.processEvents()
                self.recommendations = recommender.generate(
                    weights,
                    100,
                    20,
                    300,
                    True,
                    category,
                    fixed_numbers=fixed_numbers,
                    excluded_numbers=excluded_numbers,
                    candidate_numbers=candidate_numbers,
                    strategy=strategy,
                )
            if not self.recommendations:
                QMessageBox.information(
                    self, "결과 없음",
                    "추천 가능한 조합이 없습니다. 입력번호를 10개 이상으로 확인해 주세요."
                )
                return

            key_map = Recommender.CATEGORY_NAMES
            selected_key = key_map.get(category, "composite")
            self.rec_table.setRowCount(len(self.recommendations))

            for r, (score, combo, metrics) in enumerate(self.recommendations, 1):
                pair_text = ", ".join(
                    f"{a}↔{b} {count}회"
                    for (a, b), count in recommender.pair_details(combo, 3)
                )
                reason = recommender.recommendation_reason(
                    metrics,
                    fixed_numbers,
                    candidate_numbers,
                    combo,
                )
                if metrics.get("pattern_names"):
                    reason = (
                        "패턴투표: " + ", ".join(metrics["pattern_names"])
                        + " / " + reason
                    )
                if metrics.get("performance_reasons"):
                    reason = (
                        "성과최적: " + " / ".join(metrics["performance_reasons"][:2])
                    )
                confidence = recommender.confidence_score(score, metrics)
                grade = recommender.confidence_grade(confidence)
                rank_text = f"TOP {r}" if r <= 10 else str(r)
                values = [
                    rank_text,
                    " · ".join(map(str, combo)),
                    reason,
                    f"{confidence:.1f}",
                    grade,
                    f"{score:.1f}",
                    f"{metrics['composite']:.1f}",
                    f"{metrics['input']:.1f}",
                    f"{metrics['pair']:.1f}",
                    f"{metrics['triple']:.1f}",
                    f"{metrics['recent']:.1f}",
                    pair_text,
                    str(sum(combo)),
                ]
                for c, value in enumerate(values):
                    self.rec_table.setItem(r - 1, c, QTableWidgetItem(value))

                # 현재 선택한 카테고리의 점수 칸을 강조
                metric_column = {
                    "composite": 6, "input": 7, "pair": 8,
                    "triple": 9, "recent": 10, "mixed": 5,
                    "pattern": 5, "performance": 5, "self": 5,
                }[selected_key]
                metric_value = metrics.get(selected_key, score)
                item = self.rec_table.item(r - 1, metric_column)
                if metric_value >= 70:
                    color = QColor("#2E7D32")
                elif metric_value >= 50:
                    color = QColor("#9A7B16")
                else:
                    color = QColor("#5A3A3A")
                item.setBackground(QBrush(color))
                item.setForeground(QBrush(QColor("#FFFFFF")))

                if r <= 10:
                    for top_col in range(self.rec_table.columnCount()):
                        top_item = self.rec_table.item(r - 1, top_col)
                        if top_item is not None:
                            top_item.setBackground(QBrush(QColor("#3A3215")))

                combo_item = self.rec_table.item(r - 1, 1)
                if metric_value >= 70:
                    combo_item.setForeground(QBrush(QColor("#7CFF8A")))
                elif metric_value >= 50:
                    combo_item.setForeground(QBrush(QColor("#FFD95A")))

            self.rec_table.resizeColumnsToContents()
            if category == "자체추천":
                filter_text = "역대 전체 데이터 자동분석"
            else:
                mode_values = {
                    str(row_metrics.get("filter_mode", "자동"))
                    for _, _, row_metrics in self.recommendations
                }
                filter_text = ", ".join(sorted(mode_values))

            self.rec_status.setText(
                f"{category} 기준 {len(self.recommendations)}조합 계산 완료 · "
                f"적용 필터: {filter_text} · 유사조합 자동분산 적용"
            )
            self.statusBar().showMessage(
                f"{category} 추천 {len(self.recommendations)}개 완료"
            )
        except Exception as e:
            QMessageBox.warning(self, "추천 오류", str(e))



    def show_performance_report(self) -> None:
        result = TKPerformanceEngine.OPTIMIZATION_RESULT
        weights = result["weights"]
        lines = [
            "TK 성과 최적화 연구결과",
            "",
            f"실제 시험 설정: {result['tested_settings']:,}개",
            "검증 방식: 각 회차 직전까지만 사용한 순차 워크포워드",
            f"최종 보류검증: {result['holdout']['round_start']}~{result['holdout']['round_end']}회",
            "",
            "[선택된 가중치]",
        ]
        for name, value in sorted(weights.items(), key=lambda item: -item[1]):
            lines.append(f"{name}: {value * 100:.2f}%")
        hold = result["holdout"]
        lines.extend([
            "",
            "[최종 보류구간 TOP15 번호 포함 성적]",
            f"평균 포함번호: {hold['average_top15_hits']:.3f}개",
            f"3개 이상 포함률: {hold['three_plus_rate'] * 100:.1f}%",
            f"4개 이상 포함률: {hold['four_plus_rate'] * 100:.1f}%",
            f"최고 포함번호: {hold['max_hits']}개",
            f"무작위 TOP15 기대값: {hold['random_expected_hits']:.1f}개",
            "",
            "※ 이 결과는 과거 데이터상의 검증 성적이며 미래 당첨을 보장하지 않습니다.",
            "※ 추천 점수는 실제 당첨확률이 아니라 엔진 내부 비교점수입니다.",
        ])
        QMessageBox.information(self, "성과 최적화 결과", "\n".join(lines))

    def _draw_data_path(self) -> Path:
        root = Path(os.getenv("LOCALAPPDATA", Path.home())) / "TaegyeongLottoLab"
        root.mkdir(parents=True, exist_ok=True)
        return root / "manual_draws.json"

    def _backup_draw_data(self) -> None:
        path = self._draw_data_path()
        if not path.exists():
            return
        backup_dir = path.parent / "backup"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(path, backup_dir / f"manual_draws_{stamp}.json")

    def manual_add_draw(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(self, "데이터 필요", "먼저 역대 Excel을 불러오세요.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("수동 회차 추가")
        form = QFormLayout(dialog)
        round_input = QLineEdit(str(self.analyzer.draws[-1].round_no + 1))
        numbers_input = QLineEdit()
        numbers_input.setPlaceholderText("예: 3 11 17 24 33 42")
        bonus_input = QLineEdit()
        bonus_input.setPlaceholderText("예: 7")
        form.addRow("회차", round_input)
        form.addRow("당첨번호 6개", numbers_input)
        form.addRow("보너스번호", bonus_input)
        save_btn = QPushButton("검사 후 저장")
        form.addRow(save_btn)

        def save():
            try:
                round_no = int(round_input.text().strip())
                nums = sorted(parse_numbers(numbers_input.text()))
                bonus_values = parse_numbers(bonus_input.text())
                if len(nums) != 6 or len(set(nums)) != 6:
                    raise ValueError("서로 다른 당첨번호 6개를 입력하세요.")
                if len(bonus_values) != 1:
                    raise ValueError("보너스번호 1개를 입력하세요.")
                bonus = bonus_values[0]
                if bonus in nums:
                    raise ValueError("보너스번호가 당첨번호와 중복됩니다.")
                existing = {draw.round_no for draw in self.analyzer.draws}
                if round_no in existing:
                    raise ValueError("이미 등록된 회차입니다.")
                expected = self.analyzer.draws[-1].round_no + 1
                if round_no != expected:
                    answer = QMessageBox.question(
                        dialog, "회차 확인",
                        f"다음 예상 회차는 {expected}회입니다. {round_no}회로 저장할까요?"
                    )
                    if answer != QMessageBox.Yes:
                        return
                self._backup_draw_data()
                path = self._draw_data_path()
                records = []
                if path.exists():
                    records = json.loads(path.read_text(encoding="utf-8"))
                records.append({
                    "round_no": round_no,
                    "numbers": nums,
                    "bonus": bonus,
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                })
                path.write_text(
                    json.dumps(records, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                self.analyzer.draws.append(
                    Draw(round_no, tuple(nums), bonus)
                )
                self.analyzer.draws.sort(key=lambda draw: draw.round_no)
                self.analyzer._analyze()
                self.pattern_cache.clear()
                dialog.accept()
                self.refresh_all()
                QMessageBox.information(
                    self, "저장 완료",
                    f"{round_no}회가 추가되었습니다. 분석 캐시도 새 데이터 기준으로 갱신했습니다."
                )
            except Exception as exc:
                QMessageBox.warning(dialog, "입력 오류", str(exc))

        save_btn.clicked.connect(save)
        dialog.exec()

    @staticmethod
    def _parse_official_latest_html(html: str):
        rounds = [int(value) for value in re.findall(r"(\d{1,4})회", html)]
        if not rounds:
            raise ValueError("공식 페이지에서 최신 회차를 찾지 못했습니다.")
        latest = max(rounds)
        # 공식 결과 페이지의 최신 회차 주변에서 번호 6개와 보너스를 탐색
        marker = html.find(f"{latest}회")
        segment = html[marker:marker + 12000] if marker >= 0 else html[:12000]
        clean = re.sub(r"<[^>]+>", " ", segment)
        values = [int(v) for v in re.findall(r"(?<!\d)([1-9]|[1-3]\d|4[0-5])(?!\d)", clean)]
        # 중복을 보존하되 첫 정상적인 7개 연속 범위를 탐색
        for i in range(max(0, len(values) - 30)):
            candidate = values[i:i + 7]
            if len(candidate) == 7 and len(set(candidate[:6])) == 6 and candidate[6] not in candidate[:6]:
                return latest, sorted(candidate[:6]), candidate[6]
        raise ValueError("공식 페이지에서 번호를 안전하게 해석하지 못했습니다.")

    def check_latest_draw(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(self, "데이터 필요", "먼저 역대 Excel을 불러오세요.")
            return
        try:
            request = urllib.request.Request(
                "https://www.dhlottery.co.kr/lt645/result",
                headers={"User-Agent": "Mozilla/5.0"},
            )
            with urllib.request.urlopen(request, timeout=15) as response:
                html = response.read().decode("utf-8", errors="ignore")
            latest, numbers, bonus = self._parse_official_latest_html(html)
            current = self.analyzer.draws[-1].round_no
            if latest <= current:
                QMessageBox.information(
                    self, "최신 회차 확인",
                    f"현재 보유 데이터 {current}회가 최신입니다."
                )
                return
            QMessageBox.information(
                self, "새 회차 발견",
                f"현재 보유: {current}회\n공식 페이지 최신: {latest}회\n"
                f"당첨번호: {' · '.join(map(str, numbers))}\n보너스: {bonus}\n\n"
                "번호를 확인한 뒤 데이터 관리 > 수동 회차 추가에서 저장하세요. "
                "초기 안정화 기간에는 자동저장 대신 확인 절차를 유지합니다."
            )
        except Exception as exc:
            QMessageBox.warning(
                self, "자동 확인 실패",
                "공식 페이지 구조 변경 또는 인터넷 연결 문제로 자동 확인하지 못했습니다.\n"
                f"상세: {exc}\n\n수동 회차 추가 기능을 이용할 수 있습니다."
            )

    def show_pattern_briefing(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(
                self, "데이터 필요", "먼저 역대 Excel을 불러오세요."
            )
            return
        try:
            recommender = Recommender(self.analyzer)
            cache_key = f"briefing:{len(self.analyzer.draws)}"
            cached = self.pattern_cache.get(cache_key)
            if cached is None:
                briefing = recommender.pattern_briefing()
                board, reliability = recommender.pattern_board()
                self.pattern_cache[cache_key] = (briefing, board, reliability)
            else:
                briefing, board, reliability = cached
            lines = [briefing, "", "번호별 패턴투표 TOP20"]
            for rank, row in enumerate(board[:20], 1):
                patterns = ", ".join(row["patterns"]) or "단독 점수"
                reasons = " / ".join(row["reasons"][:2])
                lines.append(
                    f"{rank:02d}. {row['number']}번 | {row['votes']}표 | "
                    f"{row['score']:.1f}점 | {patterns}"
                    + (f" | {reasons}" if reasons else "")
                )
            self.detail_box.setPlainText("\n".join(lines))
            self.rec_status.setText(
                "이번 주 패턴 브리핑 완료 · 상세창에서 핵심번호와 제외번호를 확인하세요."
            )
        except Exception as exc:
            QMessageBox.critical(self, "패턴 분석 오류", str(exc))

    def search_rounds(self) -> None:
        if not self.analyzer.draws:
            QMessageBox.information(
                self, "데이터 필요", "먼저 역대 Excel을 불러오세요."
            )
            return
        try:
            values = parse_numbers(self.round_search_input.text())
            if not values:
                raise ValueError("검색할 번호를 입력하세요.")
            target = set(values)
            rows = []
            for draw in reversed(self.analyzer.draws):
                matched = sorted(target & set(draw.numbers))
                if matched:
                    rows.append(
                        f"{draw.round_no}회 | {' · '.join(map(str, draw.numbers))} | "
                        f"일치 {len(matched)}개: {', '.join(map(str, matched))}"
                    )
                if len(rows) >= 100:
                    break
            if not rows:
                rows = ["일치하는 회차가 없습니다."]
            self.detail_box.setPlainText(
                f"번호 검색: {', '.join(map(str, values))}\n\n"
                + "\n".join(rows)
            )
        except Exception as exc:
            QMessageBox.warning(self, "회차 검색", str(exc))

    def run_strategy_battle(self) -> None:
        if len(self.analyzer.draws) < 120:
            QMessageBox.information(
                self,
                "데이터 부족",
                "전략 배틀은 최소 120회 이상의 역대 데이터가 필요합니다.",
            )
            return

        answer = QMessageBox.question(
            self,
            "전략 배틀",
            "최근 100회를 과거 데이터만 사용해 백테스트합니다.\n"
            "PC 성능에 따라 1~3분 정도 걸릴 수 있습니다. 실행할까요?",
        )
        if answer != QMessageBox.Yes:
            return

        self.strategy_battle_btn.setEnabled(False)
        self.rec_status.setText("전략 배틀 준비 중...")
        QApplication.processEvents()

        try:
            strategies = list(Recommender.STRATEGY_WEIGHTS)
            stats = {
                name: {
                    "three_plus": 0,
                    "four_plus": 0,
                    "five_plus": 0,
                    "six": 0,
                    "hit_sum": 0,
                    "rounds": 0,
                }
                for name in strategies
            }

            draws = self.analyzer.draws
            start_index = max(20, len(draws) - 100)

            for test_no, target_index in enumerate(
                range(start_index, len(draws)), 1
            ):
                history = draws[:target_index]
                target = set(draws[target_index].numbers)

                temp = LottoAnalyzer()
                temp.draws = list(history)
                temp._analyze()
                recommender = Recommender(temp)

                # 과거 출현빈도 상위 15개를 입력번호로 가정해 전략별 TOP10 생성
                ranked_numbers = [
                    n for n, _ in temp.number_counts.most_common(15)
                ]
                source_weights = Counter({
                    n: max(1, temp.number_counts[n])
                    for n in ranked_numbers
                })

                for strategy in strategies:
                    rows = recommender.generate(
                        source_weights,
                        10,
                        20,
                        300,
                        True,
                        "추천조합",
                        strategy=strategy,
                    )
                    best_hit = max(
                        (len(target & set(combo)) for _, combo, _ in rows),
                        default=0,
                    )
                    item = stats[strategy]
                    item["rounds"] += 1
                    item["hit_sum"] += best_hit
                    if best_hit >= 3:
                        item["three_plus"] += 1
                    if best_hit >= 4:
                        item["four_plus"] += 1
                    if best_hit >= 5:
                        item["five_plus"] += 1
                    if best_hit >= 6:
                        item["six"] += 1

                if test_no % 5 == 0:
                    self.rec_status.setText(
                        f"전략 배틀 진행 중 · {test_no}/100회"
                    )
                    QApplication.processEvents()

            rows = []
            for strategy, item in stats.items():
                rounds = max(1, item["rounds"])
                average = item["hit_sum"] / rounds
                battle_score = (
                    item["three_plus"] * 1
                    + item["four_plus"] * 3
                    + item["five_plus"] * 10
                    + item["six"] * 50
                    + average * 10
                )
                rows.append((battle_score, strategy, item, average))

            rows.sort(reverse=True)

            lines = [
                "최근 100회 전략 배틀 결과",
                "각 회차마다 이전 회차 데이터만 사용하고 전략별 TOP10 중 최고 적중을 비교합니다.",
                "",
            ]
            for rank, (_, strategy, item, average) in enumerate(rows, 1):
                lines.append(
                    f"{rank}위 {strategy} | 평균 최고적중 {average:.2f}개 | "
                    f"3개+ {item['three_plus']}회 | 4개+ {item['four_plus']}회 | "
                    f"5개+ {item['five_plus']}회 | 6개 {item['six']}회"
                )

            result_text = "\n".join(lines)
            self.detail_box.setPlainText(result_text)
            self.rec_status.setText(
                f"전략 배틀 완료 · 1위: {rows[0][1]}"
            )
            QMessageBox.information(
                self,
                "전략 배틀 완료",
                f"최근 100회 기준 1위 전략은 '{rows[0][1]}'입니다.\n"
                "자세한 결과는 추천 결과 아래 상세창에서 확인하세요.",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "전략 배틀 오류",
                f"{exc}\n\n{traceback.format_exc(limit=3)}",
            )
        finally:
            self.strategy_battle_btn.setEnabled(True)

    def show_recommend_detail(self, row: int, _column: int) -> None:
        if row < 0 or row >= len(self.recommendations):
            return

        recommender = Recommender(self.analyzer)
        score, combo, metrics = self.recommendations[row]
        category = self.rec_category.currentText()

        if category == "자체추천":
            fixed_numbers = ()
            candidate_numbers = ()
        else:
            fixed_numbers = self.fixed_numbers()
            candidate_numbers = self.candidate_numbers()

        confidence = recommender.confidence_score(score, metrics)
        grade = recommender.confidence_grade(confidence)
        reason = recommender.recommendation_reason(
            metrics, fixed_numbers, candidate_numbers, combo
        )

        pair_lines = [
            f"{a}↔{b}: {count}회"
            for (a, b), count in recommender.pair_details(combo, 5)
        ]
        similar_lines = [
            f"{round_no}회 · 유사도 {similarity:.1f}% · "
            + " · ".join(map(str, numbers))
            for round_no, similarity, numbers
            in recommender.historical_similar_draws(combo, 5)
        ]
        triple_lines = [
            f"{a}-{b}-{c}: {count}회"
            for (a, b, c), count in recommender.triple_details(combo, 5)
        ]

        mixed_detail = ""
        if metrics.get("mixed_preset"):
            mixed_detail = (
                f"통합비율: {metrics['mixed_preset']} "
                f"(입력 {metrics['mixed_input_weight']:.0%} / "
                f"100회 {metrics['mixed_100_weight']:.0%} / "
                f"500회 {metrics['mixed_500_weight']:.0%} / "
                f"1000회 {metrics['mixed_1000_weight']:.0%})\n"
                f"구간점수: 100회 {metrics['score100']:.1f} / "
                f"500회 {metrics['score500']:.1f} / "
                f"1000회 {metrics['score1000']:.1f}\n"
            )

        text = (
            f"추천조합: {' · '.join(map(str, combo))}\n"
            f"추천신뢰도: {confidence:.1f}점 ({grade}등급)\n"
            f"추천전략: {metrics.get('strategy', '자동')}\n"
            f"추천이유: {reason}\n"
            f"{mixed_detail}"
            f"합계: {sum(combo)} / 홀수 {sum(n % 2 for n in combo)}개 / "
            f"고번호 {sum(n >= 23 for n in combo)}개\n\n"
            f"[동반출현 횟수]\n" + "\n".join(pair_lines) +
            f"\n\n[트리플 출현 횟수]\n" + "\n".join(triple_lines) +
            f"\n\n[세부 점수]\n"
            f"나온횟수 {metrics.get('input', 0):.1f} / "
            f"동반수 {metrics.get('pair', 0):.1f} / "
            f"트리플 {metrics.get('triple', 0):.1f} / "
            f"최근패턴 {metrics.get('recent', 0):.1f} / "
            f"조합균형 {metrics.get('structure', 0):.1f}"
            + (
                f"\n\n[특이패턴]\n"
                f"패턴투표 {metrics.get('pattern_votes', 0)}표 / "
                f"패턴점수 {metrics.get('pattern_score', 0):.1f}\n"
                f"주요패턴: {', '.join(metrics.get('pattern_names', []))}\n"
                f"추천근거: {' / '.join(metrics.get('pattern_reasons', []))}"
                if metrics.get("pattern_names") else ""
            )
            + (
                "\n\n[성과최적 엔진 근거]\n"
                + "\n".join(metrics.get("performance_reasons", []))
                if metrics.get("performance_reasons") else ""
            )
            + "\n\n[역대 유사 회차]\n" + "\n".join(similar_lines)
        )
        self.detail_box.setPlainText(text)

    def export_results(self) -> None:
        if not self.recommendations:
            QMessageBox.information(self, "저장할 결과 없음", "먼저 추천조합을 생성하세요.")
            return
        category = self.rec_category.currentText()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "추천 결과 저장",
            f"Taegyeong_Lotto_{category}_추천결과.xlsx",
            "Excel (*.xlsx)",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        recommender = Recommender(self.analyzer)
        if category == "자체추천":
            fixed_numbers = ()
            excluded_numbers = ()
            candidate_numbers = ()
        else:
            fixed_numbers = self.fixed_numbers()
            excluded_numbers = self.excluded_numbers()
            candidate_numbers = self.candidate_numbers()

        rows = []
        for rank, (score, combo, metrics) in enumerate(self.recommendations, 1):
            rows.append({
                "카테고리": category,
                "추천전략": metrics.get("strategy", "자체추천"),
                "통합프리셋": metrics.get("mixed_preset", ""),
                "입력비중": metrics.get("mixed_input_weight", ""),
                "최근100회비중": metrics.get("mixed_100_weight", ""),
                "최근500회비중": metrics.get("mixed_500_weight", ""),
                "최근1000회비중": metrics.get("mixed_1000_weight", ""),
                "패턴모드": metrics.get("pattern_mode", ""),
                "패턴투표수": metrics.get("pattern_votes", ""),
                "패턴점수": metrics.get("pattern_score", ""),
                "주요패턴": ", ".join(metrics.get("pattern_names", [])),
                "패턴추천근거": " / ".join(metrics.get("pattern_reasons", [])),
                "성과최적점수": metrics.get("performance", ""),
                "성과최적근거": " / ".join(metrics.get("performance_reasons", [])),
                "순위": rank,
                "번호1": combo[0], "번호2": combo[1], "번호3": combo[2],
                "번호4": combo[3], "번호5": combo[4], "번호6": combo[5],
                "필수번호": ", ".join(map(str, fixed_numbers)),
                "제외번호": ", ".join(map(str, excluded_numbers)),
                "후보번호": ", ".join(map(str, candidate_numbers)),
                "추천이유": recommender.recommendation_reason(
                    metrics,
                    fixed_numbers,
                    candidate_numbers,
                    combo,
                ),
                "추천신뢰도": recommender.confidence_score(score, metrics),
                "등급": recommender.confidence_grade(
                    recommender.confidence_score(score, metrics)
                ),
                "카테고리점수": round(score, 1),
                "종합점수": round(metrics["composite"], 1),
                "입력횟수점수": round(metrics["input"], 1),
                "동반수점수": round(metrics["pair"], 1),
                "트리플점수": round(metrics["triple"], 1),
                "최근패턴점수": round(metrics["recent"], 1),
                "자체추천점수": round(metrics.get("self", 0.0), 1),
                "동반출현횟수": ", ".join(
                    f"{a}-{b}({count}회)"
                    for (a, b), count in recommender.pair_details(combo, 3)
                ),
                "합계": sum(combo),
                "역대1등동일": "아니오",
                "역대2등동일": "아니오",
            })
        pd.DataFrame(rows).to_excel(path, index=False)
        QMessageBox.information(self, "저장 완료", path)

    def apply_theme(self) -> None:
        self.setStyleSheet("""
        QMainWindow, QWidget {
            background:#111111; color:#F4F0E6;
            font-family:"Malgun Gothic"; font-size:14px;
        }
        #sidebar { background:#080808; border-right:1px solid #4A3A12; }
        #logo { color:#D4AF37; font-size:48px; font-weight:800; }
        #subtitle { color:#E8D9A7; font-weight:700; }
        #pageTitle { color:#D4AF37; font-size:28px; font-weight:800; padding:8px; }
        #card { background:#1A1A1A; border:1px solid #4A3A12;
                border-radius:12px; padding:16px; }
        QPushButton {
            background:#252525; color:#F4F0E6; border:1px solid #3A3A3A;
            border-radius:8px; padding:11px; text-align:left;
        }
        QPushButton:hover { border-color:#D4AF37; background:#302817; }
        #primary { background:#D4AF37; color:#111111; font-weight:800; text-align:center; }
        QPlainTextEdit, QLineEdit, QListWidget, QTableWidget, QSpinBox, QComboBox {
            background:#181818; color:#F4F0E6; border:1px solid #404040;
            border-radius:7px; padding:6px;
        }
        QHeaderView::section {
            background:#2A2416; color:#F0D980; padding:8px; border:0;
        }
        QProgressBar { background:#222; border:1px solid #444; border-radius:7px; text-align:center; }
        QProgressBar::chunk { background:#D4AF37; border-radius:6px; }
        """)


def main() -> int:
    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 10))
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
