#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NGB 外国年金請求書（兼 請求明細書）CSV変換ツール

使い方:
    python ngb_extract.py invoice.pdf
    python ngb_extract.py *.pdf
    python ngb_extract.py invoice.pdf --out-dir output

必要ライブラリ:
    pip install PyMuPDF

出力:
    請求書番号_full_extract.csv
    例: B001784512_full_extract.csv
"""

from __future__ import annotations

import argparse
import csv
import glob
import re
import statistics
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence

import pymupdf


CSV_COLUMNS = [
    "請求書番号",
    "得意先コード",
    "作成日",
    "ページ",
    "識別",
    "登録番号",
    "年次",
    "納付期限",
    "分割比率",
    "貴社整理番号",
    "メモ",
    "共願先1",
    "共願先2",
    "請求明細番号",
    "外貨額",
    "費用円貨額",
    "外貨コード",
    "（レート）",
    "手数料",
    "税抜き金額",
]

G_REF_RE = re.compile(r"^G\d{8}[A-Z]{2}-\d{2}$")
DETAIL_NO_RE = re.compile(r"^F[A-Z0-9]{10,}$")
INVOICE_RE = re.compile(r"B\d{9}")
CUSTOMER_RE = re.compile(r"^\d{7}$")
TNO_RE = re.compile(r"T\d{13}")
DATE_RE = re.compile(r"\d{4}年\d{1,2}月\d{1,2}日")
MONTH_RE = re.compile(r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\.$")
NUMBER_RE = re.compile(r"^\d{1,3}(?:,\d{3})*(?:\.\d+)?$|^\d+(?:\.\d+)?$")


@dataclass(slots=True)
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    block: int
    line: int
    word: int

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass(slots=True)
class Header:
    invoice_no: str = ""
    customer_code: str = ""
    create_date: str = ""
    t_number: str = ""


@dataclass(slots=True)
class Detail:
    page: int = 0
    country: str = ""
    kind: str = ""
    register_no: str = ""
    year: str = ""
    due_date: str = ""
    ratio: str = ""
    customer_ref: str = ""
    memo: str = ""
    joint1: str = ""
    joint2: str = ""
    detail_no: str = ""
    foreign_amount: str = ""
    yen_amount: str = ""
    currency: str = ""
    rate: str = ""
    fee: str = ""
    subtotal: str = ""

    def to_csv_row(self, header: Header) -> list[str]:
        return [
            header.invoice_no,
            header.customer_code,
            header.create_date,
            str(self.page),
            self.kind,
            self.register_no,
            self.year,
            self.due_date,
            self.ratio,
            self.customer_ref,
            self.memo,
            self.joint1,
            self.joint2,
            self.detail_no,
            self.foreign_amount,
            self.yen_amount,
            self.currency,
            self.rate,
            self.fee,
            self.subtotal,
        ]


@dataclass(slots=True)
class DetailBlock:
    page: int
    top: float
    bottom: float
    words: list[Word]


def is_number(text: str) -> bool:
    return bool(NUMBER_RE.fullmatch(text))


def to_int_amount(text: str) -> int:
    if not text:
        return 0
    return int(text.replace(",", "").split(".")[0])


def words_to_text(words: Sequence[Word]) -> str:
    return " ".join(w.text for w in sorted(words, key=lambda w: (w.y0, w.x0)))


def group_lines(words: Sequence[Word], y_tolerance: float = 3.0) -> list[list[Word]]:
    """座標付き単語をY座標で行にまとめる。"""
    groups: list[list[Word]] = []

    for word in sorted(words, key=lambda w: (w.cy, w.x0)):
        for group in groups:
            group_y = statistics.mean(w.cy for w in group)
            if abs(word.cy - group_y) <= y_tolerance:
                group.append(word)
                break
        else:
            groups.append([word])

    return [
        sorted(group, key=lambda w: w.x0)
        for group in sorted(groups, key=lambda g: statistics.mean(w.cy for w in g))
    ]


def read_words(page: pymupdf.Page) -> list[Word]:
    raw_words = page.get_text("words")
    words: list[Word] = []

    for item in raw_words:
        words.append(
            Word(
                x0=float(item[0]),
                y0=float(item[1]),
                x1=float(item[2]),
                y1=float(item[3]),
                text=str(item[4]).strip(),
                block=int(item[5]),
                line=int(item[6]),
                word=int(item[7]),
            )
        )

    return [w for w in words if w.text]


def parse_header(first_page_words: Sequence[Word]) -> Header:
    text = " ".join(w.text for w in first_page_words)
    header = Header()

    invoice = INVOICE_RE.search(text)
    if invoice:
        header.invoice_no = invoice.group(0)

    date = DATE_RE.search(text)
    if date:
        header.create_date = date.group(0)

    t_number = TNO_RE.search(text)
    if t_number:
        header.t_number = t_number.group(0)

    customer_candidates = [w for w in first_page_words if CUSTOMER_RE.fullmatch(w.text)]
    if customer_candidates:
        # 得意先コードは通常ページ左上付近にあるため、その位置を優先する。
        customer_candidates.sort(
            key=lambda w: (
                0 if w.x0 < 160 and w.y0 < 80 else 1,
                w.y0,
                w.x0,
            )
        )
        header.customer_code = customer_candidates[0].text

    return header


def find_detail_blocks(page_no: int, words: Sequence[Word]) -> list[DetailBlock]:
    """
    請求明細番号（Fで始まる11文字程度の番号）をアンカーとして1明細を切り出す。

    通常明細にはGで始まる貴社整理番号があるが、放棄手数料明細などではG番号が
    存在しない場合があるため、全明細に存在する請求明細番号を基準にする。
    """
    anchors = sorted(
        [w for w in words if DETAIL_NO_RE.fullmatch(w.text) and w.x0 > 450],
        key=lambda w: w.y0,
    )

    blocks: list[DetailBlock] = []

    for index, anchor in enumerate(anchors):
        if index == 0:
            top = max(0.0, anchor.y0 - 38.0)
        else:
            top = (anchors[index - 1].y0 + anchor.y0) / 2

        if index == len(anchors) - 1:
            bottom = anchor.y0 + 45.0
        else:
            bottom = (anchor.y0 + anchors[index + 1].y0) / 2

        block_words = [
            w for w in words
            if top <= w.cy <= bottom
        ]

        blocks.append(
            DetailBlock(
                page=page_no,
                top=top,
                bottom=bottom,
                words=block_words,
            )
        )

    return blocks


def parse_top_line(line: Sequence[Word]) -> tuple[str, str, str, str, str, str]:
    """明細1行目から、国名・識別・登録番号・年次・納付期限・分割比率を抽出する。"""
    words = sorted(line, key=lambda w: w.x0)
    country = ""
    kind = ""
    register_no = ""
    year = ""
    due_date = ""
    ratio = ""

    kind_index = None
    for index, word in enumerate(words):
        if word.text in {"PATENT", "DESIGN", "UTILITY", "TRADEMARK"}:
            kind = word.text
            kind_index = index
            break

    if kind_index is not None:
        country = " ".join(w.text for w in words[:kind_index]).strip()

    th_index = None
    for index, word in enumerate(words):
        if word.text == "TH":
            th_index = index
            if index > 0 and re.fullmatch(r"\d+", words[index - 1].text):
                year = f"{words[index - 1].text} TH"
            else:
                year = "TH"
            break

    if kind_index is not None:
        end = th_index - 1 if th_index is not None and th_index > 0 else None
        register_words = words[kind_index + 1:end]
        register_words = [w for w in register_words if w.x0 < 310]
        register_no = " ".join(w.text for w in register_words).strip()

    for index, word in enumerate(words):
        if MONTH_RE.fullmatch(word.text) and index + 2 < len(words):
            due_date = f"{words[index].text} {words[index + 1].text} {words[index + 2].text}"
            break

    for word in words:
        if re.fullmatch(r"\d+/\d+", word.text):
            ratio = word.text
            break

    return country, kind, register_no, year, due_date, ratio


def extract_joint_applicants(
    lines: Sequence[Sequence[Word]],
    top_line: Sequence[Word] | None,
    g_words: Sequence[Word],
    f_words: Sequence[Word],
) -> tuple[str, str]:
    """共願先1・共願先2を抽出する。存在しなければ空欄。"""
    if not f_words:
        return "", ""

    f_y = f_words[0].cy
    top_y = statistics.mean(w.cy for w in top_line) if top_line else -1.0
    g_y = g_words[0].cy if g_words else None

    candidates: list[list[Word]] = []

    for line in lines:
        line_y = statistics.mean(w.cy for w in line)

        if line_y <= top_y + 5:
            continue

        if line_y > f_y + 2:
            continue

        if g_y is not None and abs(line_y - g_y) <= 5:
            continue

        if any(DETAIL_NO_RE.fullmatch(w.text) or G_REF_RE.fullmatch(w.text) for w in line):
            line = [w for w in line if not DETAIL_NO_RE.fullmatch(w.text) and not G_REF_RE.fullmatch(w.text)]

        text = " ".join(w.text for w in line)
        if not text:
            continue

        if any(mark in text for mark in ["US$", "YEN", "+", "=", "放棄手数料"]):
            continue

        # 共願先は日本語の会社名などで記載される想定。
        if re.search(r"[\u3040-\u30ff\u3400-\u9fff]", text):
            candidates.append([w for w in line if w.x0 < 500])

    if not candidates:
        return "", ""

    line = candidates[0]
    joint1 = " ".join(w.text for w in line if w.x0 < 260).strip()
    joint2 = " ".join(w.text for w in line if 260 <= w.x0 < 500).strip()

    return joint1, joint2


def parse_money_lines(lines: Sequence[Sequence[Word]]) -> tuple[str, str, str, str, str, str]:
    currency = ""
    foreign_amount = ""
    yen_amount = ""
    rate = ""
    fee = ""
    subtotal = ""

    for line in lines:
        words = sorted(line, key=lambda w: w.x0)
        texts = [w.text for w in words]

        # 外貨行: US$ 664.87 + = 664.87 ( 161.00 )
        if any(text.endswith("$") for text in texts):
            for index, text in enumerate(texts):
                if text.endswith("$"):
                    currency = text
                    for next_text in texts[index + 1:]:
                        if is_number(next_text):
                            foreign_amount = next_text
                            break
                    break

            for word in words:
                if is_number(word.text) and 360 <= word.x0 <= 430:
                    rate = word.text
                    break

        # 円貨行: YEN 107,044 + = 107,044 3,600 110,644
        if texts and texts[0] == "YEN":
            nums = [text for text in texts[1:] if is_number(text)]
            if nums:
                yen_amount = nums[0]
            if len(nums) >= 3:
                fee = nums[-2]
                subtotal = nums[-1]
            elif len(nums) == 2:
                fee = nums[0]
                subtotal = nums[1]

        # 放棄手数料行: 放棄手数料 3,600 3,600
        if any("放棄手数料" in text for text in texts):
            nums = [text for text in texts if is_number(text)]
            if len(nums) >= 2:
                fee = nums[-2]
                subtotal = nums[-1]
            elif len(nums) == 1:
                fee = nums[0]
                subtotal = nums[0]

    return currency, foreign_amount, yen_amount, rate, fee, subtotal


def parse_detail(block: DetailBlock) -> Detail:
    lines = group_lines(block.words)
    detail = Detail(page=block.page)

    top_line: list[Word] | None = None
    for line in lines:
        if any(w.text in {"PATENT", "DESIGN", "UTILITY", "TRADEMARK"} for w in line):
            top_line = list(line)
            break

    if top_line:
        (
            detail.country,
            detail.kind,
            detail.register_no,
            detail.year,
            detail.due_date,
            detail.ratio,
        ) = parse_top_line(top_line)

    g_words = sorted(
        [w for w in block.words if G_REF_RE.fullmatch(w.text)],
        key=lambda w: (w.y0, w.x0),
    )
    if g_words:
        detail.customer_ref = g_words[0].text

    f_words = sorted(
        [w for w in block.words if DETAIL_NO_RE.fullmatch(w.text)],
        key=lambda w: (w.y0, w.x0),
    )
    if f_words:
        detail.detail_no = f_words[0].text

    if any("放棄手数料" in w.text for w in block.words):
        detail.memo = "放棄手数料"
    elif g_words:
        g_y = g_words[0].cy
        memo_words = [
            w.text
            for w in block.words
            if abs(w.cy - g_y) <= 4 and 230 <= w.x0 <= 330
        ]
        detail.memo = " ".join(memo_words).strip()

    detail.joint1, detail.joint2 = extract_joint_applicants(
        lines=lines,
        top_line=top_line,
        g_words=g_words,
        f_words=f_words,
    )

    (
        detail.currency,
        detail.foreign_amount,
        detail.yen_amount,
        detail.rate,
        detail.fee,
        detail.subtotal,
    ) = parse_money_lines(lines)

    return detail


def output_path_for(pdf_path: Path, header: Header, out_dir: Path | None) -> Path:
    base_dir = out_dir if out_dir else pdf_path.parent
    invoice_no = header.invoice_no if header.invoice_no else "invoice"
    return base_dir / f"{invoice_no}_full_extract.csv"


def write_csv(path: Path, header: Header, details: Sequence[Detail]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_COLUMNS)
        for detail in details:
            writer.writerow(detail.to_csv_row(header))


def validate_totals(details: Sequence[Detail]) -> tuple[int, int, int]:
    yen_total = sum(to_int_amount(d.yen_amount) for d in details)
    fee_total = sum(to_int_amount(d.fee) for d in details)
    subtotal_total = sum(to_int_amount(d.subtotal) for d in details)
    return yen_total, fee_total, subtotal_total


def extract_pdf(pdf_path: Path, out_dir: Path | None = None) -> Path:
    doc = pymupdf.open(pdf_path)

    try:
        all_details: list[Detail] = []

        first_page_words = read_words(doc[0])
        header = parse_header(first_page_words)

        for page_index in range(doc.page_count):
            page_no = page_index + 1
            page_words = read_words(doc[page_index])
            blocks = find_detail_blocks(page_no, page_words)

            print(f"Page {page_no:>2}: {len(blocks):>3} detail(s)")

            for block in blocks:
                all_details.append(parse_detail(block))

        out_path = output_path_for(pdf_path, header, out_dir)
        write_csv(out_path, header, all_details)

        yen_total, fee_total, subtotal_total = validate_totals(all_details)

        print("-" * 60)
        print(f"Invoice     : {header.invoice_no or '(not found)'}")
        print(f"Customer    : {header.customer_code or '(not found)'}")
        print(f"Date        : {header.create_date or '(not found)'}")
        print(f"T Number    : {header.t_number or '(not found)'}")
        print(f"Details     : {len(all_details)}")
        print(f"YEN total   : {yen_total:,}")
        print(f"Fee total   : {fee_total:,}")
        print(f"Subtotal    : {subtotal_total:,}")
        print(f"CSV saved   : {out_path}")

        return out_path

    finally:
        doc.close()


def expand_inputs(patterns: Iterable[str]) -> list[Path]:
    files: list[Path] = []

    for pattern in patterns:
        matched = glob.glob(pattern)
        if matched:
            files.extend(Path(p) for p in matched)
        else:
            files.append(Path(pattern))

    unique: list[Path] = []
    seen: set[Path] = set()

    for path in files:
        resolved = path.resolve()
        if resolved not in seen:
            unique.append(path)
            seen.add(resolved)

    return unique


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="NGB外国年金請求書PDFをCSV化します。",
    )
    parser.add_argument(
        "pdf",
        nargs="+",
        help="入力PDF。ワイルドカードも指定できます。例: *.pdf",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="CSVの出力先フォルダ。省略時はPDFと同じフォルダです。",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)
    pdf_paths = expand_inputs(args.pdf)

    if not pdf_paths:
        print("入力PDFがありません。", file=sys.stderr)
        return 1

    exit_code = 0

    for pdf_path in pdf_paths:
        print("=" * 60)
        print(f"Processing: {pdf_path}")

        if not pdf_path.exists():
            print(f"PDFが見つかりません: {pdf_path}", file=sys.stderr)
            exit_code = 1
            continue

        try:
            extract_pdf(pdf_path, args.out_dir)
        except Exception as exc:
            print(f"エラー: {pdf_path}: {exc}", file=sys.stderr)
            exit_code = 1

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
