"""
Data models for NGB pension invoice parser.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


# ==========================================================
# Header
# ==========================================================

@dataclass(slots=True)
class Header:

    invoice_no: str = ""
    customer_code: str = ""
    create_date: str = ""
    t_number: str = ""


# ==========================================================
# Word
# ==========================================================

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


# ==========================================================
# Detail Block
# ==========================================================

@dataclass(slots=True)
class DetailBlock:

    page: int

    top: float
    bottom: float

    words: List[Word] = field(default_factory=list)


# ==========================================================
# Invoice Detail
# ==========================================================

@dataclass(slots=True)
class Detail:

    page: int = 0

    country: str = ""
    kind: str = ""

    register_no: str = ""

    year: str = ""

    due_date: str = ""

    ratio: str = ""

    ngb_ref: str = ""

    customer_ref: str = ""

    memo: str = ""

    joint1: str = ""
    joint2: str = ""

    detail_no: str = ""

    currency: str = ""

    foreign_amount: str = ""

    yen_amount: str = ""

    rate: str = ""

    fee: str = ""

    subtotal: str = ""

    def to_csv_row(self, header: Header):

        return [

            header.invoice_no,
            header.customer_code,
            header.create_date,

            self.page,

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