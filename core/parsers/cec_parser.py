import re
from typing import List

from core.utils import pdf_to_text

from . import BaseParser, Transaction

DATE_REGEX = r"\d{2}[./-]\d{2}[./-]\d{4}"
RATA_REGEX = re.compile(r"Rata (\d*) din (\d*)")
TOTAL_TRANZACTIE_REGEX = re.compile(r"\s+comerciant\s+([\d\.,]+)\s+RON")

AMOUNT_RE = r"[\d\.,]+"

# Heuristic block regex similar to the one used in main2.py
BLOCK_RE = re.compile(
    f"(?P<header_date>{DATE_REGEX})"  # first date (sometimes present)
    r"\s*\n\s*\n?"
    rf"(?P<date>{DATE_REGEX})"  # transaction date (the one we want)
    r"\s*\n"
    r"(?P<info>.*?)"  # details (non-greedy, can include newlines)
    r"\n"
    r"(?P<ref>\d+?)"  # reference token (non-space)
    r"\s*\n"
    r"(?P<number>[A-Z]{2,}\d+)\s*(?P<sign>[-+])?\s*"
    f"(?P<amount>{AMOUNT_RE})",
    re.DOTALL | re.VERBOSE,
)


class CecParser(BaseParser):
    """Improved CEC parser adapted from main2.py's parser logic."""

    def get_name(self) -> str:
        return "CEC Parser"

    def get_description(self) -> str:
        return "Heuristic parser for CEC Bank statements."

    def validate_pdf(self, pdf_path: str) -> bool:
        try:
            raw = pdf_to_text(pdf_path)
            content = raw.get("content", "") or ""
            content_upper = content.upper()
            cec_indicators = ["CEC", "CASA DE ECONOMII", "EXTRAS DE CONT", "RON"]
            matches = sum(
                1 for indicator in cec_indicators if indicator in content_upper
            )
            date_matches = len
            (re.findall(DATE_REGEX, content))
            return matches >= 2 and date_matches > 0
        except Exception:
            return False

    def parse_pdf(self, pdf_path: str) -> List[Transaction]:
        content = pdf_to_text(pdf_path)
        return self.parse_text(content)

    def get_columns(self, language: str = "en"):
        # columns: key, header label
        from core.translations import get_translation

        return [
            ("date", get_translation("data", language)),
            ("details", get_translation("details", language)),
            ("installment", get_translation("rate_nr", language)),
            ("installment_count", get_translation("num_rates", language)),
            ("store", get_translation("store", language)),
            ("category", get_translation("category", language)),
            ("transaction_nr", get_translation("transaction_nr", language)),
            ("total_transaction", get_translation("total_transaction", language)),
            ("amount", get_translation("amount_to_return", language)),
        ]

    @staticmethod
    def __normalize_amount(s: str) -> float:
        return float(s.replace(",", "").replace(".", "").strip()) / 100.0

    def parse_text(self, text: str) -> List[Transaction]:
        results: List[Transaction] = []
        for m in BLOCK_RE.finditer(text):
            header_date = m.group("header_date")
            date = m.group("date")
            info = m.group("info").strip()
            ref = m.group("ref").strip()
            number = m.group("number")
            sign = m.group("sign")
            amount = self.__normalize_amount(m.group("amount"))

            installment_match = RATA_REGEX.search(info)
            if installment_match:
                installment = int(installment_match.group(1))
                installment_count = int(installment_match.group(2))
            else:
                installment = None
                installment_count = None

            transaction_total = 0.0
            total_match = TOTAL_TRANZACTIE_REGEX.search(info)
            if total_match:
                transaction_total = self.__normalize_amount(total_match.group(1))

            # Extract merchant/vendor name
            try:
                payload_after_date = re.split(DATE_REGEX, info)[1].strip()
                payee = " ".join(
                    filter(
                        lambda token: not re.match(r"\d", token),
                        payload_after_date.split(),
                    )
                )
            except Exception:
                payee = ""

            # Build canonical 'data' list expected by downstream code
            # We'll use the same column order used in main.py: DATA_TRANZACTIEI, DETALII, RATA, MAGAZIN, NR_TRANZACTIE, TOTAL_TRANZACTIE, SUMA
            # Some fields are approximated from parsed values
            # Provide explicit named fields to Transaction
            results.append(
                Transaction(
                    date=date,
                    details=info,
                    rata=(
                        f"Rata {installment} din {installment_count}"
                        if installment_count
                        else ""
                    ),
                    store=payee,
                    transaction_nr=number,
                    total_transaction=(
                        transaction_total if sign == "+" else -transaction_total
                    ),
                    amount=(amount if sign == "+" else -amount),
                    vendor=payee,
                    ref=ref,
                    number=number,
                    transaction_total=transaction_total,
                    header_date=header_date,
                    installment=installment,
                    installment_count=installment_count,
                    sign=sign,
                )
            )

        return results
