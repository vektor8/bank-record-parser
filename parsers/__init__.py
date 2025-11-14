import importlib
import inspect
import os
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Type


class Transaction:
    """Transaction container using explicit keyword-only fields."""

    def __init__(
        self,
        *,
        date: Optional[str] = None,
        details: Optional[str] = None,
        rata: Optional[str] = None,
        store: Optional[str] = None,
        transaction_nr: Optional[str] = None,
        total_transaction: Optional[float] = None,
        amount: Optional[float] = None,
        vendor: Optional[str] = None,
        ref: Optional[str] = None,
        number: Optional[str] = None,
        installment: Optional[int] = None,
        installment_count: Optional[int] = None,
        transaction_total: Optional[float] = None,
        header_date: Optional[str] = None,
        sign: Optional[str] = None,
    ):
        self.date = date
        self.details = details
        self.rata = rata
        self.store = store
        self.transaction_nr = transaction_nr
        self.total_transaction = total_transaction
        self.amount = amount
        self.vendor = vendor
        self.ref = ref
        self.number = number
        self.installment = installment
        self.installment_count = installment_count
        self.transaction_total = transaction_total
        self.header_date = header_date
        self.sign = sign


class BaseParser(ABC):
    """Abstract base class for all PDF parsers

    Parsers should implement:
    - parse_pdf(pdf_path) -> List[Transaction]
    - get_name() -> str
    - get_description() -> str
    - validate_pdf(pdf_path) -> bool
    """

    @abstractmethod
    def parse_pdf(self, pdf_path: str) -> List[Transaction]:
        """Parse PDF file and return list of transactions"""

    @abstractmethod
    def get_name(self) -> str:
        """Return display name for this parser"""

    @abstractmethod
    def get_description(self) -> str:
        """Return description of what this parser does"""

    @abstractmethod
    def validate_pdf(self, pdf_path: str) -> bool:
        """Check if PDF file matches this parser's expected format"""

    def get_columns(self, language: str = "en"):
        """Return a list of (key, header_label) pairs describing columns the parser emits.

        Parsers may override to declare their own column order and names. Default
        returns the canonical columns used by the app.
        """
        # import lazily to avoid circular imports
        from lib.translations import get_translation

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


class ParserRegistry:
    """Registry for managing available parsers"""

    def __init__(self):
        self._parsers: Dict[str, Type[BaseParser]] = {}
        self._load_parsers()

    def _load_parsers(self):
        """Auto-discover parser classes in the parsers directory"""
        parsers_dir = os.path.dirname(__file__)

        for filename in os.listdir(parsers_dir):
            if filename.endswith(".py") and filename != "__init__.py":
                module_name = filename[:-3]
                try:
                    module = importlib.import_module(f"parsers.{module_name}")

                    # Find all classes that inherit from BaseParser
                    for name, obj in inspect.getmembers(module, inspect.isclass):
                        if (
                            issubclass(obj, BaseParser)
                            and obj != BaseParser
                            and not inspect.isabstract(obj)
                        ):

                            # Create instance to get name
                            try:
                                instance = obj()
                                parser_name = instance.get_name()
                                self._parsers[parser_name] = obj
                            except Exception as e:
                                print(
                                    f"Warning: Could not instantiate parser {name}: {e}"
                                )

                except Exception as e:
                    print(f"Warning: Could not load parser from {filename}: {e}")

    def get_parsers(self) -> Dict[str, Type[BaseParser]]:
        """Get all available parsers"""
        return self._parsers.copy()

    def get_parser(self, name: str) -> Type[BaseParser]:
        """Get a specific parser by name"""
        return self._parsers.get(name)

    def create_parser(self, name: str) -> BaseParser:
        """Create an instance of a parser by name"""
        parser_class = self._parsers.get(name)
        if parser_class:
            return parser_class()
        raise ValueError(f"Parser '{name}' not found")

    def auto_detect_parser(self, pdf_path: str) -> Optional[str]:
        """Auto-detect which parser should be used for a PDF"""
        for name, parser_class in self._parsers.items():
            try:
                instance = parser_class()
                if instance.validate_pdf(pdf_path):
                    return name
            except Exception:
                continue
        return None


# Global registry instance
registry = ParserRegistry()
