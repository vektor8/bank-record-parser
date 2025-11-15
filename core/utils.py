from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple

import pypdf


def pdf_to_text(pdf_path: Path, password: str = None) -> str:
    reader = pypdf.PdfReader(pdf_path)
    if reader.is_encrypted:
        if password == None:
            raise ValueError("Encrypted PDF. Send the password as well")
        else:
            reader.decrypt(password)
    return "\n".join([page.extract_text() for page in reader.pages])


def decrypt_pdf(input_path: Path, output_path: Path, password: str) -> None:
    with open(input_path, "rb") as input_file, open(output_path, "wb") as output_file:
        reader = pypdf.PdfReader(input_file)
        reader.decrypt(password)

        writer = pypdf.PdfWriter()

        for i in range(len(reader.pages)):
            writer.add_page(reader.pages[i])

        writer.write(output_file)


def load_rules(path: str) -> List[Tuple[str, str]]:
    """Load rules from CSV or TXT file"""
    loaded: List[Tuple[str, str]] = []
    with open(path, "r") as f:
        for line in f.readlines():
            elements = line.split(",")
            if len(elements) < 2:
                raise ValueError("Bad rules file")
            loaded.append((elements[0], elements[1]))
    return loaded
