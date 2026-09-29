"""Pure text normalization. Identifiers never pass through float conversion."""
import math
import re
import unicodedata


FIELDS = {"name": "Nom et prénom", "ddp": "DDP / DDR", "cin": "CIN", "rib": "RIB"}


def text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def identifier(value):
    value = text(value).strip()
    return re.sub(r"^([0-9]+)\.0+$", r"\1", value)


def name(value):
    value = "".join(c for c in unicodedata.normalize("NFKD", text(value)) if not unicodedata.combining(c))
    return " ".join(value.upper().split())


def cin(value):
    return "".join(identifier(value).upper().split())


def rib(value):
    return "".join(identifier(value).split())


def build_rib(bank, city, account, key):
    parts = [rib(v) for v in (bank, city, account, key)]
    if not all(parts):
        return ""
    return "".join(p.zfill(width) if p.isascii() and p.isdigit() else p
                   for p, width in zip(parts, (3, 3, 16, 2)))


def normalize(field, value):
    return {"name": name, "ddp": identifier, "cin": cin, "rib": rib}[field](value)
