"""
Compatibility shim: exposes a zxingcpp.read_barcodes() API backed by pyzbar.
Only the attributes used by ticket_processor.py are implemented:
  result.valid  -> bool
  result.text   -> str
"""
from pyzbar.pyzbar import decode as _pyzbar_decode


class _Result:
    def __init__(self, data: str):
        self.valid = True
        self.text = data


def read_barcodes(image):
    """Accept a PIL image and return a list of _Result objects."""
    decoded = _pyzbar_decode(image)
    results = []
    for item in decoded:
        try:
            text = item.data.decode("utf-8")
        except (UnicodeDecodeError, AttributeError):
            text = str(item.data)
        results.append(_Result(text))
    return results
