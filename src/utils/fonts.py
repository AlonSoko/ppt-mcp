"""The names a font file answers to, read with the standard library.

PowerPoint for Mac carries the Office fonts inside its own bundle, where the
system's font list does not see them, so the names are read from the files.
"""

import os
import struct

# The name table's family, full name and typographic family records.
_NAME_IDS = (1, 4, 16)


def font_names(data: bytes) -> set:
    """Every family and full name in a TrueType or OpenType file, or a
    collection of them. An unreadable file answers an empty set."""
    names = set()
    try:
        if data[:4] == b"ttcf":
            count = struct.unpack(">I", data[8:12])[0]
            offsets = struct.unpack(f">{count}I", data[12:12 + 4 * count])
        else:
            offsets = (0,)
        for offset in offsets:
            names |= _names_at(data, offset)
    except (struct.error, IndexError, UnicodeDecodeError):
        pass
    return names


def _names_at(data, offset):
    tables = struct.unpack(">H", data[offset + 4:offset + 6])[0]
    for i in range(tables):
        entry = offset + 12 + 16 * i
        tag = data[entry:entry + 4]
        if tag != b"name":
            continue
        start = struct.unpack(">I", data[entry + 8:entry + 12])[0]
        count, strings = struct.unpack(">HH", data[start + 2:start + 6])
        names = set()
        for n in range(count):
            record = start + 6 + 12 * n
            platform, _, _, name_id, length, at = struct.unpack(
                ">HHHHHH", data[record:record + 12])
            if name_id not in _NAME_IDS:
                continue
            raw = data[start + strings + at:start + strings + at + length]
            text = raw.decode("utf-16-be") if platform in (0, 3) else raw.decode("mac_roman")
            if text.strip():
                names.add(text.strip())
        return names
    return set()


def names_in_folder(folder) -> set:
    """The names of every font file in a folder, not its subfolders."""
    names = set()
    try:
        entries = os.listdir(folder)
    except OSError:
        return names
    for entry in entries:
        if not entry.lower().endswith((".ttf", ".otf", ".ttc", ".otc")):
            continue
        try:
            with open(os.path.join(folder, entry), "rb") as handle:
                names |= font_names(handle.read())
        except OSError:
            continue
    return names
