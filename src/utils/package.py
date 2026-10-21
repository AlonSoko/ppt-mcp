"""Facts read out of a saved .pptx, for what PowerPoint will not say.

PowerPoint gives no way to ask a picture for its own proportions short of
resetting its scale, which changes the deck. The saved file holds them: each
picture's image is in the package, and its crop is in the slide. Reading
them needs nothing past the standard library.
"""

import posixpath
import struct
import zipfile
from xml.etree import ElementTree

_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_EMBED = "{%s}embed" % _NS["r"]
_ID = "{%s}id" % _NS["r"]


def image_size(data: bytes):
    """An image's width and height in pixels, or None.

    PNG, JPEG, GIF, BMP and WebP. A JPEG its camera marked as turned is left
    out, since whether PowerPoint turns it as well is not known here.
    """
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data[:2] == b"BM" and len(data) >= 26:
        width, height = struct.unpack("<ii", data[18:26])
        return abs(width), abs(height)
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP" and len(data) >= 30:
        kind = data[12:16]
        if kind == b"VP8 ":
            width, height = struct.unpack("<HH", data[26:30])
            return width & 0x3FFF, height & 0x3FFF
        if kind == b"VP8L":
            bits = int.from_bytes(data[21:25], "little")
            return (bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1
        if kind == b"VP8X":
            return (int.from_bytes(data[24:27], "little") + 1,
                    int.from_bytes(data[27:30], "little") + 1)
        return None
    if data[:2] == b"\xff\xd8":
        return _jpeg_size(data)
    return None


def _jpeg_size(data):
    i, size, turned = 2, None, False
    while i + 4 <= len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        length = struct.unpack(">H", data[i + 2:i + 4])[0]
        segment = data[i + 4:i + 2 + length]
        if marker == 0xE1 and segment[:6] == b"Exif\x00\x00":
            turned = _exif_turned(segment[6:])
        elif 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC) and len(segment) >= 5:
            height, width = struct.unpack(">HH", segment[1:5])
            size = (width, height)
            break
        i += 2 + length
    return None if turned else size


def _exif_turned(tiff):
    """Whether an EXIF block says the picture is turned a quarter."""
    if len(tiff) < 8 or tiff[:2] not in (b"II", b"MM"):
        return False
    order = "<" if tiff[:2] == b"II" else ">"
    offset = struct.unpack(order + "I", tiff[4:8])[0]
    if offset + 2 > len(tiff):
        return False
    count = struct.unpack(order + "H", tiff[offset:offset + 2])[0]
    for n in range(count):
        entry = offset + 2 + 12 * n
        if entry + 12 > len(tiff):
            break
        tag, = struct.unpack(order + "H", tiff[entry:entry + 2])
        if tag == 0x0112:
            value, = struct.unpack(order + "H", tiff[entry + 8:entry + 10])
            return value >= 5
    return False


def _rels(archive, part):
    """A part's relationships, id to the target's path in the package."""
    folder, name = posixpath.split(part)
    path = posixpath.join(folder, "_rels", name + ".rels")
    if path not in archive.namelist():
        return {}
    root = ElementTree.fromstring(archive.read(path))
    return {
        rel.get("Id"): posixpath.normpath(posixpath.join(folder, rel.get("Target")))
        for rel in root.findall("rel:Relationship", _NS)
        if rel.get("TargetMode") != "External"
    }


def picture_aspects(path, slide_indices):
    """Each picture's own proportions, width over height, as cropped.

    ``{slide_index: {shape name: aspect}}``. A picture whose image cannot be
    measured, one filled as a tile, and names used twice on a slide are left
    out, so an absent entry always means unknown.
    """
    found = {}
    with zipfile.ZipFile(path) as archive:
        presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        targets = _rels(archive, "ppt/presentation.xml")
        order = [targets.get(s.get(_ID))
                 for s in presentation.findall("p:sldIdLst/p:sldId", _NS)]
        for index in slide_indices:
            if not 1 <= index <= len(order) or not order[index - 1]:
                continue
            part = order[index - 1]
            slide = ElementTree.fromstring(archive.read(part))
            media = _rels(archive, part)
            aspects, seen = {}, set()
            for pic in slide.iter("{%s}pic" % _NS["p"]):
                props = pic.find("p:nvPicPr/p:cNvPr", _NS)
                fill = pic.find("p:blipFill", _NS)
                if props is None or fill is None:
                    continue
                name = props.get("name")
                if name in seen:
                    aspects.pop(name, None)
                    continue
                seen.add(name)
                blip = fill.find("a:blip", _NS)
                if blip is None or fill.find("a:tile", _NS) is not None:
                    continue
                image = media.get(blip.get(_EMBED))
                if not image or image not in archive.namelist():
                    continue
                size = image_size(archive.read(image))
                if not size or not size[0] or not size[1]:
                    continue
                crop = fill.find("a:srcRect", _NS)
                cut = {side: 0.0 for side in "ltrb"}
                if crop is not None:
                    for side in "ltrb":
                        cut[side] = int(crop.get(side, "0")) / 100000
                width = size[0] * (1 - cut["l"] - cut["r"])
                height = size[1] * (1 - cut["t"] - cut["b"])
                if width > 0 and height > 0:
                    aspects[name] = width / height
            found[index] = aspects
    return found
