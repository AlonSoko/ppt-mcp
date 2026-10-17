"""Layout checks: what is wrong with a slide, measured rather than looked at.

A slide preview shows a model what it built, but only a model that can read
images, and only when it thinks to ask. Most of what goes wrong on a generated
slide is geometry, and geometry can be answered with numbers: text that runs
into another shape, text that does not fit its box, a shape hanging off the
slide, three cards whose gaps are 24, 24 and 31 points, a title that sits three
points lower than the one on the slide before it.

This module answers those questions in two places. ``ppt_check_layout`` checks
one slide or the whole deck on request. The tools that create, move, resize or
retext a shape run the per-shape checks on the shape they just touched and put
what they find in their own result, so a problem is reported by the call that
caused it, while the caller still knows why it made that call.

The arithmetic is kept apart from the reading, so it can be tested without
PowerPoint and so both platforms answer the same way. ``_slide_facts_impl``
reads a slide into plain dicts; everything after that is pure.
"""

import json
import logging
import math
from typing import Optional

from pydantic import BaseModel, ConfigDict, Field

from backend import ppt
from utils.color import int_to_hex
from utils.navigation import goto_slide
from utils.offload import run_offloaded
from ppt_com.constants import (
    msoChart, msoEmbeddedOLEObject, msoFillSolid, msoGroup, msoLine,
    msoLinkedOLEObject, msoLinkedPicture, msoMedia, msoPicture,
    msoPlaceholder, msoSmartArt, msoTable, msoTextBox, msoTrue,
    ppAlignCenter, ppAlignRight,
    ppPlaceholderCenterTitle, ppPlaceholderDate, ppPlaceholderFooter,
    ppPlaceholderHeader, ppPlaceholderSlideNumber, ppPlaceholderTitle,
    ppPlaceholderVerticalTitle,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Thresholds, in points unless they say otherwise
# ---------------------------------------------------------------------------

# A shape may sit right on an edge; a fraction of a point past it is rounding.
# The same allowance ppt_check_typography gives an off-slide edge.
_EDGE_TOLERANCE = 0.5

# Edges this close are one edge. PowerPoint stores positions in EMU and reads
# them back in points, so two shapes placed at the same number can differ in
# the second decimal.
_ALIGNED = 0.5

# Edges this close, and not closer, were meant to line up and do not. Further
# apart than this is a decision rather than a slip.
_NEAR_MISS = 4.0

# Text has to cover this much of another shape, in both directions, before it
# counts as running into it. Less is a glyph's side bearing touching a border.
_MIN_OVERLAP = 2.0

# How far apart the gaps in a row of like shapes can be before it shows.
_GAP_TOLERANCE = 2.0

# Shapes count as the same kind of thing, for spacing, when their widths and
# heights are within this fraction of each other.
_SIMILAR_SIZE = 0.15

# Two shapes are in the same row (or column) when they share at least this
# fraction of the smaller one's height (or width).
_SHARED_SPAN = 0.5

# A shape covering this much of the slide is a backdrop, not part of the
# layout. It is not lined up against anything and it does not make the shapes
# on it its children.
_BACKDROP_SHARE = 0.9

# A fill more see-through than this mixes with whatever is under it, so its
# own colour is not the colour behind the text.
_OPAQUE_ENOUGH = 0.3

# WCAG 2.x contrast minimums. Large text is 18 pt, or 14 pt bold.
_CONTRAST_TEXT = 4.5
_CONTRAST_LARGE = 3.0
_LARGE_TEXT = 18.0
_LARGE_BOLD_TEXT = 14.0

# A deck's titles are compared once at least this many slides have one.
_MIN_TITLED_SLIDES = 3

# Title sizes this close are the same size.
_TITLE_SIZE_TOLERANCE = 1.0

# Without a title placeholder, a slide's title is its largest text in the top
# quarter of the slide, at this size or more.
_TITLE_ZONE = 0.25
_MIN_TITLE_SIZE = 24.0

# PowerPoint's default text frame insets, for a frame whose margins could not
# be read.
_DEFAULT_MARGINS = {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}

# Placeholders the layout owns. They are not part of the slide's design, so
# they are not lined up, spaced, or reported as left empty.
_LAYOUT_PLACEHOLDERS = frozenset({
    ppPlaceholderSlideNumber, ppPlaceholderHeader,
    ppPlaceholderFooter, ppPlaceholderDate,
})

_TITLE_PLACEHOLDERS = frozenset({
    ppPlaceholderTitle, ppPlaceholderCenterTitle, ppPlaceholderVerticalTitle,
})

# Shapes that draw their whole box whatever their fill says.
_OPAQUE_TYPES = frozenset({
    msoPicture, msoLinkedPicture, msoTable, msoChart, msoGroup, msoMedia,
    msoSmartArt, msoEmbeddedOLEObject, msoLinkedOLEObject,
})

# The ones whose colour cannot be known, so text on them is not given a
# contrast verdict.
_UNKNOWN_BACKDROP_TYPES = frozenset({
    msoPicture, msoLinkedPicture, msoChart, msoGroup, msoMedia, msoSmartArt,
    msoEmbeddedOLEObject, msoLinkedOLEObject,
})

# The finding types the write tools report on the shape they touched. The rest
# describe a slide or a deck rather than one shape, and a slide being built is
# allowed to be unfinished, so those wait for ppt_check_layout.
SHAPE_FINDINGS = frozenset({
    "off_slide", "overflow", "overlap", "low_contrast", "misaligned",
})

# Reported in this order, worst first.
_ORDER = (
    "off_slide", "overflow", "overlap", "low_contrast", "misaligned",
    "uneven_spacing", "empty_placeholder", "inconsistent_titles",
)


# ---------------------------------------------------------------------------
# Input model
# ---------------------------------------------------------------------------
class CheckLayoutInput(BaseModel):
    """Input for checking the layout of one slide or the whole deck."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: Optional[int] = Field(
        default=None, ge=1,
        description=(
            "1-based slide index to check. Omit to check every slide, which "
            "also compares the slides' titles with each other."
        ),
    )


# ---------------------------------------------------------------------------
# Boxes
# ---------------------------------------------------------------------------
def _pt(value) -> str:
    """A number of points the way a person writes it: 12 pt, 12.5 pt."""
    return f"{round(value, 1):g} pt"


def _box(left, top, width, height):
    return (left, top, left + width, top + height)


def _width(box):
    return box[2] - box[0]


def _height(box):
    return box[3] - box[1]


def _shared(a, b):
    """Width and height of the area two boxes share, zero when they do not."""
    return (
        max(0.0, min(a[2], b[2]) - max(a[0], b[0])),
        max(0.0, min(a[3], b[3]) - max(a[1], b[1])),
    )


def _inside(inner, outer, tolerance=_EDGE_TOLERANCE) -> bool:
    return (
        inner[0] >= outer[0] - tolerance and inner[1] >= outer[1] - tolerance
        and inner[2] <= outer[2] + tolerance and inner[3] <= outer[3] + tolerance
    )


def _turned(shape) -> bool:
    """Whether a shape is rotated off the axes.

    Half a turn leaves the box where it was, so 180 degrees is not turned.
    """
    rotation = (shape.get("rotation") or 0.0) % 180.0
    return min(rotation, 180.0 - rotation) > 0.5


def shape_box(shape):
    """The shape's own box, or the box around it when it is rotated."""
    left, top = shape["left"], shape["top"]
    width, height = shape["width"], shape["height"]
    if not _turned(shape):
        return _box(left, top, width, height)
    angle = math.radians(shape.get("rotation") or 0.0)
    cos, sin = abs(math.cos(angle)), abs(math.sin(angle))
    turned_w = width * cos + height * sin
    turned_h = width * sin + height * cos
    cx, cy = left + width / 2, top + height / 2
    return _box(cx - turned_w / 2, cy - turned_h / 2, turned_w, turned_h)


def text_box(shape):
    """Where the shape's text is drawn, or None when it has none.

    PowerPoint reports this directly, and that is what is used. When it would
    not, the position is worked out from the frame: the margins, the vertical
    anchor and the first paragraph's alignment place a block of the measured
    width and height inside the box. Overflowing text is placed the way
    PowerPoint draws it, past the bottom for a top anchor, past both edges for
    a middle one.
    """
    text = shape.get("text")
    if not text or not text.get("chars"):
        return None
    width, height = text.get("width"), text.get("height")
    if width is None or height is None:
        return None
    if text.get("left") is not None and text.get("top") is not None:
        return _box(text["left"], text["top"], width, height)

    margins = text.get("margins") or _DEFAULT_MARGINS
    inner_left = shape["left"] + margins["left"]
    inner_top = shape["top"] + margins["top"]
    usable_w = shape["width"] - margins["left"] - margins["right"]
    usable_h = shape["height"] - margins["top"] - margins["bottom"]

    alignment = text.get("alignment")
    if alignment == ppAlignCenter:
        left = inner_left + (usable_w - width) / 2
    elif alignment == ppAlignRight:
        left = inner_left + usable_w - width
    else:
        left = inner_left

    anchor = text.get("vertical_anchor")
    if anchor == "middle":
        top = inner_top + (usable_h - height) / 2
    elif anchor == "bottom":
        top = inner_top + usable_h - height
    else:
        top = inner_top
    return _box(left, top, width, height)


def paints_box(shape) -> bool:
    """Whether the shape draws its whole box rather than only its text.

    A filled shape, an outlined one, a picture or a table can be seen edge to
    edge. A text box with neither fill nor line is only its text, so a box
    drawn larger than its text covers nothing.

    PowerPoint for Mac does not say whether a line is drawn, so there the
    shape's kind decides. A text box or placeholder has no line unless one is
    added, and an autoshape has one unless it is taken away.
    """
    if shape["type"] in _OPAQUE_TYPES:
        return True
    fill = shape.get("fill") or {}
    if fill.get("visible"):
        return True
    line = shape.get("line_visible")
    if line is None:
        return shape["type"] not in (msoTextBox, msoPlaceholder)
    return bool(line)


def _is_line(shape) -> bool:
    return shape["type"] == msoLine or bool(shape.get("connector"))


def ink_box(shape):
    """What the shape puts on the slide: its box, its text, or nothing."""
    if not shape.get("visible", True) or _is_line(shape):
        return None
    if paints_box(shape):
        return shape_box(shape)
    return text_box(shape)


def _is_empty_placeholder(shape) -> bool:
    text = shape.get("text")
    return (
        shape.get("placeholder_type") is not None
        and text is not None
        and not text.get("chars")
    )


def _is_backdrop(shape, slide_w, slide_h) -> bool:
    if not slide_w or not slide_h:
        return False
    box = shape_box(shape)
    width, height = _shared(box, (0.0, 0.0, slide_w, slide_h))
    return width * height >= _BACKDROP_SHARE * slide_w * slide_h


def _arranged(shapes, slide_w, slide_h):
    """The shapes that make up the layout: the ones lined up and spaced.

    Not lines, not rotated shapes, not backdrops, not the layout's own footer
    placeholders, and not empty placeholders, which are reported on their own
    and draw nothing once the slide is shown.
    """
    arranged = []
    for shape in shapes:
        if not shape.get("visible", True) or _is_line(shape) or _turned(shape):
            continue
        if shape.get("placeholder_type") in _LAYOUT_PLACEHOLDERS:
            continue
        if _is_empty_placeholder(shape):
            continue
        if shape["width"] <= 0 or shape["height"] <= 0:
            continue
        if _is_backdrop(shape, slide_w, slide_h):
            continue
        arranged.append(shape)
    return arranged


# ---------------------------------------------------------------------------
# Per shape findings
# ---------------------------------------------------------------------------
def _finding(slide_index, kind, shapes, message, fix, severity="error", **extra):
    finding = {
        "slide_index": slide_index,
        "type": kind,
        "severity": severity,
        "shapes": list(shapes),
        "message": message,
        "fix": fix,
    }
    finding.update(extra)
    return finding


def _off_slide(slide_index, shapes, slide_w, slide_h):
    """Shapes, or the text of shapes, that reach past an edge of the slide.

    Measured on what is drawn. A text box whose frame hangs over the edge with
    its text well inside is not wrong, and a box that fits on the slide with
    its text spilling off the bottom is.
    """
    if not slide_w or not slide_h:
        return []
    findings = []
    for shape in shapes:
        if not shape.get("visible", True):
            continue
        box = shape_box(shape) if _is_line(shape) else ink_box(shape)
        if box is None:
            continue
        past = {
            "left": -box[0],
            "top": -box[1],
            "right": box[2] - slide_w,
            "bottom": box[3] - slide_h,
        }
        past = {edge: round(amount, 1) for edge, amount in past.items()
                if amount > _EDGE_TOLERANCE}
        if not past:
            continue
        parts = [f"{_pt(amount)} past the {edge} edge" for edge, amount in past.items()]
        moves = []
        if "left" in past:
            moves.append(f"right by {_pt(past['left'])}")
        if "right" in past:
            moves.append(f"left by {_pt(past['right'])}")
        if "top" in past:
            moves.append(f"down by {_pt(past['top'])}")
        if "bottom" in past:
            moves.append(f"up by {_pt(past['bottom'])}")
        name = shape["name"]
        findings.append(_finding(
            slide_index, "off_slide", [name],
            f"'{name}' runs {' and '.join(parts)} of the slide.",
            f"Move '{name}' {' and '.join(moves)}, or make it smaller.",
            edges=past,
        ))
    return findings


def _overflow(slide_index, shapes):
    """Text that needs more room than its box gives it.

    The same verdict ppt_get_text(measure=True) gives, from the same
    arithmetic, so the two never disagree about a box.
    """
    from ppt_com.text import build_measurement

    findings = []
    for shape in shapes:
        text = shape.get("text")
        if not shape.get("visible", True) or not text or not text.get("chars"):
            continue
        measurement = build_measurement(
            [], text.get("width"), text.get("height"),
            shape["width"], shape["height"], text.get("margins"),
            text.get("word_wrap"), text.get("autofit"), text.get("orientation"),
        )
        if not measurement["overflows"]:
            continue
        name = shape["name"]
        needed = measurement["text_height_pt"]
        room = measurement["usable_height_pt"]
        if text.get("orientation") in ("vertical", "upward", "downward"):
            needed = measurement["text_width_pt"]
            room = measurement["usable_width_pt"]
        elif needed is not None and room is not None and needed <= room + 0.5:
            # Wrapping is off and the overflow is a line too long for the box.
            needed = measurement["text_width_pt"]
            room = measurement["usable_width_pt"]
            findings.append(_finding(
                slide_index, "overflow", [name],
                f"A line of text in '{name}' is {_pt(needed)} long and the "
                f"box has {_pt(room)} for it, so it runs out of the side.",
                f"Make '{name}' {_pt(needed - room)} wider, turn word wrap "
                "on, or shorten the line.",
                needed_pt=needed, available_pt=room,
            ))
            continue
        findings.append(_finding(
            slide_index, "overflow", [name],
            f"The text in '{name}' needs {_pt(needed)} and the box has "
            f"{_pt(room)}, so it spills out of the box.",
            f"Make '{name}' {_pt(needed - room)} taller, cut the text, or "
            "make the font smaller.",
            needed_pt=needed, available_pt=room,
        ))
    return findings


def _overlap(slide_index, shapes):
    """Text that runs into another shape, or into other text.

    Only text is judged. Two filled shapes crossing may be a Venn diagram or a
    badge on a card, and nothing about the slide says which, but text drawn
    over something, or under it, is never meant. Text that sits wholly inside
    another shape is on it, like the words on a card, and is fine.
    """
    findings = []
    seen = set()
    for a in shapes:
        a_text = text_box(a) if a.get("visible", True) else None
        if a_text is None or _turned(a):
            continue
        for b in shapes:
            # Text spilling out of its own box is the overflow check's.
            if b is a or _turned(b):
                continue
            b_ink = ink_box(b)
            if b_ink is None or _inside(a_text, b_ink):
                continue
            width, height = _shared(a_text, b_ink)
            if width <= _MIN_OVERLAP or height <= _MIN_OVERLAP:
                continue
            pair = frozenset((a["name"], b["name"]))
            if pair in seen:
                continue
            seen.add(pair)
            findings.append(_overlap_finding(slide_index, a, a_text, b, b_ink, width, height))
    return findings


def _overlap_finding(slide_index, a, a_text, b, b_ink, width, height):
    a_name, b_name = a["name"], b["name"]
    shared = {"width": round(width, 1), "height": round(height, 1)}

    # Text mostly on a shape underneath it is meant to be on it, like the
    # words on a card, and the problem is that it does not fit. Saying where
    # it runs past is more use than saying the two collide.
    area = _width(a_text) * _height(a_text)
    if (paints_box(b) and b.get("z", 0) < a.get("z", 0)
            and area and width * height >= 0.5 * area):
        past = {
            "left": b_ink[0] - a_text[0],
            "top": b_ink[1] - a_text[1],
            "right": a_text[2] - b_ink[2],
            "bottom": a_text[3] - b_ink[3],
        }
        past = {edge: round(amount, 1) for edge, amount in past.items()
                if amount > _EDGE_TOLERANCE}
        parts = " and ".join(f"{_pt(amount)} past its {edge} edge"
                             for edge, amount in past.items())
        return _finding(
            slide_index, "overlap", [a_name, b_name],
            f"The text of '{a_name}' sits on '{b_name}' and runs {parts}.",
            f"Make '{b_name}' larger by that much, or make the text of "
            f"'{a_name}' shorter or smaller.",
            overlap_pt=shared, past_edges=past,
        )

    # What b puts on the slide is its text when it does not paint its box.
    target = f"'{b_name}'" if paints_box(b) else f"the text of '{b_name}'"
    # The cheaper way apart: along whichever axis they share less of.
    if height <= width:
        clear = height + 4.0
        if (b_ink[1] + b_ink[3]) >= (a_text[1] + a_text[3]):
            fix = (f"Move '{b_name}' down by {_pt(clear)}, or move '{a_name}' "
                   f"up or shorten its text by that much.")
        else:
            fix = (f"Move '{b_name}' up by {_pt(clear)}, or move '{a_name}' "
                   f"down by that much.")
    else:
        clear = width + 4.0
        if (b_ink[0] + b_ink[2]) >= (a_text[0] + a_text[2]):
            fix = (f"Move '{b_name}' right by {_pt(clear)}, or make "
                   f"'{a_name}' narrower by that much.")
        else:
            fix = (f"Move '{b_name}' left by {_pt(clear)}, or move "
                   f"'{a_name}' right by that much.")
    return _finding(
        slide_index, "overlap", [a_name, b_name],
        f"The text of '{a_name}' runs into {target}: they share an area "
        f"{_pt(width)} wide and {_pt(height)} tall.",
        fix,
        overlap_pt=shared,
    )


def _backdrop_colour(shape, shapes, background_hex):
    """The colour the shape's text is drawn on, or None when it cannot be told.

    The shape's own fill when it has a solid one, otherwise the nearest shape
    underneath that holds the whole of the text and paints something, and
    otherwise the slide's background. A picture, a gradient or a see-through
    fill in the way means the colour is not one colour, so there is no answer.
    """
    def solid(fill):
        if not fill or not fill.get("visible"):
            return False, None
        if not fill.get("solid") or (fill.get("transparency") or 0.0) > _OPAQUE_ENOUGH:
            return True, None
        return True, fill.get("color_hex")

    painted, colour = solid(shape.get("fill"))
    if painted:
        return colour
    words = text_box(shape)
    if words is None:
        return None
    below = [s for s in shapes if s.get("z", 0) < shape.get("z", 0)]
    for other in sorted(below, key=lambda s: s.get("z", 0), reverse=True):
        if not other.get("visible", True) or _is_line(other):
            continue
        if not _inside(words, shape_box(other)):
            continue
        if other["type"] in _UNKNOWN_BACKDROP_TYPES:
            return None
        painted, colour = solid(other.get("fill"))
        if painted:
            return colour
    return background_hex


def _low_contrast(slide_index, shapes, background_hex):
    from ppt_com.themes import _contrast_ratio

    findings = []
    for shape in shapes:
        text = shape.get("text")
        if not shape.get("visible", True) or not text or not text.get("chars"):
            continue
        backdrop = _backdrop_colour(shape, shapes, background_hex)
        if not backdrop:
            continue
        worst = None
        for run in text.get("runs") or []:
            colour = run.get("color_hex")
            if not colour:
                continue
            size = run.get("size") or 0.0
            large = size >= _LARGE_TEXT or (run.get("bold") and size >= _LARGE_BOLD_TEXT)
            needed = _CONTRAST_LARGE if large else _CONTRAST_TEXT
            ratio = _contrast_ratio(colour, backdrop)
            if ratio < needed and (worst is None or ratio / needed < worst[0] / worst[1]):
                worst = (ratio, needed, colour)
        if worst is None:
            continue
        ratio, needed, colour = worst
        name = shape["name"]
        dark, light = _contrast_ratio("#000000", backdrop), _contrast_ratio("#FFFFFF", backdrop)
        suggestion = "#000000" if dark >= light else "#FFFFFF"
        findings.append(_finding(
            slide_index, "low_contrast", [name],
            f"Text in '{name}' is {colour} on {backdrop}, a contrast of "
            f"{ratio:.1f}:1 where {needed:g}:1 is the minimum for its size.",
            f"Make the text {suggestion} (contrast "
            f"{max(dark, light):.1f}:1), or change what is behind it.",
            text_color=colour, backdrop_color=backdrop,
            contrast=round(ratio, 2), minimum=needed,
        ))
    return findings


# Edge names per axis, in the order they are preferred when two tie.
_X_EDGES = ("left", "center", "right")
_Y_EDGES = ("top", "middle", "bottom")


def _edges(box, axis):
    if axis == "x":
        return dict(zip(_X_EDGES, (box[0], (box[0] + box[2]) / 2, box[2])))
    return dict(zip(_Y_EDGES, (box[1], (box[1] + box[3]) / 2, box[3])))


def _related(a_box, b_box, axis) -> bool:
    """Shapes stacked in a column are lined up left and right; shapes in a row,
    top and bottom. Shapes that are neither have nothing to line up."""
    if axis == "x":
        shared = _shared(a_box, b_box)[0]
        return shared >= _SHARED_SPAN * min(_width(a_box), _width(b_box))
    shared = _shared(a_box, b_box)[1]
    return shared >= _SHARED_SPAN * min(_height(a_box), _height(b_box))


def _misaligned(slide_index, shapes, slide_w, slide_h, only=None):
    """Edges a few points apart that were meant to be one edge.

    Each pair of shapes in a row is compared top, middle and bottom, and each
    pair in a column left, centre and right. A pair already sharing one of
    those lines is aligned. A pair whose nearest lines are less than
    ``_NEAR_MISS`` apart, but not together, is the slip this reports.

    Of the two, the one to move is the one lined up with fewer other shapes,
    and between equals the one added later. One finding names it with every
    shape it nearly lines up with.
    """
    arranged = _arranged(shapes, slide_w, slide_h)
    boxes = {s["name"]: shape_box(s) for s in arranged}
    findings = []
    for axis in ("x", "y"):
        aligned = {s["name"]: 0 for s in arranged}
        near = []
        for i, a in enumerate(arranged):
            for b in arranged[i + 1:]:
                a_box, b_box = boxes[a["name"]], boxes[b["name"]]
                if not _related(a_box, b_box, axis):
                    continue
                a_edges, b_edges = _edges(a_box, axis), _edges(b_box, axis)
                gaps = {edge: abs(a_edges[edge] - b_edges[edge]) for edge in a_edges}
                closest = min(gaps, key=gaps.get)
                if gaps[closest] <= _ALIGNED:
                    aligned[a["name"]] += 1
                    aligned[b["name"]] += 1
                elif gaps[closest] <= _NEAR_MISS:
                    near.append((a, b, closest))

        grouped = {}
        for a, b, edge in near:
            # The one that stays is lined up with more shapes, or was there
            # first. Stacking order is creation order unless something moved
            # a shape in it, which is as close as a slide comes to a history.
            if (aligned[a["name"]], -a.get("z", 0)) > (aligned[b["name"]], -b.get("z", 0)):
                anchor, mover = a, b
            else:
                anchor, mover = b, a
            grouped.setdefault((mover["name"], edge), (mover, edge, []))[2].append(anchor)

        for (name, edge), (mover, edge, anchors) in grouped.items():
            if only is not None and name not in only and not any(
                anchor["name"] in only for anchor in anchors
            ):
                continue
            findings.append(_misaligned_finding(
                slide_index, axis, mover, edge, anchors, boxes))
    return findings


def _misaligned_finding(slide_index, axis, mover, edge, anchors, boxes):
    name = mover["name"]
    values = [_edges(boxes[a["name"]], axis)[edge] for a in anchors]
    # Where most of the others are; between equals, the nearest.
    mine = _edges(boxes[name], axis)[edge]
    target = max(
        {round(v, 2) for v in values},
        key=lambda v: (sum(1 for x in values if abs(x - v) <= _ALIGNED), -abs(v - mine)),
    )
    delta = mine - target
    # Moving the whole shape by the difference puts that edge on the target,
    # whichever edge it is.
    if axis == "x":
        field, value = "left", round(mover["left"] - delta, 2)
        direction = "right of" if delta > 0 else "left of"
    else:
        field, value = "top", round(mover["top"] - delta, 2)
        direction = "below" if delta > 0 else "above"
    others = ", ".join(f"'{a['name']}'" for a in anchors)
    return _finding(
        slide_index, "misaligned", [name] + [a["name"] for a in anchors],
        f"The {edge} of '{name}' is {_pt(abs(delta))} {direction} the "
        f"{edge} of {others}, close enough to look like a slip.",
        f"Set {field} of '{name}' to {_pt(value)} (ppt_update_shape "
        f"{field}={value:g}), or ppt_align_shapes with align_to='{edge}'.",
        severity="warning", edge=edge, off_by_pt=round(abs(delta), 2),
    )


def _connected(members, joined):
    """Split members into groups, two sharing a group when ``joined`` says so,
    directly or through others."""
    parent = list(range(len(members)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, a in enumerate(members):
        for j in range(i + 1, len(members)):
            if joined(a, members[j]):
                parent[root(i)] = root(j)
    groups = {}
    for i, member in enumerate(members):
        groups.setdefault(root(i), []).append(member)
    return list(groups.values())


def _uneven_spacing(slide_index, shapes, slide_w, slide_h):
    """Rows and columns of like shapes whose gaps differ.

    Three or more shapes of about the same size, side by side in a row or
    stacked in a column, are one group: cards, steps, columns of icons. Their
    gaps are meant to be equal. Shapes sitting inside another shape belong to
    it, so the text on three cards is not a fourth row to space.
    """
    arranged = _arranged(shapes, slide_w, slide_h)
    boxes = {s["name"]: shape_box(s) for s in arranged}

    def area(box):
        return _width(box) * _height(box)

    def belongs_to(s, other):
        # Mostly inside, rather than wholly: the text on a card still belongs
        # to the card when it runs past its edge, which is its own finding,
        # and it is not a fourth card in the row. A text box set to grow with
        # its text can end up with a larger frame than the card it is on, so
        # a frame that only holds text belongs to a shape that paints, at any
        # size.
        mine, theirs = boxes[s["name"]], boxes[other["name"]]
        width, height = _shared(mine, theirs)
        if width * height < 0.5 * area(mine):
            return False
        if paints_box(s) != paints_box(other):
            # Never the other way round: a card is not inside a text frame.
            return paints_box(other)
        return area(theirs) > area(mine)

    outer = [
        s for s in arranged
        if not any(other is not s and belongs_to(s, other) for other in arranged)
    ]

    def similar(a, b):
        aw, ah = _width(boxes[a["name"]]), _height(boxes[a["name"]])
        bw, bh = _width(boxes[b["name"]]), _height(boxes[b["name"]])
        return (abs(aw - bw) <= _SIMILAR_SIZE * max(aw, bw)
                and abs(ah - bh) <= _SIMILAR_SIZE * max(ah, bh))

    findings = []
    # A row shares its height and is spaced along x; a column the other way.
    for axis, direction, start in (("y", "horizontal", 0), ("x", "vertical", 1)):
        def joined(a, b, axis=axis):
            return similar(a, b) and _related(boxes[a["name"]], boxes[b["name"]], axis)

        for members in _connected(outer, joined):
            if len(members) < 3:
                continue
            members.sort(key=lambda s: boxes[s["name"]][start])
            gaps = [
                boxes[b["name"]][start] - boxes[a["name"]][start + 2]
                for a, b in zip(members, members[1:])
            ]
            if min(gaps) <= 0 or max(gaps) - min(gaps) <= _GAP_TOLERANCE:
                continue
            names = [s["name"] for s in members]
            even = sum(gaps) / len(gaps)
            listed = ", ".join(f"'{n}'" for n in names)
            findings.append(_finding(
                slide_index, "uneven_spacing", names,
                f"The gaps between {listed} are "
                f"{', '.join(_pt(g) for g in gaps)}; a "
                f"{'row' if direction == 'horizontal' else 'column'} of like "
                "shapes reads as one group when they are equal.",
                f"ppt_distribute_shapes with these shape names and "
                f"direction='{direction}', or make every gap {_pt(even)}.",
                severity="warning", gaps_pt=[round(g, 1) for g in gaps],
            ))
    return findings


def _empty_placeholders(slide_index, shapes):
    findings = []
    for shape in shapes:
        if not _is_empty_placeholder(shape):
            continue
        if shape.get("placeholder_type") in _LAYOUT_PLACEHOLDERS:
            continue
        name = shape["name"]
        findings.append(_finding(
            slide_index, "empty_placeholder", [name],
            f"'{name}' is an empty placeholder. It draws nothing in the slide "
            "show, but in the editor it shows its prompt text, and it still "
            "takes the space the layout gave it.",
            f"Put text in '{name}', or delete it with ppt_delete_shape.",
            severity="warning",
        ))
    return findings


def slide_findings(slide, slide_w, slide_h, only=None):
    """Every finding for one slide, worst first.

    ``slide`` is one entry of what ``_slide_facts_impl`` returns. With
    ``only``, a set of shape names, the answer is the per shape findings that
    involve one of those shapes, which is what a write tool reports.
    """
    index = slide["slide_index"]
    shapes = slide["shapes"]
    findings = (
        _off_slide(index, shapes, slide_w, slide_h)
        + _overflow(index, shapes)
        + _overlap(index, shapes)
        + _low_contrast(index, shapes, slide.get("background_hex"))
        + _misaligned(index, shapes, slide_w, slide_h, only=only)
    )
    if only is not None:
        return [
            f for f in findings
            if f["type"] in SHAPE_FINDINGS and any(name in only for name in f["shapes"])
        ]
    findings += (
        _uneven_spacing(index, shapes, slide_w, slide_h)
        + _empty_placeholders(index, shapes)
    )
    return _sorted(findings)


def _sorted(findings):
    rank = {kind: i for i, kind in enumerate(_ORDER)}
    return sorted(findings, key=lambda f: (f["slide_index"] or 0, rank.get(f["type"], len(rank))))


# ---------------------------------------------------------------------------
# Deck findings
# ---------------------------------------------------------------------------
def _largest_size(shape):
    sizes = [run.get("size") for run in (shape.get("text") or {}).get("runs") or []
             if run.get("size")]
    return max(sizes) if sizes else None


def find_title(slide, slide_h):
    """The shape that is this slide's title, or None.

    A title placeholder with text when there is one. Generated decks often
    have no placeholders at all, so otherwise it is the largest text in the
    top quarter of the slide, at a title's size. A cover slide's title sits
    lower than that and is left out, which is right: it is meant to differ.
    """
    for shape in slide["shapes"]:
        text = shape.get("text")
        if shape.get("placeholder_type") in _TITLE_PLACEHOLDERS and text and text.get("chars"):
            return shape
    if not slide_h:
        return None
    best = None
    for shape in slide["shapes"]:
        text = shape.get("text")
        if not shape.get("visible", True) or not text or not text.get("chars"):
            continue
        if shape["top"] > _TITLE_ZONE * slide_h:
            continue
        size = _largest_size(shape)
        if size is None or size < _MIN_TITLE_SIZE:
            continue
        if best is None or (size, -shape["top"]) > (_largest_size(best), -best["top"]):
            best = shape
    return best


def _clusters(values, tolerance):
    """Group (slide, value) pairs whose values are within tolerance of the
    group's first value, largest group first."""
    groups = []
    for slide_index, value in sorted(values, key=lambda pair: pair[1]):
        for group in groups:
            if abs(group[0][1] - value) <= tolerance:
                group.append((slide_index, value))
                break
        else:
            groups.append([(slide_index, value)])
    return sorted(groups, key=len, reverse=True)


def _slides(numbers):
    numbers = sorted(numbers)
    return ("slide " if len(numbers) == 1 else "slides ") + ", ".join(str(n) for n in numbers)


def deck_findings(slides, slide_h):
    """Titles that differ between slides that share a layout.

    Flipping from one slide to the next, a title that is a size larger, or a
    few points lower, is the thing that jumps. Slides are compared within their
    layout, so a section header is not held to a content slide's title.
    """
    by_layout = {}
    for slide in slides:
        title = find_title(slide, slide_h)
        if title is None:
            continue
        by_layout.setdefault(slide.get("layout"), []).append((slide["slide_index"], title))

    findings = []
    for layout, titled in by_layout.items():
        if len(titled) < _MIN_TITLED_SLIDES:
            continue
        for label, read, tolerance, field in (
            ("size", _largest_size, _TITLE_SIZE_TOLERANCE, None),
            ("left edge", lambda s: s["left"], _NEAR_MISS, "left"),
            ("top edge", lambda s: s["top"], _NEAR_MISS, "top"),
        ):
            values = [(index, read(shape)) for index, shape in titled if read(shape) is not None]
            groups = _clusters(values, tolerance)
            if len(groups) < 2:
                continue
            usual = groups[0]
            usual_value = sorted(v for _, v in usual)[len(usual) // 2]
            unit = "pt"
            described = "; ".join(
                f"{round(sorted(v for _, v in g)[len(g) // 2], 1):g} {unit} on "
                f"{_slides(i for i, _ in g)}"
                for g in groups
            )
            odd = [i for g in groups[1:] for i, _ in g]
            names = [shape["name"] for index, shape in titled if index in odd]
            if field is None:
                fix = (f"Set the titles on {_slides(odd)} to "
                       f"{round(usual_value, 1):g} pt.")
            else:
                fix = (f"Set {field} of the titles on {_slides(odd)} to "
                       f"{round(usual_value, 1):g} pt.")
            findings.append(_finding(
                None, "inconsistent_titles", names,
                f"Slide titles{f' on the {layout} layout' if layout else ''} "
                f"differ in {label}: {described}.",
                fix, severity="warning", slides=sorted(odd),
            ))
    return findings


# ---------------------------------------------------------------------------
# What a reader has to read
# ---------------------------------------------------------------------------
def within_reach(shape, touched) -> bool:
    """Whether a write that touched ``touched`` needs this shape's fill and text.

    Geometry is all the alignment check needs, and it comes with every shape.
    Fill and text decide only overlap and contrast, and a shape can only be in
    the way of a touched shape, or behind its text, where the two meet. So the
    rest are left unread, which on macOS is most of the cost of a check.

    Placeholders are always read, since whether one is empty decides whether it
    is lined up against anything. Text overflowing its own frame far enough to
    reach a touched shape is the one meeting this misses, and that shape's own
    write already reported the overflow.
    """
    if shape.get("placeholder_type") is not None:
        return True
    box = shape_box(shape)
    for other in touched:
        for reach in (shape_box(other), text_box(other)):
            if reach is None:
                continue
            width, height = _shared(box, reach)
            if width > 0 and height > 0:
                return True
    return False


# ---------------------------------------------------------------------------
# Reading a slide (Windows)
# ---------------------------------------------------------------------------
def _fill_facts(shape):
    try:
        fill = shape.Fill
    except Exception:
        return None
    facts = {"visible": None, "solid": None, "color_hex": None, "transparency": None}
    try:
        facts["visible"] = fill.Visible == msoTrue
    except Exception:
        pass
    try:
        facts["solid"] = fill.Type == msoFillSolid
    except Exception:
        pass
    try:
        facts["color_hex"] = int_to_hex(fill.ForeColor.RGB)
    except Exception:
        pass
    try:
        facts["transparency"] = round(fill.Transparency, 2)
    except Exception:
        pass
    return facts


def _text_facts(shape, full):
    """What a layout check needs to know about a shape's text.

    Every read stands on its own, so a frame that will not report one thing
    still reports the rest, and a missing number stays None rather than
    becoming a zero that would read as a measurement.

    Where the text is drawn is always read, because any shape's text can be
    in another's way. The frame's settings and the text's colours and sizes
    are read only when ``full``: they answer overflow, contrast and titles,
    which a write tool asks only about the shape it touched.
    """
    from ppt_com.shapes import _text_frame_state

    tr = shape.TextFrame.TextRange
    try:
        content = tr.Text or ""
    except Exception:
        content = ""
    facts = {
        "chars": len(content.strip()),
        "left": None, "top": None, "width": None, "height": None,
        "alignment": None, "runs": [],
        "margins": None, "word_wrap": None, "autofit": None,
        "orientation": None, "vertical_anchor": None,
    }
    if full:
        state = _text_frame_state(shape) or {}
        for key in ("margins", "word_wrap", "autofit", "orientation", "vertical_anchor"):
            facts[key] = state.get(key)
    if not facts["chars"]:
        return facts

    # An empty frame has no bounds, which is why this waits for text.
    for key, attribute in (("left", "BoundLeft"), ("top", "BoundTop"),
                           ("width", "BoundWidth"), ("height", "BoundHeight")):
        try:
            facts[key] = round(getattr(tr, attribute), 2)
        except Exception:
            pass
    # Alignment only places text that PowerPoint would not place itself.
    if facts["left"] is None or facts["top"] is None:
        try:
            facts["alignment"] = tr.Paragraphs(1).ParagraphFormat.Alignment
        except Exception:
            pass
    if not full:
        return facts

    styles = set()
    try:
        run_count = tr.Runs().Count
    except Exception:
        run_count = 0
    for i in range(1, run_count + 1):
        try:
            run = tr.Runs(i)
            if not run.Text.strip():
                continue
            font = run.Font
            styles.add((int_to_hex(font.Color.RGB), round(font.Size, 1),
                        font.Bold == msoTrue))
        except Exception:
            continue
    facts["runs"] = [
        {"color_hex": colour, "size": size, "bold": bold}
        for colour, size, bold in sorted(styles, key=lambda s: (s[0] or "", s[1] or 0))
    ]
    return facts


def _geometry_facts(shape, z):
    """Where a shape is and what kind of shape it is."""
    facts = {
        "name": shape.Name,
        "z": z,
        "type": shape.Type,
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
        "rotation": None,
        "visible": True,
        "connector": False,
        "placeholder_type": None,
        "fill": None,
        "line_visible": None,
        "text": None,
    }
    try:
        facts["rotation"] = round(shape.Rotation, 2)
    except Exception:
        pass
    try:
        facts["visible"] = shape.Visible == msoTrue
    except Exception:
        pass
    try:
        facts["connector"] = shape.Connector == msoTrue
    except Exception:
        pass
    if facts["type"] == msoPlaceholder:
        try:
            facts["placeholder_type"] = shape.PlaceholderFormat.Type
        except Exception:
            pass
    return facts


def _read_appearance(facts, shape, full):
    """Add what the shape draws: its fill and line, and its text."""
    if facts["type"] != msoGroup:
        facts["fill"] = _fill_facts(shape)
        try:
            facts["line_visible"] = shape.Line.Visible == msoTrue
        except Exception:
            pass
    try:
        has_frame = shape.HasTextFrame == msoTrue
    except Exception:
        has_frame = False
    if has_frame:
        try:
            facts["text"] = _text_facts(shape, full)
        except Exception:
            logger.debug("Cannot read the text of '%s'", facts["name"], exc_info=True)
    return facts


def _shape_facts(shape, z, detail):
    """One shape read in full, for a caller that has no slide around it."""
    return _read_appearance(_geometry_facts(shape, z), shape,
                            detail is None or shape.Name in detail)


def needs_background(shapes, detail) -> bool:
    """Whether the slide's background has to be read.

    It is the backdrop of last resort for the contrast check, and nothing else
    uses it, so a write that touched no text does not pay for reading it.
    """
    if detail is None:
        return True
    return any(
        (shape.get("text") or {}).get("chars")
        for shape in shapes if shape["name"] in detail
    )


def read_shapes(refs, geometry, appearance, detail):
    """Read a slide's shapes, as much of each as the check needs.

    ``geometry(ref, z)`` and ``appearance(facts, ref, full)`` are the
    platform's own readers. With no ``detail`` every shape is read in full.
    With names, those shapes are read in full first, and then the others
    only where they meet one of them; the rest stay geometry.
    """
    shapes = [geometry(ref, z) for z, ref in enumerate(refs, start=1)]
    if detail is None:
        for facts, ref in zip(shapes, refs):
            appearance(facts, ref, True)
        return shapes
    touched = []
    for facts, ref in zip(shapes, refs):
        if facts["name"] in detail:
            touched.append(appearance(facts, ref, True))
    for facts, ref in zip(shapes, refs):
        if facts["name"] not in detail and within_reach(facts, touched):
            appearance(facts, ref, False)
    return shapes


def _background_hex(slide):
    """The slide's background colour when it is one solid colour, else None.

    A slide that follows its layout's background, which follows the master's,
    is painted by whichever of them stopped following.
    """
    try:
        source = slide
        if slide.FollowMasterBackground == msoTrue:
            source = slide.CustomLayout
            if source.FollowMasterBackground == msoTrue:
                source = slide.Master
        fill = source.Background.Fill
        if fill.Type == msoFillSolid:
            return int_to_hex(fill.ForeColor.RGB)
    except Exception:
        logger.debug("Cannot read the slide background", exc_info=True)
    return None


def _slide_facts_impl(slide_indices, navigate, detail):
    """Read the slides into the plain dicts the checks work on.

    ``detail`` is None to read every shape in full, or the names of the
    shapes a write touched; see ``read_shapes``.
    """
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    total = pres.Slides.Count
    facts = {
        "slide_width": pres.PageSetup.SlideWidth,
        "slide_height": pres.PageSetup.SlideHeight,
        "slides": [],
    }
    for slide_index in slide_indices:
        if slide_index < 1 or slide_index > total:
            raise ValueError(
                f"Slide index {slide_index} is out of range. "
                f"The presentation has {total} slides (1-based)."
            )
        if navigate:
            goto_slide(app, slide_index)
        slide = pres.Slides(slide_index)
        try:
            layout = slide.CustomLayout.Name
        except Exception:
            layout = None
        refs = [slide.Shapes(z) for z in range(1, slide.Shapes.Count + 1)]
        shapes = read_shapes(refs, _geometry_facts, _read_appearance, detail)
        facts["slides"].append({
            "slide_index": slide_index,
            "layout": layout,
            "background_hex": (
                _background_hex(slide) if needs_background(shapes, detail) else None
            ),
            "shapes": shapes,
        })
    return facts


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------
def check_layout(params: CheckLayoutInput) -> str:
    """Check one slide, or every slide and the deck as a whole."""
    # The text module's, which is already the Apple Event one on macOS.
    from ppt_com.text import _slide_count_impl

    try:
        if params.slide_index is not None:
            indices = [params.slide_index]
        else:
            indices = list(range(1, ppt.execute(_slide_count_impl) + 1))
        facts = ppt.execute(_slide_facts_impl, indices, len(indices) == 1, None)
        slide_w, slide_h = facts["slide_width"], facts["slide_height"]
        findings = []
        for slide in facts["slides"]:
            findings += slide_findings(slide, slide_w, slide_h)
        if params.slide_index is None:
            findings += deck_findings(facts["slides"], slide_h)
        result = {
            "issues": findings,
            "total": len(findings),
            "slides_checked": len(facts["slides"]),
        }
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def report_layout(result: dict, slide_index: int, shape_names) -> dict:
    """Add what is now wrong with the shapes a write tool just touched.

    Called by those tools after their write has landed, with the names out of
    their own result. Nothing is added when nothing is wrong. When the check
    itself could not run, the result says that instead, so an absent
    ``layout_issues`` always means the shapes were checked and are fine.
    """
    names = {name for name in shape_names if name}
    if not names or not isinstance(result, dict) or "error" in result:
        return result
    try:
        facts = ppt.execute(_slide_facts_impl, [slide_index], False, sorted(names))
        slide = facts["slides"][0]
        issues = slide_findings(
            slide, facts["slide_width"], facts["slide_height"], only=names,
        )
    except Exception as e:  # noqa: BLE001 - the write itself already succeeded
        logger.debug("Layout check after a write failed", exc_info=True)
        result["layout_check_failed"] = str(e)
        return result
    if issues:
        result["layout_issues"] = issues
    return result


def touched_names(result: dict) -> list:
    """The shape names a write tool's result reports, one or several."""
    if not isinstance(result, dict):
        return []
    names = []
    if result.get("shape_name"):
        names.append(result["shape_name"])
    for entry in result.get("updated") or []:
        if isinstance(entry, dict) and entry.get("shape_name"):
            names.append(entry["shape_name"])
    return names


# ---------------------------------------------------------------------------
# Tool registration
# ---------------------------------------------------------------------------
def register_tools(mcp):
    """Register the layout check tool with the MCP server."""

    @mcp.tool(
        name="ppt_check_layout",
        annotations={
            "title": "Check Slide Layout",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_check_layout(params: CheckLayoutInput) -> str:
        """Find layout problems on a slide, or across the deck, by measuring.

        Reports what is wrong as numbers rather than as a picture, so it works
        whether or not the caller can read an image, and it does not miss a
        three point slip the way an eye does. Each issue names the shapes and
        says how to fix it, worst first:

        - off_slide: a shape, or its text, reaches past a slide edge
        - overflow: text needs more room than its box gives it
        - overlap: text runs into another shape or other text
        - low_contrast: text colour against what is behind it is below the
          WCAG minimum for its size
        - misaligned: edges a few points apart that were meant to line up
        - uneven_spacing: a row or column of like shapes with unequal gaps
        - empty_placeholder: a placeholder left without text
        - inconsistent_titles (whole deck only): titles that differ in size or
          position between slides with the same layout

        Run it when a slide is finished, and on the whole deck before calling
        it done. A slide with no issues is finished as far as geometry goes;
        ppt_get_slide_preview is still the way to judge how it looks, and
        ppt_check_typography covers line breaks.

        The tools that add, move, resize or retext a shape run the per shape
        checks on that shape themselves and report what they find in
        ``layout_issues``. No ``layout_issues`` means it was checked and is
        fine; ``layout_check_failed`` means the check could not run.
        """
        return await run_offloaded(check_layout, params)


# On macOS the slide is read over Apple Events. Only _slide_facts_impl is
# swapped; the checks, and what decides how much of a slide to read, are shared.
from backend import IS_MACOS, use_mac_impls  # noqa: E402

if IS_MACOS:  # pragma: no cover - platform specific
    from ppt_mac import layout_check as _mac_layout_check
    use_mac_impls(globals(), _mac_layout_check)
