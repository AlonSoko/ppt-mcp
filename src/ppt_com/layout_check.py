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
import re
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

# Centres this close are one centre. A text box that grows or shrinks to fit
# its text keeps its top, so its middle moves by half the change, and a point
# between two centres cannot be seen the way it can between two edges.
_ALIGNED_CENTRE = 1.5

# Text this close to a side of its frame's inset keeps to that side: it is set
# left, right or centred there, or fills the frame.
_SNUG = 1.5

# Edges this close, and not closer, were meant to line up and do not. Further
# apart than this is a decision rather than a slip.
_NEAR_MISS = 4.0

# Text has to cover this much of another shape, in both directions, before it
# counts as running into it. Less is a glyph's side bearing touching a border.
_MIN_OVERLAP = 2.0

# How far apart a fix sets two things that collided: enough to read as two.
_CLEARANCE = 6.0

# How far apart the gaps in a row of like shapes can be before it shows.
_GAP_TOLERANCE = 2.0

# Shapes count as the same kind of thing, for spacing, when their widths and
# heights are within this fraction of each other.
_SIMILAR_SIZE = 0.15

# Two shapes are in the same row (or column) when they share at least this
# fraction of the smaller one's height (or width).
_SHARED_SPAN = 0.5

# Like shapes in a row whose heights (or in a column, widths) are this close
# are the same size.
_SIZE_TOLERANCE = 2.0

# Text closer than this to an edge of the card it sits on is pressed against
# it, when it has twice as much room the other way. Equally little room all
# round is a badge or a pill, and is meant.
_MIN_ROOM = 12.0

# A slide's side margin is the distance most of its shapes keep from the left
# and right edges. Only distances within this share of the slide's width
# count; a shape further in is placed, not kept to a margin.
_MARGIN_ZONE = 0.2

# Margins this close are one margin, and a shape further inside it than this
# is not keeping it.
_MARGIN_TOLERANCE = 2.0

# How far inside the margin a shape can be and still have been meant to keep
# it: this much, or a quarter of the margin when that is more. Further in is a
# decision, a corner logo against a centred composition, not a slip.
_MARGIN_NEAR = 10.0

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

# How much of a shape's text the reader keeps, for the checks that read it.
_CONTENT_KEPT = 500

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
    "off_slide", "overflow", "overlap", "low_contrast", "cramped_text",
    "narrow_margin", "misaligned", "uneven_sizes", "sample_text",
})

# Reported in this order, worst first.
_ORDER = (
    "off_slide", "overflow", "overlap", "low_contrast", "cramped_text",
    "narrow_margin", "misaligned", "uneven_spacing", "uneven_sizes",
    "empty_placeholder", "sample_text", "inconsistent_titles",
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
# How a slide is put together: what sits on what, rows, margins
# ---------------------------------------------------------------------------
def _num(value) -> str:
    """A number of points as a tool call takes it: 640, 219.2."""
    return f"{round(value, 2):g}"


def _belongs_to(s, other, boxes) -> bool:
    """Whether ``s`` sits on ``other``, the way text sits on a card.

    Mostly inside, rather than wholly: the text on a card still belongs to the
    card when it runs past its edge, which is its own finding. A text box set
    to grow with its text can end up with a larger frame than the card it is
    on, so a frame that only holds text belongs to a shape that paints, at any
    size.
    """
    mine, theirs = boxes[s["name"]], boxes[other["name"]]
    width, height = _shared(mine, theirs)
    if width * height < 0.5 * _width(mine) * _height(mine):
        return False
    if paints_box(s) != paints_box(other):
        # Never the other way round: a card is not inside a text frame.
        return paints_box(other)
    return _width(theirs) * _height(theirs) > _width(mine) * _height(mine)


def _depths(arranged, boxes):
    """How many shapes each one sits on: none for a card, one for its text."""
    return {
        s["name"]: sum(
            1 for other in arranged
            if other is not s and _belongs_to(s, other, boxes)
        )
        for s in arranged
    }


def _similar(a_box, b_box) -> bool:
    """Whether two boxes are about one size, in both directions."""
    aw, ah, bw, bh = _width(a_box), _height(a_box), _width(b_box), _height(b_box)
    return (abs(aw - bw) <= _SIMILAR_SIZE * max(aw, bw)
            and abs(ah - bh) <= _SIMILAR_SIZE * max(ah, bh))


def _like_groups(members, boxes):
    """Rows and columns of like shapes among ``members``, two or more each.

    Shapes are alike when their widths and heights are within
    ``_SIMILAR_SIZE`` of each other. A row shares its height and runs left to
    right, a column shares its width and runs down, and each comes sorted that
    way, as ``("horizontal", row)`` or ``("vertical", column)``.
    """
    groups = []
    for axis, direction, start in (("y", "horizontal", 0), ("x", "vertical", 1)):
        def joined(a, b, axis=axis):
            a_box, b_box = boxes[a["name"]], boxes[b["name"]]
            return _similar(a_box, b_box) and _related(a_box, b_box, axis)

        for group in _connected(members, joined):
            if len(group) >= 2:
                group.sort(key=lambda s, start=start: boxes[s["name"]][start])
                groups.append((direction, group))
    return groups


def _text_anchor(shape):
    """Which sides of its frame a text box's text keeps to, as PowerPoint drew it.

    A pair, across and down: ``"left"``, ``"center"``, ``"right"`` or
    ``"full"`` when the text fills the frame's width, and ``"top"``,
    ``"middle"``, ``"bottom"`` or ``"full"`` down it. Either is None when the
    text sits nowhere in particular, and the pair is None when the text was not
    measured. Read off where the text was drawn rather than off the paragraph
    settings, which a reader does not always fetch and which do not say how a
    frame grew around its text.
    """
    words = text_box(shape)
    if words is None:
        return None
    margins = (shape.get("text") or {}).get("margins") or _DEFAULT_MARGINS
    box = shape_box(shape)
    inner = (box[0] + margins["left"], box[1] + margins["top"],
             box[2] - margins["right"], box[3] - margins["bottom"])

    def keeps(low_gap, high_gap, low, middle, high):
        if low_gap <= _SNUG and high_gap <= _SNUG:
            return "full"
        if low_gap <= _SNUG:
            return low
        if high_gap <= _SNUG:
            return high
        # Measured text can sit a few points off the centre it was set to, so
        # the two gaps only have to be close for their size.
        if abs(low_gap - high_gap) <= max(_SNUG, 0.2 * (low_gap + high_gap)):
            return middle
        return None

    return (keeps(words[0] - inner[0], inner[2] - words[2], "left", "center", "right"),
            keeps(words[1] - inner[1], inner[3] - words[3], "top", "middle", "bottom"))


def _seen_edges(shape):
    """The edges of a shape a reader can see, or None for all of them.

    All of a shape that paints. Of a frame that only holds text, the ones its
    text keeps to: the left of left-set text, the centre of centred text. The
    rest of the frame cannot be seen, and lining it up means nothing.
    """
    if paints_box(shape):
        return None
    anchor = _text_anchor(shape)
    if anchor is None:
        return None
    across, down = anchor
    seen = set()
    seen |= {"left", "center", "right"} if across in (None, "full") else {across}
    seen |= {"top", "middle", "bottom"} if down in (None, "full") else {down}
    return seen


def _margins(arranged, boxes, depth, slide_w, slide_h):
    """The margins the slide's shapes keep from its edges, or None.

    The side margin is a distance top level shapes keep from both the left and
    the right edge, the one the most of them keep. A shape touching an edge
    bleeds off it on purpose and keeps no margin, so it does not count. Of a
    frame holding only text, the sides its text keeps to count, and a wide one
    set to a side counts on both, the way a title frame is drawn across the
    slide. A centred one counts on neither: centred text keeps no margin.

    A title and a footer rarely keep the same distance from the top and the
    bottom, so the margin there is the one the topmost shape keeps, usually
    the title, and no more than the side margin.
    """
    if not slide_w or not slide_h:
        return None
    lefts, rights, tops = [], [], []
    for shape in arranged:
        if depth[shape["name"]]:
            continue
        box = boxes[shape["name"]]
        if paints_box(shape):
            sides = ("left", "right")
        else:
            across = (_text_anchor(shape) or (None, None))[0]
            if across == "center" or across is None:
                sides = ()
            elif across == "full" or _width(box) >= 0.5 * slide_w:
                sides = ("left", "right")
            else:
                sides = (across,)
        for side in sides:
            distance = box[0] if side == "left" else slide_w - box[2]
            if _EDGE_TOLERANCE < distance < _MARGIN_ZONE * slide_w:
                (lefts if side == "left" else rights).append(distance)
        if _EDGE_TOLERANCE < box[1] < _MARGIN_ZONE * slide_h:
            tops.append(box[1])
    kept = []
    for value in {round(d, 1) for d in lefts}:
        on_left = sum(1 for d in lefts if abs(d - value) <= _MARGIN_TOLERANCE)
        on_right = sum(1 for d in rights if abs(d - value) <= _MARGIN_TOLERANCE)
        if on_right:
            # The most shapes keep it; between equals, the narrower.
            kept.append((on_left + on_right, -value))
    if not kept:
        return None
    side = -max(kept)[1]
    return {"horizontal": side, "vertical": round(min([side] + tops), 1)}


def _near_margin(distance, margin) -> bool:
    """Whether a shape this far from the edge was meant to keep the margin."""
    return margin - max(_MARGIN_NEAR, 0.25 * margin) <= distance < margin - _MARGIN_TOLERANCE


# Narrower than this, a frame cannot hold the words of a label.
_MIN_LABEL_WIDTH = 60.0

# A label sits this close above or below the dot or icon it labels.
_LABEL_GAP = 30.0


def _centred_text(shape) -> bool:
    """Whether a shape is a frame of centred text, a label under an icon or a dot."""
    return not paints_box(shape) and (_text_anchor(shape) or (None, None))[0] == "center"


def _centred_shorten(shape, low, high):
    """The fix for centred text too near the edge for a narrower frame to help."""
    box = shape_box(shape)
    centre = (box[0] + box[2]) / 2
    room = max(0.0, 2 * min(centre - low, high - centre))
    return (f"Shorten the text of '{shape['name']}': centred where it is, at "
            f"{_pt(centre)}, it has {_pt(room)} across before the margin. Moving "
            "it would take it off what it labels; to make room, move that in "
            "from the edge as well.")


def _centred_fit(shape, low, high):
    """The fix that keeps centred text between ``low`` and ``high``, or None.

    Centred text labels whatever it is centred on, so moving it would take it
    off its mark. Narrowing its frame around the same centre makes the text
    wrap inside the limits and leaves it where it was.
    """
    box = shape_box(shape)
    centre = (box[0] + box[2]) / 2
    half = min(centre - low, high - centre)
    if 2 * half < _MIN_LABEL_WIDTH:
        return None
    return (f"Narrow '{shape['name']}' to {_pt(2 * half)} around its centre, so "
            f"its text wraps inside the margin and stays centred where it is "
            f"(ppt_update_shape left={_num(centre - half)} width={_num(2 * half)}), "
            "or shorten the text.")


def _layout(shapes, slide_w, slide_h):
    """What the checks on how a slide is put together share.

    The shapes that make up the layout and their boxes, how many shapes each
    sits on, the rows and columns of like shapes, and the slide's margins.
    ``groups`` are the top level ones, which are spaced and kept to the
    margins. ``rows`` and ``columns`` map each shape that paints to the like
    shapes beside it on its own level, which are kept one size.

    Only shapes that paint make rows. A text frame's size follows its text,
    so text set beside a row of icons is spaced by the icons it labels, and
    whether it lines up with them is the alignment check's question.
    """
    arranged = _arranged(shapes, slide_w, slide_h)
    boxes = {s["name"]: shape_box(s) for s in arranged}
    depth = _depths(arranged, boxes)
    rows, columns = {}, {}
    for level in sorted(set(depth.values())):
        painted = [s for s in arranged if depth[s["name"]] == level and paints_box(s)]
        for direction, group in _like_groups(painted, boxes):
            names = [s["name"] for s in group]
            for name in names:
                (rows if direction == "horizontal" else columns)[name] = names
    top = [s for s in arranged if depth[s["name"]] == 0 and paints_box(s)]
    return {
        "arranged": arranged, "boxes": boxes, "depth": depth,
        "groups": _like_groups(top, boxes), "rows": rows, "columns": columns,
        "margins": _margins(arranged, boxes, depth, slide_w, slide_h),
    }


def _spans(row, boxes, slide_w) -> bool:
    """Whether a row reaches across to both sides, as a row of cards under a
    title does, rather than sitting somewhere on the slide."""
    zone = _MARGIN_ZONE * slide_w
    return boxes[row[0]["name"]][0] < zone and slide_w - boxes[row[-1]["name"]][2] < zone


def _row_fix(row, boxes, slide_w, margin):
    """The fix that spreads ``row`` evenly between the side margins, or None
    when it is already there or too wide to fit."""
    widths = [_width(boxes[s["name"]]) for s in row]
    room = slide_w - 2 * margin - sum(widths)
    if room <= 0:
        return None
    gap = room / (len(row) - 1)
    moves, x = [], margin
    for shape, width in zip(row, widths, strict=True):
        if abs(boxes[shape["name"]][0] - x) > 0.05:
            moves.append(f"of '{shape['name']}' to {_pt(x)}")
        x += width + gap
    if not moves:
        return None
    return (f"Set left {' and '.join(moves)} (ppt_update_shape). Every gap is "
            f"then {_pt(gap)}, and the row keeps the {_pt(margin)} side "
            f"margins the rest of the slide keeps.")


_OPPOSITE = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
_WHERE = {"top": "above", "bottom": "below", "left": "on the left", "right": "on the right"}


def _room(words, card):
    """How far text is from each edge of the card it sits on, negative past it."""
    return {
        "top": words[1] - card[1], "bottom": card[3] - words[3],
        "left": words[0] - card[0], "right": card[2] - words[2],
    }


def _wanted_room(room, side):
    """The room text should have on ``side`` of its card.

    What it has on the opposite side, the way a card is padded evenly, but no
    more than it has on the closer of the other two, so text set at the top
    of a tall card is not told to move to the middle of it. The closer one,
    because ragged lines leave more room on their own side than the card was
    padded with.
    """
    across = ("left", "right") if side in ("top", "bottom") else ("top", "bottom")
    return max(0.0, min(room[_OPPOSITE[side]], min(room[s] for s in across)))


def _padding(room, side, want) -> str:
    """Why the room asked for is that much: ', as it has above', when it is."""
    if abs(room[_OPPOSITE[side]] - want) <= 0.05:
        return f", as it has {_WHERE[_OPPOSITE[side]]}"
    return ""


def _card_of(shape, layout):
    """The shape that paints under this one's text, the innermost when they
    nest, or None."""
    boxes = layout["boxes"]
    if shape["name"] not in boxes:
        return None
    cards = [other for other in layout["arranged"]
             if other is not shape and paints_box(other)
             and _belongs_to(shape, other, boxes)]
    if not cards:
        return None
    return min(cards, key=lambda c: _width(boxes[c["name"]]) * _height(boxes[c["name"]]))


def _grow(card, side, amount, layout):
    """The fix that gives ``card`` ``amount`` more room on ``side``.

    The rest of its row grows with it when the room is height, and the rest
    of its column when it is width, so like cards stay one size.
    """
    name, box = card["name"], layout["boxes"][card["name"]]
    if side in ("top", "bottom"):
        word, group = "taller", "row"
        call = f"height={_num(_height(box) + amount)}"
        mates = layout["rows"].get(name, [])
    else:
        word, group = "wider", "column"
        call = f"width={_num(_width(box) + amount)}"
        mates = layout["columns"].get(name, [])
    if side == "top":
        call += f" top={_num(box[1] - amount)}"
    elif side == "left":
        call += f" left={_num(box[0] - amount)}"
    others = [m for m in mates if m != name]
    if not others:
        return f"make '{name}' {_pt(amount)} {word} (ppt_update_shape {call})"
    listed = ", ".join(f"'{m}'" for m in others)
    return (f"make '{name}' {_pt(amount)} {word}, and {listed} with it so the "
            f"{group} stays one size (ppt_update_shape shape_names="
            f"{[name] + others} {call})")


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


def _off_slide(slide_index, shapes, slide_w, slide_h, layout):
    """Shapes, or the text of shapes, that reach past an edge of the slide.

    Measured on what is drawn. A text box whose frame hangs over the edge with
    its text well inside is not wrong, and a box that fits on the slide with
    its text spilling off the bottom is.

    The fix brings the shape back inside the margin the rest of the slide
    keeps, not just onto it: moved by the overshoot alone, it would end up
    flush against the edge.
    """
    if not slide_w or not slide_h:
        return []
    margins = layout["margins"] or {}
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
        name = shape["name"]
        moves, calls, kept, too_big = [], [], [], []
        for low, high, field, keep, words, bigger in (
            ("left", "right", "left", margins.get("horizontal"), ("right", "left"), "wider"),
            ("top", "bottom", "top", margins.get("vertical"), ("down", "up"), "taller"),
        ):
            if low in past and high in past:
                too_big.append(bigger)
                continue
            for edge, sign, word in ((low, 1, words[0]), (high, -1, words[1])):
                if edge not in past:
                    continue
                amount = past[edge] + (keep or 0.0)
                moves.append(f"{word} by {'' if keep else 'at least '}{_pt(amount)}")
                calls.append(f"{field}={_num(shape[field] + sign * amount)}")
                if keep:
                    kept.append(_pt(keep))
        keep = margins.get("horizontal") or 0.0
        centred = None
        if _centred_text(shape) and not {"top", "bottom"} & set(past):
            # Moved, centred text would leave what it is centred on.
            centred = _centred_fit(shape, keep, slide_w - keep) or _centred_shorten(
                shape, keep, slide_w - keep)
        if too_big:
            fix = f"'{name}' is {' and '.join(too_big)} than the slide: make it smaller."
        elif centred:
            fix = centred
        else:
            keeping = (f", to keep the {' and '.join(kept)} margin"
                       f"{'s' if len(kept) > 1 else ''} the rest of the slide keeps"
                       if kept else "")
            fix = (f"Move '{name}' {' and '.join(moves)}{keeping} "
                   f"(ppt_update_shape {' '.join(calls)}), or make it smaller.")
        findings.append(_finding(
            slide_index, "off_slide", [name],
            f"'{name}' runs {' and '.join(parts)} of the slide.",
            fix, edges=past,
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
                f"Make '{name}' {_pt(needed - room)} wider (ppt_update_shape "
                f"width={_num(shape['width'] + needed - room)}), turn word "
                "wrap on (ppt_set_textframe word_wrap=true), or shorten the line.",
                needed_pt=needed, available_pt=room,
            ))
            continue
        vertical = text.get("orientation") in ("vertical", "upward", "downward")
        field = "width" if vertical else "height"
        findings.append(_finding(
            slide_index, "overflow", [name],
            f"The text in '{name}' needs {_pt(needed)} and the box has "
            f"{_pt(room)}, so it spills out of the box.",
            f"Make '{name}' {_pt(needed - room)} {'wider' if vertical else 'taller'} "
            f"(ppt_update_shape {field}={_num(shape[field] + needed - room)}), "
            "cut the text, or make the font smaller.",
            needed_pt=needed, available_pt=room,
        ))
    return findings


def _overlap(slide_index, shapes, layout):
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
            findings.append(_overlap_finding(
                slide_index, a, a_text, b, b_ink, width, height, layout))
    return findings


def _overlap_finding(slide_index, a, a_text, b, b_ink, width, height, layout):
    a_name, b_name = a["name"], b["name"]
    shared = {"width": round(width, 1), "height": round(height, 1)}

    # Text mostly on a shape underneath it is meant to be on it, like the
    # words on a card, and the problem is that it does not fit. Saying where
    # it runs past is more use than saying the two collide.
    area = _width(a_text) * _height(a_text)
    if (paints_box(b) and b.get("z", 0) < a.get("z", 0)
            and area and width * height >= 0.5 * area):
        room = _room(a_text, b_ink)
        past = {edge: round(-amount, 1) for edge, amount in room.items()
                if -amount > _EDGE_TOLERANCE}
        parts = " and ".join(f"{_pt(amount)} past its {edge} edge"
                             for edge, amount in past.items())
        # Grown by the overshoot alone, the card would end exactly where the
        # text does. It grows by as much again as the text has on the other
        # side, so the text is padded evenly.
        if b_name in layout["boxes"]:
            grows = []
            for edge, amount in past.items():
                want = _wanted_room(room, edge)
                grows.append(
                    _grow(b, edge, amount + want, layout)
                    + (f", which leaves the text {_pt(want)} {_WHERE[edge]}"
                       f"{_padding(room, edge, want)}" if want else "")
                )
            fix = f"{'; '.join(grows)}. Or make the text of '{a_name}' shorter or smaller."
            fix = fix[0].upper() + fix[1:]
        else:
            fix = (f"Make '{b_name}' larger by that much, or make the text of "
                   f"'{a_name}' shorter or smaller.")
        return _finding(
            slide_index, "overlap", [a_name, b_name],
            f"The text of '{a_name}' sits on '{b_name}' and runs {parts}.",
            fix, overlap_pt=shared, past_edges=past,
        )

    # What b puts on the slide is its text when it does not paint its box.
    target = f"'{b_name}'" if paints_box(b) else f"the text of '{b_name}'"
    # The cheaper way apart: along whichever axis they share less of, moving
    # the smaller of the two, the caption rather than the table it captions.
    a_box, b_box = shape_box(a), shape_box(b)
    if _width(a_box) * _height(a_box) <= _width(b_box) * _height(b_box):
        mover, mine, theirs = a_name, a_text, b_ink
    else:
        mover, mine, theirs = b_name, b_ink, a_text
    if height <= width:
        clear = height + _CLEARANCE
        sign = 1 if mine[1] + mine[3] >= theirs[1] + theirs[3] else -1
        way, call = ("down" if sign > 0 else "up"), f"dtop={_num(sign * clear)}"
    else:
        clear = width + _CLEARANCE
        sign = 1 if mine[0] + mine[2] >= theirs[0] + theirs[2] else -1
        way, call = ("right" if sign > 0 else "left"), f"dleft={_num(sign * clear)}"
    other = b_name if mover == a_name else a_name
    fix = (f"Move '{mover}' {way} by {_pt(clear)} (ppt_update_shape {call}), "
           f"which leaves {_pt(_CLEARANCE)} between it and '{other}', or "
           f"shorten the text of '{a_name}'.")
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
                worst = (ratio, needed, colour, size)
        if worst is None:
            continue
        ratio, needed, colour, size = worst
        name = shape["name"]
        # The colour the slide's other text already uses keeps it looking like
        # one slide; black or white is the answer when none of them is enough.
        ours = [c for c in _text_colours(shapes, shape, size)
                if _contrast_ratio(c, backdrop) >= needed]
        if ours:
            suggestion, source = ours[0], ", the colour of the other text on the slide"
        else:
            dark, light = _contrast_ratio("#000000", backdrop), _contrast_ratio("#FFFFFF", backdrop)
            suggestion, source = ("#000000" if dark >= light else "#FFFFFF"), ""
        findings.append(_finding(
            slide_index, "low_contrast", [name],
            f"Text in '{name}' is {colour} on {backdrop}, a contrast of "
            f"{ratio:.1f}:1 where {needed:g}:1 is the minimum for its size.",
            f"Make the text {suggestion}{source} (contrast "
            f"{_contrast_ratio(suggestion, backdrop):.1f}:1): ppt_format_text "
            f"color='{suggestion}'. Or change what is behind it.",
            text_color=colour, backdrop_color=backdrop,
            contrast=round(ratio, 2), minimum=needed,
        ))
    return findings


def _text_colours(shapes, exclude, size=None):
    """The colours the slide's other text is set in.

    The colour of text about ``size`` first, body text for body text rather
    than the colour of the headings, and then the most used.
    """
    counts = {}
    for shape in shapes:
        if shape is exclude or not shape.get("visible", True):
            continue
        for run in (shape.get("text") or {}).get("runs") or []:
            colour = (run.get("color_hex") or "").upper()
            if not colour:
                continue
            alike = bool(size) and abs((run.get("size") or 0.0) - size) <= 2.0
            seen = counts.setdefault(colour, [0, 0])
            seen[0] += alike
            seen[1] += 1
    return sorted(counts, key=lambda c: (-counts[c][0], -counts[c][1]))


# Edge names per axis, in the order they are preferred when two tie.
_X_EDGES = ("left", "center", "right")
_Y_EDGES = ("top", "middle", "bottom")


def _edges(box, axis):
    if axis == "x":
        return dict(zip(_X_EDGES, (box[0], (box[0] + box[2]) / 2, box[2]), strict=True))
    return dict(zip(_Y_EDGES, (box[1], (box[1] + box[3]) / 2, box[3]), strict=True))


def _related(a_box, b_box, axis) -> bool:
    """Shapes stacked in a column are lined up left and right; shapes in a row,
    top and bottom. Shapes that are neither have nothing to line up."""
    if axis == "x":
        shared = _shared(a_box, b_box)[0]
        return shared >= _SHARED_SPAN * min(_width(a_box), _width(b_box))
    shared = _shared(a_box, b_box)[1]
    return shared >= _SHARED_SPAN * min(_height(a_box), _height(b_box))


def _within(edge, gap, loose=True) -> bool:
    """Whether two lines this far apart are one line.

    Centres get the looser allowance only between shapes of different sizes,
    whose edges cannot meet anyway. Between two of one size, a centre off is
    every edge off by as much.
    """
    centre = loose and edge in ("center", "middle")
    return gap <= (_ALIGNED_CENTRE if centre else _ALIGNED)


def _misaligned(slide_index, layout, only=None):
    """Edges a few points apart that were meant to be one edge.

    Each pair of shapes in a row is compared top, middle and bottom, and each
    pair in a column left, centre and right. A pair already sharing one of
    those lines is aligned. A pair whose nearest lines are less than
    ``_NEAR_MISS`` apart, but not together, is the slip this reports.

    Of the two, the one to move is the one lined up with fewer other shapes,
    and between equals the one added later. One finding names it with the
    shapes on the line it nearly meets.

    Only shapes on the same level are compared: cards with cards, and the text
    on one card with the text on the next. Text is placed on its card, inset
    from its edges, rather than lined up with it or with the cards beside it.
    Of a text frame only the edges its text keeps to are compared: the centre
    of a centred name, not the sides of a frame no one can see.

    Levels are judged from the cards down. The icon and text on a card that
    slipped slipped with it, so they are judged where the card's fix will put
    them, and the card's fix moves them along.
    """
    arranged, boxes, depth = layout["arranged"], layout["boxes"], layout["depth"]
    seen = {s["name"]: _seen_edges(s) for s in arranged}
    placed = dict(boxes)
    slips = []
    for level in sorted(set(depth.values())):
        members = [s for s in arranged if depth[s["name"]] == level]
        found = []
        for axis in ("x", "y"):
            found += _slips_along(axis, members, placed, seen)
        for slip in found:
            mover = slip["mover"]
            slip["carried"] = [
                s for s in arranged
                if depth[s["name"]] > level and paints_box(mover)
                and _belongs_to(s, mover, boxes)
            ]
            for shape in slip["carried"]:
                placed[shape["name"]] = _moved(placed[shape["name"]], slip["axis"], -slip["delta"])
        slips += found

    findings = []
    for slip in slips:
        finding = _misaligned_finding(slide_index, slip, placed)
        if only is None or set(finding["shapes"]) & set(only):
            findings.append(finding)
    return findings


def _off_centre_labels(slide_index, layout):
    """Centred labels set well off the centre of the mark they label.

    A label under a dot or an icon is centred on it. One a few points off is
    an ordinary slip and the alignment check reports it; this finds the label
    whose frame was drawn wider than the rest from the same left offset, so
    its text is centred a long way to one side, up to a quarter of its own
    width, beyond which it is not that mark's label at all.
    """
    boxes, depth = layout["boxes"], layout["depth"]
    findings = []
    for label in layout["arranged"]:
        words = text_box(label) if _centred_text(label) else None
        if words is None:
            continue
        mark, nearest = None, None
        for shape in layout["arranged"]:
            if shape is label or not paints_box(shape) or depth[shape["name"]] != depth[label["name"]]:
                continue
            box = boxes[shape["name"]]
            gap = max(box[1] - words[3], words[1] - box[3])
            if not _related(words, box, "x") or not -_EDGE_TOLERANCE <= gap <= _LABEL_GAP:
                continue
            if nearest is None or gap < nearest:
                mark, nearest = shape, gap
        if mark is None:
            continue
        frame, mark_box = boxes[label["name"]], boxes[mark["name"]]
        off = (frame[0] + frame[2]) / 2 - (mark_box[0] + mark_box[2]) / 2
        if not _NEAR_MISS < abs(off) <= 0.25 * _width(words):
            continue
        name, left = label["name"], round(frame[0] - off, 2)
        findings.append(_finding(
            slide_index, "misaligned", [name, mark["name"]],
            f"The center of '{name}' is {_pt(abs(off))} "
            f"{'right' if off > 0 else 'left'} of the center of '{mark['name']}', "
            "the mark it labels.",
            f"Set left of '{name}' to {_pt(left)} so it is centred on "
            f"'{mark['name']}' (ppt_update_shape left={left:g}).",
            severity="warning", edge="center", off_by_pt=round(abs(off), 2),
        ))
    return findings


def _moved(box, axis, by):
    """A box moved along one axis."""
    if axis == "x":
        return (box[0] + by, box[1], box[2] + by, box[3])
    return (box[0], box[1] + by, box[2], box[3] + by)


def _slips_along(axis, members, boxes, seen):
    """The slips among one level's shapes along one axis."""
    aligned = {s["name"]: 0 for s in members}
    near = []
    for i, a in enumerate(members):
        for b in members[i + 1:]:
            a_box, b_box = boxes[a["name"]], boxes[b["name"]]
            if not _related(a_box, b_box, axis):
                continue
            a_edges, b_edges = _edges(a_box, axis), _edges(b_box, axis)
            gaps = {
                edge: abs(a_edges[edge] - b_edges[edge]) for edge in a_edges
                if all(seen[s["name"]] is None or edge in seen[s["name"]] for s in (a, b))
            }
            if not gaps:
                continue
            span = _width if axis == "x" else _height
            loose = abs(span(a_box) - span(b_box)) > 2 * _ALIGNED
            if any(_within(edge, gap, loose) for edge, gap in gaps.items()):
                aligned[a["name"]] += 1
                aligned[b["name"]] += 1
                continue
            closest = min(gaps, key=gaps.get)
            if gaps[closest] <= _NEAR_MISS:
                near.append((a, b, closest))

    grouped = {}
    for a, b, edge in near:
        # The one that stays is lined up with more shapes, or was there
        # first. Stacking order is creation order unless something moved a
        # shape in it, which is as close as a slide comes to a history.
        if (aligned[a["name"]], -a.get("z", 0)) > (aligned[b["name"]], -b.get("z", 0)):
            anchor, mover = a, b
        else:
            anchor, mover = b, a
        grouped.setdefault((mover["name"], edge), (mover, edge, []))[2].append(anchor)
    return [_slip(axis, mover, edge, anchors, boxes) for mover, edge, anchors in grouped.values()]


def _slip(axis, mover, edge, anchors, boxes):
    """How far ``mover``'s edge is from the line most of ``anchors`` share."""
    values = [_edges(boxes[a["name"]], axis)[edge] for a in anchors]
    # Where most of the others are; between equals, the nearest.
    mine = _edges(boxes[mover["name"]], axis)[edge]
    target = max(
        {round(v, 2) for v in values},
        key=lambda v: (sum(1 for x in values if _within(edge, abs(x - v))), -abs(v - mine)),
    )
    # Only the shapes on that line are named. One that is off it as well is
    # off against them too, and has its own finding.
    on_line = [a for a, v in zip(anchors, values, strict=True) if _within(edge, abs(v - target))]
    return {"axis": axis, "mover": mover, "edge": edge, "anchors": on_line,
            "delta": mine - target}


def _misaligned_finding(slide_index, slip, placed):
    mover, axis, edge, delta = slip["mover"], slip["axis"], slip["edge"], slip["delta"]
    carried = slip.get("carried") or []
    name = mover["name"]
    field = "left" if axis == "x" else "top"
    # Where it goes is worked out from where it will be once the card it sits
    # on has been put back, which is where it was judged.
    value = round(placed[name][0 if axis == "x" else 1] - delta, 2)
    if axis == "x":
        direction, way = ("right of", "left") if delta > 0 else ("left of", "right")
    else:
        direction, way = ("below", "up") if delta > 0 else ("above", "down")
    others = ", ".join(f"'{a['name']}'" for a in slip["anchors"])
    if carried:
        names = [name] + [c["name"] for c in carried]
        listed = ", ".join(f"'{c['name']}'" for c in carried)
        fix = (f"Move '{name}' {way} by {_pt(abs(delta))} together with the shapes "
               f"on it, {listed} (ppt_update_shape shape_names={names} "
               f"d{field}={_num(-delta)}).")
    else:
        names = [name]
        fix = (f"Set {field} of '{name}' to {_pt(value)} (ppt_update_shape "
               f"{field}={value:g}), or ppt_align_shapes with align_to='{edge}'.")
    return _finding(
        slide_index, "misaligned", names[:1] + [a["name"] for a in slip["anchors"]] + names[1:],
        f"The {edge} of '{name}' is {_pt(abs(delta))} {direction} the "
        f"{edge} of {others}, close enough to look like a slip.",
        fix, severity="warning", edge=edge, off_by_pt=round(abs(delta), 2),
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


def _uneven_spacing(slide_index, layout, slide_w):
    """Rows and columns of like shapes whose gaps differ.

    Three or more shapes of about the same size, side by side in a row or
    stacked in a column, are one group: cards, steps, columns of icons. Their
    gaps are meant to be equal. Shapes sitting inside another shape belong to
    it, so the text on three cards is not a fourth row to space.

    A row that reaches across the slide is spread between the side margins
    rather than between its own two ends. When an end card is the one out of
    place, evening out the gaps around it leaves the row off centre.
    """
    boxes = layout["boxes"]
    margin = (layout["margins"] or {}).get("horizontal")
    findings = []
    for direction, members in layout["groups"]:
        if len(members) < 3:
            continue
        # A row is spaced along x, a column along y.
        start = 0 if direction == "horizontal" else 1
        gaps = [
            boxes[b["name"]][start] - boxes[a["name"]][start + 2]
            for a, b in zip(members, members[1:], strict=False)
        ]
        if min(gaps) <= 0 or max(gaps) - min(gaps) <= _GAP_TOLERANCE:
            continue
        names = [s["name"] for s in members]
        fix = None
        if direction == "horizontal" and margin is not None and _spans(members, boxes, slide_w):
            fix = _row_fix(members, boxes, slide_w, margin)
        if fix is None:
            even = sum(gaps) / len(gaps)
            fix = (f"ppt_distribute_shapes with these shape names and "
                   f"direction='{direction}', or make every gap {_pt(even)}.")
        listed = ", ".join(f"'{n}'" for n in names)
        findings.append(_finding(
            slide_index, "uneven_spacing", names,
            f"The gaps between {listed} are "
            f"{', '.join(_pt(g) for g in gaps)}; a "
            f"{'row' if direction == 'horizontal' else 'column'} of like "
            "shapes reads as one group when they are equal.",
            fix, severity="warning", gaps_pt=[round(g, 1) for g in gaps],
        ))
    return findings


def _uneven_sizes(slide_index, layout):
    """Like shapes side by side that are not quite one size.

    Cards in a row are meant to be one height, and in a column one width.
    Shapes far enough apart in size are different kinds of thing and are not
    grouped at all, so what this finds is a card grown to hold its text while
    the rest of its row was left as it was. Only shapes that paint are
    judged, since a text box's frame cannot be seen.
    """
    boxes = layout["boxes"]
    findings, seen = [], set()
    for groups, measure, field, word, kind in (
        (layout["rows"], _height, "height", "tall", "row"),
        (layout["columns"], _width, "width", "wide", "column"),
    ):
        for names in groups.values():
            if tuple(names) in seen:
                continue
            seen.add(tuple(names))
            sizes = {name: measure(boxes[name]) for name in names}
            biggest = max(sizes.values())
            if biggest - min(sizes.values()) <= _SIZE_TOLERANCE:
                continue
            clusters = sorted(
                _clusters(list(sizes.items()), _SIZE_TOLERANCE),
                key=lambda g: -max(v for _, v in g),
            )
            described = "; ".join(
                f"{_pt(max(v for _, v in g))} for "
                f"{', '.join(repr(n) for n in names if n in dict(g))}"
                for g in clusters
            )
            findings.append(_finding(
                slide_index, "uneven_sizes", names,
                f"A {kind} of like shapes is not one {field}: {described}.",
                f"Make them all {_pt(biggest)} {word}, unless one is meant to "
                f"stand out: ppt_update_shape shape_names={names} "
                f"{field}={_num(biggest)}.",
                severity="warning",
                sizes_pt={name: round(size, 1) for name, size in sizes.items()},
            ))
    return findings


def _cramped_text(slide_index, layout):
    """Text pressed against an edge of the card it sits on.

    The text does not cross the edge, so nothing overlaps, but it has next to
    no room on one side and plenty on the other, and reads as cut off. A card
    grown by exactly as much as its text overflowed leaves it just like this.
    """
    boxes = layout["boxes"]
    findings = []
    for shape in layout["arranged"]:
        if paints_box(shape):
            # Text in a shape's own frame is inset by the frame's margins, and
            # whether it fits there is the overflow check's question.
            continue
        words = text_box(shape)
        card = _card_of(shape, layout) if words else None
        if card is None:
            continue
        room = _room(words, boxes[card["name"]])
        if min(room.values()) < -_EDGE_TOLERANCE:
            continue  # It runs past the edge, which overlap reports.
        for side in ("bottom", "top", "right", "left"):
            want = _wanted_room(room, side)
            have = max(room[side], 0.0)
            if have >= _MIN_ROOM or have >= 0.5 * want:
                continue
            name, card_name, short = shape["name"], card["name"], want - have
            if side == "bottom":
                fix = _grow(card, "bottom", short, layout)
            elif side == "top":
                fix = f"move '{name}' down by {_pt(short)} (ppt_update_shape dtop={_num(short)})"
                if room["bottom"] - short < want:
                    fix += f", and {_grow(card, 'bottom', want - room['bottom'] + short, layout)}"
            elif side == "left":
                fix = f"move '{name}' right by {_pt(short)} (ppt_update_shape dleft={_num(short)})"
            else:
                fix = (f"make '{name}' {_pt(short)} narrower so its lines wrap "
                       f"sooner (ppt_update_shape dwidth={_num(-short)}), or "
                       f"{_grow(card, 'right', short, layout)}")
            opposite = _OPPOSITE[side]
            findings.append(_finding(
                slide_index, "cramped_text", [name, card_name],
                f"The text of '{name}' is {_pt(have)} from the {side} edge of "
                f"'{card_name}', with {_pt(room[opposite])} {_WHERE[opposite]}, "
                "so it looks pressed against the edge.",
                f"{fix[0].upper()}{fix[1:]}, which leaves it {_pt(want)} "
                f"{_WHERE[side]}{_padding(room, side, want)}.",
                severity="warning", room_pt={k: round(v, 1) for k, v in room.items()},
            ))
            break
    return findings


def _narrow_margins(slide_index, layout, slide_w):
    """Shapes closer to the side of the slide than the rest of it keeps.

    A shape touching the edge bleeds off it on purpose and is left alone, and
    one past it is off the slide, which is its own finding. What is left is a
    shape that stops short of the edge but inside the margin the other shapes
    keep: a row one card too far to the right, or a box pushed back onto the
    slide only just far enough. Only a near miss counts. A logo in the corner
    of a slide whose text is centred well in from the sides is placed there.
    """
    margins = layout["margins"]
    if not margins:
        return []
    side = margins["horizontal"]
    boxes = layout["boxes"]
    rows = {s["name"]: members for direction, members in layout["groups"]
            if direction == "horizontal" for s in members}
    findings = []
    for shape in layout["arranged"]:
        # Centred text is placed by what it labels, not by the margin. Off the
        # slide altogether, it is reported as such.
        if layout["depth"][shape["name"]] or _centred_text(shape):
            continue
        ink = ink_box(shape)
        if ink is None:
            continue
        for edge, distance in (("left", ink[0]), ("right", slide_w - ink[2])):
            if distance <= _EDGE_TOLERANCE or not _near_margin(distance, side):
                continue
            name, box = shape["name"], boxes[shape["name"]]
            row = rows.get(name)
            fix = _row_fix(row, boxes, slide_w, side) if row and _spans(row, boxes, slide_w) else None
            if fix is None and _width(box) > slide_w - 2 * side:
                fix = (f"Make '{name}' {_pt(_width(box) - slide_w + 2 * side)} "
                       f"narrower and move it to the margin (ppt_update_shape "
                       f"left={_num(side)} width={_num(slide_w - 2 * side)}).")
            elif fix is None:
                left = min(max(box[0], side), slide_w - side - _width(box))
                fix = (f"Move '{name}' {'right' if left > box[0] else 'left'} by "
                       f"{_pt(abs(left - box[0]))} (ppt_update_shape left={_num(left)}).")
            findings.append(_finding(
                slide_index, "narrow_margin", [name],
                f"'{name}' comes within {_pt(distance)} of the {edge} edge of "
                f"the slide, where the rest of the slide keeps {_pt(side)}.",
                fix, severity="warning", edge=edge,
                distance_pt=round(distance, 1), margin_pt=side,
            ))
            break
    return findings


# Text a template or a draft leaves for someone to replace. The words in any
# case. Stand-in figures only in capitals and only with a currency sign, a
# decimal point or a unit, so that neither a word nor a Roman numeral is one.
_SAMPLE_WORDS = re.compile(
    r"lorem\s+ipsum|click\s+to\s+(?:add|edit)\b"
    r"|\b(?:your|insert)\s+(?:text|title|name|company|logo|content|subtitle|tagline"
    r"|caption|image|photo|picture)\b"
    r"|(?:^|\n)\s*(?:add\s+)?(?:text|title|subtitle|caption|content)\s+(?:goes\s+)?here\b"
    r"|\[[^\]\n]{0,40}\b(?:insert|your|name|date|title|company|client|logo|image"
    r"|placeholder|amount|number)\b[^\]\n]{0,40}\]",
    re.IGNORECASE,
)
_SAMPLE_FIGURES = re.compile(
    r"(?<![A-Za-z0-9])(?:[$\u20ac\u00a3]XX+(?:[.,]X+)*[%KMB]?|XX+(?:[.,]X+)+[%KMB]?|XX+[%KMB])"
    r"(?![A-Za-z0-9])"
)


def _sample_text(slide_index, shapes):
    """Text a template or a draft left behind to be replaced.

    Lorem ipsum, "Your title here", "[Company name]", "XX%". A model fills a
    slide in passes and loses track of one, and on a finished slide it reads
    as unfinished work.
    """
    findings = []
    for shape in shapes:
        content = (shape.get("text") or {}).get("content") or ""
        if not content or not shape.get("visible", True):
            continue
        found = _SAMPLE_WORDS.search(content) or _SAMPLE_FIGURES.search(content)
        if not found:
            continue
        name = shape["name"]
        findings.append(_finding(
            slide_index, "sample_text", [name],
            f"'{name}' still has stand-in text: \"{found.group(0)}\".",
            f"Replace it with the real content (ppt_find_replace_text, or "
            f"ppt_set_text for all of '{name}'), or delete the shape if it is "
            "not needed.",
            severity="warning", sample=found.group(0),
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
    layout = _layout(shapes, slide_w, slide_h)
    findings = (
        _off_slide(index, shapes, slide_w, slide_h, layout)
        + _overflow(index, shapes)
        + _overlap(index, shapes, layout)
        + _low_contrast(index, shapes, slide.get("background_hex"))
        + _cramped_text(index, layout)
        + _narrow_margins(index, layout, slide_w)
        + _misaligned(index, layout, only=only)
        + _off_centre_labels(index, layout)
        + _uneven_sizes(index, layout)
        + _sample_text(index, shapes)
    )
    if only is not None:
        return _sorted([
            f for f in findings
            if f["type"] in SHAPE_FINDINGS and any(name in only for name in f["shapes"])
        ])
    spacing = _uneven_spacing(index, layout, slide_w)
    # A row whose gaps are reported is put back between the margins by that
    # finding's fix, so its cards are not told about the margin as well.
    spaced = {name for f in spacing for name in f["shapes"]}
    findings = [f for f in findings
                if not (f["type"] == "narrow_margin" and f["shapes"][0] in spaced)]
    findings += spacing + _empty_placeholders(index, shapes)
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
            # The title slide's own title, and a section header's, set lower,
            # are meant to differ from the rest. Where the text is drawn says
            # how low: a title placeholder's frame is often far taller than
            # the line in it.
            words = text_box(shape) or shape_box(shape)
            if shape.get("placeholder_type") == ppPlaceholderCenterTitle or (
                slide_h and words[1] > _TITLE_ZONE * slide_h
            ):
                return None
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
    """Titles that differ between slides.

    Flipping from one slide to the next, a title that is a size larger, a few
    points lower, or set in another weight is the thing that jumps, whatever
    layout each slide is on: a slide on the layout's own title placeholder
    among slides with a title box of their own jumps just the same. Only
    titles at the top of a slide are compared, so a cover's or a section
    header's title, set lower and meant to differ, is not held to the rest.
    Colours are compared only between titles on the same backdrop, since a
    light title on a dark slide is meant.
    """
    by_index = {slide["slide_index"]: slide for slide in slides}
    titled = []
    for slide in slides:
        title = find_title(slide, slide_h)
        if title is not None:
            titled.append((slide["slide_index"], title))
    if len(titled) < _MIN_TITLED_SLIDES:
        return []

    findings = []
    for label, read, tolerance, field in (
        ("size", _largest_size, _TITLE_SIZE_TOLERANCE, None),
        ("left edge", lambda s: _title_edge(s, "left"), _NEAR_MISS, "left"),
        ("top edge", lambda s: _title_edge(s, "top"), _NEAR_MISS, "top"),
    ):
        values = [(index, read(shape)) for index, shape in titled if read(shape) is not None]
        groups = _clusters(values, tolerance)
        if len(groups) < 2:
            continue
        usual = groups[0]
        usual_value = sorted(v for _, v in usual)[len(usual) // 2]
        described = "; ".join(
            f"{round(sorted(v for _, v in g)[len(g) // 2], 1):g} pt on "
            f"{_slides(i for i, _ in g)}"
            for g in groups
        )
        odd = [i for g in groups[1:] for i, _ in g]
        names = [shape["name"] for index, shape in titled if index in odd]
        if field is None:
            fix = (f"Set the titles on {_slides(odd)} to "
                   f"{round(usual_value, 1):g} pt (ppt_format_text "
                   f"font_size={round(usual_value, 1):g}).")
        else:
            moves = _title_moves(field, usual_value, odd, titled, by_index)
            fix = (f"Set {field} of the titles on {_slides(odd)} to "
                   f"{round(usual_value, 1):g} pt: {moves}.")
        findings.append(_finding(
            None, "inconsistent_titles", names,
            f"Slide titles differ in {label}: {described}.",
            fix, severity="warning", slides=sorted(odd),
        ))

    findings += _title_differences(
        "weight", [(index, title, _title_weight(title)) for index, title in titled],
        lambda value: f"in {value}" if value == "bold" else "in regular weight",
        lambda value: f"bold={'true' if value == 'bold' else 'false'}",
    )
    # A layout's title placeholder is set in the theme's heading font, and a
    # title box of one's own in the body font, so a deck mixing the two has
    # titles in two typefaces.
    findings += _title_differences(
        "font", [(index, title, _title_font(title)) for index, title in titled],
        lambda value: f"in {value}", lambda value: f"font_name='{value}'",
    )
    by_backdrop = {}
    for index, title in titled:
        slide = by_index[index]
        backdrop = _backdrop_colour(title, slide["shapes"], slide.get("background_hex"))
        if backdrop:
            by_backdrop.setdefault(backdrop, []).append((index, title, _title_colour(title)))
    for entries in by_backdrop.values():
        findings += _title_differences(
            "colour", entries, lambda value: f"to {value}",
            lambda value: f"color='{value}'",
        )
    return findings


def _title_edge(title, field):
    """Where a title's text starts, which is what jumps from slide to slide.

    The frame can say otherwise: a layout's title placeholder is often far
    taller than its line, with the line anchored at its middle or bottom.
    """
    box = text_box(title) or shape_box(title)
    return box[0] if field == "left" else box[1]


def _title_weight(title):
    flags = [bool(run.get("bold")) for run in (title.get("text") or {}).get("runs") or []
             if run.get("bold") is not None]
    if not flags:
        return None
    return "bold" if 2 * sum(flags) >= len(flags) else "regular"


def _title_font(title):
    fonts = [run["font"] for run in (title.get("text") or {}).get("runs") or []
             if run.get("font")]
    return max(set(fonts), key=fonts.count) if fonts else None


def _title_colour(title):
    colours = [run["color_hex"].upper() for run in (title.get("text") or {}).get("runs") or []
               if run.get("color_hex")]
    return max(set(colours), key=colours.count) if colours else None


def _title_differences(label, entries, as_words, as_call):
    """Titles that differ in a setting with no in-between, a weight or a colour."""
    groups = {}
    for index, title, value in entries:
        if value is not None:
            groups.setdefault(value, []).append((index, title))
    if len(groups) < 2 or sum(len(g) for g in groups.values()) < _MIN_TITLED_SLIDES:
        return []
    ordered = sorted(groups.items(), key=lambda item: (-len(item[1]), min(i for i, _ in item[1])))
    usual = ordered[0][0]
    odd = sorted(i for _, members in ordered[1:] for i, _ in members)
    names = [t["name"] for _, members in ordered[1:] for _, t in members]
    described = "; ".join(f"{value} on {_slides(i for i, _ in members)}" for value, members in ordered)
    return [_finding(
        None, "inconsistent_titles", names,
        f"Slide titles differ in {label}: {described}.",
        f"Set the titles on {_slides(odd)} {as_words(usual)} (ppt_format_text "
        f"{as_call(usual)}).",
        severity="warning", slides=odd,
    )]


def _title_moves(field, value, odd, titled, by_index):
    """The calls that put each odd title on ``value``, with what lines up with it.

    A slide's body is often set to its title's left edge. Moving the title
    alone would leave the body out of line with it, so whatever shares the
    edge moves with the title.
    """
    parts = []
    for index, title in titled:
        if index not in odd:
            continue
        by = value - _title_edge(title, field)
        if field == "top":
            # A title placeholder in the right place with its line anchored at
            # the bottom of a tall frame: anchoring it at the top puts the text
            # where the others are, without moving the frame off the slide.
            down = (_text_anchor(title) or (None, None))[1]
            margins = (title.get("text") or {}).get("margins") or _DEFAULT_MARGINS
            if down not in ("top", "full") and abs(title["top"] + margins["top"] - value) <= _NEAR_MISS:
                parts.append(f"on slide {index}, ppt_set_textframe vertical_anchor='top' "
                             f"on '{title['name']}', which sets its text at the top of "
                             "its frame, where the other titles' text is")
                continue
        along = [shape["name"] for shape in by_index[index]["shapes"]
                 if shape is not title and abs(shape[field] - title[field]) <= _ALIGNED]
        if along:
            names = [title["name"]] + along
            listed = ", ".join(f"'{n}'" for n in along)
            parts.append(
                f"on slide {index}, ppt_update_shape shape_names={names} "
                f"d{field}={_num(by)}, which moves {listed}, lined up with the "
                "title, along with it")
        else:
            parts.append(f"on slide {index}, ppt_update_shape {field}="
                         f"{_num(title[field] + by)} on '{title['name']}'")
    return "; ".join(parts)


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
    is lined up against anything. So are like shapes in a touched shape's row
    or column, since whether they paint decides whether the row has to be one
    size. Text overflowing its own frame far enough to reach a touched shape is
    the one meeting this misses, and that shape's own write already reported
    the overflow.
    """
    if shape.get("placeholder_type") is not None:
        return True
    box = shape_box(shape)
    for other in touched:
        other_box = shape_box(other)
        if _similar(box, other_box) and (
            _related(box, other_box, "x") or _related(box, other_box, "y")
        ):
            return True
        for reach in (other_box, text_box(other)):
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
        "content": content[:_CONTENT_KEPT],
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
                        font.Bold == msoTrue, getattr(font, "Name", None) or None))
        except Exception:
            continue
    facts["runs"] = [
        {"color_hex": colour, "size": size, "bold": bold, "font": name}
        for colour, size, bold, name in sorted(styles, key=lambda s: (s[0] or "", s[1] or 0))
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
        for facts, ref in zip(shapes, refs, strict=True):
            appearance(facts, ref, True)
        return shapes
    touched = []
    for facts, ref in zip(shapes, refs, strict=True):
        if facts["name"] in detail:
            touched.append(appearance(facts, ref, True))
    for facts, ref in zip(shapes, refs, strict=True):
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
        - cramped_text: text on a card pressed against one of its edges
        - narrow_margin: a shape closer to the side of the slide than the
          margin the rest of the slide keeps
        - misaligned: edges a few points apart that were meant to line up
        - uneven_spacing: a row or column of like shapes with unequal gaps
        - uneven_sizes: like shapes in a row (or column) that are not one
          height (or width)
        - empty_placeholder: a placeholder left without text
        - sample_text: text a template or a draft left to be replaced, such
          as lorem ipsum, "Your title here", "[Company name]" or "XX%"
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
