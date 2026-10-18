"""Tests for the layout checks behind ppt_check_layout and layout_issues.

Pure Python for the checks themselves: a slide is a list of plain dicts, the
shape ``_slide_facts_impl`` returns on both platforms, so these pin the
answers rather than the walk that gathers them. The Windows reader is then
run over fake COM objects, since it is the half CI exercises.
"""

import json
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, "src")

import pytest

from ppt_com import layout_check
from ppt_com.constants import (
    msoAutoShape, msoFalse, msoFillSolid, msoPicture, msoPlaceholder,
    msoTable, msoTextBox, msoTrue,
)
from ppt_com.layout_check import (
    deck_findings, find_title, ink_box, paints_box, report_layout,
    slide_findings, text_box, touched_names,
)

W, H = 960.0, 540.0
MARGINS = {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}


def text(width=150.0, height=30.0, left=None, top=None, chars=12,
         runs=None, alignment=None, anchor="top", word_wrap=True,
         autofit="none", orientation="horizontal", margins=MARGINS):
    return {
        "chars": chars, "left": left, "top": top, "width": width,
        "height": height, "alignment": alignment, "vertical_anchor": anchor,
        "runs": runs if runs is not None else [
            {"color_hex": "#222222", "size": 20.0, "bold": False}],
        "margins": margins, "word_wrap": word_wrap, "autofit": autofit,
        "orientation": orientation,
    }


def shape(name, left, top, width, height, z=1, type=msoAutoShape,
          fill=None, line=None, words=None, **extra):
    facts = {
        "name": name, "z": z, "type": type,
        "left": left, "top": top, "width": width, "height": height,
        "rotation": 0.0, "visible": True, "connector": False,
        "placeholder_type": None, "fill": fill, "line_visible": line,
        "text": words,
    }
    facts.update(extra)
    return facts


def textbox(name, left, top, width, height, z=1, text_width=150.0,
            text_height=30.0, **words):
    """A text box with no fill or line, its text at the top left inset."""
    return shape(
        name, left, top, width, height, z=z, type=msoTextBox,
        fill={"visible": False, "solid": True, "color_hex": "#FFFFFF", "transparency": 0.0},
        line=False,
        words=text(width=text_width, height=text_height,
                   left=left + 7.2, top=top + 3.6, **words),
    )


def card(name, left, top, width, height, z=1, colour="#FFFFFF", words=None):
    return shape(
        name, left, top, width, height, z=z,
        fill={"visible": True, "solid": True, "color_hex": colour, "transparency": 0.0},
        line=False, words=words,
    )


def slide(*shapes, index=1, background="#FFFFFF", layout="Blank"):
    return {"slide_index": index, "layout": layout,
            "background_hex": background, "shapes": list(shapes)}


def kinds(findings):
    return [f["type"] for f in findings]


def check(*shapes, **kwargs):
    return slide_findings(slide(*shapes, **kwargs), W, H)


def titled_row(lefts=(40, 340, 640), heights=(180, 180, 180)):
    """A title across the slide and a row of three cards under it, 40 pt in
    from both sides, as the live test slide was built."""
    title = textbox("title", 40, 30, 880, 55.74, z=1, text_width=300, text_height=48)
    return [title] + [card(f"card{i}", left, 140, 280, height, z=2 + i)
                      for i, (left, height) in enumerate(zip(lefts, heights, strict=True))]


def spill():
    """The paragraph on the first card, where PowerPoint for Mac drew it:
    its text from 153.6 to 345.6 pt, 25.6 pt past a 180 pt card."""
    return textbox("words", 52, 150, 256, 201.15, z=9,
                   text_width=245.07, text_height=192)


# ---------------------------------------------------------------------------
# What a shape puts on the slide
# ---------------------------------------------------------------------------
class TestWhatAShapeDraws:
    def test_a_filled_shape_draws_its_box(self):
        assert ink_box(card("c", 10, 20, 100, 50)) == (10, 20, 110, 70)

    def test_a_text_box_draws_only_its_text(self):
        box = textbox("t", 100, 100, 400, 200, text_width=150, text_height=30)
        assert ink_box(box) == pytest.approx((107.2, 103.6, 257.2, 133.6))

    def test_an_empty_text_box_draws_nothing(self):
        box = textbox("t", 100, 100, 400, 200)
        box["text"]["chars"] = 0
        assert ink_box(box) is None

    def test_a_hidden_shape_draws_nothing(self):
        assert ink_box(card("c", 10, 20, 100, 50) | {"visible": False}) is None

    def test_a_picture_draws_its_box_whatever_its_fill_says(self):
        assert paints_box(shape("p", 0, 0, 10, 10, type=msoPicture, fill={"visible": False}))

    def test_an_autoshape_whose_line_cannot_be_read_is_taken_as_outlined(self):
        # macOS cannot say whether a line is drawn, and a new autoshape has one.
        assert paints_box(shape("s", 0, 0, 10, 10, fill={"visible": False}, line=None))

    def test_a_text_box_whose_line_cannot_be_read_is_taken_as_not(self):
        assert not paints_box(shape("t", 0, 0, 10, 10, type=msoTextBox,
                                    fill={"visible": False}, line=None))


class TestWhereTextIsPlacedWhenPowerPointWillNotSay:
    def frame(self, **words):
        return shape("t", 100, 100, 400, 200, type=msoTextBox,
                     fill={"visible": False}, line=False,
                     words=text(**{"width": 100, "height": 40, **words}))

    def test_top_left_is_the_inset_corner(self):
        assert text_box(self.frame()) == pytest.approx((107.2, 103.6, 207.2, 143.6))

    def test_centred_text_is_centred_in_the_usable_width(self):
        left = text_box(self.frame(alignment=2))[0]
        assert left == pytest.approx(107.2 + (385.6 - 100) / 2)

    def test_right_aligned_text_ends_at_the_right_inset(self):
        assert text_box(self.frame(alignment=3))[2] == pytest.approx(500 - 7.2)

    def test_middle_anchored_text_is_centred_in_the_usable_height(self):
        top = text_box(self.frame(anchor="middle"))[1]
        assert top == pytest.approx(103.6 + (192.8 - 40) / 2)

    def test_bottom_anchored_text_ends_at_the_bottom_inset(self):
        assert text_box(self.frame(anchor="bottom"))[3] == pytest.approx(300 - 3.6)

    def test_reported_bounds_win_over_the_arithmetic(self):
        assert text_box(self.frame(left=5, top=6)) == (5, 6, 105, 46)

    def test_without_a_measured_size_there_is_no_answer(self):
        assert text_box(self.frame(width=None)) is None


# ---------------------------------------------------------------------------
# Off the slide
# ---------------------------------------------------------------------------
class TestOffSlide:
    def test_a_shape_past_the_right_edge_says_by_how_much(self):
        [finding] = check(card("c", 900, 100, 100, 50))
        assert finding["type"] == "off_slide"
        assert finding["edges"] == {"right": 40.0}
        # With nothing else on the slide there is no margin to keep.
        assert "left by at least 40 pt (ppt_update_shape left=860)" in finding["fix"]

    def test_it_comes_back_inside_the_margin_not_onto_the_edge(self):
        # Live: moved by the overshoot alone, the box ended flush against the
        # slide's right edge.
        [finding] = check(*titled_row(), card("box", 900, 470, 120, 50, z=9))
        assert finding["edges"] == {"right": 60.0}
        assert "left by 100 pt, to keep the 40 pt margin" in finding["fix"]
        assert "ppt_update_shape left=800" in finding["fix"]

    def test_a_frame_hanging_over_the_edge_with_its_text_inside_is_fine(self):
        assert check(textbox("t", 800, 100, 300, 50, text_width=100)) == []

    def test_text_spilling_off_the_bottom_is_not(self):
        box = textbox("t", 100, 480, 300, 50, text_height=90,
                      autofit="shape_to_fit")
        assert "off_slide" in kinds(check(box))

    def test_half_a_point_is_rounding(self):
        assert check(card("c", 860.4, 100, 100, 50)) == []

    def test_a_rotated_shape_is_measured_by_the_box_around_it(self):
        bar = card("bar", 380, 200, 200, 20) | {"rotation": 90.0}
        assert check(bar) == []
        bar["top"] = 5
        assert kinds(check(bar)) == ["off_slide"]


# ---------------------------------------------------------------------------
# Overflow
# ---------------------------------------------------------------------------
class TestOverflow:
    def test_text_taller_than_its_box_says_how_much_more_it_needs(self):
        box = textbox("t", 100, 100, 300, 40, text_height=60)
        [finding] = [f for f in check(box) if f["type"] == "overflow"]
        assert finding["needed_pt"] == 60
        assert finding["available_pt"] == pytest.approx(32.8)
        assert "27.2 pt taller" in finding["fix"]

    def test_text_that_fits_is_not_reported(self):
        assert check(textbox("t", 100, 100, 300, 40, text_height=30)) == []

    def test_a_line_too_long_for_an_unwrapped_box_is_a_width_problem(self):
        box = textbox("t", 100, 100, 200, 40, text_width=300, text_height=30,
                      word_wrap=False)
        [finding] = [f for f in check(box) if f["type"] == "overflow"]
        assert "wider" in finding["fix"]
        assert finding["needed_pt"] == 300


# ---------------------------------------------------------------------------
# Overlap
# ---------------------------------------------------------------------------
class TestOverlap:
    def test_two_texts_running_into_each_other_are_reported_once(self):
        found = check(textbox("a", 100, 100, 200, 40, z=1),
                      textbox("b", 100, 120, 200, 40, z=2))
        overlaps = [f for f in found if f["type"] == "overlap"]
        assert len(overlaps) == 1
        assert set(overlaps[0]["shapes"]) == {"a", "b"}
        assert overlaps[0]["overlap_pt"] == {"width": 150.0, "height": 10.0}

    def test_text_on_a_card_is_on_it_not_in_its_way(self):
        found = check(card("card", 80, 80, 300, 200, z=1),
                      textbox("words", 100, 100, 260, 60, z=2))
        assert "overlap" not in kinds(found)

    def test_text_running_past_the_bottom_of_its_card_says_how_far(self):
        found = check(card("card", 80, 80, 300, 100, z=1),
                      textbox("words", 100, 100, 260, 60, z=2, text_height=110))
        [finding] = [f for f in found if f["type"] == "overlap"]
        assert finding["past_edges"] == {"bottom": pytest.approx(33.6)}
        assert "sits on 'card'" in finding["message"]

    def test_an_icon_dropped_on_a_paragraph_is_in_the_way(self):
        found = check(textbox("para", 100, 100, 400, 100, z=1,
                              text_width=380, text_height=90),
                      shape("icon", 200, 120, 40, 40, z=2, type=msoPicture))
        [finding] = [f for f in found if f["type"] == "overlap"]
        assert "runs into 'icon'" in finding["message"]

    def test_the_card_grows_with_padding_and_its_row_with_it(self):
        # Live: "larger by that much" left the last line on the card's edge
        # and the card taller than the other two.
        [finding] = [f for f in check(*titled_row(), spill()) if f["type"] == "overlap"]
        assert finding["past_edges"] == {"bottom": 25.6}
        assert "make 'card0' 39.2 pt taller" in finding["fix"].lower()
        assert "shape_names=['card0', 'card1', 'card2'] height=219.2" in finding["fix"]
        assert "13.6 pt below, as it has above" in finding["fix"]

    def test_two_filled_shapes_crossing_are_left_alone(self):
        # A Venn diagram looks exactly like this.
        assert check(card("a", 100, 100, 200, 200, z=1),
                     card("b", 250, 100, 200, 200, z=2)) == []

    def test_touching_is_not_overlapping(self):
        found = check(textbox("a", 100, 100, 200, 40, z=1),
                      textbox("b", 100, 132, 200, 40, z=2))
        assert "overlap" not in kinds(found)


# ---------------------------------------------------------------------------
# Contrast
# ---------------------------------------------------------------------------
WHITE_TEXT = [{"color_hex": "#FFFFFF", "size": 20.0, "bold": False}]


class TestContrast:
    def test_white_text_on_a_white_slide_is_reported(self):
        [finding] = check(textbox("t", 100, 100, 300, 50, runs=WHITE_TEXT))
        assert finding["type"] == "low_contrast"
        assert finding["contrast"] == 1.0
        assert "#000000" in finding["fix"]

    def test_dark_text_on_a_white_slide_is_fine(self):
        assert check(textbox("t", 100, 100, 300, 50)) == []

    def test_the_card_underneath_is_what_the_text_is_read_against(self):
        found = check(card("card", 80, 80, 400, 200, z=1, colour="#1F3864"),
                      textbox("t", 100, 100, 300, 50, z=2, runs=WHITE_TEXT))
        assert "low_contrast" not in kinds(found)

    def test_text_on_a_picture_gets_no_verdict(self):
        found = check(shape("photo", 0, 0, 960, 540, z=1, type=msoPicture),
                      textbox("t", 100, 100, 300, 50, z=2, runs=WHITE_TEXT))
        assert "low_contrast" not in kinds(found)

    def test_a_see_through_fill_gets_no_verdict(self):
        glass = card("glass", 80, 80, 400, 200, z=1, colour="#FFFFFF")
        glass["fill"]["transparency"] = 0.6
        found = check(glass, textbox("t", 100, 100, 300, 50, z=2, runs=WHITE_TEXT))
        assert "low_contrast" not in kinds(found)

    def test_the_fix_names_the_call(self):
        [finding] = check(textbox("t", 100, 100, 300, 50, runs=WHITE_TEXT))
        assert "ppt_format_text color='#000000'" in finding["fix"]

    def test_the_fix_prefers_the_colour_the_other_text_uses(self):
        found = check(textbox("title", 40, 30, 880, 60, z=1),
                      textbox("note", 40, 440, 400, 40, z=2, runs=WHITE_TEXT))
        [finding] = found
        assert "#222222, the colour of the other text on the slide" in finding["fix"]
        assert "ppt_format_text color='#222222'" in finding["fix"]

    def test_large_text_is_held_to_the_lower_minimum(self):
        grey = [{"color_hex": "#949494", "size": 24.0, "bold": False}]
        assert check(textbox("t", 100, 100, 300, 50, runs=grey)) == []
        small = [{"color_hex": "#949494", "size": 12.0, "bold": False}]
        assert kinds(check(textbox("t", 100, 100, 300, 50, runs=small))) == ["low_contrast"]


# ---------------------------------------------------------------------------
# Alignment
# ---------------------------------------------------------------------------
def three_cards(second_top=200.0, third_left=560.0):
    return (card("one", 100, 200, 200, 100, z=1),
            card("two", 330, second_top, 200, 100, z=2),
            card("three", third_left, 200, 200, 100, z=3))


class TestMisaligned:
    def test_a_card_three_points_low_is_the_one_named(self):
        found = check(*three_cards(second_top=203))
        [finding] = [f for f in found if f["type"] == "misaligned"]
        assert finding["shapes"][0] == "two"
        assert set(finding["shapes"][1:]) == {"one", "three"}
        assert finding["edge"] == "top"
        assert "top=200" in finding["fix"]

    def test_ten_points_is_a_decision_not_a_slip(self):
        assert "misaligned" not in kinds(check(*three_cards(second_top=210)))

    def test_aligned_cards_are_fine(self):
        assert check(*three_cards()) == []

    def test_shapes_in_neither_a_row_nor_a_column_are_not_compared(self):
        found = check(card("a", 100, 100, 100, 50, z=1),
                      card("b", 400, 302, 100, 50, z=2))
        assert "misaligned" not in kinds(found)

    def test_between_two_the_later_one_moves(self):
        found = check(card("title", 40, 40, 880, 60, z=1),
                      card("body", 43, 120, 400, 300, z=2))
        [finding] = [f for f in found if f["type"] == "misaligned"]
        assert finding["shapes"] == ["body", "title"]
        assert "left=40" in finding["fix"]

    def test_text_on_a_card_is_not_lined_up_with_the_row(self):
        # Live on PowerPoint for Mac, once the cards had grown to hold this
        # text: the middle of its frame was 0.6 pt off the middle of the
        # cards, and the text was reported as out of line with them.
        found = check(card("one", 40, 140, 280, 220, z=1),
                      card("two", 340, 140, 280, 220, z=2),
                      card("three", 650, 140, 280, 220, z=3),
                      textbox("words", 52, 150, 256, 201.15, z=4,
                              text_width=245.07, text_height=192))
        assert "misaligned" not in kinds(found)

    def test_the_text_on_one_card_is_lined_up_with_the_next(self):
        lefts = (40, 340, 640)
        cards = [card(f"card{i}", left, 140, 280, 220, z=i + 1)
                 for i, left in enumerate(lefts)]
        heads = [textbox(f"head{i}", left + 20, 163 if i == 1 else 160, 240, 40,
                         z=4 + i, text_width=100)
                 for i, left in enumerate(lefts)]
        [finding] = [f for f in check(*cards, *heads) if f["type"] == "misaligned"]
        assert finding["shapes"][0] == "head1"
        assert set(finding["shapes"][1:]) == {"head0", "head2"}

    def test_only_the_shapes_on_the_line_are_named(self):
        found = check(card("one", 100, 200, 100, 100, z=1),
                      card("two", 230, 203, 100, 100, z=2),
                      card("three", 360, 200, 100, 100, z=3),
                      card("four", 490, 201.5, 100, 100, z=4))
        by_mover = {f["shapes"][0]: f for f in found if f["type"] == "misaligned"}
        assert set(by_mover) == {"two", "four"}
        assert set(by_mover["four"]["shapes"][1:]) == {"one", "three"}
        assert "'two'" not in by_mover["four"]["message"]


# ---------------------------------------------------------------------------
# Spacing
# ---------------------------------------------------------------------------
class TestUnevenSpacing:
    def test_gaps_of_thirty_and_twenty_three_are_reported(self):
        found = check(*three_cards(third_left=560))
        assert "uneven_spacing" not in kinds(found)
        found = check(card("one", 100, 200, 200, 100, z=1),
                      card("two", 324, 200, 200, 100, z=2),
                      card("three", 555, 200, 200, 100, z=3))
        [finding] = [f for f in found if f["type"] == "uneven_spacing"]
        assert finding["shapes"] == ["one", "two", "three"]
        assert finding["gaps_pt"] == [24.0, 31.0]
        assert "direction='horizontal'" in finding["fix"]

    def test_the_text_on_the_cards_is_not_a_second_row(self):
        cards = [card("one", 100, 200, 200, 100, z=1),
                 card("two", 324, 200, 200, 100, z=2),
                 card("three", 555, 200, 200, 100, z=3)]
        words = [textbox(f"t{i}", c["left"] + 10, 210, 180, 40, z=4 + i,
                         text_width=100)
                 for i, c in enumerate(cards)]
        found = check(*cards, *words)
        assert kinds(found).count("uneven_spacing") == 1

    def test_text_spilling_off_a_card_still_belongs_to_the_card(self):
        # Live on PowerPoint for Mac this hid the row: the spilt text box was
        # counted as a member, overlapped its card, and the row was dropped.
        found = check(card("one", 40, 140, 280, 180, z=1),
                      card("two", 340, 140, 280, 180, z=2),
                      card("three", 650, 140, 280, 180, z=3),
                      textbox("spill", 52, 150, 256, 201.15, z=4,
                              text_width=245.07, text_height=192))
        [finding] = [f for f in found if f["type"] == "uneven_spacing"]
        assert finding["gaps_pt"] == [20.0, 30.0]
        # The frame grew with its text until it was larger than the card,
        # 256 by 201 against 280 by 180, which is what PowerPoint did.

    def test_shapes_of_different_sizes_are_not_a_group(self):
        found = check(card("one", 100, 200, 100, 100, z=1),
                      card("two", 224, 200, 300, 100, z=2),
                      card("three", 555, 200, 50, 100, z=3))
        assert "uneven_spacing" not in kinds(found)

    def test_a_row_across_the_slide_is_spread_between_the_margins(self):
        # Live: the end card was the one 10 pt out, and distributing kept it
        # there, with the gaps even and the row off centre.
        found = check(*titled_row(lefts=(40, 340, 650)))
        [finding] = [f for f in found if f["type"] == "uneven_spacing"]
        assert finding["gaps_pt"] == [20.0, 30.0]
        assert "Set left of 'card2' to 640 pt" in finding["fix"]
        assert "Every gap is then 20 pt" in finding["fix"]
        # The same card is not reported a second time for the margin.
        assert "narrow_margin" not in kinds(found)


# ---------------------------------------------------------------------------
# Sizes, room on a card, margins
# ---------------------------------------------------------------------------
class TestUnevenSizes:
    def test_one_card_taller_than_its_row_is_reported(self):
        found = check(*titled_row(heights=(205.6, 180, 180)))
        [finding] = [f for f in found if f["type"] == "uneven_sizes"]
        assert finding["sizes_pt"] == {"card0": 205.6, "card1": 180.0, "card2": 180.0}
        assert "205.6 pt for 'card0'; 180 pt for 'card1', 'card2'" in finding["message"]
        assert "shape_names=['card0', 'card1', 'card2'] height=205.6" in finding["fix"]

    def test_a_point_or_two_is_the_same_size(self):
        assert "uneven_sizes" not in kinds(check(*titled_row(heights=(181.5, 180, 180))))

    def test_text_boxes_of_different_heights_are_not_judged(self):
        # Their frames cannot be seen, and text runs as long as it runs.
        found = check(textbox("a", 40, 140, 280, 100, z=1, text_height=90),
                      textbox("b", 340, 140, 280, 90, z=2, text_height=80),
                      textbox("c", 640, 140, 280, 110, z=3, text_height=100))
        assert "uneven_sizes" not in kinds(found)

    def test_a_write_to_one_card_hears_about_the_row(self):
        facts = slide(*titled_row(heights=(205.6, 180, 180)))
        found = slide_findings(facts, W, H, only={"card0"})
        assert kinds(found) == ["uneven_sizes"]


class TestCrampedText:
    def test_text_on_the_edge_of_its_card_is_reported(self):
        # Live, after the card grew by the overshoot alone.
        found = check(*titled_row(heights=(205.6, 205.6, 205.6)), spill())
        [finding] = [f for f in found if f["type"] == "cramped_text"]
        assert finding["shapes"] == ["words", "card0"]
        assert "0 pt from the bottom edge of 'card0', with 13.6 pt above" in finding["message"]
        assert "shape_names=['card0', 'card1', 'card2'] height=219.2" in finding["fix"]
        assert "overlap" not in kinds(found)

    def test_evenly_padded_text_is_fine(self):
        assert check(*titled_row(heights=(220, 220, 220)), spill()) == []

    def test_text_at_the_top_edge_moves_down_by_the_side_padding(self):
        found = check(card("card", 100, 100, 300, 200, z=1),
                      textbox("words", 112.8, 97.4, 280, 40, z=2,
                              text_width=200, text_height=40))
        [finding] = [f for f in found if f["type"] == "cramped_text"]
        assert "Move 'words' down by 19 pt (ppt_update_shape dtop=19)" in finding["fix"]

    def test_little_room_all_round_is_a_badge(self):
        found = check(card("pill", 100, 100, 80, 24, z=1),
                      textbox("label", 95.8, 99.4, 88.4, 25.2, z=2,
                              text_width=74, text_height=18))
        assert "cramped_text" not in kinds(found)


class TestNarrowMargins:
    def test_a_card_inside_the_margin_moves_the_row(self):
        # Live, after ppt_distribute_shapes: gaps of 25 pt, the row from 40
        # to 930 pt, so 10 pt short of the margin on the right.
        found = check(*titled_row(lefts=(40, 345, 650)))
        [finding] = found
        assert finding["type"] == "narrow_margin"
        assert finding["shapes"] == ["card2"]
        assert (finding["edge"], finding["distance_pt"], finding["margin_pt"]) == ("right", 30.0, 40.0)
        assert "Set left of 'card1' to 340 pt and of 'card2' to 640 pt" in finding["fix"]

    def test_a_shape_on_its_own_moves_in(self):
        found = check(*titled_row(), card("box", 808, 470, 120, 50, z=9))
        [finding] = found
        assert finding["distance_pt"] == 32.0
        assert "Move 'box' left by 8 pt (ppt_update_shape left=800)" in finding["fix"]

    def test_a_corner_element_well_inside_the_margin_is_placed_there(self):
        # A logo or a page number often keeps a narrower margin than the
        # content, and 20 pt against 40 is a decision rather than a slip.
        assert check(*titled_row(), card("logo", 820, 480, 120, 40, z=9)) == []

    def test_a_shape_bleeding_off_the_edge_is_meant(self):
        photo = shape("photo", 760, 0, 200, 540, z=9, type=msoPicture)
        found = check(textbox("title", 40, 30, 680, 60, z=1),
                      card("card", 40, 140, 280, 180, z=2), photo)
        assert "narrow_margin" not in kinds(found)

    def test_one_shape_is_not_a_margin(self):
        assert check(card("card", 10, 100, 300, 100)) == []

    def test_between_two_margins_the_narrower_is_kept(self):
        # Two edges keep 40 pt and two keep 30: nothing says which is meant.
        a, b = card("a", 40, 100, 880, 100, z=1), card("b", 30, 300, 900, 100, z=2)
        layout = layout_check._layout([a, b], W, H)
        assert layout["margins"]["horizontal"] == 30.0
        assert "narrow_margin" not in kinds(check(a, b))


# ---------------------------------------------------------------------------
# Placeholders
# ---------------------------------------------------------------------------
class TestEmptyPlaceholders:
    def placeholder(self, kind):
        return shape("Text Placeholder 2", 50, 150, 860, 300, type=msoPlaceholder,
                     fill={"visible": False}, line=False,
                     words=text(chars=0), placeholder_type=kind)

    def test_an_empty_body_placeholder_is_reported(self):
        [finding] = check(self.placeholder(2))
        assert finding["type"] == "empty_placeholder"
        assert "ppt_delete_shape" in finding["fix"]

    def test_an_empty_footer_belongs_to_the_layout(self):
        assert check(self.placeholder(15)) == []


# ---------------------------------------------------------------------------
# Titles across the deck
# ---------------------------------------------------------------------------
def titled(index, size=40.0, top=30.0, left=40.0, layout="Blank"):
    words = [{"color_hex": "#222222", "size": size, "bold": True}]
    title = textbox(f"Title {index}", left, top, 880, 60, runs=words)
    return slide(title, index=index, layout=layout)


class TestTitles:
    def test_one_title_a_size_smaller_is_reported(self):
        slides = [titled(2), titled(3), titled(4, size=36), titled(5)]
        [finding] = deck_findings(slides, H)
        assert finding["type"] == "inconsistent_titles"
        assert finding["slides"] == [4]
        assert "40 pt" in finding["fix"]

    def test_one_title_six_points_lower_is_reported(self):
        slides = [titled(2), titled(3, top=36), titled(4)]
        [finding] = deck_findings(slides, H)
        assert "top edge" in finding["message"]
        assert finding["slides"] == [3]

    def test_a_cover_title_in_the_middle_of_the_slide_is_not_a_title(self):
        cover = titled(1, size=60, top=220)
        assert find_title(cover, H) is None
        assert deck_findings([cover, titled(2), titled(3), titled(4)], H) == []

    def test_slides_on_different_layouts_are_compared(self):
        # Live: one slide on the Title and Content layout, its 44 pt title
        # placeholder among slides with 32 pt title boxes of their own.
        slides = [titled(2), titled(3), titled(4, size=36, layout="Title and Content")]
        [finding] = deck_findings(slides, H)
        assert finding["slides"] == [4]

    def test_a_title_placeholder_is_the_title_whatever_its_size(self):
        heading = shape("Title 1", 40, 30, 880, 60, type=msoPlaceholder,
                        placeholder_type=1, words=text())
        assert find_title(slide(heading), H) is heading

    def test_a_section_headers_title_lower_on_the_slide_is_not_compared(self):
        section = shape("Title 1", 40, 220, 880, 80, type=msoPlaceholder,
                        placeholder_type=1, words=text())
        assert find_title(slide(section, layout="Section Header"), H) is None
        assert deck_findings([titled(2), slide(section, index=3), titled(4), titled(5)], H) == []


# ---------------------------------------------------------------------------
# What a write tool reports
# ---------------------------------------------------------------------------
class TestOnlyTheTouchedShape:
    def test_findings_are_limited_to_the_shapes_named(self):
        shapes = (card("off", 900, 10, 100, 50, z=1), *three_cards(second_top=203))
        found = slide_findings(slide(*shapes), W, H, only={"two"})
        assert kinds(found) == ["misaligned"]

    def test_and_to_the_kinds_a_shape_can_cause(self):
        shapes = (card("one", 100, 200, 200, 100, z=1),
                  card("two", 324, 200, 200, 100, z=2),
                  card("three", 555, 200, 200, 100, z=3))
        assert slide_findings(slide(*shapes), W, H, only={"three"}) == []


class FakePpt:
    def __init__(self, answer):
        self.answer = answer
        self.calls = []

    def execute(self, func, *args):
        self.calls.append((func.__name__, args))
        if isinstance(self.answer, Exception):
            raise self.answer
        return self.answer


def facts(*shapes):
    return {"slide_width": W, "slide_height": H, "slides": [slide(*shapes)]}


class TestReportLayout:
    def test_a_problem_is_added_to_the_result(self, monkeypatch):
        fake = FakePpt(facts(card("c", 900, 100, 100, 50)))
        monkeypatch.setattr(layout_check, "ppt", fake)
        result = report_layout({"shape_name": "c"}, 1, ["c"])
        assert kinds(result["layout_issues"]) == ["off_slide"]
        assert fake.calls == [("_slide_facts_impl", ([1], False, ["c"]))]

    def test_nothing_is_added_when_nothing_is_wrong(self, monkeypatch):
        monkeypatch.setattr(layout_check, "ppt", FakePpt(facts(card("c", 10, 10, 50, 50))))
        assert report_layout({"shape_name": "c"}, 1, ["c"]) == {"shape_name": "c"}

    def test_a_check_that_could_not_run_says_so(self, monkeypatch):
        monkeypatch.setattr(layout_check, "ppt", FakePpt(RuntimeError("busy")))
        result = report_layout({"shape_name": "c"}, 1, ["c"])
        assert result["layout_check_failed"] == "busy"
        assert "layout_issues" not in result

    def test_a_failed_write_is_left_as_it_is(self, monkeypatch):
        fake = FakePpt(facts())
        monkeypatch.setattr(layout_check, "ppt", fake)
        assert report_layout({"error": "nope"}, 1, ["c"]) == {"error": "nope"}
        assert fake.calls == []

    def test_the_names_come_from_one_shape_or_several(self):
        assert touched_names({"shape_name": "a"}) == ["a"]
        assert touched_names({"updated": [{"shape_name": "a"}, {"shape_name": "b"}]}) == ["a", "b"]
        assert touched_names("not a dict") == []


# ---------------------------------------------------------------------------
# Reading a slide on Windows, over fake COM objects
# ---------------------------------------------------------------------------
class FakeRange:
    def __init__(self, content, runs, bounds=(107.2, 103.6, 150.0, 30.0), alignment=1):
        self.Text = content
        self._runs = runs
        if bounds is not None:
            self.BoundLeft, self.BoundTop, self.BoundWidth, self.BoundHeight = bounds
        self._alignment = alignment

    def Runs(self, i=None):
        if i is None:
            return SimpleNamespace(Count=len(self._runs))
        text, rgb, size, bold = self._runs[i - 1]
        return SimpleNamespace(Text=text, Font=SimpleNamespace(
            Color=SimpleNamespace(RGB=rgb), Size=size, Bold=bold,
            Name="Calibri Light" if size >= 30 else "Calibri"))

    def Paragraphs(self, i):
        return SimpleNamespace(ParagraphFormat=SimpleNamespace(Alignment=self._alignment))


def fake_shape(name="TextBox 1", content="Hello there", runs=None, bounds=(107.2, 103.6, 150.0, 30.0)):
    runs = runs if runs is not None else [("Hello there", 0x222222, 20.0, msoFalse)]
    tr = FakeRange(content, runs, bounds)
    frame = SimpleNamespace(
        TextRange=tr, WordWrap=msoTrue, VerticalAnchor=1, Orientation=1,
        MarginLeft=7.2, MarginRight=7.2, MarginTop=3.6, MarginBottom=3.6,
        HasText=msoTrue,
    )
    return SimpleNamespace(
        Name=name, Type=msoTextBox, Left=100.0, Top=100.0, Width=300.0, Height=50.0,
        Rotation=0.0, Visible=msoTrue, Connector=msoFalse, HasTextFrame=msoTrue,
        TextFrame=frame, TextFrame2=SimpleNamespace(AutoSize=0),
        Fill=SimpleNamespace(Visible=msoFalse, Type=msoFillSolid,
                             ForeColor=SimpleNamespace(RGB=0xFFFFFF), Transparency=0.0),
        Line=SimpleNamespace(Visible=msoFalse),
    )


class TestTheWindowsReader:
    def test_text_bounds_come_from_powerpoint(self):
        facts = layout_check._shape_facts(fake_shape(), 1, None)
        words = facts["text"]
        assert (words["left"], words["top"], words["width"], words["height"]) == (
            107.2, 103.6, 150.0, 30.0)
        assert words["margins"] == MARGINS
        assert facts["line_visible"] is False
        assert facts["fill"]["visible"] is False

    def test_runs_are_folded_into_distinct_styles(self):
        runs = [("Big ", 0x000000, 40.0, msoTrue), ("small", 0x000000, 12.0, msoFalse),
                ("   ", 0xFFFFFF, 99.0, msoFalse), ("again", 0x000000, 12.0, msoFalse)]
        words = layout_check._shape_facts(fake_shape(runs=runs), 1, None)["text"]
        assert words["runs"] == [
            {"color_hex": "#000000", "size": 12.0, "bold": False, "font": "Calibri"},
            {"color_hex": "#000000", "size": 40.0, "bold": True, "font": "Calibri Light"},
        ]

    def test_a_bound_powerpoint_will_not_give_stays_unknown(self):
        words = layout_check._shape_facts(fake_shape(bounds=None), 1, None)["text"]
        assert words["width"] is None and words["left"] is None

    def test_an_empty_frame_is_not_measured(self):
        words = layout_check._shape_facts(fake_shape(content="  ", runs=[]), 1, None)["text"]
        assert words["chars"] == 0
        assert words["width"] is None

    def test_a_shape_not_asked_about_is_read_only_for_where_its_text_is(self):
        words = layout_check._shape_facts(fake_shape(), 1, ["Another"])["text"]
        assert (words["left"], words["width"]) == (107.2, 150.0)
        assert words["runs"] == [] and words["margins"] is None

    def test_the_shape_asked_about_is_read_in_full(self):
        words = layout_check._shape_facts(fake_shape(), 1, ["TextBox 1"])["text"]
        assert words["runs"] and words["margins"] == MARGINS

    def test_alignment_is_read_only_when_powerpoint_will_not_place_the_text(self):
        assert layout_check._shape_facts(fake_shape(), 1, None)["text"]["alignment"] is None
        unplaced = layout_check._shape_facts(fake_shape(bounds=None), 1, None)["text"]
        assert unplaced["alignment"] == 1


class TestWhatAWriteReads:
    def test_a_shape_meeting_the_touched_one_is_read(self):
        touched = card("moved", 100, 100, 200, 100)
        assert layout_check.within_reach(card("near", 250, 150, 100, 100), [touched])

    def test_a_shape_far_from_it_is_not(self):
        touched = card("moved", 100, 100, 200, 100)
        assert not layout_check.within_reach(card("far", 600, 300, 100, 100), [touched])

    def test_the_touched_text_reaches_as_far_as_it_is_drawn(self):
        spilling = textbox("words", 100, 100, 200, 40, text_height=200)
        below = card("below", 100, 250, 200, 50)
        assert layout_check.within_reach(below, [spilling])

    def test_a_like_shape_in_its_row_is_read(self):
        touched = card("moved", 40, 140, 280, 180)
        assert layout_check.within_reach(card("third", 640, 140, 280, 180), [touched])
        assert not layout_check.within_reach(card("below", 640, 400, 280, 100), [touched])

    def test_a_placeholder_is_always_read(self):
        far = shape("Body", 600, 300, 100, 100, type=msoPlaceholder, placeholder_type=2)
        assert layout_check.within_reach(far, [card("moved", 0, 0, 10, 10)])

    def test_the_reading_follows_that(self):
        refs = ["moved", "near", "far"]
        boxes = {"moved": (100, 100), "near": (250, 150), "far": (600, 300)}
        read = []

        def geometry(ref, z):
            left, top = boxes[ref]
            return card(ref, left, top, 200, 100, z=z) | {"fill": None}

        def appearance(facts, ref, full):
            read.append((ref, full))
            return facts

        layout_check.read_shapes(refs, geometry, appearance, ["moved"])
        assert read == [("moved", True), ("near", False)]
        read.clear()
        layout_check.read_shapes(refs, geometry, appearance, None)
        assert read == [("moved", True), ("near", True), ("far", True)]

    def test_the_background_is_read_only_when_touched_text_needs_it(self):
        words = textbox("words", 100, 100, 200, 40)
        assert layout_check.needs_background([words], None)
        assert layout_check.needs_background([words], ["words"])
        assert not layout_check.needs_background([words, card("c", 0, 0, 5, 5)], ["c"])


# ---------------------------------------------------------------------------
# A real deck
# ---------------------------------------------------------------------------
# Fourteen slides built through the MCP server on PowerPoint for Mac the way
# a model builds them, read back as the checks read them, before and after
# their findings were fixed. Every finding below was checked against a
# screenshot of the slide, and so was every slide reported clean.
FIXTURES = Path(__file__).parent / "fixtures" / "layout"


def recorded(name):
    return json.loads((FIXTURES / f"{name}.json").read_text(encoding="utf-8"))


def findings_by_slide(deck):
    return {
        s["slide_index"]: [(f["type"], f["shapes"][0])
                           for f in slide_findings(s, deck["slide_width"], deck["slide_height"])]
        for s in deck["slides"]
    }


class TestARealDeck:
    SEEN = {
        1: [],                                   # cover: centred placeholders, corner logo
        2: [("misaligned", "TextBox 7")],         # agenda item 3, 2 pt low
        3: [("cramped_text", "TextBox 13")],      # the long KPI label on the card's edge
        4: [("misaligned", "Shape_234")],         # one card 2 pt low, with its contents
        5: [("off_slide", "TextBox 17"),          # the last timeline label, off the edge
            ("misaligned", "TextBox 17")],        # and 20 pt off its dot
        6: [],                                   # section band bleeding off both sides
        7: [],                                   # two columns
        8: [("overlap", "TextBox 3")],           # the caption into the table
        9: [],                                   # chart and takeaways
        10: [("misaligned", "TextBox 6")],       # a name 3 pt off its photo
        11: [],                                  # centred quote, corner logo
        12: [("low_contrast", "TextBox 10")],    # dark text on a dark card
        13: [],                                  # its title differs from the deck's
        14: [],                                  # closing slide, corner logo
    }

    def test_each_slide_reports_what_its_screenshot_shows(self):
        assert findings_by_slide(recorded("realdeck_broken_mac")) == self.SEEN

    def test_the_odd_title_is_reported_in_every_way_it_differs(self):
        deck = recorded("realdeck_broken_mac")
        found = deck_findings(deck["slides"], deck["slide_height"])
        assert {f["message"].split(":")[0] for f in found} == {
            f"Slide titles differ in {what}"
            for what in ("size", "left edge", "top edge", "weight", "colour")
        }
        assert all(f["slides"] == [13] for f in found)
        # Its frame was in place; its line was anchored at the bottom.
        [top] = [f for f in found if "top edge" in f["message"]]
        assert "vertical_anchor='top'" in top["fix"]

    def test_the_fixes_say_what_was_done(self):
        deck = recorded("realdeck_broken_mac")
        fixes = {}
        for s in deck["slides"]:
            for f in slide_findings(s, deck["slide_width"], deck["slide_height"]):
                fixes[(s["slide_index"], f["type"])] = f["fix"]
        # The KPI row grows as one.
        assert "shape_names=['Shape_225', 'Shape_222', 'Shape_223', 'Shape_224']" in fixes[(3, "cramped_text")]
        # The card moves with what is on it.
        assert "'Shape_235', 'TextBox 20', 'TextBox 21'" in fixes[(4, "misaligned")]
        # Centred on its dot, the label is shortened, never moved off it.
        assert fixes[(5, "off_slide")].startswith("Shorten the text of 'TextBox 17'")
        # The caption moves, not the table.
        assert fixes[(8, "overlap")].startswith("Move 'TextBox 3' down")
        # Body text takes the other body text's colour, not the headings'.
        assert "#CBD5E1" in fixes[(12, "low_contrast")]

    def test_the_fixed_deck_is_clean(self):
        deck = recorded("realdeck_fixed_mac")
        assert all(found == [] for found in findings_by_slide(deck).values())
        assert deck_findings(deck["slides"], deck["slide_height"]) == []


# ---------------------------------------------------------------------------
# What text keeps to
# ---------------------------------------------------------------------------
def placed(name, left, top, width, height, text_left, text_top, text_width, text_height, z=1):
    """A text box whose text PowerPoint drew at the given place."""
    box = textbox(name, left, top, width, height, z=z)
    box["text"].update(left=text_left, top=text_top, width=text_width, height=text_height)
    return box


class TestWhereTextKeeps:
    def test_left_set_text_hugs_the_left_inset(self):
        title = placed("t", 40, 30, 880, 46, 47.2, 33.6, 195, 38.4)
        assert layout_check._text_anchor(title) == ("left", "full")

    def test_centred_text_is_centred_even_a_few_points_off(self):
        # PowerPoint for Mac reported this label 3 pt off the centre it was set to.
        label = placed("l", 0, 300, 160, 23.02, 33.47, 303.6, 96.07, 15.6)
        assert layout_check._text_anchor(label) == ("center", "full")

    def test_right_set_text_hugs_the_right_inset(self):
        note = placed("n", 160, 320, 640, 30, 600, 323.6, 192.8, 22.8)
        assert layout_check._text_anchor(note)[0] == "right"

    def test_text_at_the_top_of_a_tall_frame(self):
        body = placed("b", 40, 100, 400, 300, 47.2, 103.6, 200, 60)
        assert layout_check._text_anchor(body) == ("left", "top")

    def test_a_reader_sees_only_the_edges_text_keeps_to(self):
        label = placed("l", 0, 300, 160, 23.02, 33.47, 303.6, 96.07, 15.6)
        assert layout_check._seen_edges(label) == {"center", "top", "middle", "bottom"}
        assert layout_check._seen_edges(card("c", 0, 0, 10, 10)) is None


# ---------------------------------------------------------------------------
# Margins, from more kinds of slide
# ---------------------------------------------------------------------------
class TestWhatAMarginIs:
    def test_centred_text_keeps_no_margin(self):
        # Live: a centred quote 160 pt in from both sides made a "160 pt
        # margin", and the corner logo was reported against it.
        quote = placed("quote", 160, 150, 640, 160, 230, 153.6, 500, 70)
        attribution = placed("by", 160, 320, 640, 30, 600, 323.6, 192.8, 22.8, z=2)
        logo = card("logo", 840, 468, 80, 32, z=3)
        assert check(quote, attribution, logo) == []

    def test_a_cover_on_its_placeholders_has_no_margin_to_break(self):
        title = placed("title", 120, 88, 720, 188, 290, 150, 380, 70)
        title.update(type=msoPlaceholder, placeholder_type=3)
        logo = card("logo", 840, 468, 80, 32, z=2)
        assert check(title, logo) == []

    def test_a_margin_is_kept_on_both_sides(self):
        # Live, the agenda: four circles 60 pt in on the left made 60 pt the
        # most common distance, and the title, 40 pt in, was reported.
        title = placed("title", 40, 30, 880, 46, 47.2, 33.6, 120, 38.4)
        dots = [card(f"dot{i}", 60, 130 + 80 * i, 44, 44, z=2 + i) for i in range(4)]
        layout = layout_check._layout([title, *dots], W, H)
        assert layout["margins"]["horizontal"] == 40.0
        assert "narrow_margin" not in kinds(check(title, *dots))


# ---------------------------------------------------------------------------
# Labels
# ---------------------------------------------------------------------------
class TestLabels:
    def test_a_label_off_the_edge_is_narrowed_around_its_centre(self):
        # Centred on its dot at 820 pt, 300 pt of text runs 10 pt off the edge.
        dot = card("dot", 810, 270, 20, 20, z=8)
        label = placed("label", 620, 300, 400, 23, 670, 303.6, 300, 15.6, z=9)
        found = check(*titled_row(heights=(100, 100, 100)), dot, label)
        [finding] = [f for f in found if f["type"] == "off_slide"]
        assert finding["fix"].startswith("Narrow 'label' to ")
        assert "around its centre" in finding["fix"]

    def test_a_label_well_off_its_dot_is_reported(self):
        dot = card("dot", 870, 270, 20, 20, z=1)
        label = placed("label", 780 + 20, 300, 200, 23, 821.5, 303.6, 160, 15.6, z=2)
        [finding] = [f for f in check(dot, label) if f["type"] == "misaligned"]
        assert finding["shapes"] == ["label", "dot"]
        assert "left=780" in finding["fix"]

    def test_a_centred_name_is_compared_by_its_centre_only(self):
        # Live: a name 3 pt off its photo was also reported against the
        # left edge of the frame under it, which no one can see.
        photo = card("photo", 310, 120, 110, 110, z=1)
        name = placed("name", 288, 245, 160, 29, 326.2, 248.6, 87.8, 21.6, z=2)
        role = placed("role", 285, 275, 160, 24, 343, 278.6, 47.1, 16.8, z=3)
        [finding] = [f for f in check(photo, name, role) if f["type"] == "misaligned"]
        assert finding["edge"] == "center"
        assert set(finding["shapes"]) == {"name", "photo", "role"}


# ---------------------------------------------------------------------------
# What else a real deck gets wrong
# ---------------------------------------------------------------------------
class TestSpacingIsForShapesThatPaint:
    def test_text_beside_a_row_of_icons_is_not_spaced(self):
        dots = [card(f"dot{i}", 60, 130 + 80 * i, 44, 44, z=1 + i) for i in range(4)]
        items = [textbox(f"item{i}", 124, 134 + 80 * i + (3 if i == 2 else 0), 600, 34,
                         z=5 + i, text_width=200, text_height=26.8) for i in range(4)]
        assert "uneven_spacing" not in kinds(check(*dots, *items))


class TestCollisions:
    def test_the_smaller_shape_moves_and_clears(self):
        table = shape("table", 40, 110, 880, 250, z=1, type=msoTable)
        caption = textbox("caption", 40, 350, 880, 30, z=2, text_width=160, text_height=15)
        [finding] = [f for f in check(table, caption) if f["type"] == "overlap"]
        assert finding["fix"].startswith("Move 'caption' down by ")
        assert "6 pt between it and 'table'" in finding["fix"]


class TestContrastSuggestions:
    def test_body_text_takes_the_colour_of_body_text(self):
        headings = [textbox(f"h{i}", 60 + 300 * i, 150, 240, 36, z=1 + i,
                            runs=[{"color_hex": "#F8FAFC", "size": 20.0, "bold": True}])
                    for i in range(3)]
        bodies = [textbox(f"b{i}", 60 + 300 * i, 196, 240, 100, z=4 + i,
                          runs=[{"color_hex": "#CBD5E1", "size": 16.0, "bold": False}])
                  for i in range(2)]
        dim = textbox("dim", 660, 196, 240, 100, z=9,
                      runs=[{"color_hex": "#475569", "size": 16.0, "bold": False}])
        found = check(*headings, *bodies, dim, background="#0F172A")
        [finding] = [f for f in found if f["type"] == "low_contrast"]
        assert "color='#CBD5E1'" in finding["fix"]


class TestDeckTitles:
    @staticmethod
    def titled_with(index, runs, background="#FFFFFF"):
        title = textbox(f"Title {index}", 40, 30, 880, 60, runs=runs)
        return slide(title, index=index, background=background)

    def test_a_title_in_another_typeface_is_reported(self):
        body = [{"color_hex": "#1F2937", "size": 32.0, "bold": True, "font": "Calibri"}]
        light = [{"color_hex": "#1F2937", "size": 32.0, "bold": True, "font": "Calibri Light"}]
        slides = [self.titled_with(i, body) for i in (2, 3, 4)] + [self.titled_with(5, light)]
        [finding] = deck_findings(slides, H)
        assert "font_name='Calibri'" in finding["fix"]

    def test_a_light_title_on_a_dark_slide_is_meant(self):
        dark = [{"color_hex": "#1F2937", "size": 32.0, "bold": True}]
        light = [{"color_hex": "#F8FAFC", "size": 32.0, "bold": True}]
        slides = [self.titled_with(i, dark) for i in (2, 3, 4)]
        slides.append(self.titled_with(5, light, background="#0F172A"))
        assert deck_findings(slides, H) == []

    def test_a_title_moving_takes_the_body_lined_up_with_it(self):
        def page(index, left):
            words = [{"color_hex": "#222222", "size": 32.0, "bold": True}]
            title = textbox(f"Title {index}", left, 30, 840, 60, runs=words)
            body = textbox(f"Body {index}", left, 140, 840, 300, z=2)
            return slide(title, body, index=index)
        slides = [page(2, 40), page(3, 40), page(4, 66)]
        [finding] = deck_findings(slides, H)
        assert "shape_names=['Title 4', 'Body 4'] dleft=-26" in finding["fix"]


class TestSampleText:
    @pytest.mark.parametrize("words, sample", [
        ("Lorem ipsum dolor sit amet", "Lorem ipsum"),
        ("Click to add text", "Click to add"),
        ("Your title here", "Your title"),
        ("Prepared for [Client name]", "[Client name]"),
        ("Revenue grew XX% this year", "XX%"),
        ("Margin of XX.X points", "XX.X"),
    ])
    def test_stand_in_text_is_reported(self, words, sample):
        box = textbox("t", 100, 100, 400, 40)
        box["text"]["content"] = words
        [finding] = check(box)
        assert finding["type"] == "sample_text"
        assert finding["sample"] == sample

    @pytest.mark.parametrize("words", [
        "Super Bowl XX", "Sizes up to XXL", "See note [1]", "Our text here is final",
        "The name of the game",
    ])
    def test_real_text_is_not(self, words):
        box = textbox("t", 100, 100, 400, 40)
        box["text"]["content"] = words
        assert check(box) == []
