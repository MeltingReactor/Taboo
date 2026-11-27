#!/usr/bin/env python3
"""Generate a Taboo! PowerPoint deck from a JSON card file.

Install dependencies:
    pip install python-pptx tqdm

`tqdm` is optional; the script has a built-in fallback bar. Use
`--fallback-bar` to force the fallback even when tqdm is installed.

Usage:
    python taboo_generator.py [config.json] [options]

The JSON file controls the template, output name, shuffle settings and every
section/card. Use either "categories" or "sections" as the top-level array.
Paths inside the JSON are resolved relative to the JSON file.

Minimal JSON example:
{
  "settings": {
    "template": "Taboo!.pptx",
    "output": "My_Taboo_Deck.pptx",
    "seed": 42,
    "shuffle_categories": true,
    "shuffle_cards": true
  },
  "sections": [
    {
      "name": "General",
      "cards": [
        {"word": "Umbrella", "banned": ["Rain", "Canopy", "Wet", "Collapsible"]}
      ]
    }
  ]
}

Optional per-card fields:
  "source_word": "Park"   - reuse an original template slide whose title is "Park"
  "source_slide": 5       - reuse template slide 5 (1-based) instead of cloning
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import random
import sys
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from pptx import Presentation

try:
    from tqdm import tqdm
except ImportError:  # Fallback progress bar keeps the script dependency-free.
    tqdm = None

DEFAULT_CONFIG = "taboo_sections.json"
DEFAULT_BANNED_COUNT = 4
DEFAULT_NUMBER_BOX_WIDTH = 600000
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


class ProgressBar:
    """Tiny dependency-free progress bar written to stderr."""

    def __init__(self, label: str, total: int, enabled: bool = True, width: int = 28):
        self.label = label
        self.total = max(1, int(total))
        self.width = width
        self.current = 0
        self.enabled = enabled
        self.stream = sys.stderr
        if self.enabled:
            self._render()

    def _render(self) -> None:
        fraction = min(1.0, self.current / self.total)
        filled = int(self.width * fraction)
        bar = "#" * filled + "." * (self.width - filled)
        self.stream.write(
            f"\r{self.label}: [{bar}] {self.current}/{self.total} ({fraction * 100:5.1f}%)"
        )
        self.stream.flush()

    def advance(self, amount: int = 1) -> None:
        self.current = min(self.total, self.current + amount)
        if self.enabled:
            self._render()

    def close(self) -> None:
        if self.enabled:
            if self.current < self.total:
                self._render()
            self.stream.write("\n")
            self.stream.flush()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()


class TqdmBar:
    """Small adapter so tqdm supports our advance()/context-manager API."""

    def __init__(self, label: str, total: int, enabled: bool):
        self.bar = None
        if enabled:
            self.bar = tqdm(
                total=total,
                desc=label,
                unit="step",
                file=sys.stderr,
                leave=True,
            )

    def advance(self, amount: int = 1) -> None:
        if self.bar is not None:
            self.bar.update(amount)

    def close(self) -> None:
        if self.bar is not None:
            self.bar.close()
            self.bar = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()


def progress_bar(
    label: str,
    total: int,
    enabled: bool = True,
    force: bool = False,
    use_fallback: bool = False,
):
    """Create a progress bar.

    Uses tqdm by default. Pass ``use_fallback=True`` (or ``--fallback-bar`` on
    the CLI) to force the dependency-free ProgressBar instead.
    """
    show = enabled and (force or use_fallback or sys.stderr.isatty())
    if use_fallback or tqdm is None:
        return ProgressBar(label, total, enabled=show)
    return TqdmBar(label, total, show)


def clear_screen() -> None:
    """Clear the terminal before printing the final summary."""
    os.system("cls" if os.name == "nt" else "clear")


def progress_was_shown(
    enabled: bool, force: bool = False, fallback: bool = False
) -> bool:
    """Mirror progress_bar()'s visibility decision for summary clearing."""
    return enabled and (force or fallback or sys.stderr.isatty())


# ---------------------------------------------------------------------------
# Small shape/text helpers
# ---------------------------------------------------------------------------

def shape_by_name(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise ValueError(f"Shape {name!r} not found on slide")


def set_line(shape, text: str) -> None:
    """Replace text while keeping the shape's existing single-run formatting."""
    tf = shape.text_frame
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)
    p = tf.paragraphs[0]
    runs = list(p.runs)
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        p.add_run().text = text


def find_number_shape(slide, slide_width: int, slide_height: int):
    """Return the bottom-right card-number textbox, if the slide has one."""
    for shape in slide.shapes:
        if shape.name == "TextBox 8" and shape.has_text_frame:
            return shape
    candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        bottom_right = (
            shape.left >= slide_width * 0.85 and shape.top >= slide_height * 0.85
        )
        if bottom_right and (text.isdigit() or text == ""):
            candidates.append(shape)
    return max(candidates, key=lambda s: (s.left, s.top)) if candidates else None


def widen_number_shape(shape, slide_width: int, width: int) -> None:
    """Widen the number box while keeping its original right margin."""
    width = max(1, int(width))
    right_margin = slide_width - (shape.left + shape.width)
    if right_margin < 0:
        right_margin = 62818
    shape.width = width
    shape.left = max(0, slide_width - right_margin - width)


def renumber_number_boxes(
    prs,
    start: int,
    width: int | None,
    progress: Any | None = None,
) -> list[tuple[int, str, int]]:
    """Renumber all numbered slides in final slide order; optionally widen boxes."""
    changes = []
    number = start
    for slide_index, slide in enumerate(prs.slides, start=1):
        if progress is not None:
            progress.advance()
        shape = find_number_shape(slide, prs.slide_width, prs.slide_height)
        if shape is None:
            continue
        old = shape.text_frame.text.strip()
        if width is not None:
            widen_number_shape(shape, prs.slide_width, width)
        set_line(shape, str(number))
        changes.append((slide_index, old, number))
        number += 1
    return changes


def set_transition(slide, effect: str, direction: str, speed: str = "slow") -> None:
    """Apply a standard PresentationML slide transition."""
    sld = slide._element
    qn = lambda tag: f"{{{P_NS}}}{tag}"
    existing = sld.find(qn("transition"))
    if existing is not None:
        speed = existing.get("spd", speed)
        position = sld.index(existing)
        sld.remove(existing)
    else:
        position = None

    transition = etree.Element(qn("transition"), spd=speed)
    etree.SubElement(transition, qn(effect), dir=direction)

    if position is None:
        position = 0
        for idx, child in enumerate(sld):
            if etree.QName(child).localname in ("cSld", "clrMapOvr"):
                position = idx + 1
    sld.insert(position, transition)


def existing_section_names(prs) -> tuple[str, str]:
    """Read the template's first/last section names, falling back to defaults."""
    sections = prs.part._element.findall(f".//{{{P14_NS}}}section")
    names = [sec.get("name", "") for sec in sections if sec.get("name")]
    if len(names) >= 2:
        return names[0], names[-1]
    return "Intro", "End"


def set_sections(prs, sections: list[tuple[str, list]]) -> None:
    """Replace the presentation's PowerPoint section list."""
    pres = prs.part._element
    ns = {"p": P_NS}

    ext_lst = pres.find("p:extLst", ns)
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, f"{{{P_NS}}}extLst")

    section_ext = None
    for ext in ext_lst.findall("p:ext", ns):
        if ext.get("uri") == SECTION_EXT_URI:
            section_ext = ext
            break
    if section_ext is None:
        section_ext = etree.SubElement(ext_lst, f"{{{P_NS}}}ext", uri=SECTION_EXT_URI)

    for old in section_ext.findall(f"{{{P14_NS}}}sectionLst"):
        section_ext.remove(old)
    section_lst = etree.SubElement(
        section_ext, f"{{{P14_NS}}}sectionLst", nsmap={"p14": P14_NS}
    )

    sld_id_lst = pres.find("p:sldIdLst", ns)
    id_by_part = {
        prs.part.rels[sid.get(f"{{{REL_NS}}}id")].target_part: sid.get("id")
        for sid in sld_id_lst
    }

    for index, (name, slides) in enumerate(sections, start=1):
        guid = "{%08X-0000-4000-8000-%012X}" % (index, index)
        section = etree.SubElement(section_lst, f"{{{P14_NS}}}section", name=name, id=guid)
        id_list = etree.SubElement(section, f"{{{P14_NS}}}sldIdLst")
        for slide in slides:
            slide_id = id_by_part.get(slide.part)
            if slide_id is not None:
                etree.SubElement(id_list, f"{{{P14_NS}}}sldId", id=slide_id)


def clean_cloned_ids(slide) -> None:
    """Strip duplicated creationId metadata from cloned shapes."""
    targets = [el for el in slide._element.iter() if el.tag.endswith("}creationId")]
    for el in targets:
        parent = el.getparent()
        if parent is None:
            continue
        parent.remove(el)
        if len(parent) == 0 and parent.tag.endswith("}extLst"):
            grandparent = parent.getparent()
            if grandparent is not None:
                grandparent.remove(parent)


def clone_slide(prs, source_slide):
    """Clone a slide's shapes onto a new slide using the same layout."""
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)
    for shape in source_slide.shapes:
        new_slide.shapes._spTree.insert_element_before(
            copy.deepcopy(shape._element), "p:extLst"
        )
    clean_cloned_ids(new_slide)
    return new_slide


def slide_at(prs, index: int):
    """Resolve a 1-based slide index; negative values count from the end."""
    count = len(prs.slides)
    if index == 0:
        raise ValueError("Slide indices are 1-based; 0 is not valid")
    if index > 0:
        if index > count:
            raise ValueError(f"Slide {index} does not exist (template has {count})")
        return prs.slides[index - 1]
    resolved = count + index + 1
    if resolved < 1:
        raise ValueError(f"Slide {index} is out of range")
    return prs.slides[resolved - 1]


def is_word_slide(slide) -> bool:
    names = {s.name for s in slide.shapes}
    return {"Title 1", "TextBox 7", "TextBox 8"} <= names


def is_divider_slide(slide) -> bool:
    names = {s.name for s in slide.shapes}
    return "Title 1" in names and "TextBox 8" not in names and len(slide.shapes) <= 2


def slide_title(slide) -> str:
    for shape in slide.shapes:
        if shape.name in ("Title 1", "Title 2"):
            return shape.text_frame.text
    return ""


# ---------------------------------------------------------------------------
# Config loading and validation
# ---------------------------------------------------------------------------

def resolve_path(value: str | Path, config_dir: Path) -> Path:
    path = Path(value).expanduser()
    return path if path.is_absolute() else (config_dir / path).resolve()


def load_config(config_path: Path) -> tuple[dict[str, Any], Path]:
    if not config_path.is_file():
        raise FileNotFoundError(f"Config not found: {config_path}")
    with config_path.open("r", encoding="utf-8") as fh:
        config = json.load(fh)
    if not isinstance(config, dict):
        raise ValueError("Config root must be a JSON object")
    config["_section_label"] = "section" if "sections" in config else "category"
    if "categories" not in config and "sections" in config:
        config["categories"] = config["sections"]
    if not isinstance(config.get("categories"), list):
        raise ValueError("Config must contain a 'categories' or 'sections' array")
    return config, config_path.parent.resolve()


def validate_categories(config: dict[str, Any], banned_count: int) -> list[dict[str, Any]]:
    categories = []
    seen_words: set[str] = set()
    for ci, category in enumerate(config["categories"], start=1):
        if not isinstance(category, dict):
            raise ValueError(f"Category #{ci} must be an object")
        name = category.get("name")
        cards = category.get("cards")
        if not isinstance(name, str) or not name.strip():
            raise ValueError(f"Category #{ci} needs a non-empty 'name'")
        if not isinstance(cards, list) or not cards:
            raise ValueError(f"Category {name!r} needs a non-empty 'cards' array")

        clean_cards = []
        for ki, card in enumerate(cards, start=1):
            if not isinstance(card, dict):
                raise ValueError(f"Card #{ki} in {name!r} must be an object")
            word = card.get("word")
            bans = card.get("banned")
            if not isinstance(word, str) or not word.strip():
                raise ValueError(f"Card #{ki} in {name!r} needs a non-empty 'word'")
            if not isinstance(bans, list) or len(bans) != banned_count:
                raise ValueError(
                    f"Card {word!r} in {name!r} needs exactly {banned_count} banned words"
                )
            if any(not isinstance(b, str) or not b.strip() for b in bans):
                raise ValueError(f"Card {word!r} has an empty banned word")
            if word in seen_words:
                print(f"Warning: duplicate guess word {word!r}", file=sys.stderr)
            seen_words.add(word)
            clean_card = dict(card)
            clean_card["word"] = word.strip()
            clean_card["banned"] = [b.strip() for b in bans]
            clean_cards.append(clean_card)
        categories.append({"name": name.strip(), "cards": clean_cards})
    return categories


# ---------------------------------------------------------------------------
# Deck generation
# ---------------------------------------------------------------------------

def discover_template(prs, settings: dict[str, Any]):
    front_indices = settings.get("front_slide_indices", [1, 2, 3])
    final_index = settings.get("final_slide_index", -1)
    if not isinstance(front_indices, list) or not front_indices:
        raise ValueError("'front_slide_indices' must be a non-empty array")
    front = [slide_at(prs, int(i)) for i in front_indices]
    final = slide_at(prs, int(final_index))
    reserved = {id(s) for s in front} | {id(final)}

    word_template = None
    divider_template = None
    word_by_title: dict[str, Any] = {}
    word_by_index: dict[int, Any] = {}
    divider_by_title: dict[str, Any] = {}

    for idx, slide in enumerate(prs.slides, start=1):
        if id(slide) in reserved:
            continue
        if is_word_slide(slide):
            if word_template is None:
                word_template = slide
            word_by_title[slide_title(slide)] = slide
            word_by_index[idx] = slide
        elif is_divider_slide(slide):
            if divider_template is None:
                divider_template = slide
            divider_by_title[slide_title(slide)] = slide

    if word_template is None:
        raise ValueError(
            "No word slide found in the template. Use the original Taboo!.pptx as template "
            "or adjust 'front_slide_indices'."
        )
    if divider_template is None:
        raise ValueError(
            "No category divider slide found in the template. Use the original Taboo!.pptx."
        )
    return front, final, word_template, divider_template, word_by_title, word_by_index, divider_by_title


def fill_word_slide(slide, card: dict[str, Any], category: str, footer_format: str) -> None:
    set_line(shape_by_name(slide, "Title 1"), card["word"])
    for offset, ban in enumerate(card["banned"], start=3):
        set_line(shape_by_name(slide, f"Rectangle: Rounded Corners {offset}"), ban)
    set_line(
        shape_by_name(slide, "TextBox 7"),
        footer_format.format(category=category, section=category),
    )


def reorder_slides(prs, ordered_slides) -> None:
    sld_id_lst = prs.slides._sldIdLst
    part_to_sid = {prs.part.rels[sid.rId].target_part: sid for sid in list(sld_id_lst)}
    for sid in list(sld_id_lst):
        sld_id_lst.remove(sid)
    for slide in ordered_slides:
        sld_id_lst.append(part_to_sid[slide.part])


def relink_begin_button(prs, first_card, button_text: str) -> bool:
    needle = button_text.lower()
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if needle in shape.text_frame.text.lower() and "hlinkClick" in shape._element.xml:
                shape.click_action.target_slide = first_card
                return True
    return False


def patch_app_props(path: Path, slide_count: int, titles: list[str]) -> None:
    """Keep PowerPoint's stored slide count/title list in sync."""
    ns = {"ep": EP_NS, "vt": VT_NS}
    with ZipFile(path, "r") as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]

    patched = []
    for info, data in entries:
        if info.filename == "docProps/app.xml":
            root = etree.fromstring(data)
            slides_el = root.find("ep:Slides", ns)
            if slides_el is not None:
                slides_el.text = str(slide_count)
            heading = root.find("ep:HeadingPairs/vt:vector", ns)
            if heading is not None:
                variants = heading.findall("vt:variant", ns)
                if len(variants) >= 6:
                    count_el = variants[5].find("vt:i4", ns)
                    if count_el is not None:
                        count_el.text = str(slide_count)
            tvec = root.find("ep:TitlesOfParts/vt:vector", ns)
            if tvec is not None:
                preserved = [el.text or "" for el in tvec.findall("vt:lpstr", ns)[:5]]
                for child in list(tvec):
                    tvec.remove(child)
                for value in preserved + titles:
                    etree.SubElement(tvec, f"{{{VT_NS}}}lpstr").text = value
                tvec.set("size", str(len(preserved) + len(titles)))
            data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        patched.append((info, data))

    with ZipFile(path, "w", ZIP_DEFLATED) as zout:
        for info, data in patched:
            zout.writestr(info, data)


def generate(
    config: dict[str, Any],
    config_dir: Path,
    overrides: dict[str, Any],
    progress_enabled: bool = True,
    force_progress: bool = False,
    fallback_bar: bool = False,
) -> dict[str, Any]:
    settings = dict(config.get("settings", {}))
    settings.update({k: v for k, v in overrides.items() if v is not None})

    template = resolve_path(settings.get("template", "Taboo!.pptx"), config_dir)
    output = resolve_path(settings.get("output", "Taboo!_Generated.pptx"), config_dir)
    if not template.is_file():
        raise FileNotFoundError(f"Template not found: {template}")
    output.parent.mkdir(parents=True, exist_ok=True)

    banned_count = int(settings.get("banned_count", DEFAULT_BANNED_COUNT))
    categories = validate_categories(config, banned_count)
    footer_format = str(settings.get("footer_format", "Taboo • {category}"))
    number_start = int(settings.get("number_start", 1))
    number_box_width = settings.get("number_box_width")
    number_box_width = None if number_box_width is None else int(number_box_width)

    total_cards = sum(len(c["cards"]) for c in categories)
    progress_enabled = progress_enabled and bool(settings.get("show_progress", True))
    force_progress = force_progress or bool(settings.get("force_progress", False))
    use_fallback = fallback_bar or bool(settings.get("fallback_bar", False))
    # 1 template load + one step per divider/card + 7 finalization steps.
    total_steps = 1 + len(categories) + total_cards + 7

    bar = progress_bar(
        "Generating deck", total_steps, progress_enabled, force_progress, use_fallback
    )
    try:
        prs = Presentation(template)
        bar.advance()
        (front, final, word_template, divider_template,
         word_by_title, word_by_index, divider_by_title) = discover_template(prs, settings)
        template_intro_name, template_end_name = existing_section_names(prs)
        intro_name = str(settings.get("intro_section_name", template_intro_name))
        end_name = str(settings.get("end_section_name", template_end_name))

        rng = random.Random(settings.get("seed"))
        shuffle_categories = bool(
            settings.get("shuffle_sections", settings.get("shuffle_categories", True))
        )
        shuffle_cards = bool(settings.get("shuffle_cards", True))
        if shuffle_categories:
            rng.shuffle(categories)
        if shuffle_cards:
            for category in categories:
                rng.shuffle(category["cards"])

        ordered_slides = list(front)
        titles: list[str] = []
        card_slides = []
        divider_slides = []
        section_plan: list[tuple[str, list]] = [(intro_name, list(front))]
        card_count = 0
        reused = 0

        for category in categories:
            divider = divider_by_title.get(category["name"])
            if divider is None:
                divider = clone_slide(prs, divider_template)
            else:
                reused += 1
            set_line(shape_by_name(divider, "Title 1"), category["name"])
            ordered_slides.append(divider)
            titles.append(category["name"])
            divider_slides.append(divider)
            category_slides = [divider]
            bar.advance()

            for card in category["cards"]:
                source = None
                if card.get("source_slide") is not None:
                    source = word_by_index.get(int(card["source_slide"]))
                    if source is None:
                        raise ValueError(
                            f"source_slide {card['source_slide']!r} for {card['word']!r} "
                            "is not a word slide in the template"
                        )
                elif card.get("source_word") is not None:
                    source = word_by_title.get(str(card["source_word"]))
                    if source is None:
                        raise ValueError(
                            f"source_word {card['source_word']!r} for {card['word']!r} "
                            "was not found in the template"
                        )
                if source is None:
                    source = clone_slide(prs, word_template)
                else:
                    reused += 1

                fill_word_slide(source, card, category["name"], footer_format)
                ordered_slides.append(source)
                titles.append(card["word"])
                card_slides.append(source)
                category_slides.append(source)
                card_count += 1
                bar.advance()

            section_plan.append((category["name"], category_slides))

        ordered_slides.append(final)
        titles.append(slide_title(final) or "Well done!")
        section_plan.append((end_name, [final]))

        reorder_slides(prs, ordered_slides)
        bar.advance()
        first_card = ordered_slides[len(front) + 1]
        if not relink_begin_button(
            prs, first_card, str(settings.get("begin_button_text", "Begin!"))
        ):
            print("Warning: no 'Begin!' hyperlink button found in template", file=sys.stderr)
        bar.advance()

        renumbered = renumber_number_boxes(prs, number_start, number_box_width)
        bar.advance()

        # PowerPoint names the origin edge: "From Right" means the incoming
        # slide travels left, so the OOXML direction is "l". "From Top" -> "d".
        for slide in card_slides:
            set_transition(slide, "cover", "l")
        for slide in divider_slides:
            set_transition(slide, "push", "d")
        bar.advance()

        set_sections(prs, section_plan)
        bar.advance()

        prs.save(output)
        bar.advance()
        patch_app_props(output, len(ordered_slides), titles)
        bar.advance()
    finally:
        bar.close()

    return {
        "output": str(output),
        "slides": len(ordered_slides),
        "categories": len(categories),
        "cards": card_count,
        "reused_template_slides": reused,
        "renumbered": len(renumbered),
        "sections": [name for name, _ in section_plan],
        "section_label": str(config.get("_section_label", "category")),
        "section_order": [c["name"] for c in categories],
    }


def fix_existing_deck(
    source: Path,
    output: Path | None,
    in_place: bool,
    start: int,
    width: int | None,
    progress_enabled: bool = True,
    force_progress: bool = False,
    fallback_bar: bool = False,
) -> tuple[list[tuple[int, str, int]], Path]:
    """Renumber and widen number boxes in an existing deck."""
    if not source.is_file():
        raise FileNotFoundError(f"Deck not found: {source}")
    prs = Presentation(source)
    with progress_bar(
        "Renumbering", len(prs.slides), progress_enabled, force_progress, fallback_bar
    ) as bar:
        changes = renumber_number_boxes(prs, start, width, bar)
    if not changes:
        raise ValueError("No number boxes found in the deck")

    if output is None:
        output = source.with_name(f"{source.stem}_FixedNumbers{source.suffix}")
    output = output.resolve()
    if output == source:
        in_place = True
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_name(output.name + ".tmp")
    try:
        prs.save(temp)
        temp.replace(source if in_place else output)
    finally:
        if temp.exists():
            temp.unlink()
    return changes, source if in_place else output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("config", nargs="?", default=DEFAULT_CONFIG, help="JSON config file")
    parser.add_argument("--template", help="Override template .pptx path")
    parser.add_argument("--output", help="Override output .pptx path")
    parser.add_argument("--seed", type=int, help="Override shuffle seed")
    parser.add_argument("--no-shuffle", action="store_true", help="Keep the JSON category/card order")
    parser.add_argument(
        "--fix-numbers",
        metavar="PPTX",
        help="Renumber an existing deck instead of generating from JSON",
    )
    parser.add_argument("--in-place", action="store_true", help="Overwrite input with --fix-numbers")
    parser.add_argument("--start", type=int, default=1, help="First card number (default: 1)")
    parser.add_argument(
        "--number-box-width",
        type=int,
        default=None,
        help=f"Number box width in EMU (default: {DEFAULT_NUMBER_BOX_WIDTH} from JSON; 0 = leave unchanged)",
    )
    parser.add_argument("--no-progress", action="store_true", help="Hide the progress bar")
    parser.add_argument(
        "--force-progress",
        action="store_true",
        help="Show the progress bar even when output is redirected",
    )
    parser.add_argument(
        "--fallback-bar",
        action="store_true",
        help="Force the built-in fallback bar instead of tqdm",
    )
    args = parser.parse_args(argv)

    progress_enabled = not args.no_progress
    force_progress = args.force_progress
    fallback_bar = args.fallback_bar
    clear_before_summary = progress_was_shown(
        progress_enabled, force_progress, fallback_bar
    )

    if args.fix_numbers:
        requested_width = (
            DEFAULT_NUMBER_BOX_WIDTH
            if args.number_box_width is None
            else args.number_box_width
        )
        width = None if requested_width == 0 else requested_width
        try:
            changes, target = fix_existing_deck(
                Path(args.fix_numbers).expanduser().resolve(),
                Path(args.output).expanduser().resolve() if args.output else None,
                args.in_place,
                args.start,
                width,
                progress_enabled,
                force_progress,
                fallback_bar,
            )
        except Exception as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 1
        if clear_before_summary:
            clear_screen()
        print(f"Renumbered {len(changes)} slides: {changes[0][2]}–{changes[-1][2]}")
        print(f"Saved {target}")
        return 0

    overrides = {
        "template": args.template,
        "output": args.output,
        "seed": args.seed,
        "number_box_width": args.number_box_width,
    }
    if args.no_shuffle:
        overrides["shuffle_categories"] = False
        overrides["shuffle_cards"] = False

    try:
        config_path = Path(args.config).expanduser().resolve()
        config, config_dir = load_config(config_path)
        result = generate(
            config,
            config_dir,
            overrides,
            progress_enabled,
            force_progress,
            fallback_bar,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if clear_before_summary:
        clear_screen()
    label = result["section_label"]
    plural = "categories" if label == "category" else "sections"
    print(f"Created {result['output']}")
    print(
        f"{result['renumbered']} cards | "
        f"{len(result['sections'])} sections"
    )
    print("Press any key to continue...")
    input()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
