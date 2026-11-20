#!/usr/bin/env python3
"""Weekly SMT deck update script.

This keeps the last week's deck as the base, updates the first slide date, cleans
out stale screenshots, captures the current Looker tiles, and updates the named
slides in the presentation.
"""

from __future__ import annotations

import argparse
import json
from copy import deepcopy
from io import BytesIO
import shutil
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

from PIL import Image, ImageChops, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt
from pptx.oxml.xmlchemy import OxmlElement

# Dashboard URLs, browser defaults, and chart-title-to-slide assignments.
FILE_PREFIX = "SMT Working Version - Weekly Service Update - OPS Week Comm"
DEFAULT_DASHBOARD_URL = (
    "https://eshopworldgcp.cloud.looker.com/dashboards/4274"
    "?SLA+End+Date=1+week+ago+for+1+week"
)
NIKE_VOLUME_REPORT_URL = (
    "https://eshopworldgcp.cloud.looker.com/dashboards/2639"
    "?SLA+End+Week=1+week+ago+for+1+week&Order+Creation+Date=&Hub="
    "&Brand+Name=Nike&Destination+Country="
)
# "23.05.03 - Recon Overview - Last 30 Days Monitor" (slides 9 and 22).
RECON_OVERVIEW_URL = (
    "https://eshopworldgcp.cloud.looker.com/dashboards/4239"
    "?EShop+Package+Reference=&Brand+Name=-iBlues%2C-Levi+Strauss+%26+Co.+Europe%2C-The+Irish+Store%2C-Denby+Pottery"
    "&Carrier+Name=&Hub=&Country=&Carrier+Reference=&SLA+End+Date=30+days&Order+Creation+Date=&Brand+Package+Reference="
)
# "1.04 - Weekly Logistics Update" (slides 27, 28 and 29: highest volume of failed SLA packages).
WEEKLY_LOGISTICS_URL = "https://eshopworldgcp.cloud.looker.com/dashboards/379"
# Dashboard holding "Nike - By Country (Week)" (slide 40).
NIKE_BY_COUNTRY_URL = "https://eshopworldgcp.cloud.looker.com/dashboards/2292"
DEFAULT_PROFILE_DIR =Path.home() / ".smt_deck_browser_profile_auto"
# SharePoint folder the finished deck is uploaded to: Operations SharePoint >
# Shared Documents > Weekly Exec Reporting > <year> - Weekly Exec Reporting > SMT Weekly Deck.
SHAREPOINT_DECK_FOLDER = (
    "/sites/OperationsSharepoint/Shared Documents/Weekly Exec Reporting/"
    "{year} - Weekly Exec Reporting/SMT Weekly Deck"
)
DEFAULT_BROWSER_CHANNEL = "msedge"
TARGET_TILE_REPLACEMENTS = {
    3: ["NIKE BY HUB", "Nike SLA 97% and Avg. Transit Days Per Lane"],
    4: ["NIKE VOLUMES SPLIT BY CARRIER AND HUB", "NIKE VOLUMES SPLIT BY COUNTRY AND HUB"],
    5: ["NIKE NL"],
    6: ["NIKE SG"],
    7: ["NIKE MX"],
    8: ["NIKE HK"],
    9: ["Last 30 Days Recon Overview Nike"],
    11: ["WOTC GROUP"],
    12: ["LULULEMON - EMEA"],
    13: ["LULULEMON - MEXICO"],
    15: ["CALVIN KLEIN", "TOMMY HILFIGER"],
    16: ["NORDSTROM", "ALOHAS"],
    17: ["MAX MARA", "DSQUARED2"],
    18: ["RGG", "FOOTASYLUM GROUP"],
    19: ["KNITWELL GROUP", "ALEXANDER WANG"],
    20: ["RAY-BAN", "JULIAN FASHION"],
    21: ["TODs Group", "MOON BOOT"],
    22: ["Last 30 Days Recon Overview CORE"],
    # Slide 24 shows only the Tier 2 Retailers chart (Tier 3 is no longer captured).
    24: ["TIER 2 RETAILERS"],
    27: ["NIKE Highest Volume of Failed SLA Packages by Hub - Previous 4 Complete Weeks - OPS Manager Focus"],
    28: ["CORE Highest Volume of Failed SLA Packages - Previous 4 Complete Weeks - OPS Manager Focus"],
    29: ["Highest Volume of Failed SLA Packages - Previous 4 Complete Weeks - OPS Manager Focus (Brand Level)"],
    31: ["NIKE By HUB - % CT Delay Email Sent", "Top 15 Core Brands - % CT Delay Email Sent"],
    38: ["Top 10 Wismo's By Retailer"],
    # Slide 39: Nike chart on top, Core chart underneath.
    39: ["Nike Wismo Trend (8 Complete Weeks)", "Core Wismo Trend (8 Complete Weeks)"],
    # Slide 40: taken from its own dashboard (NIKE_BY_COUNTRY_URL).
    40: ["Nike - By Country (Week)"],
}
TILE_REPORT_URLS = {
    (3, 1): NIKE_VOLUME_REPORT_URL,
    (4, 0): NIKE_VOLUME_REPORT_URL,
    (4, 1): NIKE_VOLUME_REPORT_URL,
    (9, 0): RECON_OVERVIEW_URL,
    (22, 0): RECON_OVERVIEW_URL,
    (27, 0): WEEKLY_LOGISTICS_URL,
    (28, 0): WEEKLY_LOGISTICS_URL,
    (29, 0): WEEKLY_LOGISTICS_URL,
    (40, 0): NIKE_BY_COUNTRY_URL,
}
SERVICE_ISSUE_HEADERS = [
    "Service Issues",
    "Owner",
    "Action Taken",
    "Next Steps",
    "Support Needed",
    "Resolvers",
]
STATIC_SLIDE_SCREENSHOTS = {
    33: Path("output") / "nike_wh_reverse_eod_report.png",
    35: Path("output") / "client_returns_nike_tier1.png",
    36: Path("output") / "client_returns_tier2_thom_browne.png",
}


def this_week_monday() -> date:
    """Return Monday for the current calendar week."""
    today = date.today()
    return today - timedelta(days=today.weekday())


def ordinal_day(value: date) -> str:
    """Format a date's day number for the title slide (for example, '2nd')."""
    day = value.day
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


# All Looker screenshots are saved here (the project's "output" folder), as before.
SCREENSHOT_DIR = Path.cwd() / "output"


def resolve_output_dir(week: date) -> Path:
    """Build the generated-output folder path for the requested week."""
    output_dir = Path.cwd() / "output" / "SMT Weekly Deck" / f"{week.year} - Weekly Exec Reporting" / "SMT Weekly Deck"
    return output_dir


def deck_file_path(output_dir: Path) -> Path:
    """Name the deck with the date it is built on (for example '... OPS Week Comm 20.11.2026.pptx')."""
    return output_dir / f"{FILE_PREFIX} {date.today().strftime('%d.%m.%Y')}.pptx"


def find_latest_deck(base_dir: Path) -> Path | None:
    """Find the most recently modified deck to use as the next week's template."""
    if not base_dir.exists():
        return None
    decks = [p for p in base_dir.rglob(f"{FILE_PREFIX}*.pptx") if p.is_file()]
    return max(decks, key=lambda p: p.stat().st_mtime) if decks else None


def clean_old_screenshots(root: Path) -> list[Path]:
    """Remove generated dashboard and slide captures before starting a fresh run."""
    removed: list[Path] = []
    if not root.exists():
        return removed
    for pattern in ("dashboard_*.png", "dashboard_*_full_page.png", "slide_*.png", "slide_*_full_page.png"):
        for path in root.rglob(pattern):
            if path.is_file():
                path.unlink()
                removed.append(path)
    return removed


def build_capture_command(
    args: argparse.Namespace,
    out_image: Path,
    tile: str | None = None,
    url: str | None = None,
    header_only: bool = False,
) -> list[str]:
    """Build the command that asks the separate Playwright helper to capture a chart."""
    cmd = [
        sys.executable,
        str(Path(__file__).with_name("looker_capture.py")),
        "--url",
        url or args.url,
        "--profile",
        str(args.profile),
        "--channel",
        args.channel,
        "--out",
        str(out_image),
        "--wait",
        str(args.wait),
    ]
    if tile:
        cmd.extend(["--tile", tile])
    if header_only:
        cmd.append("--header-only")
    if args.login:
        cmd.append("--login")
    if args.headed:
        cmd.append("--headed")
    return cmd


def capture_tile(
    args: argparse.Namespace,
    tile_name: str,
    out_path: Path,
    url: str | None = None,
    header_only: bool = False,
) -> Path:
    """Run a Looker capture and surface its output if the capture process fails."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = build_capture_command(args, out_path, tile_name, url, header_only)
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.stdout:
        print(proc.stdout, end="")
    if proc.stderr:
        print(proc.stderr, end="", file=sys.stderr)
    if proc.returncode != 0:
        raise RuntimeError(f"Failed to capture '{tile_name}' (exit {proc.returncode}).")
    return out_path


def screenshot_path(slide_index: int, idx: int) -> Path:
    return SCREENSHOT_DIR / f"slide_{slide_index}_{idx}.png"


def capture_all_screenshots(args: argparse.Namespace, monday: date) -> Path:
    """Step 1: open Looker ONCE and save every chart screenshot to the output folder.

    Each dashboard is opened a single time and all of its charts are taken from it,
    instead of starting a new browser for every chart.
    """
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    dashboard_path = SCREENSHOT_DIR / f"dashboard_{monday.isoformat()}.png"
    jobs = [{"url": args.url, "tile": "NIKE NL", "out": str(dashboard_path), "header_only": False}]
    for slide_index, tile_names in TARGET_TILE_REPLACEMENTS.items():
        for idx, tile_name in enumerate(tile_names):
            jobs.append({
                "url": TILE_REPORT_URLS.get((slide_index, idx)) or args.url,
                "tile": tile_name,
                "out": str(screenshot_path(slide_index, idx)),
                "header_only": True,
            })

    jobs_file = SCREENSHOT_DIR / "capture_jobs.json"
    jobs_file.write_text(json.dumps(jobs, indent=1), encoding="utf-8")
    cmd = [
        sys.executable, str(Path(__file__).with_name("looker_capture.py")),
        "--jobs", str(jobs_file),
        "--profile", str(args.profile),
        "--channel", args.channel,
        "--wait", str(args.wait),
    ]
    if args.headed:
        cmd.append("--headed")
    print(f"Step 1 of 2: capturing {len(jobs)} Looker screenshot(s) into {SCREENSHOT_DIR}", flush=True)
    proc = subprocess.run(cmd)  # output shows live
    if proc.returncode != 0:
        raise RuntimeError("Looker capture failed - see the messages above. The PowerPoint was not built.")
    return dashboard_path


def update_slide_one_date(slide, monday: date) -> None:
    """Put today's date (the day the deck is built) under the title on slide 1.

    The deck's file name also uses today's date (see ``deck_file_path``).
    """
    today = date.today()
    for shape in slide.shapes:
        if not hasattr(shape, "text"):
            continue
        text = shape.text.strip()
        # Old decks say "WK Commencing ..."; newer ones just carry the title and a date.
        if "Commencing" in text or "SMT Weekly Service Management" in text:
            shape.text = (
                "SMT Weekly Service Management\n\n"
                f"{ordinal_day(today)} {today.strftime('%B %Y')}"
            )
            return


def normalize_service_issue_tables(prs: Presentation) -> int:
    """Keep only the standard service-issue headers in matching tables on every slide."""
    updated = 0

    def add_column(table, width: int) -> None:
        # Extend the PowerPoint table XML as well as its visible rows and columns.
        grid_column = OxmlElement("a:gridCol")
        grid_column.set("w", str(width))
        table._tbl.tblGrid.append(grid_column)
        for row in table._tbl.tr_lst:
            row.append(deepcopy(row.tc_lst[-1]))

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.TABLE:
                continue

            table = shape.table
            if not is_service_issue_table(table):
                continue

            total_width = sum(column.width for column in table.columns)
            column_widths = [column.width for column in table.columns]
            while len(column_widths) < len(SERVICE_ISSUE_HEADERS):
                column_widths.append(max(1, total_width // len(SERVICE_ISSUE_HEADERS)))
                add_column(table, column_widths[-1])

            width_sum = sum(column_widths)
            scaled_widths = [total_width * width // width_sum for width in column_widths]
            scaled_widths[-1] += total_width - sum(scaled_widths)
            for index, width in enumerate(scaled_widths):
                table.columns[index].width = width

            if len(table.columns) != len(SERVICE_ISSUE_HEADERS):
                continue

            for row_index, row in enumerate(table.rows):
                for column_index, cell in enumerate(row.cells):
                    cell.text = SERVICE_ISSUE_HEADERS[column_index] if row_index == 0 else ""
            updated += 1

    return updated


def is_service_issue_table(table) -> bool:
    if not table.rows:
        return False
    header_names = {header.casefold() for header in SERVICE_ISSUE_HEADERS}
    first_row = [cell.text.strip().casefold() for cell in table.rows[0].cells]
    return sum(value in header_names for value in first_row) >= 3


def is_eod_status_table(table) -> bool:
    return bool(table.rows) and table.cell(0, 0).text.strip().casefold() == "eod - rc status"


def prevent_screenshot_table_overlaps(prs: Presentation) -> tuple[int, int]:
    """Keep screenshots clear of service tables and fit moved tables on their slides."""
    gap = 914400 // 20
    min_horizontal_overlap = 914400 // 2
    min_vertical_overlap = 914400 // 100
    adjusted_slides: set[int] = set()
    removed_rows = 0

    for slide_number, slide in enumerate(prs.slides, start=1):
        pictures = [
            shape for shape in slide.shapes
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.image.size[0] >= 1000
        ]
        tables = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.TABLE]

        for picture in pictures:
            for table_shape in tables:
                horizontal_overlap = min(
                    picture.left + picture.width,
                    table_shape.left + table_shape.width,
                ) - max(picture.left, table_shape.left)
                if horizontal_overlap < min_horizontal_overlap:
                    continue

                vertical_overlap = min(
                    picture.top + picture.height,
                    table_shape.top + table_shape.height,
                ) - max(picture.top, table_shape.top)
                if vertical_overlap <= min_vertical_overlap:
                    continue

                if is_eod_status_table(table_shape.table):
                    picture.top = max(0, table_shape.top)
                    table_shape.top = picture.top + picture.height + gap
                    following_tops = [
                        shape.top for shape in slide.shapes
                        if shape._element is not picture._element
                        and shape._element is not table_shape._element
                        and shape.top >= table_shape.top
                        and (
                            shape.shape_type == MSO_SHAPE_TYPE.TABLE
                            or (getattr(shape, "has_text_frame", False) and shape.text.strip())
                        )
                    ]
                    bottom_limit = min(following_tops, default=prs.slide_height)
                    available_height = bottom_limit - table_shape.top - gap
                    row_heights = [row.height for row in table_shape.table.rows]
                    total_row_height = sum(row_heights)
                    target_height = min(table_shape.height, available_height)
                    if total_row_height <= 0 or target_height <= 0:
                        raise ValueError(f"EOD status table on slide {slide_number} cannot fit below its screenshot.")
                    if target_height < table_shape.height:
                        scaled_heights = [
                            round(height * target_height / total_row_height)
                            for height in row_heights
                        ]
                        scaled_heights[-1] += target_height - sum(scaled_heights)
                        for row, height in zip(table_shape.table.rows, scaled_heights):
                            row.height = height
                        table_shape.height = target_height
                    adjusted_slides.add(slide_number)
                    continue

                if is_service_issue_table(table_shape.table):
                    continue
                if (
                    table_shape.top < 914400 * 1.5
                    and vertical_overlap > min_vertical_overlap
                ):
                    picture.top = table_shape.top + table_shape.height + gap
                    adjusted_slides.add(slide_number)

        for table_shape in tables:
            table = table_shape.table
            if not is_service_issue_table(table):
                continue

            overlapping_pictures = [
                picture for picture in pictures
                if min(picture.left + picture.width, table_shape.left + table_shape.width)
                - max(picture.left, table_shape.left) >= min_horizontal_overlap
                and min(
                    picture.top + picture.height,
                    table_shape.top + table_shape.height,
                ) - max(picture.top, table_shape.top) > min_vertical_overlap
            ]
            if overlapping_pictures:
                table_shape.top = max(picture.top + picture.height for picture in overlapping_pictures) + gap
                adjusted_slides.add(slide_number)
            available_height = prs.slide_height - table_shape.top
            original_row_count = len(table.rows)
            while len(table.rows) > 1 and sum(row.height for row in table.rows) > available_height:
                table._tbl.remove(table.rows[-1]._tr)
                removed_rows += 1

            row_heights = sum(row.height for row in table.rows)
            if row_heights > available_height:
                raise ValueError(
                    f"Service-issue table on slide {slide_number} cannot fit below its screenshots."
                )
            if len(table.rows) != original_row_count:
                table_shape.height = row_heights
                adjusted_slides.add(slide_number)
            elif table_shape.height > available_height:
                table_shape.height = row_heights
                adjusted_slides.add(slide_number)

    return len(adjusted_slides), removed_rows


def drop_image_relation_if_unused(slide, old_rid: str, replaced_shape) -> None:
    """Keep an image relationship while another picture still references it."""
    for shape in slide.shapes:
        if shape is replaced_shape or shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        if shape._element.blipFill.blip.rEmbed == old_rid:
            return
    slide.part.drop_rel(old_rid)


def replace_picture(
    shape,
    image_path: Path,
    slide,
    fit_image: bool = True,
) -> None:
    """Replace a picture and center the new image inside its existing frame."""
    frame = (shape.left, shape.top, shape.width, shape.height)
    old_rid = shape._element.blipFill.blip.rEmbed
    image_part, new_rid = slide.part.get_or_add_image_part(str(image_path))
    shape._element.blipFill.blip.rEmbed = new_rid
    if old_rid != new_rid:
        drop_image_relation_if_unused(slide, old_rid, shape)
    shape.crop_left = shape.crop_right = shape.crop_top = shape.crop_bottom = 0
    if fit_image:
        fit_picture_to_frame(shape, frame, image_part.image.size)
    else:
        shape.left, shape.top, shape.width, shape.height = frame


def fit_picture_to_frame(shape, frame, image_size) -> None:
    """Contain an image in a PowerPoint frame, preserving aspect ratio and centering it."""
    left, top, frame_width, frame_height = frame
    image_width, image_height = image_size
    scale = min(frame_width / image_width, frame_height / image_height)
    width = round(image_width * scale)
    height = round(image_height * scale)
    shape.left = left + (frame_width - width) // 2
    shape.top = top + (frame_height - height) // 2
    shape.width = width
    shape.height = height


def equalize_slide_pictures(slide) -> None:
    """Give all slide pictures one shared position and size (used for slide 12)."""
    pictures = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if len(pictures) < 2:
        return

    left = min(shape.left for shape in pictures)
    width = max(shape.width for shape in pictures)
    height = max(shape.height for shape in pictures)
    for shape in pictures:
        shape.left = left
        shape.width = width
        shape.height = height


def match_slide_picture_sizes(slide) -> None:
    """Align slide 3 pictures to one left edge and matching dimensions."""
    pictures = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if len(pictures) < 2:
        return

    reference = pictures[0]
    for shape in pictures[1:]:
        shape.left = reference.left
        shape.width = reference.width
        shape.height = reference.height


def match_slide_four_picture_sizes(prs: Presentation) -> None:
    """Size slide 4 pictures to match slide 3 while preserving their separate positions."""
    if len(prs.slides) < 4:
        return

    slide_three_pictures = [
        shape for shape in prs.slides[2].shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    slide_four_pictures = [
        shape for shape in prs.slides[3].shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    if not slide_three_pictures or not slide_four_pictures:
        return

    reference = slide_three_pictures[0]
    for shape in slide_four_pictures:
        shape.width = reference.width
        shape.height = reference.height


def trim_white_border(
    image: Image.Image,
    padding: int = 12,
    threshold: int = 20,
    min_fraction: float = 0.03,
) -> Image.Image:
    """Crop the blank (white) margin around a chart, keeping a small padding.

    A column (or row) only counts as chart content when at least ``min_fraction``
    of it is non-white. This stops thin grid/border lines that run across the
    whole capture (common in Looker tables) from blocking the trim.
    """
    rgb = image.convert("RGB")
    width, height = rgb.size
    if width < 20 or height < 20:
        return image
    difference = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).convert("L")
    mask = difference.point(lambda value: 255 if value > threshold else 0)

    limit = 255 * min_fraction
    column_profile = list(mask.resize((width, 1), Image.BOX).getdata())
    row_profile = list(mask.resize((1, height), Image.BOX).getdata())
    columns = [index for index, value in enumerate(column_profile) if value > limit]
    rows = [index for index, value in enumerate(row_profile) if value > limit]
    if not columns or not rows:
        return image

    left, right = columns[0], columns[-1] + 1
    upper, lower = rows[0], rows[-1] + 1
    # Never cut dark text such as the chart title: thin title text is too sparse to pass
    # the min_fraction test, so widen the crop to include every strongly non-white pixel.
    # Light grid/border lines stay below this threshold and still get trimmed.
    strong = difference.point(lambda value: 255 if value > 80 else 0).getbbox()
    if strong:
        left, upper = min(left, strong[0]), min(upper, strong[1])
        right, lower = max(right, strong[2]), max(lower, strong[3])
    return image.crop(
        (
            max(0, left - padding),
            max(0, upper - padding),
            min(width, right + padding),
            min(height, lower + padding),
        )
    )


EMU_PER_INCH = 914400


def center_chart_title(image: Image.Image, padding: int = 12, min_gap: int = 6) -> Image.Image:
    """Move the chart's title so it sits in the middle, above the chart content.

    Looker centres a title over the whole tile, so once a narrow table is cropped the
    title ends up off to one side. This cuts out the title band (the dark text at the top,
    followed by a blank gap), trims the chart body below it, and re-centres the title.
    Images without a clear title band at the top are returned unchanged.
    """
    rgb = image.convert("RGB")
    width, height = rgb.size
    if width < 50 or height < 50:
        return image
    strong = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).convert("L")
    strong = strong.point(lambda value: 255 if value > 80 else 0)
    row_has_ink = [value > 0 for value in strong.resize((1, height), Image.BOX).getdata()]

    # Title band = first run of rows with dark text, near the top, followed by a blank gap.
    search_limit = min(height, max(200, height // 4))
    top = next((y for y in range(search_limit) if row_has_ink[y]), None)
    if top is None:
        return image
    bottom = top
    while bottom < height and row_has_ink[bottom]:
        bottom += 1
    if bottom - top > 120:  # too tall for a one-line title; leave the image alone
        return image
    gap_end = bottom
    while gap_end < height and not row_has_ink[gap_end]:
        gap_end += 1
    if gap_end - bottom < min_gap or gap_end >= height:
        return image

    title_box = strong.crop((0, top, width, bottom)).getbbox()
    if not title_box:
        return image
    title = image.crop((title_box[0], top, title_box[2], bottom))
    body = trim_white_border(image.crop((0, bottom + (gap_end - bottom) // 2, width, height)), padding=0)

    canvas_width = max(body.width, title.width) + 2 * padding
    canvas_height = padding + title.height + (gap_end - bottom) + body.height + padding
    background = (255, 255, 255, 255) if image.mode == "RGBA" else "white"
    canvas = Image.new(image.mode, (canvas_width, canvas_height), background)
    content_left = padding + (canvas_width - 2 * padding - body.width) // 2
    canvas.paste(title, ((canvas_width - title.width) // 2, padding))
    canvas.paste(body, (content_left, padding + title.height + (gap_end - bottom)))
    return canvas


def crop_chart_image(source: Image.Image, slide_number: int, index: int) -> Image.Image:
    """Crop the chart screenshot, then centre its title above the chart."""
    return center_chart_title(_crop_chart_margins(source, slide_number, index))


def _crop_chart_margins(source: Image.Image, slide_number: int, index: int) -> Image.Image:
    """Cut the empty white margin from a chart screenshot (slide-specific rules kept)."""
    if slide_number in {22, 24} or (slide_number == 31 and index == 1):
        difference = ImageChops.difference(
            source.convert("RGB"),
            Image.new("RGB", source.size, "white"),
        ).convert("L")
        bounds = difference.point(lambda value: 255 if value > 20 else 0).getbbox()
        if slide_number == 24 and bounds:
            strong_bounds = difference.point(lambda value: 255 if value > 80 else 0).getbbox()
            if strong_bounds:
                bounds = (strong_bounds[0], bounds[1], strong_bounds[2], bounds[3])
        if bounds:
            padding = 12
            left, upper, right, lower = bounds
            return source.crop(
                (
                    max(0, left - padding),
                    max(0, upper - padding),
                    min(source.width, right + padding),
                    min(source.height, lower + padding),
                )
            )
        return source
    return trim_white_border(source)


def chart_content_area(prs: Presentation, slide, chart_shapes) -> tuple[int, int, int, int]:
    """Return (left, top, right, bottom) of the free area below the title bar.

    The area runs from just under the title bar down to the first table, status
    bar or text below it (e.g. the Returns bar or the Service Issues table), or to
    the bottom of the slide when there is nothing underneath the charts.
    """
    slide_width, slide_height = prs.slide_width, prs.slide_height
    margin = EMU_PER_INCH // 3
    gap = EMU_PER_INCH // 10
    chart_ids = {shape.shape_id for shape in chart_shapes}
    others = [
        shape for shape in slide.shapes
        if shape.shape_id not in chart_ids and shape.shape_type != MSO_SHAPE_TYPE.PICTURE
    ]

    # Title bar: anything starting in the top inch of the slide.
    header = [shape for shape in others if shape.top < EMU_PER_INCH]
    header_bottom = max((shape.top + shape.height for shape in header), default=margin)
    top = header_bottom + gap

    # Line the charts up with the title bar's edges when it spans the slide.
    wide_header = [shape for shape in header if shape.width >= slide_width * 0.8]
    if wide_header:
        left = min(shape.left for shape in wide_header)
        right = min(slide_width - margin // 2, max(shape.left + shape.width for shape in wide_header))
    else:
        left, right = margin, slide_width - margin

    # First real shape below the charts (ignore tiny decorative shapes).
    def is_blocker(shape) -> bool:
        if shape.top < header_bottom or shape.width < EMU_PER_INCH:
            return False
        if shape.shape_type in (MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.AUTO_SHAPE):
            return True
        return bool(getattr(shape, "has_text_frame", False) and shape.text.strip())

    blocker_tops = [shape.top for shape in others if is_blocker(shape)]
    bottom = min(blocker_tops, default=slide_height - margin) - gap
    if bottom - top < EMU_PER_INCH * 1.5:
        # Not enough room above the next shape; use the slide height and let the
        # table-overlap step move the table below the charts.
        bottom = slide_height - margin
    return left, top, right, bottom


def stack_charts_in_area(shapes_and_sizes, area) -> None:
    """Stack charts top-to-bottom, as wide as possible, filling the area's height.

    Each chart keeps its aspect ratio. If they don't all fit at full width, the
    tallest ones are shrunk first (and centred), so short charts stay full width.
    """
    left, top, right, bottom = area
    width = right - left
    gap = EMU_PER_INCH // 10
    available = (bottom - top) - gap * (len(shapes_and_sizes) - 1)
    ratios = [height / width_px for _, (width_px, height) in shapes_and_sizes]
    natural = [width * ratio for ratio in ratios]

    if sum(natural) <= available:
        heights = natural
    else:
        low, high = 0.0, max(natural)
        for _ in range(60):
            cap = (low + high) / 2
            if sum(min(h, cap) for h in natural) > available:
                high = cap
            else:
                low = cap
        heights = [min(h, low) for h in natural]

    y = top
    for (shape, _), ratio, height in zip(shapes_and_sizes, ratios, heights):
        height = max(1, round(height))
        item_width = max(1, round(height / ratio))
        shape.left = left + (width - item_width) // 2
        shape.top = y
        shape.width = item_width
        shape.height = height
        y += height + gap


def normalize_targeted_screenshot_layout(
    prs: Presentation,
    slide_numbers: set[int] | None = None,
) -> None:
    """Trim each chart screenshot and resize it to fill the slide below the title bar."""
    slide_width = prs.slide_width
    if slide_numbers is None:
        slide_numbers = set(TARGET_TILE_REPLACEMENTS) | set(STATIC_SLIDE_SCREENSHOTS)
    for slide_number in sorted(slide_numbers):
        if slide_number > len(prs.slides):
            continue

        slide = prs.slides[slide_number - 1]
        pictures = sorted(
            [
                shape for shape in slide.shapes
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                and shape.width >= slide_width * 0.25
            ],
            key=lambda shape: (shape.top, shape.left),
        )
        if not pictures:
            continue

        area = chart_content_area(prs, slide, pictures)
        shapes_and_sizes = []
        for index, shape in enumerate(pictures):
            source = ImageOps.exif_transpose(Image.open(BytesIO(shape.image.blob))).convert("RGBA")
            source = crop_chart_image(source, slide_number, index)

            image_bytes = BytesIO()
            source.convert("RGB").save(image_bytes, format="PNG")
            image_bytes.seek(0)

            old_rid = shape._element.blipFill.blip.rEmbed
            _, new_rid = slide.part.get_or_add_image_part(image_bytes)
            shape._element.blipFill.blip.rEmbed = new_rid
            if old_rid != new_rid:
                drop_image_relation_if_unused(slide, old_rid, shape)
            shape.crop_left = shape.crop_right = shape.crop_top = shape.crop_bottom = 0
            shapes_and_sizes.append((shape, source.size))

        stack_charts_in_area(shapes_and_sizes, area)


# ---------------------------------------------------------------------------
# ESW theme: restyles the deck in ESW's brand look (colours and font from esw.com).
# Runs after the charts and tables are updated, so it is re-applied every week.
# ---------------------------------------------------------------------------
APPLY_ESW_THEME = True
THEME_FONT = "DM Sans"  # ESW's website font (free on Google Fonts). Change to e.g. "Segoe UI" if not installed.
ESW_INK = RGBColor(0x1E, 0x22, 0x21)        # near-black (headers, text)
ESW_DEEP_GREEN = RGBColor(0x06, 0x3E, 0x2D)  # table header rows
ESW_GREEN = RGBColor(0x00, 0xD1, 0x80)       # bright accent green
ESW_MINT = RGBColor(0xE6, 0xEF, 0xEA)        # banded table rows
ESW_PAGE = RGBColor(0xF2, 0xF6, 0xF3)        # slide background
ESW_LINE = RGBColor(0xD0, 0xD5, 0xD2)        # thin borders
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_NAME = "ESW accent"


def style_text_frame(text_frame, color=None, bold=None, size=None, font=THEME_FONT) -> None:
    """Apply font, colour, weight and size to every run (and empty paragraph) in a text frame."""
    for paragraph in text_frame.paragraphs:
        targets = [run.font for run in paragraph.runs] + [paragraph.font]
        for target in targets:
            if font:
                target.name = font
            if color is not None:
                target.color.rgb = color
            if bold is not None:
                target.bold = bold
            if size is not None:
                target.size = size


def set_cell_border(cell, color: RGBColor, width: int = 9525) -> None:
    """Give a table cell a thin solid border on all four sides."""
    tc_pr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for existing in tc_pr.findall(qn(tag)):
            tc_pr.remove(existing)
    for position, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        line = OxmlElement(tag)
        line.set("w", str(width))
        fill = OxmlElement("a:solidFill")
        colour = OxmlElement("a:srgbClr")
        colour.set("val", str(color))
        fill.append(colour)
        line.append(fill)
        dash = OxmlElement("a:prstDash")
        dash.set("val", "solid")
        line.append(dash)
        tc_pr.insert(position, line)


def fill_cell(cell, color: RGBColor) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = color


def is_title_bar(shape, prs: Presentation) -> bool:
    return (
        shape.shape_type == MSO_SHAPE_TYPE.TABLE
        and shape.top < EMU_PER_INCH
        and shape.width >= prs.slide_width * 0.8
        and not is_eod_status_table(shape.table)
    )


def is_commentary_table(table) -> bool:
    if not table.rows:
        return False
    header = " ".join(cell.text for cell in table.rows[0].cells).casefold()
    return any(word in header for word in ("commentary", "opportunit"))


def is_divider_slide(slide, prs: Presentation) -> bool:
    """Section-divider slides: a full-slide background group, text, and no tables."""
    shapes = list(slide.shapes)
    has_table = any(shape.shape_type == MSO_SHAPE_TYPE.TABLE for shape in shapes)
    has_backdrop = any(
        shape.shape_type == MSO_SHAPE_TYPE.GROUP and shape.width >= prs.slide_width * 0.9
        for shape in shapes
    )
    return has_backdrop and not has_table


def style_divider_slide(slide) -> None:
    """Left-aligned ESW hero style: small green eyebrow line, large bold white title."""
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False) or not shape.text.strip():
            continue
        paragraphs = [p for p in shape.text_frame.paragraphs if p.text.strip()]
        for index, paragraph in enumerate(paragraphs):
            is_eyebrow = index == 0 and len(paragraphs) > 1 and "service management" in paragraph.text.casefold()
            paragraph.alignment = PP_ALIGN.LEFT
            for target in [run.font for run in paragraph.runs] + [paragraph.font]:
                target.name = THEME_FONT
                target.bold = not is_eyebrow
                target.color.rgb = ESW_GREEN if is_eyebrow else WHITE
                target.size = Pt(36) if is_eyebrow else Pt(72)
        for paragraph in shape.text_frame.paragraphs:
            if not paragraph.text.strip():
                paragraph.alignment = PP_ALIGN.LEFT
                paragraph.font.size = Pt(20)


def style_title_bar(slide, shape) -> None:
    """Near-black title bar: white bold section name, bright-green subtitle, green underline."""
    table = shape.table
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            fill_cell(cell, ESW_INK)
            set_cell_border(cell, ESW_INK)
            style_text_frame(cell.text_frame, color=WHITE if index == 0 else ESW_GREEN, bold=index == 0)

    if not any(other.name == ACCENT_NAME for other in slide.shapes):
        accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            shape.left,
            shape.top + shape.height,
            shape.width,
            EMU_PER_INCH // 16,
        )
        accent.name = ACCENT_NAME
        accent.fill.solid()
        accent.fill.fore_color.rgb = ESW_GREEN
        accent.line.fill.background()
        accent.shadow.inherit = False


def style_data_table(table) -> None:
    """Deep-green header row with white bold text, banded white/mint rows, thin borders."""
    for row_index, row in enumerate(table.rows):
        for cell in row.cells:
            set_cell_border(cell, ESW_LINE)
            if row_index == 0:
                fill_cell(cell, ESW_DEEP_GREEN)
                style_text_frame(cell.text_frame, color=WHITE, bold=True)
            else:
                fill_cell(cell, WHITE if row_index % 2 else ESW_MINT)
                style_text_frame(cell.text_frame, color=ESW_INK, bold=False)


def style_status_label(table) -> None:
    """Restyle the 'Returns' label cell of a RAG status bar (RAG colours are kept)."""
    for row in table.rows:
        for cell in row.cells:
            style_text_frame(cell.text_frame, color=None)
            if cell.text.strip().casefold() == "returns":
                fill_cell(cell, ESW_DEEP_GREEN)
                style_text_frame(cell.text_frame, color=WHITE, bold=True)


def apply_esw_theme(prs: Presentation) -> int:
    """Restyle every slide in ESW's look. Returns the number of slides styled."""
    styled = 0
    for slide in prs.slides:
        if is_divider_slide(slide, prs):
            style_divider_slide(slide)
            styled += 1
            continue

        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = ESW_PAGE

        for shape in list(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                if is_title_bar(shape, prs):
                    style_title_bar(slide, shape)
                elif is_service_issue_table(table) or is_commentary_table(table):
                    style_data_table(table)
                elif table.rows and table.cell(0, 0).text.strip().casefold() == "returns":
                    style_status_label(table)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.width >= prs.slide_width * 0.25:
                # Charts sit on the page like white cards with a fine border.
                shape.line.color.rgb = ESW_LINE
                shape.line.width = Pt(0.75)
            elif (
                getattr(shape, "has_text_frame", False)
                and shape.top < EMU_PER_INCH
                and shape.text.strip()
            ):
                # Slides with a plain text title (e.g. slide 31).
                style_text_frame(shape.text_frame, color=ESW_INK, bold=True)
        styled += 1
    return styled


def saved_screenshot(slide_index: int, idx: int, tile_name: str) -> Path:
    """Return the screenshot saved in step 1 for this slide/chart."""
    path = screenshot_path(slide_index, idx)
    if not path.is_file():
        raise FileNotFoundError(f"Missing screenshot for slide {slide_index} ('{tile_name}'): {path}")
    return path


def replace_slide_images(prs: Presentation, slide_index: int, tile_names: list[str], output_dir: Path, args: argparse.Namespace) -> None:
    """Capture each named chart and place it in the corresponding slide picture slot."""
    slide = list(prs.slides)[slide_index - 1]
    picture_shapes = sorted(
        (
            shape for shape in slide.shapes
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE
        ),
        key=lambda shape: (shape.top, shape.left),
    )
    if not picture_shapes:
        return

    # If the slide has more chart pictures than charts to capture (e.g. slide 24 used
    # to show Tier 2 and Tier 3), remove the extra old chart pictures so only the
    # requested charts remain. Small pictures such as logos are left alone.
    chart_pictures = [shape for shape in picture_shapes if shape.width >= prs.slide_width * 0.25]
    for extra in chart_pictures[len(tile_names):]:
        old_rid = extra._element.blipFill.blip.rEmbed
        extra._element.getparent().remove(extra._element)
        drop_image_relation_if_unused(slide, old_rid, None)
        picture_shapes.remove(extra)

    if len(picture_shapes) == 1 and len(tile_names) == 2:
        # Split a single existing frame into two stacked slots for paired chart captures.
        original = picture_shapes[0]
        left, top, width, height = original.left, original.top, original.width, original.height
        gap = 91440
        slot_height = (height - gap) // 2
        original.height = slot_height
        for idx, tile_name in enumerate(tile_names):
            out_path = saved_screenshot(slide_index, idx, tile_name)
            if idx == 0:
                replace_picture(original, out_path, slide)
            else:
                picture = slide.shapes.add_picture(
                    str(out_path), left, top + slot_height + gap, width=width, height=slot_height
                )
                fit_picture_to_frame(
                    picture,
                    (left, top + slot_height + gap, width, slot_height),
                    picture.image.size,
                )
        return

    for idx, tile_name in enumerate(tile_names[: len(picture_shapes)]):
        out_path = saved_screenshot(slide_index, idx, tile_name)
        replace_picture(picture_shapes[idx], out_path, slide)


def replace_static_slide_image(prs: Presentation, slide_index: int, image_path: Path) -> None:
    """Replace the existing image on a slide with a locally generated spreadsheet screenshot."""
    if not image_path.is_file():
        raise FileNotFoundError(f"Required slide {slide_index} screenshot is missing: {image_path}")
    if slide_index > len(prs.slides):
        raise ValueError(f"Cannot add the slide {slide_index} screenshot: presentation has only {len(prs.slides)} slides.")

    slide = prs.slides[slide_index - 1]
    pictures = sorted(
        (shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE),
        key=lambda shape: shape.width * shape.height,
        reverse=True,
    )
    if not pictures:
        raise ValueError(f"Slide {slide_index} has no picture frame for {image_path.name}.")
    replace_picture(pictures[0], image_path, slide, fit_image=False)


def update_deck(prs: Presentation, monday: date, output_dir: Path, args: argparse.Namespace) -> None:
    """Apply the date, requested chart replacements, and standard table formatting."""
    slide_1 = list(prs.slides)[0]
    update_slide_one_date(slide_1, monday)

    for slide_no, tile_names in TARGET_TILE_REPLACEMENTS.items():
        replace_slide_images(prs, slide_no, tile_names, output_dir, args)

    for slide_no, image_path in STATIC_SLIDE_SCREENSHOTS.items():
        replace_static_slide_image(prs, slide_no, image_path)

    normalize_targeted_screenshot_layout(prs)
    print(f"Standardized tables on {normalize_service_issue_tables(prs)} slide(s).")
    adjusted_slides, removed_rows = prevent_screenshot_table_overlaps(prs)
    if adjusted_slides:
        print(
            f"Adjusted screenshot/table layout on {adjusted_slides} slide(s); "
            f"removed {removed_rows} excess table row(s)."
        )
    if APPLY_ESW_THEME:
        print(f"Applied ESW theme to {apply_esw_theme(prs)} slide(s).")


def create_new_deck(img_path: Path, out_path: Path, monday: date) -> None:
    """Create a minimal one-slide presentation when no previous deck is available."""
    prs = Presentation()
    prs.slide_width = 12192000
    prs.slide_height = 6858000
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(0.5 * 914400, 0.3 * 914400, 10 * 914400, 1.2 * 914400)
    box.text_frame.text = (
        "SMT Weekly Service Management\n\n"
        f"WK Commencing {ordinal_day(monday)} {monday.strftime('%B')}"
    )
    image_frame = (int(0.6 * 914400), int(1.3 * 914400), int(11.5 * 914400), int(6.3 * 914400))
    image = slide.shapes.add_picture(str(img_path), *image_frame)
    fit_picture_to_frame(image, image_frame, image.image.size)
    prs.save(str(out_path))


def build_weekly_deck(args: argparse.Namespace, monday: date) -> Path:
    """Clean old captures, take current screenshots, and save the refreshed deck."""
    output_dir = resolve_output_dir(monday)
    output_dir.mkdir(parents=True, exist_ok=True)

    project_output_root = Path.cwd() / "output"
    removed = clean_old_screenshots(project_output_root)
    if output_dir != project_output_root:
        removed.extend(clean_old_screenshots(output_dir))
    if removed:
        print(f"Removed {len(removed)} old screenshot(s) before capture.")

    # Step 1: all screenshots, in one Looker session, saved to the output folder.
    dashboard_path = capture_all_screenshots(args, monday)

    # Step 2: build the PowerPoint from the saved screenshots.
    print("Step 2 of 2: building the PowerPoint from the saved screenshots", flush=True)
    latest_deck = find_latest_deck(output_dir)
    out_path = deck_file_path(output_dir)
    if latest_deck and latest_deck.exists():
        prs = Presentation(str(latest_deck))
        update_deck(prs, monday, output_dir, args)
        prs.save(str(out_path))
        return out_path

    create_new_deck(dashboard_path, out_path, monday)
    return out_path


def upload_to_sharepoint(args: argparse.Namespace, deck_path: Path, monday: date) -> int:
    """Step 3: copy the finished deck to the SMT Weekly Deck folder on SharePoint."""
    folder = SHAREPOINT_DECK_FOLDER.format(year=monday.year)
    print(f"Step 3: uploading the deck to SharePoint ({folder})", flush=True)
    cmd = [
        sys.executable, str(Path(__file__).with_name("sharepoint_upload.py")),
        "--file", str(deck_path),
        "--folder", folder,
        "--profile", str(args.profile),
        "--channel", args.channel,
    ]
    proc = subprocess.run(cmd)  # output shows live
    if proc.returncode != 0:
        print(f"The deck was NOT uploaded to SharePoint. It is saved locally: {deck_path}", file=sys.stderr)
    return proc.returncode


def parse_args() -> argparse.Namespace:
    """Read browser, Looker, and build options from the command line."""
    ap = argparse.ArgumentParser(description="Capture updated Looker tiles and refresh the weekly SMT deck.")
    ap.add_argument("--url", default=DEFAULT_DASHBOARD_URL)
    ap.add_argument("--profile", default=str(DEFAULT_PROFILE_DIR))
    ap.add_argument("--channel", default=DEFAULT_BROWSER_CHANNEL)
    ap.add_argument("--wait", type=int, default=15)
    ap.add_argument("--login", action="store_true")
    ap.add_argument("--headed", action="store_true")
    ap.add_argument("--build", action="store_true", help="Build the weekly deck (the default behavior).")
    ap.add_argument("--no-upload", action="store_true", help="Build the deck but don't upload it to SharePoint.")
    ap.add_argument("--upload-only", action="store_true", help="Upload the most recently built deck to SharePoint; no capture or build.")
    return ap.parse_args()


def main() -> int:
    """Handle the login-only option or run the complete weekly deck workflow."""
    args = parse_args()
    monday = this_week_monday()
    if args.login:
        login_path = resolve_output_dir(monday) / "login_placeholder.png"
        login_path.parent.mkdir(parents=True, exist_ok=True)
        cmd = build_capture_command(args, login_path)
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.stdout:
            print(proc.stdout, end="")
        if proc.stderr:
            print(proc.stderr, end="", file=sys.stderr)
        return 0 if proc.returncode == 0 else proc.returncode

    output_dir = resolve_output_dir(monday)
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.upload_only:
        # Upload today's deck; if none was built today, the most recently built one.
        deck_path = deck_file_path(output_dir)
        if not deck_path.is_file():
            deck_path = find_latest_deck(output_dir)
        if deck_path is None:
            print(f"No built deck to upload in {output_dir}", file=sys.stderr)
            return 1
        return upload_to_sharepoint(args, deck_path, monday)

    deck_path = build_weekly_deck(args, monday)
    print(f"Deck updated: {deck_path}")
    if args.no_upload:
        return 0
    return upload_to_sharepoint(args, deck_path, monday)


if __name__ == "__main__":
    sys.exit(main())
