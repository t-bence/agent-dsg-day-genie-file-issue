import io
from dataclasses import dataclass

import pandas as pd
from agents import function_tool
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from agent_server.tools.volume_files import VolumeFileCache

MAX_CACHED_PRESENTATIONS = 10
MAX_READ_SLIDES = 10


@dataclass
class Slide:
    number: int
    title: str
    text: str


def _slide_text(slide, number: int) -> str:
    parts = [f"--- Slide {number} ---"]

    # --- text content ---
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if text:
                    parts.append(text)

    # --- tables ---
    for shape in slide.shapes:
        if shape.has_table:
            parts.append("\n[Table]")
            for row in shape.table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))

    # --- chart data ---
    chart_num = 0
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.CHART:
            chart_num += 1
            chart = shape.chart
            title = chart.chart_title.text_frame.text if chart.has_title else f"Chart {chart_num}"
            chart_type = str(chart.chart_type)
            plot = chart.plots[0]
            categories = [str(c) for c in plot.categories]

            series_data = {}
            for si, series in enumerate(plot.series):
                series_data[f"Series_{si}"] = list(series.values)

            n_rows = len(next(iter(series_data.values())))
            df = pd.DataFrame(series_data, index=categories[:n_rows])
            df.index.name = "Category"

            parts.append(f"\n[{title} | {chart_type}]")
            parts.append(df.to_string())

    # --- presenter notes ---
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            parts.append(f"\n[Presenter Notes]\n{notes}")

    return "\n".join(parts)


def _slide_title(slide) -> str:
    if slide.shapes.title is not None and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            return shape.text_frame.text.strip().splitlines()[0]
    return "(no title)"


def _parse_slides(data: bytes) -> list[Slide]:
    prs = Presentation(io.BytesIO(data))
    return [Slide(i, _slide_title(slide), _slide_text(slide, i)) for i, slide in enumerate(prs.slides, 1)]


slide_cache = VolumeFileCache((".pptx",), _parse_slides, MAX_CACHED_PRESENTATIONS)


@function_tool
async def read_slides(filename: str, slide_numbers: list[int]) -> str:
    """Return the content of specific slides of a PPTX file from the energy reports volume.

    Use the slide numbers from search_documents results. Reads at most 10 slides per call.
    Each slide includes its text, tables, chart data (as formatted tables) and presenter notes.
    """
    slides = await slide_cache.get(filename)
    if not slide_numbers:
        return "Pass at least one slide number."
    if len(slide_numbers) > MAX_READ_SLIDES:
        return f"Request at most {MAX_READ_SLIDES} slides per call."
    invalid = [n for n in slide_numbers if not 1 <= n <= len(slides)]
    if invalid:
        return f"Slides {invalid} do not exist. {filename} has {len(slides)} slides."
    return "\n\n".join(slides[n - 1].text for n in slide_numbers)
