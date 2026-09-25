import base64
import io
from dataclasses import dataclass

from agents import function_tool
from databricks_openai import AsyncDatabricksOpenAI
from pypdf import PdfReader, PdfWriter

from agent_server.tools.volume_files import VolumeFileCache

PDF_MODEL = "databricks-claude-opus-5"
MAX_CACHED_PDFS = 5
MAX_READ_PAGES = 15
MAX_READ_CHARS = 60000
MAX_ANALYZE_PAGES = 5
LOW_TEXT_CHARS = 300


@dataclass
class ParsedPdf:
    reader: PdfReader
    page_texts: list[str]
    # The table of contents section that each page belongs to, used as the page title in search results
    page_sections: list[str]


def _outline_entries(reader: PdfReader, items: list) -> list[tuple[int, str]]:
    """Flatten the table of contents into (page index, title) pairs."""
    entries = []
    for item in items:
        if isinstance(item, list):
            entries += _outline_entries(reader, item)
        else:
            page_index = reader.get_destination_page_number(item)
            if page_index is not None:
                entries.append((page_index, item.title))
    return entries


def _page_sections(reader: PdfReader) -> list[str]:
    entries = sorted(_outline_entries(reader, reader.outline))
    sections = []
    current = "(no section)"
    for page_index in range(len(reader.pages)):
        while entries and entries[0][0] <= page_index:
            current = entries.pop(0)[1]
        sections.append(current)
    return sections


def _parse_pdf(data: bytes) -> ParsedPdf:
    reader = PdfReader(io.BytesIO(data))
    page_texts = [page.extract_text() or "" for page in reader.pages]
    return ParsedPdf(reader, page_texts, _page_sections(reader))


pdf_cache = VolumeFileCache((".pdf",), _parse_pdf, MAX_CACHED_PDFS)


def _check_pages(pdf: ParsedPdf, first_page: int, last_page: int, max_pages: int) -> None:
    page_count = len(pdf.page_texts)
    if not 1 <= first_page <= last_page <= page_count:
        raise ValueError(f"Pages must be between 1 and {page_count}, with first_page <= last_page.")
    if last_page - first_page + 1 > max_pages:
        raise ValueError(f"Request at most {max_pages} pages per call.")


def _outline_lines(reader: PdfReader, items: list, depth: int = 0) -> list[str]:
    lines = []
    for item in items:
        if isinstance(item, list):
            lines += _outline_lines(reader, item, depth + 1)
        else:
            page_index = reader.get_destination_page_number(item)
            page = page_index + 1 if page_index is not None else "unknown"
            lines.append(f"{'  ' * depth}- {item.title} (page {page})")
    return lines


@function_tool
async def pdf_overview(filename: str) -> str:
    """Describe a PDF from the energy reports volume: page count, metadata and table of contents.

    Also lists pages with little extractable text. Those pages are usually figures, charts
    or scans; use pdf_analyze_pages to look at them. Call this first for a new PDF.
    Use a file name returned by get_files_in_volume.
    """
    pdf = await pdf_cache.get(filename)
    reader = pdf.reader
    metadata = {key.lstrip("/"): str(value) for key, value in (reader.metadata or {}).items()}
    lines = [
        f"PDF: {filename}",
        f"Pages: {len(pdf.page_texts)}",
        f"Metadata: {metadata or 'none'}",
        "",
        "Table of contents:",
    ]
    lines += _outline_lines(reader, reader.outline) or ["- none"]
    low_text = [str(i) for i, text in enumerate(pdf.page_texts, 1) if len(text.strip()) < LOW_TEXT_CHARS]
    lines += ["", f"Pages with little text (figures, charts or blank pages): {', '.join(low_text) or 'none'}"]
    return "\n".join(lines)


@function_tool
async def pdf_read_pages(filename: str, first_page: int, last_page: int) -> str:
    """Return the extracted text of a page range of a PDF (page numbers start at 1).

    Reads at most 15 pages per call. Tables lose their layout in extracted text,
    so for tables where the columns matter, and for charts, use pdf_analyze_pages instead.
    """
    pdf = await pdf_cache.get(filename)
    _check_pages(pdf, first_page, last_page, MAX_READ_PAGES)
    parts = []
    total = 0
    for page_number in range(first_page, last_page + 1):
        text = pdf.page_texts[page_number - 1].strip() or "(no extractable text)"
        parts.append(f"--- Page {page_number} ---\n{text}")
        total += len(text)
        if total >= MAX_READ_CHARS and page_number < last_page:
            parts.append(f"(Stopped after page {page_number} because the text is long. Continue from page {page_number + 1}.)")
            break
    return "\n\n".join(parts)


@function_tool
async def pdf_analyze_pages(filename: str, first_page: int, last_page: int, question: str) -> str:
    """Let a vision model look at PDF pages as they appear, and answer a question about them.

    Use this for tables, charts, figures and layouts that the extracted text does not show well.
    Analyzes at most 5 pages per call. The question says what to look for,
    for example "Extract Table 3 as a markdown table."
    """
    pdf = await pdf_cache.get(filename)
    _check_pages(pdf, first_page, last_page, MAX_ANALYZE_PAGES)
    writer = PdfWriter()
    for page_number in range(first_page, last_page + 1):
        writer.add_page(pdf.reader.pages[page_number - 1])
    buffer = io.BytesIO()
    writer.write(buffer)
    data = base64.b64encode(buffer.getvalue()).decode()

    # The Agents SDK drops files from tool outputs on the chat completions API,
    # so the tool sends the pages to the model itself and returns the text answer
    completion = await AsyncDatabricksOpenAI().chat.completions.create(
        model=PDF_MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"These are pages {first_page}-{last_page} of {filename}. {question}"},
                    # "document" is a Databricks content type for Claude, so the OpenAI type hints do not include it
                    {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": data}},  # type: ignore[list-item]
                ],
            }
        ],
        extra_body={"thinking": {"type": "disabled"}},
    )
    return completion.choices[0].message.content or "The model returned no answer."
