import asyncio
import logging
import re
from dataclasses import dataclass

import bm25s
import Stemmer
from agents import function_tool

from agent_server.tools.list_files import volume_file_names
from agent_server.tools.pdf_reader import pdf_cache
from agent_server.tools.pptx_reader import slide_cache

MAX_RESULTS = 20
SNIPPET_CHARS = 150

_stemmer = Stemmer.Stemmer("english")
logging.getLogger("bm25s").setLevel(logging.WARNING)


@dataclass
class SearchUnit:
    filename: str
    location: str  # "slide 7" or "page 46"
    title: str
    text: str
    read_hint: str


@dataclass
class SearchIndex:
    units: list[SearchUnit]
    retriever: bm25s.BM25


# The files in the volume do not change, so each set of files is indexed once per app start
_indexes: dict[frozenset[str], SearchIndex] = {}
_index_lock = asyncio.Lock()


async def _load_units(filename: str) -> list[SearchUnit]:
    if filename.lower().endswith(".pptx"):
        slides = await slide_cache.get(filename)
        return [
            SearchUnit(filename, f"slide {s.number}", s.title, s.text, f"read_slides('{filename}', [{s.number}])")
            for s in slides
        ]
    pdf = await pdf_cache.get(filename)
    return [
        SearchUnit(
            filename, f"page {n}", pdf.page_sections[n - 1], text, f"pdf_read_pages('{filename}', {n}, {n})"
        )
        for n, text in enumerate(pdf.page_texts, 1)
    ]


def _tokenize(texts: list[str]):
    return bm25s.tokenize(texts, stopwords="en", stemmer=_stemmer.stemWords, show_progress=False)


def _build_index(units: list[SearchUnit]) -> SearchIndex:
    retriever = bm25s.BM25()
    retriever.index(_tokenize([f"{u.title}\n{u.text}" for u in units]), show_progress=False)
    return SearchIndex(units, retriever)


async def _get_index(filenames: list[str]) -> SearchIndex:
    key = frozenset(filenames)
    async with _index_lock:
        if key not in _indexes:
            unit_lists = await asyncio.gather(*(_load_units(name) for name in sorted(key)))
            units = [unit for unit_list in unit_lists for unit in unit_list]
            _indexes[key] = await asyncio.to_thread(_build_index, units)
        return _indexes[key]


def _snippet(text: str, query: str) -> str:
    """Cut the text around the place where the most query words appear close together."""
    flat = " ".join(text.split())
    lower = flat.lower()
    stems = {_stemmer.stemWord(word) for word in re.findall(r"\w{3,}", query.lower())}
    hits = sorted((m.start(), stem) for stem in stems for m in re.finditer(re.escape(stem), lower))
    center = 0
    if hits:
        # Pick the hit whose window contains the most different query words
        center = max(
            hits,
            key=lambda hit: len({stem for pos, stem in hits if abs(pos - hit[0]) <= SNIPPET_CHARS}),
        )[0]
    start = max(0, center - SNIPPET_CHARS)
    return ("..." if start > 0 else "") + flat[start : center + SNIPPET_CHARS] + "..."


@function_tool
async def search_documents(query: str, filenames: list[str] | None = None, max_results: int = 10) -> str:
    """Keyword search across the slides and pages of the PPTX and PDF files in the energy reports volume.

    Returns the best matching slides and pages, ranked with BM25, each with a short snippet.
    It does not return full content: read the relevant hits afterwards with read_slides
    or pdf_read_pages. Search all files by default, or pass file names to search only those.
    Use specific keywords, for example "combined cycle overnight cost".
    """
    available = [name for name in await asyncio.to_thread(volume_file_names) if name.lower().endswith((".pptx", ".pdf"))]
    if filenames:
        unknown = sorted(set(filenames) - set(available))
        if unknown:
            return f"Files not found or not searchable: {unknown}. Searchable files: {available}"
        available = filenames
    if not available:
        return "There are no PPTX or PDF files to search."

    index = await _get_index(available)
    query_tokens = _tokenize([query])
    if not query_tokens.vocab:
        return "The query has no searchable words. Use more specific keywords."

    k = min(max(1, min(max_results, MAX_RESULTS)), len(index.units))
    doc_ids, scores = index.retriever.retrieve(query_tokens, k=k, show_progress=False)
    lines = []
    for doc_id, score in zip(doc_ids[0], scores[0]):
        if score <= 0:
            continue
        unit = index.units[doc_id]
        lines.append(
            f"- {unit.filename} | {unit.location} | {unit.title} | score {score:.2f}\n"
            f"  {_snippet(unit.text, query)}\n"
            f"  Read with: {unit.read_hint}"
        )
    if not lines:
        return f"No slides or pages match {query!r}. Try other keywords."
    return f"Top {len(lines)} matches for {query!r}:\n" + "\n".join(lines)
