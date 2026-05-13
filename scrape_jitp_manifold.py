#!/usr/bin/env python3
import argparse
import concurrent.futures
import datetime as dt
import html as html_lib
import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from lxml import html

BASE = os.environ.get("JITP_BASE_URL", "https://cuny.manifoldapp.org").rstrip("/")
JOURNAL_ID = os.environ.get("JITP_JOURNAL_ID", "9002e7fe-82b4-4549-9ba4-9fbcce811ce2")
ISSUES_THROUGH = int(os.environ.get("JITP_ISSUES_THROUGH", "27"))
OUT_DIR = Path(os.environ.get("JITP_OUTPUT_DIR", Path.cwd() / "outputs" / "jitp_manifold_metadata"))
JSON_OUT = OUT_DIR / "jitp_manifold_metadata.json"

SHORT_FORM_LABELS = {
    "Assignments",
    "Teaching Fails",
    "Reviews",
    "Tool Tips",
    "Blueprints",
}

HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6"}
AUTHOR_BIO_HEADINGS = {
    "about the author",
    "about the authors",
    "author bio",
    "author bios",
    "author biography",
    "author biographies",
    "about the contributor",
    "about the contributors",
    "about the editor",
    "about the editors",
    "about the guest editors",
}

CONTENT_TAGS = "self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::h6 or self::p or self::li"

EXISTING_FALLBACK_FIELDS = {
    "article_url",
    "page_title",
    "page_h1",
    "author_byline_lines",
    "author_affiliations",
    "abstract",
    "keywords",
    "author_bio",
    "license",
    "normalized_full_text",
    "normalized_main_text",
    "notes_text",
    "references_text",
    "acknowledgments_text",
    "appendix_text",
    "toc_headings",
    "section_headings",
    "text_unique_identifier",
    "text_language",
    "text_doi",
    "text_rights",
    "text_metadata_keys",
    "text_metadata_json",
    "text_citations_json",
    "toc_json",
    "spine_json",
    "ingestion_source_download_url",
    "ingestion_external_source_url",
    "epub_v3_export_url",
}

EXISTING_STATUS_FIELDS = {
    "text_api_status",
    "article_html_status",
    "bio_status",
    "abstract_status",
    "keywords_status",
}


def fetch_text(url, accept_json=False, retries=3):
    headers = {
        "User-Agent": "JITP article metadata scraper",
    }
    if accept_json:
        headers["Accept"] = "application/vnd.api+json"
    req = urllib.request.Request(url, headers=headers)
    last = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=45) as resp:
                data = resp.read()
                return data.decode(resp.headers.get_content_charset() or "utf-8", errors="replace")
        except (urllib.error.URLError, TimeoutError) as exc:
            last = exc
            time.sleep(0.7 * (attempt + 1))
    raise RuntimeError(f"Failed to fetch {url}: {last}")


def fetch_json(url):
    return json.loads(fetch_text(url, accept_json=True))


def parse_initial_state(page_html):
    match = re.search(r"<script>window\.__INITIAL_STATE__=(.*?);</script>", page_html, re.S)
    if not match:
        raise ValueError("No Manifold __INITIAL_STATE__ block found")
    payload = match.group(1)
    payload = re.sub(r"(?<=[:\[,])undefined(?=[,\]}])", "null", payload)
    return json.loads(payload)


def clean_text(value):
    if value is None:
        return ""
    value = html_lib.unescape(str(value))
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def html_to_text(value):
    if not value:
        return ""
    try:
        doc = html.fromstring(f"<div>{value}</div>")
        return clean_text(doc.text_content())
    except Exception:
        return clean_text(re.sub(r"<[^>]+>", " ", value))


def get_rel_id(entity, rel):
    data = ((entity.get("relationships") or {}).get(rel) or {}).get("data")
    if isinstance(data, dict):
        return data.get("id", "")
    return ""


def get_rel_ids(entity, rel):
    data = ((entity.get("relationships") or {}).get(rel) or {}).get("data")
    if isinstance(data, list):
        return [item.get("id", "") for item in data if isinstance(item, dict)]
    if isinstance(data, dict):
        return [data.get("id", "")]
    return []


def authors_from_description(description):
    text = clean_text(description)
    if text.lower().startswith("by "):
        return clean_text(text[3:])
    return text


def split_author_display(author_display):
    if not author_display:
        return []
    text = author_display.replace(" and the ", " and the ")
    text = re.sub(r"\s+and\s+", ", ", text)
    parts = [p.strip(" ,") for p in text.split(",")]
    merged = []
    skip = False
    for i, part in enumerate(parts):
        if skip:
            skip = False
            continue
        if part.lower() == "jr." and merged:
            merged[-1] += ", Jr."
        elif i + 1 < len(parts) and parts[i + 1].lower() == "jr.":
            merged.append(f"{part}, Jr.")
            skip = True
        elif part:
            merged.append(part)
    return [p for p in merged if p]


def parse_issue_pages():
    issues = {}
    for page in range(1, 6):
        url = f"{BASE}/journals/issues" if page == 1 else f"{BASE}/journals/issues?page={page}"
        state = parse_initial_state(fetch_text(url))
        entities = state["entityStore"]["entities"]
        page_issues = entities.get("journalIssues", {})
        if not page_issues:
            break
        for issue_id, issue in page_issues.items():
            rel_journal = get_rel_id(issue, "journal")
            attrs = issue.get("attributes") or {}
            title = attrs.get("title") or ""
            number = str(attrs.get("number") or "")
            if rel_journal != JOURNAL_ID and not title.startswith("Journal of Interactive Technology and Pedagogy"):
                continue
            if number == "About the Journal":
                continue
            if number.isdigit() and not (1 <= int(number) <= ISSUES_THROUGH):
                continue
            if not number.isdigit() and number not in SHORT_FORM_LABELS:
                continue
            issues[issue_id] = issue
    ordered = sorted(
        issues.values(),
        key=lambda item: (
            0 if str((item.get("attributes") or {}).get("number") or "").isdigit() else 1,
            int((item.get("attributes") or {}).get("number") or 0)
            if str((item.get("attributes") or {}).get("number") or "").isdigit()
            else 1000,
            str((item.get("attributes") or {}).get("number") or ""),
        ),
    )
    return ordered


def extract_issue_editors(description_text):
    text = html_to_text(description_text)
    match = re.search(r"Edited by\s+(.+?)(?:[“\"']|$)", text)
    return clean_text(match.group(1)) if match else ""


def flatten_metadata(metadata):
    if not isinstance(metadata, dict):
        return {}
    out = {}
    for key, val in metadata.items():
        if isinstance(val, (dict, list)):
            out[key] = json.dumps(val, ensure_ascii=False, sort_keys=True)
        else:
            out[key] = clean_text(val)
    return out


def trim_cell(value, limit=32000):
    text = clean_text(value)
    if len(text) <= limit:
        return text
    return text[: limit - 30].rstrip() + " ... [truncated for Excel]"


def heading_label(el):
    return clean_text(el.text_content()).strip(" :").lower()


def collect_until_next_heading(start_heading):
    chunks = []
    for sibling in start_heading.itersiblings():
        if sibling.tag.lower() in HEADING_TAGS:
            break
        text = clean_text(sibling.text_content())
        if text:
            chunks.append(text)
    return chunks


def find_heading(section, labels):
    for el in section.xpath(".//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::h6]"):
        label = heading_label(el)
        if label in labels:
            return el
    return None


def collect_section_texts(section, label_matcher):
    results = []
    headings = section.xpath(".//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::h6]")
    for heading in headings:
        label = heading_label(heading)
        if label_matcher(label):
            text = trim_cell(" ".join(collect_until_next_heading(heading)))
            if text:
                results.append({"heading": clean_text(heading.text_content()), "text": text})
    return results


def normalized_article_blocks(section):
    blocks = []
    for el in section.xpath(f".//*[{CONTENT_TAGS}]"):
        tag = el.tag.lower()
        # Avoid duplicate text when paragraphs are nested inside list items.
        if tag == "li" and el.xpath(".//p"):
            continue
        if tag == "p" and el.xpath("ancestor::li"):
            continue
        text = clean_text(el.text_content())
        if not text:
            continue
        if tag in HEADING_TAGS:
            level = int(tag[1])
            blocks.append(f"{'#' * level} {text}")
        else:
            blocks.append(text)
    return blocks


def normalized_main_blocks(blocks):
    cut_labels = (
        "## notes",
        "## note",
        "## references",
        "## bibliography",
        "## works cited",
        "## about the author",
        "## about the authors",
        "## author bio",
        "## author bios",
        "## appendix",
    )
    filtered = []
    for block in blocks:
        lowered = block.lower()
        if any(lowered.startswith(label) for label in cut_labels):
            break
        filtered.append(block)
    return filtered


def word_count(text):
    return len(re.findall(r"\b[\w’'-]+\b", text))


def looks_like_bio_paragraph(text, author_names):
    lowered = text.lower()
    if len(text) < 45 or len(text) > 1800:
        return False
    if "licensed under" in lowered or "creative commons" in lowered:
        return False
    if lowered.startswith(("http://", "https://", "doi:", "figure ", "table ")):
        return False
    if re.search(r"\b\d{4}\b.*\b(journal|press|university|doi|http)", lowered):
        return False
    names = split_author_display(author_names)
    name_hit = False
    for name in names:
        parts = [p for p in re.split(r"\s+", name.replace(".", " ")) if len(p) > 1]
        if not parts:
            continue
        first = parts[0].lower()
        last = parts[-1].lower()
        full = " ".join(parts).lower()
        if lowered.startswith(full) or lowered.startswith(first + " ") or lowered.startswith(last + " "):
            name_hit = True
            break
        if full in lowered[:180] or (first in lowered[:140] and last in lowered[:180]):
            name_hit = True
            break
    bio_markers = [
        " is ",
        " are ",
        " serves as ",
        " professor",
        " lecturer",
        " teacher",
        " doctoral",
        " candidate",
        " researcher",
        " editor",
        " director",
        " coordinator",
        " fellow",
        " works ",
        " teaches ",
        " studies ",
        " holds ",
        " received ",
    ]
    return (name_hit or "@" in text) and any(marker in lowered for marker in bio_markers)


def extract_article_page(page_html, author_names=""):
    data = {
        "page_title": "",
        "page_h1": "",
        "byline_lines": [],
        "affiliations": [],
        "abstract": "",
        "keywords": "",
        "author_bio": "",
        "license": "",
        "normalized_full_text": "",
        "normalized_main_text": "",
        "notes_text": "",
        "references_text": "",
        "acknowledgments_text": "",
        "appendix_text": "",
        "normalized_full_word_count": 0,
        "normalized_main_word_count": 0,
        "paragraph_count": 0,
        "heading_count": 0,
        "toc_headings": [],
        "section_headings": [],
        "article_html_status": "parsed",
        "bio_status": "missing",
        "abstract_status": "missing",
        "keywords_status": "missing",
    }
    try:
        doc = html.fromstring(page_html)
    except Exception as exc:
        data["article_html_status"] = f"parse failed: {exc}"
        return data

    title_nodes = doc.xpath("//title")
    if title_nodes:
        data["page_title"] = clean_text(title_nodes[0].text_content())

    sections = doc.xpath('//*[contains(concat(" ", normalize-space(@class), " "), " manifold-text-section ")]')
    if not sections:
        data["article_html_status"] = "no manifold-text-section found"
        return data
    section = sections[0]

    h1 = section.xpath(".//h1[1]")
    if h1:
        data["page_h1"] = clean_text(h1[0].text_content())

    all_headings = []
    for el in section.xpath(".//*[self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::h6]"):
        text = clean_text(el.text_content())
        if text:
            all_headings.append(text)
    data["section_headings"] = all_headings

    toc_links = doc.xpath('//*[@aria-label="table of contents"]//a | //*[@role="dialog" and contains(@aria-label, "table of contents")]//a')
    toc = [clean_text(a.text_content()) for a in toc_links if clean_text(a.text_content())]
    data["toc_headings"] = toc or all_headings

    blocks = normalized_article_blocks(section)
    main_blocks = normalized_main_blocks(blocks)
    data["normalized_full_text"] = trim_cell("\n\n".join(blocks))
    data["normalized_main_text"] = trim_cell("\n\n".join(main_blocks))
    data["normalized_full_word_count"] = word_count(data["normalized_full_text"])
    data["normalized_main_word_count"] = word_count(data["normalized_main_text"])
    data["paragraph_count"] = len([block for block in blocks if not block.startswith("#")])
    data["heading_count"] = len([block for block in blocks if block.startswith("#")])

    notes_sections = collect_section_texts(section, lambda label: label == "notes" or label.startswith("notes "))
    references_sections = collect_section_texts(
        section,
        lambda label: label in {"references", "bibliography", "works cited", "reference"},
    )
    acknowledgments_sections = collect_section_texts(
        section,
        lambda label: label in {"acknowledgments", "acknowledgements", "acknowledgment", "acknowledgement"},
    )
    appendix_sections = collect_section_texts(section, lambda label: label.startswith("appendix"))
    data["notes_text"] = trim_cell(" ".join([item["text"] for item in notes_sections]))
    data["references_text"] = trim_cell(" ".join([item["text"] for item in references_sections]))
    data["acknowledgments_text"] = trim_cell(" ".join([item["text"] for item in acknowledgments_sections]))
    data["appendix_text"] = trim_cell(" ".join([item["text"] for item in appendix_sections]))

    byline_paras = section.xpath('.//h1[1]/following-sibling::p[contains(concat(" ", normalize-space(@class), " "), " byline ")]')
    if not byline_paras:
        first_h2 = section.xpath(".//h2[1]")
        if first_h2 and h1:
            candidates = []
            for sibling in h1[0].itersiblings():
                if sibling is first_h2[0] or sibling.tag.lower() in HEADING_TAGS:
                    break
                if sibling.tag.lower() == "p":
                    text = clean_text(sibling.text_content())
                    if text:
                        candidates.append(sibling)
            byline_paras = candidates[:8]
    for p in byline_paras:
        text = clean_text(p.text_content())
        if text:
            data["byline_lines"].append(text)
            if "," in text:
                data["affiliations"].append(text.split(",", 1)[1].strip())

    abstract_heading = find_heading(section, {"abstract"})
    if abstract_heading is not None:
        chunks = collect_until_next_heading(abstract_heading)
        cleaned = []
        keyword_chunks = []
        for chunk in chunks:
            if re.match(r"^keywords?\s*:", chunk, re.I):
                keyword_chunks.append(re.sub(r"^keywords?\s*:\s*", "", chunk, flags=re.I))
            else:
                cleaned.append(chunk)
        data["abstract"] = trim_cell(" ".join(cleaned))
        data["keywords"] = trim_cell("; ".join(keyword_chunks))

    if not data["abstract"]:
        abstract_blocks = section.xpath('.//*[@id="abstract" or contains(concat(" ", normalize-space(@class), " "), " abstract ")]')
        if abstract_blocks:
            text = clean_text(abstract_blocks[0].text_content())
            text = re.sub(r"^Abstract\s*", "", text, flags=re.I)
            match = re.search(r"Keywords?\s*:\s*(.+)$", text, re.I)
            if match:
                data["keywords"] = trim_cell(match.group(1))
                text = text[: match.start()]
            data["abstract"] = trim_cell(text)

    if data["abstract"]:
        data["abstract_status"] = "found"
    if data["keywords"]:
        data["keywords_status"] = "found"

    bio_heading = find_heading(section, AUTHOR_BIO_HEADINGS)
    bio_chunks = []
    if bio_heading is not None:
        for chunk in collect_until_next_heading(bio_heading):
            if re.search(r"\blicensed under\b|\bcreative commons\b", chunk, re.I):
                data["license"] = trim_cell(chunk)
            else:
                bio_chunks.append(chunk)

    if not bio_chunks:
        bio_blocks = section.xpath('.//*[@id="authorbio" or contains(concat(" ", normalize-space(@class), " "), " authorbio ")]')
        if bio_blocks:
            text = clean_text(bio_blocks[0].text_content())
            parts = re.split(r"(?i)(This entry is licensed under.*)", text, maxsplit=1)
            bio_chunks = [parts[0]] if parts and parts[0].strip() else []
            if len(parts) > 1:
                data["license"] = trim_cell(parts[1])

    if not data["license"]:
        lic = section.xpath('.//*[contains(translate(normalize-space(.), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "creative commons") or contains(translate(normalize-space(.), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "licensed under")]')
        for item in lic:
            text = clean_text(item.text_content())
            if "licensed" in text.lower() or "creative commons" in text.lower():
                data["license"] = trim_cell(text)
                break

    data["author_bio"] = trim_cell(" ".join(bio_chunks))
    if not data["author_bio"] and author_names:
        fallback_chunks = []
        paragraphs = [clean_text(p.text_content()) for p in section.xpath(".//p")]
        for text in paragraphs[-18:]:
            if looks_like_bio_paragraph(text, author_names):
                fallback_chunks.append(text)
        if fallback_chunks:
            data["author_bio"] = trim_cell(" ".join(fallback_chunks))
    if data["author_bio"]:
        data["bio_status"] = "found"
    elif data["license"]:
        data["bio_status"] = "license found, bio missing"

    return data


def scrape_text(row):
    text_id = row["text_id"]
    text_slug = row["text_slug"]
    full_api_url = f"{BASE}/api/v1/texts/{urllib.parse.quote(text_id)}?include=project,category,stylesheets"
    out = dict(row)
    out["text_api_url"] = full_api_url
    try:
        full = fetch_json(full_api_url)
        attrs = full.get("data", {}).get("attributes") or {}
        out["text_api_status"] = "fetched"
        out["text_created_at"] = attrs.get("createdAt") or out.get("text_created_at", "")
        out["text_updated_at"] = attrs.get("updatedAt") or out.get("text_updated_at", "")
        out["text_publication_date"] = attrs.get("publicationDate") or out.get("text_publication_date", "")
        out["text_unique_identifier"] = flatten_metadata(attrs.get("metadata", {})).get("uniqueIdentifier", "")
        out["text_language"] = flatten_metadata(attrs.get("metadata", {})).get("language", "")
        out["text_doi"] = flatten_metadata(attrs.get("metadata", {})).get("doi", "")
        out["text_rights"] = flatten_metadata(attrs.get("metadata", {})).get("rights", "")
        out["text_metadata_json"] = json.dumps(attrs.get("metadata") or {}, ensure_ascii=False, sort_keys=True)
        out["text_citations_json"] = json.dumps(attrs.get("citations") or {}, ensure_ascii=False, sort_keys=True)
        out["text_metadata_keys"] = ", ".join(attrs.get("metadataProperties") or [])
        out["start_text_section_id"] = attrs.get("startTextSectionId") or out.get("start_text_section_id", "")
        out["toc_json"] = json.dumps(attrs.get("toc") or [], ensure_ascii=False)
        out["spine_json"] = json.dumps(attrs.get("spine") or [], ensure_ascii=False)
        out["ingestion_source_download_url"] = attrs.get("ingestionSourceDownloadUrl") or ""
        out["ingestion_external_source_url"] = attrs.get("ingestionExternalSourceUrl") or ""
        out["epub_v3_export_url"] = attrs.get("epubV3ExportUrl") or ""
    except Exception as exc:
        out["text_api_status"] = f"failed: {exc}"

    start_section = out.get("start_text_section_id", "")
    if text_slug and start_section:
        article_url = f"{BASE}/read/{text_slug}/section/{start_section}"
    elif text_slug:
        article_url = f"{BASE}/read/{text_slug}"
    else:
        article_url = ""
    out["article_url"] = article_url

    if article_url:
        try:
            page = fetch_text(article_url)
            extracted = extract_article_page(page, out.get("author_names", ""))
            out.update(
                {
                    "article_html_status": extracted["article_html_status"],
                    "page_title": extracted["page_title"],
                    "page_h1": extracted["page_h1"],
                    "author_byline_lines": " | ".join(extracted["byline_lines"]),
                    "author_affiliations": " | ".join(extracted["affiliations"]),
                    "abstract": extracted["abstract"],
                    "keywords": extracted["keywords"],
                    "author_bio": extracted["author_bio"],
                    "license": extracted["license"],
                    "normalized_full_text": extracted["normalized_full_text"],
                    "normalized_main_text": extracted["normalized_main_text"],
                    "notes_text": extracted["notes_text"],
                    "references_text": extracted["references_text"],
                    "acknowledgments_text": extracted["acknowledgments_text"],
                    "appendix_text": extracted["appendix_text"],
                    "normalized_full_word_count": extracted["normalized_full_word_count"],
                    "normalized_main_word_count": extracted["normalized_main_word_count"],
                    "paragraph_count": extracted["paragraph_count"],
                    "heading_count": extracted["heading_count"],
                    "toc_headings": " | ".join(extracted["toc_headings"]),
                    "section_headings": " | ".join(extracted["section_headings"]),
                    "bio_status": extracted["bio_status"],
                    "abstract_status": extracted["abstract_status"],
                    "keywords_status": extracted["keywords_status"],
                }
            )
        except Exception as exc:
            out["article_html_status"] = f"failed: {exc}"
            out.setdefault("bio_status", "html fetch failed")
            out.setdefault("abstract_status", "html fetch failed")
            out.setdefault("keywords_status", "html fetch failed")
    if not out.get("author_names"):
        bylines = out.get("author_byline_lines", "")
        if bylines:
            out["author_names"] = " | ".join([line.split(",", 1)[0].strip() for line in bylines.split(" | ")])
        else:
            out["author_names"] = authors_from_description(out.get("text_description", ""))
    out["author_count_estimate"] = len(split_author_display(out.get("author_names", "")))
    out["deleted_slug_flag"] = "yes" if "deleted" in (text_slug or "").lower() else ""
    return out


def parse_args():
    parser = argparse.ArgumentParser(description="Scrape JITP metadata from CUNY Manifold.")
    parser.add_argument("--base-url", default=BASE, help="CUNY Manifold base URL.")
    parser.add_argument("--journal-id", default=JOURNAL_ID, help="JITP journal UUID in Manifold.")
    parser.add_argument("--issues-through", type=int, default=ISSUES_THROUGH, help="Last numbered issue to include.")
    parser.add_argument("--output-dir", default=str(OUT_DIR), help="Directory for jitp_manifold_metadata.json.")
    parser.add_argument(
        "--existing-json",
        default="",
        help="Optional existing JSON snapshot used as fallback when current scrape fields are blank or fail.",
    )
    parser.add_argument(
        "--no-existing-fallback",
        action="store_true",
        help="Disable fallback to an existing JSON snapshot.",
    )
    parser.add_argument("--max-workers", type=int, default=8, help="Concurrent article fetch workers.")
    return parser.parse_args()


def configure_runtime(args):
    global BASE, JOURNAL_ID, ISSUES_THROUGH, OUT_DIR, JSON_OUT
    BASE = args.base_url.rstrip("/")
    JOURNAL_ID = args.journal_id
    ISSUES_THROUGH = args.issues_through
    OUT_DIR = Path(args.output_dir).expanduser().resolve()
    JSON_OUT = OUT_DIR / "jitp_manifold_metadata.json"


def load_existing_rows(path):
    if not path:
        return {}
    existing_path = Path(path).expanduser()
    if not existing_path.exists():
        return {}
    payload = json.loads(existing_path.read_text(encoding="utf-8"))
    return {
        row.get("record_id"): row
        for row in payload.get("articles", [])
        if row.get("record_id")
    }


def merge_existing_result(row, existing):
    if not existing:
        row["existing_result_used"] = ""
        return row
    used = []
    merged = dict(row)
    for field in EXISTING_FALLBACK_FIELDS:
        if not clean_text(merged.get(field, "")) and clean_text(existing.get(field, "")):
            merged[field] = existing.get(field, "")
            used.append(field)
    for field in EXISTING_STATUS_FIELDS:
        current = clean_text(merged.get(field, "")).lower()
        previous = clean_text(existing.get(field, ""))
        if previous and (not current or current in {"pending", "missing"} or current.startswith("failed")):
            merged[field] = existing.get(field, "")
            used.append(field)
    for field in ("normalized_full_word_count", "normalized_main_word_count", "paragraph_count", "heading_count"):
        if not merged.get(field) and existing.get(field):
            merged[field] = existing.get(field)
            used.append(field)
    merged["existing_result_used"] = ", ".join(sorted(set(used)))
    return merged


def main():
    args = parse_args()
    configure_runtime(args)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    scraped_at = dt.datetime.now(dt.timezone.utc).isoformat()
    existing_json = args.existing_json or str(JSON_OUT)
    existing_by_record = {} if args.no_existing_fallback else load_existing_rows(existing_json)
    issue_entities = parse_issue_pages()

    issues = []
    rows = []
    for issue in issue_entities:
        issue_attrs = issue.get("attributes") or {}
        issue_number = str(issue_attrs.get("number") or "")
        project_slug = issue_attrs.get("projectSlug") or ""
        if not project_slug:
            continue
        project_url = f"{BASE}/projects/{project_slug}"
        project_api_url = f"{BASE}/api/v1/projects/{urllib.parse.quote(project_slug)}?include=texts,text_categories,categories,journal_issue"
        try:
            project = fetch_json(project_api_url)
        except Exception as exc:
            issues.append(
                {
                    "issue_id": issue.get("id", ""),
                    "issue_number": issue_number,
                    "project_slug": project_slug,
                    "project_url": project_url,
                    "project_api_status": f"failed: {exc}",
                }
            )
            continue

        proj = project.get("data") or {}
        proj_attrs = proj.get("attributes") or {}
        included = project.get("included") or []
        categories = {item["id"]: item for item in included if item.get("type") == "categories"}
        texts = [item for item in included if item.get("type") == "texts"]
        issue_type = "Numbered Issue" if issue_number.isdigit() else "Short Form Section"
        issue_editors = extract_issue_editors(proj_attrs.get("descriptionFormatted") or proj_attrs.get("description") or "")
        project_metadata = flatten_metadata(proj_attrs.get("metadata", {}))

        issue_row = {
            "issue_id": issue.get("id", ""),
            "journal_id": JOURNAL_ID,
            "issue_number": issue_number,
            "issue_type": issue_type,
            "issue_title": proj_attrs.get("titlePlaintext") or proj_attrs.get("title") or issue_attrs.get("title") or "",
            "issue_subtitle": proj_attrs.get("subtitlePlaintext") or "",
            "issue_editors": issue_editors,
            "issue_publication_date": proj_attrs.get("publicationDate") or issue_attrs.get("publicationDate") or "",
            "issue_created_at": proj_attrs.get("createdAt") or issue_attrs.get("createdAt") or "",
            "issue_updated_at": proj_attrs.get("updatedAt") or issue_attrs.get("updatedAt") or "",
            "project_id": proj.get("id", ""),
            "project_slug": project_slug,
            "project_url": project_url,
            "project_api_url": project_api_url,
            "text_count_api": len(texts),
            "publisher": project_metadata.get("publisher", ""),
            "publisher_place": project_metadata.get("publisherPlace", ""),
            "project_metadata_json": json.dumps(proj_attrs.get("metadata") or {}, ensure_ascii=False, sort_keys=True),
            "project_description": html_to_text(proj_attrs.get("descriptionFormatted") or proj_attrs.get("description") or ""),
            "cover_image_original": ((proj_attrs.get("coverStyles") or {}).get("original") or ""),
            "avatar_image_original": ((proj_attrs.get("avatarStyles") or {}).get("original") or ""),
            "project_api_status": "fetched",
        }
        issues.append(issue_row)

        for text in texts:
            text_attrs = text.get("attributes") or {}
            cat_id = get_rel_id(text, "category")
            cat = categories.get(cat_id, {})
            cat_attrs = cat.get("attributes") or {}
            author_names = authors_from_description(text_attrs.get("descriptionPlaintext") or text_attrs.get("description") or "")
            row = {
                "record_id": f"{project_slug}::{text.get('id', '')}",
                "scraped_at_utc": scraped_at,
                "journal": "Journal of Interactive Technology and Pedagogy",
                "journal_id": JOURNAL_ID,
                "issue_number": issue_number,
                "issue_type": issue_type,
                "short_form_section": issue_number if issue_type == "Short Form Section" else "",
                "issue_title": issue_row["issue_title"],
                "issue_subtitle": issue_row["issue_subtitle"],
                "issue_editors": issue_editors,
                "issue_publication_date": issue_row["issue_publication_date"],
                "project_id": proj.get("id", ""),
                "project_slug": project_slug,
                "project_url": project_url,
                "text_id": text.get("id", ""),
                "text_slug": text_attrs.get("slug") or "",
                "text_position": text_attrs.get("position") or "",
                "text_category": cat_attrs.get("title") or "",
                "text_category_position": cat_attrs.get("position") or "",
                "text_category_id": cat_id,
                "title": text_attrs.get("titlePlaintext") or text_attrs.get("title") or "",
                "subtitle": text_attrs.get("subtitlePlaintext") or "",
                "text_description": text_attrs.get("descriptionPlaintext") or html_to_text(text_attrs.get("descriptionFormatted") or text_attrs.get("description") or ""),
                "author_names": author_names,
                "author_count_estimate": len(split_author_display(author_names)),
                "text_publication_date": text_attrs.get("publicationDate") or "",
                "text_created_at": text_attrs.get("createdAt") or "",
                "text_updated_at": text_attrs.get("updatedAt") or "",
                "published_flag_api": text_attrs.get("published"),
                "annotations_count": text_attrs.get("annotationsCount"),
                "highlights_count": text_attrs.get("highlightsCount"),
                "article_url": "",
                "start_text_section_id": "",
                "text_api_status": "pending",
                "article_html_status": "pending",
                "bio_status": "pending",
                "abstract_status": "pending",
                "keywords_status": "pending",
                "page_title": "",
                "page_h1": "",
                "author_byline_lines": "",
                "author_affiliations": "",
                "abstract": "",
                "keywords": "",
                "author_bio": "",
                "license": "",
                "normalized_full_text": "",
                "normalized_main_text": "",
                "notes_text": "",
                "references_text": "",
                "acknowledgments_text": "",
                "appendix_text": "",
                "normalized_full_word_count": 0,
                "normalized_main_word_count": 0,
                "paragraph_count": 0,
                "heading_count": 0,
                "toc_headings": "",
                "section_headings": "",
                "text_unique_identifier": "",
                "text_language": "",
                "text_doi": "",
                "text_rights": "",
                "text_metadata_keys": "",
                "text_metadata_json": "",
                "text_citations_json": "",
                "toc_json": "",
                "spine_json": "",
                "ingestion_source_download_url": "",
                "ingestion_external_source_url": "",
                "epub_v3_export_url": "",
                "deleted_slug_flag": "",
                "existing_result_used": "",
            }
            rows.append(row)

    rows.sort(
        key=lambda r: (
            0 if str(r["issue_number"]).isdigit() else 1,
            int(r["issue_number"]) if str(r["issue_number"]).isdigit() else 1000,
            str(r["issue_number"]),
            int(r["text_category_position"] or 999),
            int(r["text_position"] or 999),
            r["title"],
        )
    )

    with concurrent.futures.ThreadPoolExecutor(max_workers=args.max_workers) as executor:
        rows = list(executor.map(scrape_text, rows))

    if existing_by_record:
        rows = [merge_existing_result(row, existing_by_record.get(row["record_id"])) for row in rows]

    summary = {
        "scraped_at_utc": scraped_at,
        "source_base_url": BASE,
        "journal_id": JOURNAL_ID,
        "issue_count": len(issues),
        "article_record_count": len(rows),
        "author_bio_found_count": sum(1 for r in rows if r.get("bio_status") == "found"),
        "author_bio_missing_count": sum(1 for r in rows if r.get("bio_status") != "found"),
        "abstract_found_count": sum(1 for r in rows if r.get("abstract_status") == "found"),
        "keywords_found_count": sum(1 for r in rows if r.get("keywords_status") == "found"),
        "deleted_slug_flag_count": sum(1 for r in rows if r.get("deleted_slug_flag") == "yes"),
        "existing_result_used_count": sum(1 for r in rows if r.get("existing_result_used")),
    }

    payload = {"summary": summary, "issues": issues, "articles": rows}
    JSON_OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(str(JSON_OUT))


if __name__ == "__main__":
    main()
