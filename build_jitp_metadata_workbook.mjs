import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const repoRoot = path.dirname(fileURLToPath(import.meta.url));

function argValue(name, fallback) {
  const index = process.argv.indexOf(name);
  if (index >= 0 && process.argv[index + 1]) return process.argv[index + 1];
  return fallback;
}

const outputDir = path.resolve(argValue(
  "--output-dir",
  process.env.JITP_OUTPUT_DIR || path.join(repoRoot, "outputs", "jitp_manifold_metadata"),
));
const inputPath = path.resolve(argValue(
  "--input-json",
  process.env.JITP_INPUT_JSON || path.join(outputDir, "jitp_manifold_metadata.json"),
));
const outputPath = path.resolve(argValue(
  "--output-xlsx",
  process.env.JITP_OUTPUT_XLSX || path.join(outputDir, "jitp_manifold_article_metadata.xlsx"),
));

const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
const articles = raw.articles;
const issues = raw.issues;
const summary = raw.summary;
const scrapedAtDisplay = (summary.scraped_at_utc || "")
  .replace(/\.\d+/, "")
  .replace("T", " ")
  .replace("+00:00", " UTC");

function colLetter(index) {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function asText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function compact(value) {
  return asText(value).replace(/\s+/g, " ").trim();
}

function unique(values) {
  return [...new Set(values.map(compact).filter(Boolean))];
}

function joinUnique(values, separator = " | ") {
  return unique(values).join(separator);
}

function normalizeName(name) {
  return compact(name)
    .toLowerCase()
    .replace(/[’']/g, "")
    .replace(/\s+/g, " ");
}

function dateOnly(value) {
  const text = compact(value);
  if (!text) return "";
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : text;
}

function parseJsonObject(value) {
  const text = compact(value);
  if (!text) return {};
  try {
    const parsed = JSON.parse(text);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function bestPublicationDate(row) {
  const metadata = parseJsonObject(row.text_metadata_json);
  const candidates = [
    ["text_publication_date", row.text_publication_date],
    ["issue_publication_date", row.issue_publication_date],
    ["text_metadata.originalPublicationDate", metadata.originalPublicationDate],
    ["text_created_at", row.text_created_at],
    ["text_updated_at", row.text_updated_at],
  ];
  const found = candidates.find(([, value]) => compact(value));
  if (!found) return { date: "", source: "" };
  return { date: dateOnly(found[1]), source: found[0] };
}

function noteFor(row) {
  const title = (row.title || "").toLowerCase();
  if (row.deleted_slug_flag === "yes") return "Deleted/404 text entry exposed by Manifold API.";
  if (title.includes("guidelines") || title.includes("call for submissions") || title.includes("about the journal")) {
    return "Likely journal admin/guideline text included in a Manifold issue project.";
  }
  if (row.bio_status !== "found") return "No author bio visible in rendered article page.";
  if (row.article_html_status !== "parsed") return "Article page did not parse cleanly.";
  return "";
}

function parseBylineLine(line) {
  const parts = compact(line)
    .replace(/\[\d+\]/g, "")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  if (!parts.length) return { name: "", affiliation: "" };
  if (parts.length > 1 && /^(Jr\.?|Sr\.?|II|III|IV)$/i.test(parts[1])) {
    return { name: `${parts[0]}, ${parts[1]}`, affiliation: cleanAffiliation(parts.slice(2).join(", ")) };
  }
  return { name: parts[0], affiliation: cleanAffiliation(parts.slice(1).join(", ")) };
}

function parseAuthorListSegment(value, role = "Author", source = "author_names") {
  let text = compact(value).replace(/^by\s+/i, "");
  if (!text) return [];
  text = text.replace(/\s*&\s*/g, " and ");
  text = text.replace(/,\s*(Jr\.?|Sr\.?|II|III|IV)\b/gi, "§SUFFIX§ $1");
  text = text.replace(/\s+and\s+/gi, ", ");
  return text
    .split(/\s*,\s*/)
    .map((name) => name.replace(/§SUFFIX§/g, ",").trim())
    .filter(Boolean)
    .map((name) => ({ name, role, affiliation: "", source }));
}

function parseAuthorString(value, source = "author_names") {
  const text = compact(value);
  if (!text) return [];
  return text
    .split(/\s*\|\s*/)
    .flatMap((segment) => {
      const match = segment.match(/^(.*?)(?:,?\s+in conversation with\s+)(.+)$/i);
      if (!match) return parseAuthorListSegment(segment, "Author", source);
      return [
        ...parseAuthorListSegment(match[1], "Author", source),
        ...parseAuthorListSegment(match[2], "Conversation participant", source),
      ];
    })
    .filter((record) => record.name);
}

const sectionStopPattern =
  /^(abstract|introduction|background|method|methods|results|discussion|conclusion|notes?|references|bibliography|works cited|acknowledg(?:e)?ments?|appendix|about the author|about the authors|about the editors|mission statement|review policy|masthead|how well|teaching with|the openlab|course overview|literature review)$/i;

const institutionWordsPattern =
  /\b(university|college|institute|school|library|libraries|department|program|center|centre|press|project|foundation|museum|community|cuny|suny|nyu|pratt|fordham|city tech|laguardia)\b/i;

function splitHeadingText(value) {
  return compact(value)
    .split(/\s*\|\s*/)
    .map((line) => line.replace(/\[\d+\]/g, "").trim())
    .filter(Boolean);
}

function cleanAffiliation(value) {
  return compact(value)
    .replace(/\s+(As we|As instructors|This essay|This article|This paper|In this|With the|Although existing|Though use|Open Educational|Pandemic teaching)\b.*$/i, "")
    .replace(/\s+##.*$/i, "")
    .trim();
}

function bioAuthorNames(bio) {
  const text = compact(bio);
  if (!text) return [];
  const namePattern = /(?:^|[.!?]\s+)([\p{Lu}][\p{L}\p{M}'’.-]+(?:\s+(?:[\p{Lu}]\.|[\p{Lu}][\p{L}\p{M}'’.-]+|de|del|van|von)){1,6})\s+(?:is|are|was|were|earned|received|leads|serves|completed|holds|has|works|teaches|studies|graduated|coordinates|directs|develops|creates|researches|uses|conducted|sits|currently)\b/gu;
  return unique([...text.matchAll(namePattern)].map((match) => match[1]));
}

function looksLikeAuthorCandidate(value) {
  const text = compact(value).replace(/\[\d+\]/g, "");
  if (!text.includes(",")) return false;
  const [candidate] = text.split(",");
  const name = compact(candidate);
  if (!name || institutionWordsPattern.test(name)) return false;
  if (!/[\p{Lu}]/u.test(name)) return false;
  if (name.split(/\s+/).length < 2) return false;
  return /^[\p{Lu}][\p{L}\p{M}'’.-]+(?:\s+(?:[\p{Lu}]\.|[\p{Lu}][\p{L}\p{M}'’.-]+|de|del|van|von)){1,6}$/u.test(name);
}

function splitByNamedAnchors(blob, names) {
  const text = compact(blob).replace(/\[\d+\]/g, "");
  const anchors = names
    .map((name) => ({ name, index: text.indexOf(name) }))
    .filter((anchor) => anchor.index >= 0)
    .sort((a, b) => a.index - b.index);
  if (anchors.length < 2) return [];
  return anchors
    .map((anchor, index) => {
      const end = anchors[index + 1]?.index ?? text.length;
      return text.slice(anchor.index, end).trim();
    })
    .filter(Boolean);
}

function splitBylineBlob(blob, bioNames = []) {
  const text = compact(blob).replace(/\[\d+\]/g, "");
  if (!text) return [];
  const named = splitByNamedAnchors(text, bioNames);
  if (named.length) return named;
  if (text.includes(" | ")) return splitHeadingText(text);
  if (looksLikeAuthorCandidate(text)) {
    const candidatePattern = /(?:^|\s)([\p{Lu}][\p{L}\p{M}'’.-]+(?:\s+(?:[\p{Lu}]\.|[\p{Lu}][\p{L}\p{M}'’.-]+|de|del|van|von)){1,6}),/gu;
    const matches = [...text.matchAll(candidatePattern)]
      .map((match) => ({ name: match[1], index: match.index + match[0].indexOf(match[1]) }))
      .filter((match) => !institutionWordsPattern.test(match.name));
    if (matches.length > 1) {
      return matches.map((match, index) => {
        const end = matches[index + 1]?.index ?? text.length;
        return text.slice(match.index, end).trim();
      });
    }
  }
  return [text];
}

function headingBylineEntries(row) {
  const headings = splitHeadingText(row.author_byline_lines || row.section_headings || row.toc_headings);
  if (!headings.length) return [];
  const title = compact(row.title).toLowerCase();
  const startIndex = headings.findIndex((heading) => heading.toLowerCase() === title);
  const afterTitle = headings.slice(startIndex >= 0 ? startIndex + 1 : 1);
  const bylines = [];
  for (const heading of afterTitle) {
    if (sectionStopPattern.test(heading)) break;
    if (heading.length > 220) break;
    if (heading.includes(",") || bioAuthorNames(row.author_bio).some((name) => heading.includes(name))) bylines.push(heading);
  }
  return bylines.map((line) => ({ line, source: row.author_byline_lines ? "rendered_byline" : "inferred_heading" }));
}

function textPrefixBylineEntries(row) {
  const text = compact(row.normalized_full_text);
  if (!text) return [];
  const title = compact(row.title);
  let rest = text;
  if (rest.startsWith("#")) rest = rest.replace(/^#+\s*/, "");
  if (title && rest.toLowerCase().startsWith(title.toLowerCase())) rest = rest.slice(title.length).trim();

  if (rest.startsWith("###")) {
    const beforeBody = rest.split(/\s+##\s+/)[0];
    return beforeBody
      .split(/\s+###\s+/)
      .map((line) => line.replace(/^###\s*/, "").trim())
      .filter((line) => line && !sectionStopPattern.test(line))
      .map((line) => ({ line, source: "inferred_text_heading" }));
  }

  const abstractIndex = rest.search(/\s+##\s+Abstract\b/i);
  const bioNames = bioAuthorNames(row.author_bio);
  if (abstractIndex > 0) {
    const candidate = rest.slice(0, abstractIndex).trim();
    if (candidate && candidate.length < 1200) return [{ line: candidate, source: "inferred_text_prefix" }];
  }
  if (bioNames.length) {
    const first = Math.min(...bioNames.map((name) => rest.indexOf(name)).filter((idx) => idx >= 0));
    const lastName = bioNames
      .map((name) => ({ name, index: rest.indexOf(name) }))
      .filter((match) => match.index >= 0)
      .sort((a, b) => b.index - a.index)[0];
    if (Number.isFinite(first) && lastName) {
      const afterLast = rest.slice(lastName.index + lastName.name.length);
      const tailStop = afterLast.search(/\s+(As we|As instructors|This essay|This article|This paper|In this|With the|Although existing|Though use|Open Educational|Pandemic teaching)\b/i);
      const end = tailStop >= 0 ? lastName.index + lastName.name.length + tailStop : Math.min(rest.length, first + 500);
      return [{ line: rest.slice(first, end).trim(), source: "inferred_text_prefix" }];
    }
  }
  return [];
}

function bylineEntriesForArticle(row) {
  const rendered = splitHeadingText(row.author_byline_lines).map((line) => ({ line, source: "rendered_byline" }));
  if (rendered.length) return rendered;
  const heading = headingBylineEntries(row);
  if (heading.length) return heading;
  const textPrefix = textPrefixBylineEntries(row);
  if (textPrefix.length) return textPrefix;
  return [];
}

function filledAbstract(row) {
  if (compact(row.abstract)) return { value: row.abstract, source: row.abstract_status === "found" ? "scraped_abstract" : "raw_abstract" };
  const text = compact(row.normalized_full_text);
  const match = text.match(/\s##\s+Abstract\b\s+([\s\S]*?)(?=\s##\s+[\p{Lu}0-9]|\sThis entry is licensed|\sBack to top|$)/iu);
  if (!match) return { value: "", source: "" };
  return { value: compact(match[1]), source: "normalized_text_abstract_section" };
}

function filledBio(row) {
  if (compact(row.author_bio)) return { value: row.author_bio, source: row.bio_status === "found" ? "scraped_bio" : "raw_bio" };
  const text = compact(row.normalized_full_text);
  const match = text.match(/\s##\s+About the Authors?\b\s+([\s\S]*?)(?=\s##\s+[\p{Lu}0-9]|\sThis entry is licensed|\sBack to top|$)/iu);
  if (!match) return { value: "", source: "" };
  return { value: compact(match[1]), source: "normalized_text_about_section" };
}

function biosByAuthor(row, records) {
  const bio = filledBio(row).value;
  if (!bio || !records.length) return new Map();
  if (records.length === 1) {
    const index = bio.indexOf(records[0].name);
    return new Map([[records[0].split_key, index > 0 ? bio.slice(index).trim() : bio]]);
  }
  const anchors = records
    .map((record) => ({ key: record.split_key, name: record.name, index: bio.indexOf(record.name) }))
    .filter((anchor) => anchor.index >= 0)
    .sort((a, b) => a.index - b.index);
  if (!anchors.length) return new Map();
  const out = new Map();
  anchors.forEach((anchor, index) => {
    const end = anchors[index + 1]?.index ?? bio.length;
    out.set(anchor.key, bio.slice(anchor.index, end).trim());
  });
  return out;
}

function recordsFromBylineEntries(entries, bioNames) {
  return entries.flatMap(({ line, source }) => splitBylineBlob(line, bioNames).flatMap((segment) => {
    const parsed = parseBylineLine(segment);
    const names = parseAuthorString(parsed.name, source);
    if (!names.length && parsed.name) {
      names.push({ name: parsed.name, role: "Author", affiliation: "", source });
    }
    return names.map((record) => ({
      ...record,
      affiliation: parsed.affiliation || record.affiliation || "",
      source,
    }));
  }));
}

function applyRowAffiliations(records, row) {
  const affiliations = splitHeadingText(row.author_affiliations);
  if (!affiliations.length) return records;
  return records.map((record, index) => ({
    ...record,
    affiliation: record.affiliation || affiliations[index] || (affiliations.length === 1 ? affiliations[0] : ""),
  }));
}

function authorRecordsForArticle(row) {
  const bioNames = bioAuthorNames(filledBio(row).value);
  const renderedBylines = splitHeadingText(row.author_byline_lines).map((line) => ({ line, source: "rendered_byline" }));
  let records = [];
  if (renderedBylines.length) {
    records = recordsFromBylineEntries(renderedBylines, bioNames);
  } else if (compact(row.author_names)) {
    records = applyRowAffiliations(parseAuthorString(row.author_names, "api_description"), row);
  } else {
    records = recordsFromBylineEntries(bylineEntriesForArticle(row), bioNames);
  }
  const keyedRecords = records.map((record) => ({ ...record, split_key: normalizeName(record.name) }));
  const authorBios = biosByAuthor(row, keyedRecords);
  return keyedRecords.map((record, index) => ({
    ...record,
    author_position: index + 1,
    author_count_for_article: records.length,
    record_id: row.record_id,
    text_id: row.text_id,
    project_id: row.project_id,
    issue_number: row.issue_number,
    issue_type: row.issue_type,
    short_form_section: row.short_form_section,
    text_category: row.text_category,
    text_position: row.text_position,
    title: row.title,
    subtitle: row.subtitle,
    article_url: row.article_url,
    project_url: row.project_url,
    text_publication_date: row.text_publication_date,
    issue_publication_date: row.issue_publication_date,
    best_publication_date: bestPublicationDate(row).date,
    best_publication_date_source: bestPublicationDate(row).source,
    article_author_cell: row.author_names,
    article_author_cell_filled: row.author_names || records.map((item) => item.name).join(" | "),
    article_bio_cell: row.author_bio,
    article_bio_cell_filled: filledBio(row).value,
    article_bio_source: filledBio(row).source,
    author_bio: authorBios.get(record.split_key) || "",
    abstract: filledAbstract(row).value,
    abstract_source: filledAbstract(row).source,
    keywords: row.keywords,
    bio_status: row.bio_status,
    abstract_status: row.abstract_status,
    keywords_status: row.keywords_status,
    license: row.license,
    record_note: noteFor(row),
    existing_result_used: row.existing_result_used,
  }));
}

const authorPubRecords = articles.flatMap((row) => authorRecordsForArticle(row));
const uniqueAuthorCount = new Set(authorPubRecords.map((record) => record.split_key).filter(Boolean)).size;
const recordsByArticle = new Map();
for (const record of authorPubRecords) {
  if (!recordsByArticle.has(record.record_id)) recordsByArticle.set(record.record_id, []);
  recordsByArticle.get(record.record_id).push(record);
}
const maxAuthorSlots = Math.max(1, ...[...recordsByArticle.values()].map((records) => records.length));

function issueSortKey(value) {
  const text = String(value || "");
  return /^\d+$/.test(text) ? [0, Number(text)] : [1, text];
}

function sortIssueLike(a, b) {
  const ak = issueSortKey(a.issue_number);
  const bk = issueSortKey(b.issue_number);
  if (ak[0] !== bk[0]) return ak[0] - bk[0];
  if (ak[1] < bk[1]) return -1;
  if (ak[1] > bk[1]) return 1;
  return 0;
}

function writeTable(sheet, columns, rows, tableName) {
  const data = [
    columns.map(([, label]) => label),
    ...rows.map((row) => columns.map(([key]) => asText(typeof key === "function" ? key(row) : row[key]))),
  ];
  sheet.getRangeByIndexes(0, 0, data.length, columns.length).values = data;
  sheet.tables.add(`A1:${colLetter(columns.length - 1)}${data.length}`, true, tableName);
  sheet.getRangeByIndexes(0, 0, 1, columns.length).format = {
    font: { bold: true, color: "#FFFFFF" },
    fill: "#1F4E5F",
    wrapText: true,
    verticalAlignment: "bottom",
  };
  if (data.length > 1) {
    sheet.getRangeByIndexes(1, 0, data.length - 1, columns.length).format = {
      wrapText: true,
      verticalAlignment: "top",
    };
  }
}

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1000, 1).format.columnWidthPx = width;
  });
}

const recordsByAuthor = new Map();
for (const record of authorPubRecords) {
  if (!record.split_key) continue;
  if (!recordsByAuthor.has(record.split_key)) recordsByAuthor.set(record.split_key, []);
  recordsByAuthor.get(record.split_key).push(record);
}

const authorMetadataRows = [...recordsByAuthor.entries()]
  .map(([key, records]) => {
    const sorted = [...records].sort((a, b) => {
      const ad = a.best_publication_date || a.text_publication_date || a.issue_publication_date || "";
      const bd = b.best_publication_date || b.text_publication_date || b.issue_publication_date || "";
      return ad.localeCompare(bd) || sortIssueLike(a, b);
    });
    const publicationDates = unique(records.map((record) => record.best_publication_date || record.text_publication_date || record.issue_publication_date));
    return {
      author_name: sorted[0].name,
      author_name_normalized: key,
      publication_count: new Set(records.map((record) => record.record_id)).size,
      author_record_count: records.length,
      first_publication_date: publicationDates.sort()[0] || "",
      latest_publication_date: publicationDates.sort().at(-1) || "",
      numbered_issue_records: records.filter((record) => record.issue_type === "Numbered Issue").length,
      short_form_records: records.filter((record) => record.issue_type === "Short Form Section").length,
      issues_sections: joinUnique(records.map((record) => record.issue_number), ", "),
      roles: joinUnique(records.map((record) => record.role), ", "),
      affiliations: joinUnique(records.map((record) => record.affiliation)),
      bio_statuses: joinUnique(records.map((record) => record.bio_status), ", "),
      representative_bio: sorted.find((record) => compact(record.author_bio))?.author_bio ||
        sorted.find((record) => compact(record.article_bio_cell_filled))?.article_bio_cell_filled ||
        "",
      article_titles: joinUnique(records.map((record) => record.title)),
      article_urls: joinUnique(records.map((record) => record.article_url)),
      source_author_cells: joinUnique(records.map((record) => record.article_author_cell_filled)),
      split_sources: joinUnique(records.map((record) => record.source), ", "),
      record_notes: joinUnique(records.map((record) => record.record_note)),
    };
  })
  .sort((a, b) => a.author_name.localeCompare(b.author_name));

const issueSummaryRows = issues
  .filter((issue) => /^\d+$/.test(String(issue.issue_number || "")))
  .sort(sortIssueLike)
  .map((issue) => {
    const issueArticles = articles.filter((row) => row.issue_number === issue.issue_number);
    const issueAuthors = authorPubRecords.filter((record) => record.issue_number === issue.issue_number);
    return {
      ...issue,
      text_records_found: issueArticles.length,
      author_instances_found: issueAuthors.length,
      unique_authors_found: new Set(issueAuthors.map((record) => record.split_key).filter(Boolean)).size,
      bios_found: issueArticles.filter((row) => row.bio_status === "found").length,
      abstracts_found: issueArticles.filter((row) => row.abstract_status === "found").length,
      keywords_found: issueArticles.filter((row) => row.keywords_status === "found").length,
      qa_flagged_records: issueArticles.filter((row) => noteFor(row)).length,
      article_titles: joinUnique(issueArticles.map((row) => row.title)),
    };
  });

const shortSectionRows = issues
  .filter((issue) => issue.issue_type === "Short Form Section")
  .sort((a, b) => String(a.issue_number).localeCompare(String(b.issue_number)))
  .map((issue) => {
    const shortArticles = articles.filter((row) => row.issue_number === issue.issue_number);
    const shortAuthors = authorPubRecords.filter((record) => record.issue_number === issue.issue_number);
    return {
      ...issue,
      section_name: issue.issue_number,
      text_records_found: shortArticles.length,
      author_instances_found: shortAuthors.length,
      unique_authors_found: new Set(shortAuthors.map((record) => record.split_key).filter(Boolean)).size,
      bios_found: shortArticles.filter((row) => row.bio_status === "found").length,
      abstracts_found: shortArticles.filter((row) => row.abstract_status === "found").length,
      keywords_found: shortArticles.filter((row) => row.keywords_status === "found").length,
      qa_flagged_records: shortArticles.filter((row) => noteFor(row)).length,
      article_titles: joinUnique(shortArticles.map((row) => row.title)),
    };
  });

const issueStatsByNumber = new Map(issueSummaryRows.map((row) => [row.issue_number, row]));
const shortStatsBySection = new Map(shortSectionRows.map((row) => [row.issue_number, row]));

function sortArticleRows(a, b) {
  const issueCompare = sortIssueLike(a, b);
  if (issueCompare !== 0) return issueCompare;
  return Number(a.text_position || 0) - Number(b.text_position || 0) ||
    String(a.title).localeCompare(String(b.title));
}

function publicationRow(row) {
  const authors = recordsByArticle.get(row.record_id) || [];
  const bestDate = bestPublicationDate(row);
  const abstract = filledAbstract(row);
  const bio = filledBio(row);
  const stats = row.issue_type === "Short Form Section"
    ? shortStatsBySection.get(row.issue_number)
    : issueStatsByNumber.get(row.issue_number);
  const out = {
    record_id: row.record_id,
    issue_number: row.issue_number,
    issue_title: row.issue_title,
    issue_subtitle: row.issue_subtitle,
    issue_editors: row.issue_editors,
    issue_type: row.issue_type,
    short_form_section: row.short_form_section || (/^\d+$/.test(String(row.issue_number)) ? "" : row.issue_number),
    text_category: row.text_category,
    text_position: row.text_position,
    title: row.title,
    subtitle: row.subtitle,
    best_publication_date: bestDate.date,
    best_publication_date_source: bestDate.source,
    text_publication_date: dateOnly(row.text_publication_date),
    issue_publication_date: dateOnly(row.issue_publication_date),
    author_count: authors.length,
    authors_combined: joinUnique(authors.map((record) => record.name), " | "),
    author_affiliations_combined: joinUnique(authors.map((record) => [record.name, record.affiliation].filter(Boolean).join(": "))),
    raw_author_cell: row.author_names,
    filled_author_cell: row.author_names || authors.map((record) => record.name).join(" | "),
    author_split_sources: joinUnique(authors.map((record) => record.source), ", "),
    abstract: abstract.value,
    abstract_source: abstract.source,
    keywords: row.keywords,
    keywords_status: row.keywords_status,
    article_level_bio: bio.value,
    article_bio_source: bio.source,
    bio_status: row.bio_status,
    normalized_main_text: row.normalized_main_text,
    notes_text: row.notes_text,
    references_text: row.references_text,
    acknowledgments_text: row.acknowledgments_text,
    appendix_text: row.appendix_text,
    normalized_main_word_count: row.normalized_main_word_count,
    normalized_full_word_count: row.normalized_full_word_count,
    paragraph_count: row.paragraph_count,
    heading_count: row.heading_count,
    section_headings: row.section_headings,
    toc_headings: row.toc_headings,
    license: row.license,
    article_url: row.article_url,
    project_url: row.project_url,
    text_id: row.text_id,
    project_id: row.project_id,
    text_slug: row.text_slug,
    text_doi: row.text_doi,
    text_unique_identifier: row.text_unique_identifier,
    issue_text_records_found: stats?.text_records_found || "",
    issue_author_instances_found: stats?.author_instances_found || "",
    issue_unique_authors_found: stats?.unique_authors_found || "",
    record_note: noteFor(row),
    existing_result_used: row.existing_result_used,
  };
  authors.forEach((record, index) => {
    const slot = index + 1;
    out[`author_${slot}_name`] = record.name;
    out[`author_${slot}_role`] = record.role;
    out[`author_${slot}_affiliation`] = record.affiliation;
    out[`author_${slot}_bio`] = record.author_bio;
  });
  return out;
}

const issueRows = articles
  .filter((row) => /^\d+$/.test(String(row.issue_number || "")))
  .sort(sortArticleRows)
  .map(publicationRow);

const shortRows = articles
  .filter((row) => row.issue_type === "Short Form Section")
  .sort(sortArticleRows)
  .map(publicationRow);

const workbook = Workbook.create();
const sheets = {
  authors: workbook.worksheets.add("Author Metadata"),
  authorPub: workbook.worksheets.add("Author-Pub Records"),
  issues: workbook.worksheets.add("Issues"),
  shorts: workbook.worksheets.add("Shorts"),
  field: workbook.worksheets.add("Field Notes"),
};

for (const sheet of Object.values(sheets)) sheet.showGridLines = false;

const authorMetadataColumns = [
  ["author_name", "Author Name"],
  ["author_name_normalized", "Normalized Name"],
  ["publication_count", "Publication Count"],
  ["author_record_count", "Author-Record Count"],
  ["first_publication_date", "First Publication Date"],
  ["latest_publication_date", "Latest Publication Date"],
  ["numbered_issue_records", "Numbered Issue Records"],
  ["short_form_records", "Short-Form Records"],
  ["issues_sections", "Issues / Sections"],
  ["roles", "Roles"],
  ["affiliations", "Affiliations"],
  ["bio_statuses", "Bio Statuses"],
  ["representative_bio", "Representative Bio"],
  ["article_titles", "Article Titles"],
  ["article_urls", "Article URLs"],
  ["source_author_cells", "Source Author Cells"],
  ["split_sources", "Split Sources"],
  ["record_notes", "Record Notes"],
];
writeTable(sheets.authors, authorMetadataColumns, authorMetadataRows, "AuthorMetadataTable");
sheets.authors.freezePanes.freezeRows(1);
sheets.authors.freezePanes.freezeColumns(2);
setWidths(sheets.authors, [260, 240, 115, 120, 125, 125, 115, 115, 220, 180, 360, 180, 620, 620, 420, 520, 180, 360]);
sheets.authors.getRangeByIndexes(1, 0, Math.max(authorMetadataRows.length, 1), authorMetadataColumns.length).format.rowHeightPx = 88;

const authorPubColumns = [
  ["record_id", "Record ID"],
  ["author_position", "Author Position"],
  ["author_count_for_article", "Article Author Count"],
  ["name", "Author Name"],
  ["role", "Role"],
  ["affiliation", "Affiliation"],
  ["source", "Split Source"],
  ["article_author_cell", "Original Author Cell"],
  ["article_author_cell_filled", "Filled Author Cell"],
  ["issue_number", "Issue / Section"],
  ["issue_type", "Issue Type"],
  ["short_form_section", "Short Form Section"],
  ["text_category", "Text Category"],
  ["text_position", "Text Position"],
  ["title", "Article Title"],
  ["subtitle", "Subtitle"],
  ["best_publication_date", "Best Publication Date"],
  ["best_publication_date_source", "Date Source"],
  ["text_publication_date", "Text Publication Date"],
  ["issue_publication_date", "Issue Publication Date"],
  ["abstract", "Abstract"],
  ["abstract_source", "Abstract Source"],
  ["keywords", "Keywords"],
  ["bio_status", "Bio Status"],
  ["author_bio", "Author Bio"],
  ["article_bio_cell", "Article-Level Bio"],
  ["article_bio_cell_filled", "Filled Article-Level Bio"],
  ["article_bio_source", "Bio Source"],
  ["license", "License"],
  ["record_note", "Record Note"],
  ["existing_result_used", "Existing Result Fallback Fields"],
  ["article_url", "Article URL"],
  ["project_url", "Project URL"],
  ["text_id", "Text ID"],
  ["project_id", "Project ID"],
];
writeTable(sheets.authorPub, authorPubColumns, authorPubRecords, "AuthorPubRecordsTable");
sheets.authorPub.freezePanes.freezeRows(1);
sheets.authorPub.freezePanes.freezeColumns(4);
setWidths(sheets.authorPub, [260, 95, 95, 240, 170, 280, 140, 360, 360, 95, 130, 130, 140, 90, 360, 240, 120, 160, 120, 120, 460, 170, 260, 130, 620, 620, 620, 150, 340, 320, 360, 320, 260, 260]);
sheets.authorPub.getRangeByIndexes(1, 0, Math.max(authorPubRecords.length, 1), authorPubColumns.length).format.rowHeightPx = 82;

const authorSlotColumns = Array.from({ length: maxAuthorSlots }, (_, idx) => {
  const slot = idx + 1;
  return [
    [`author_${slot}_name`, `Author ${slot} Name`],
    [`author_${slot}_role`, `Author ${slot} Role`],
    [`author_${slot}_affiliation`, `Author ${slot} Affiliation`],
    [`author_${slot}_bio`, `Author ${slot} Bio`],
  ];
}).flat();

const publicationColumns = [
  ["record_id", "Record ID"],
  ["issue_number", "Issue / Section"],
  ["issue_title", "Issue / Project Title"],
  ["issue_subtitle", "Issue Subtitle"],
  ["issue_editors", "Issue Editors"],
  ["issue_type", "Issue Type"],
  ["short_form_section", "Short Form Section"],
  ["text_category", "Text Category"],
  ["text_position", "Text Position"],
  ["title", "Article Title"],
  ["subtitle", "Subtitle"],
  ["best_publication_date", "Best Publication Date"],
  ["best_publication_date_source", "Date Source"],
  ["text_publication_date", "Text Publication Date"],
  ["issue_publication_date", "Issue Publication Date"],
  ["author_count", "Author Count"],
  ["authors_combined", "Authors Combined"],
  ["author_affiliations_combined", "Author Affiliations Combined"],
  ["raw_author_cell", "Raw Author Cell"],
  ["filled_author_cell", "Filled Author Cell"],
  ["author_split_sources", "Author Split Sources"],
  ...authorSlotColumns,
  ["abstract", "Abstract"],
  ["abstract_source", "Abstract Source"],
  ["keywords", "Keywords"],
  ["keywords_status", "Keywords Status"],
  ["article_level_bio", "Article-Level Bio"],
  ["article_bio_source", "Bio Source"],
  ["bio_status", "Bio Status"],
  ["normalized_main_text", "Normalized Main Text"],
  ["notes_text", "Notes Text"],
  ["references_text", "References Text"],
  ["acknowledgments_text", "Acknowledgments Text"],
  ["appendix_text", "Appendix Text"],
  ["normalized_main_word_count", "Main Word Count"],
  ["normalized_full_word_count", "Full Word Count"],
  ["paragraph_count", "Paragraph Count"],
  ["heading_count", "Heading Count"],
  ["section_headings", "Section Headings"],
  ["toc_headings", "TOC Headings"],
  ["license", "License"],
  ["article_url", "Article URL"],
  ["project_url", "Project URL"],
  ["text_id", "Text ID"],
  ["project_id", "Project ID"],
  ["text_slug", "Text Slug"],
  ["text_doi", "Text DOI"],
  ["text_unique_identifier", "Text Unique Identifier"],
  ["issue_text_records_found", "Project Text Records Found"],
  ["issue_author_instances_found", "Project Author Instances Found"],
  ["issue_unique_authors_found", "Project Unique Authors Found"],
  ["record_note", "Record Note"],
  ["existing_result_used", "Existing Result Fallback Fields"],
];
const issueColumns = publicationColumns;
writeTable(sheets.issues, issueColumns, issueRows, "IssuesTable");
sheets.issues.freezePanes.freezeRows(1);
sheets.issues.freezePanes.freezeColumns(10);
setWidths(sheets.issues, [
  260, 90, 280, 260, 280, 130, 130, 130, 85, 380, 220, 120, 160, 120, 120, 95, 360, 420, 360, 360, 180,
  ...Array.from({ length: maxAuthorSlots }, () => [230, 140, 280, 520]).flat(),
  520, 170, 260, 130, 620, 150, 130, 700, 520, 620, 420, 420, 110, 110, 100, 100, 520, 520, 340, 360, 340, 260, 260, 260, 220, 220, 100, 120, 120, 340,
]);
sheets.issues.getRangeByIndexes(1, 0, Math.max(issueRows.length, 1), issueColumns.length).format.rowHeightPx = 94;

const shortColumns = publicationColumns;
writeTable(sheets.shorts, shortColumns, shortRows, "ShortsTable");
sheets.shorts.freezePanes.freezeRows(1);
sheets.shorts.freezePanes.freezeColumns(10);
setWidths(sheets.shorts, [
  260, 130, 280, 260, 220, 130, 140, 130, 85, 380, 220, 120, 160, 120, 120, 95, 360, 420, 360, 360, 180,
  ...Array.from({ length: maxAuthorSlots }, () => [230, 140, 280, 520]).flat(),
  520, 170, 260, 130, 620, 150, 130, 700, 520, 620, 420, 420, 110, 110, 100, 100, 520, 520, 340, 360, 340, 260, 260, 260, 220, 220, 100, 120, 120, 340,
]);
sheets.shorts.getRangeByIndexes(1, 0, Math.max(shortRows.length, 1), shortColumns.length).format.rowHeightPx = 94;

const articlesWithRecoveredAuthors = articles.filter((row) =>
  !compact(row.author_names) && (recordsByArticle.get(row.record_id) || []).length
).length;
const articlesMissingAuthorsAfterFill = articles.filter((row) => !(recordsByArticle.get(row.record_id) || []).length).length;
const articlesWithFilledAbstract = articles.filter((row) => !compact(row.abstract) && compact(filledAbstract(row).value)).length;
const articlesWithFilledBio = articles.filter((row) => !compact(row.author_bio) && compact(filledBio(row).value)).length;
const articlesWithFallbackDate = articles.filter((row) => !compact(row.text_publication_date) && compact(bestPublicationDate(row).date)).length;
const shortSectionCounts = shortSectionRows
  .map((row) => `${row.section_name}: ${row.text_records_found}`)
  .join(" | ");

const fieldRows = [
  { field: "Workbook Scope", note: "Focused JITP Manifold metadata workbook with author-level metadata, author-publication records, article-level numbered issue records, article-level short-form records, and field notes." },
  { field: "Source Base", note: summary.source_base_url },
  { field: "Scraped At UTC", note: scrapedAtDisplay },
  { field: "Text Records", note: String(summary.article_record_count) },
  { field: "Author Instances Found", note: String(authorPubRecords.length) },
  { field: "Unique Author Names Found", note: String(uniqueAuthorCount) },
  { field: "Issues / Sections Crawled", note: String(summary.issue_count) },
  { field: "Numbered Issues", note: String(issueSummaryRows.length) },
  { field: "Numbered Issue Publications", note: String(issueRows.length) },
  { field: "Short-Form Sections", note: String(shortSectionRows.length) },
  { field: "Short-Form Publications", note: String(shortRows.length) },
  { field: "Short-Form Publication Counts", note: shortSectionCounts },
  { field: "Recovered Author Rows", note: `${articlesWithRecoveredAuthors} article rows had missing raw author cells but recoverable author names in rendered headings or normalized text.` },
  { field: "Missing Authors After Fill", note: String(articlesMissingAuthorsAfterFill) },
  { field: "Recovered Abstract Rows", note: `${articlesWithFilledAbstract} rows had abstracts filled from the normalized article text.` },
  { field: "Recovered Bio Rows", note: `${articlesWithFilledBio} rows had bios filled from the normalized article text.` },
  { field: "Fallback Date Rows", note: `${articlesWithFallbackDate} rows use an issue, metadata, created, or updated date where a text publication date was not present.` },
  { field: "Author Metadata", note: "One row per normalized author name. Article titles, URLs, affiliations, roles, bios, and record notes are aggregated across that author's publication records." },
  { field: "Author-Pub Records", note: "One row per parsed author/article relationship. Author names are split from the original single-cell byline string and include a split source for auditability." },
  { field: "Issues", note: "One row per numbered-issue article. Issue-level counts are repeated on each row, and author slots expose Author 1, Author 2, and onward without requiring users to split a combined byline cell." },
  { field: "Shorts", note: "One row per short-form article across Assignments, Blueprints, Reviews, Teaching Fails, and Tool Tips. Section totals remain in these notes and project count columns." },
  { field: "Author Split Caveat", note: "Automated splitting handles common comma, 'and', suffix, heading, text-prefix, and 'in conversation with' patterns. Group or team author names are preserved as author records when the source presents them as authors." },
  { field: "QA Notes", note: "Rows with missing author bios, admin/guideline texts, or deleted/404 Manifold slugs are surfaced in Record Notes columns rather than crowding the workbook overview." },
];
writeTable(sheets.field, [["field", "Field"], ["note", "Note"]], fieldRows, "FieldNotesTable");
sheets.field.freezePanes.freezeRows(1);
setWidths(sheets.field, [230, 900]);
sheets.field.getRangeByIndexes(1, 0, fieldRows.length, 2).format.rowHeightPx = 58;

await fs.mkdir(outputDir, { recursive: true });

for (const [sheetName, range] of [
  ["Author Metadata", "A1:R18"],
  ["Author-Pub Records", "A1:Q18"],
  ["Issues", `A1:${colLetter(Math.min(issueColumns.length - 1, 25))}28`],
  ["Shorts", `A1:${colLetter(Math.min(shortColumns.length - 1, 25))}28`],
  ["Field Notes", `A1:B${fieldRows.length + 1}`],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(
    path.join(outputDir, `${sheetName.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "")}_preview.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
