# JITP Article Scraper

Runnable scraper and processor for Journal of Interactive Technology and Pedagogy article metadata on CUNY Manifold.

This project was compiled from the JITP editorial workspace so it can be rerun on the current system and refreshed from the existing scrape results.

## What It Does

The workflow has two stages:

1. `scrape_jitp_manifold.py` collects issue, project, text, article-page, byline, author bio, abstract, notes, references, and normalized text data from CUNY Manifold.
2. `build_jitp_metadata_workbook.mjs` turns the scraper JSON into a focused Excel workbook.

Current scrape scope:

- Numbered JITP issues `1` through `27`
- Short-form sections:
  - `Assignments`
  - `Blueprints`
  - `Reviews`
  - `Teaching Fails`
  - `Tool Tips`

## Repository Layout

```text
jitp-article-scraper/
├── build_jitp_metadata_workbook.mjs
├── data/
│   └── jitp_manifold_metadata.json
├── package.json
├── pyproject.toml
├── README.md
├── requirements.txt
├── scrape_jitp_manifold.py
└── scripts/
    └── update_current_system.sh
```

`data/jitp_manifold_metadata.json` is the current seed snapshot. The scraper can use it as fallback while refreshing the scrape.

Generated outputs are written under `outputs/` and ignored by git.

## Current-System Run

This machine already has the required Codex runtimes:

- Python: `/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3`
- Node: `/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node`
- Node modules: `/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules`

Run a full refresh and workbook rebuild:

```bash
bash scripts/update_current_system.sh
```

The script:

1. Checks the current Codex Python and Node runtimes.
2. Creates a local `node_modules` symlink to the current Codex dependency bundle if needed.
3. Runs the scraper.
4. Uses `data/jitp_manifold_metadata.json` as the existing-result fallback.
5. Rebuilds the workbook.
6. Writes outputs to `outputs/jitp_manifold_metadata/`.

Expected outputs:

```text
outputs/jitp_manifold_metadata/jitp_manifold_metadata.json
outputs/jitp_manifold_metadata/jitp_manifold_article_metadata.xlsx
```

## Manual Commands

Scrape only:

```bash
/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scrape_jitp_manifold.py \
  --output-dir outputs/jitp_manifold_metadata \
  --existing-json data/jitp_manifold_metadata.json
```

Build workbook only:

```bash
/Users/zacharymuhlbauer/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node build_jitp_metadata_workbook.mjs \
  --input-json outputs/jitp_manifold_metadata/jitp_manifold_metadata.json \
  --output-dir outputs/jitp_manifold_metadata
```

## Scraper Options

```text
--base-url              Default: https://cuny.manifoldapp.org
--journal-id            Default: JITP Manifold journal UUID
--issues-through        Default: 27
--output-dir            Directory for jitp_manifold_metadata.json
--existing-json         Existing JSON snapshot used as fallback
--no-existing-fallback  Disable fallback to existing JSON
--max-workers           Concurrent article fetch workers, default 8
```

The fallback behavior matters because some older Manifold pages are inconsistent. If a current scrape leaves fields blank or returns a failed/pending/missing status, the scraper can preserve known-good fields from the existing snapshot.

Fallback use is recorded per row in `existing_result_used`.

## Workbook Options

```text
--input-json    Path to jitp_manifold_metadata.json
--output-dir    Output directory
--output-xlsx   Optional full output path for the workbook
```

The workbook builder also accepts:

```text
JITP_INPUT_JSON
JITP_OUTPUT_DIR
JITP_OUTPUT_XLSX
```

## Workbook Sheets

The generated workbook contains five focused sheets:

- `Author Metadata`: one row per normalized author name.
- `Author-Pub Records`: one row per parsed author-publication relationship.
- `Issues`: one row per numbered-issue article.
- `Shorts`: one row per short-form article.
- `Field Notes`: compact audit notes, counts, gap-fill counts, and caveats.

The `Issues` and `Shorts` sheets include wide author slots:

```text
Author 1 Name
Author 1 Role
Author 1 Affiliation
Author 1 Bio
Author 2 Name
...
```

`Author-Pub Records` is the normalized relational table for author-publication analysis.

## Sorting Rules

Source JSON rows are sorted by:

1. Numbered issues first, short-form sections second
2. Numeric issue number when available
3. Short-form section name
4. Text category position
5. Text position
6. Article title

Workbook publication rows are sorted by:

1. Issue number or short-form section
2. Text position
3. Article title

## Author Processing

The builder splits combined byline cells into separate author records. It handles:

- comma-separated names
- `and`
- `&`
- suffixes such as `Jr.`, `Sr.`, `II`, `III`, `IV`
- `in conversation with`
- rendered byline headings
- article text prefixes when the raw author field is missing

Source provenance is preserved through:

- `Raw Author Cell`
- `Filled Author Cell`
- `Author Split Sources`

The builder prefers explicit source data:

1. Rendered byline lines
2. Raw API author description
3. Inferred heading/text prefix only if the raw author cell is missing

## Gap Filling

Dates:

- `Best Publication Date` uses text publication date first, then issue date, metadata original date, created date, and updated date.
- `Date Source` records the field used.

Authors:

- Missing raw author cells can be recovered from rendered headings or normalized text.
- Rows still missing authors are counted in `Field Notes`.

Bios:

- Article-level bios are split into author-specific bios where possible.
- Noisy one-author bio blocks are trimmed to begin at the author name when the source block includes extra text.

Abstracts:

- Scraped abstracts are used directly.
- If needed, an `Abstract` section can be recovered from normalized text.

QA:

- Missing bios, admin/guideline rows, deleted/404 slugs, and parse issues are surfaced in `Record Note` columns.

## Current Snapshot Totals

From the seed JSON/workbook used to compile this repo:

- Text records: `349`
- Numbered issue publications: `233`
- Short-form publications: `116`
- Short-form section counts: `Assignments: 61 | Blueprints: 8 | Reviews: 15 | Teaching Fails: 17 | Tool Tips: 15`

## Verification

After running `scripts/update_current_system.sh`, check:

- The scraper prints a JSON summary and output path.
- The workbook builder prints the workbook path.
- The workbook builder reports `Cell search matched 0 entries`.
- `Field Notes` contains expected counts.
- `Shorts` has one header row plus short-form publication rows.
- `Issues` has one header row plus numbered-issue publication rows.

## GitHub Push

Intended remote:

```text
git@github.com:zmuhls/jitp-article-scraper.git
```

If the GitHub CLI is authenticated, create and push with:

```bash
gh repo create zmuhls/jitp-article-scraper --private --source=. --remote=origin --push
```

Or, if the repo already exists:

```bash
git remote add origin git@github.com:zmuhls/jitp-article-scraper.git
git push -u origin main
```

The local `gh` session must be authenticated for the `zmuhls` account before those commands will work.
