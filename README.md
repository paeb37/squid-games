# Slide Retrieval & Draft Generation — MVP

**What it is:** A PowerPoint-centric tool that lets users upload decks, automatically parse each slide, generate a 1–2 sentence summary, and **search** across all slides. Results show a **clean (redacted)** preview by default, with gated access to originals. This foundation powers a later “stickies → draft slide” generation flow.

# MVP Scope

* **Upload & Parse:** Ingest `.pptx`, extract per-slide text, layout, notes, and thumbnails.
* **Auto-Summaries:** Create concise, 1–2 sentence abstracts per slide (client/PII stripped).
* **Search:** Hybrid semantic + keyword search over all slides; filter by uploader/date/tags.
* **Redaction by Default:** Show sanitized previews; support “request original” access.
* **Insert Back to Deck:** From search results, insert the selected slide/content into the current presentation.

# How It Works (High Level)

1. **Ingest:** Store the original deck in object storage; parse to normalized **Slide JSON** (text, layout, assets).
2. **Summarize & Index:** Generate a short summary and embedding for each slide; persist in the search index.
3. **Search & Preview:** Users query in natural language; we rank results and show redacted thumbnails + summaries.
4. **Governance:** Originals remain access-controlled; all actions are auditable.
