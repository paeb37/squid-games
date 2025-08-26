# PowerPoint Information Extractor

A Python tool to extract various types of information from PowerPoint (.pptx) files using the `python-pptx` library.

## Installation

1. Install the required dependency:
```bash
pip install -r requirements.txt
```

Or install directly:
```bash
pip install python-pptx
```

## Usage

### Command Line Usage

Extract all information (saves to JSON file):
```bash
python powerpoint_extractor.py presentation.pptx
```

Extract only text content:
```bash
python powerpoint_extractor.py presentation.pptx --text-only
```

Extract only slide titles:
```bash
python powerpoint_extractor.py presentation.pptx --titles-only
```

Specify output file:
```bash
python powerpoint_extractor.py presentation.pptx --output my_output.json
```

### Python Code Usage

```python
from powerpoint_extractor import PowerPointExtractor

# Initialize the extractor
extractor = PowerPointExtractor("path/to/your/presentation.pptx")

# Extract all text
text_data = extractor.extract_all_text()

# Extract slide titles
titles = extractor.extract_slide_titles()

# Extract image information
images = extractor.extract_images_info()

# Extract speaker notes
notes = extractor.extract_notes()

# Extract all information
all_info = extractor.extract_all_information()

# Save to JSON file
extractor.save_to_json("output.json")
```

## Features

The extractor can extract the following information from PowerPoint files:

### 1. Basic Information
- File path
- Total number of slides
- Slide dimensions (width and height)

### 2. Text Content
- All text from each slide
- Text organized by slide number
- Combined text from all slides

### 3. Slide Titles
- Title of each slide (attempts to find title placeholder first)
- Falls back to first text element if no title placeholder found

### 4. Image Information
- Details about images in each slide
- Image dimensions and positions
- Count of images per slide

### 5. Speaker Notes
- Notes associated with each slide
- Only includes slides that have notes

### 6. Layout Information
- Layout name for each slide
- Number of shapes per slide
- Information about each shape (name, type, whether it contains text)

## Output Format

When using the full extraction mode, the tool saves information in JSON format with the following structure:

```json
{
  "basic_info": {
    "file_path": "presentation.pptx",
    "total_slides": 10,
    "slide_dimensions": {
      "width": 9144000,
      "height": 6858000
    }
  },
  "titles": [
    {
      "slide_number": 1,
      "title": "Introduction"
    }
  ],
  "text_content": {
    "slides": [
      {
        "slide_number": 1,
        "text_elements": ["Title", "Bullet point 1", "Bullet point 2"],
        "combined_text": "Title Bullet point 1 Bullet point 2"
      }
    ],
    "all_text_combined": ["Title", "Bullet point 1", "Bullet point 2"]
  },
  "images_info": [...],
  "notes": [...],
  "layout_info": [...]
}
```

## Examples

### Extract Text Only
```python
extractor = PowerPointExtractor("presentation.pptx")
text_data = extractor.extract_all_text()

for slide in text_data["slides"]:
    print(f"Slide {slide['slide_number']}: {slide['combined_text']}")
```

### Find Slides with Images
```python
extractor = PowerPointExtractor("presentation.pptx")
images_info = extractor.extract_images_info()

for slide_info in images_info:
    print(f"Slide {slide_info['slide_number']} has {slide_info['image_count']} images")
```

### Extract Speaker Notes
```python
extractor = PowerPointExtractor("presentation.pptx")
notes = extractor.extract_notes()

for note in notes:
    print(f"Slide {note['slide_number']} notes: {note['notes']}")
```

## Error Handling

The tool includes error handling for:
- Invalid file paths
- Corrupted PowerPoint files
- Missing dependencies
- File access permissions

## Requirements

- Python 3.6+
- python-pptx library

## License

This project is open source and available under the MIT License.
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
