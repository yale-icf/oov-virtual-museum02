"""
Batch OCR + translation for Origins of Value museum documents.

Reads each document's original JPEG (full-res, from JPEG Files directory),
resizes to 1500px max on the long edge, sends to Claude's vision API to extract
visible text, translates non-English text, and saves results into:
  - scripts/ocr/<id>.txt  (individual text files, one per image)
  - museum-data.json      (transcription field on each item)

Resumable: saves progress to scripts/ocr_checkpoint.json so it can be
restarted without re-processing completed documents.

Usage:
    python scripts/ocr_documents.py [--limit N]

Requires ANTHROPIC_API_KEY environment variable.
"""

import argparse
import base64
import io
import json
import os
import sys
import time

import anthropic
from PIL import Image

SITE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(SITE_DIR, "data", "museum-data.json")
CHECKPOINT_PATH = os.path.join(SITE_DIR, "scripts", "ocr_checkpoint.json")
OCR_DIR = os.path.join(SITE_DIR, "scripts", "ocr")
JPEG_DIR = r"C:\Users\ks2479\Documents\my-project\origins-of-value\JPEG Files"

MODEL = "claude-haiku-4-5-20251001"
MAX_RETRIES = 3
RETRY_DELAY = 5   # seconds base delay
MAX_IMAGE_SIZE = 1500


# ---------------------------------------------------------------------------
# Image index
# ---------------------------------------------------------------------------

def build_image_index(jpeg_dir):
    """Walk all subdirs of jpeg_dir; return {lowercase_stem: full_path}."""
    index = {}
    for root, _dirs, files in os.walk(jpeg_dir):
        for fname in files:
            if fname.lower().endswith(".jpg"):
                stem = os.path.splitext(fname)[0].lower()
                index[stem] = os.path.join(root, fname)
    print(f"Found {len(index)} JPEG files in JPEG Files directory.")
    return index


def resize_image_bytes(image_path, max_size=MAX_IMAGE_SIZE):
    """Open image, resize so long edge ≤ max_size, return JPEG bytes."""
    img = Image.open(image_path)
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    if max(w, h) > max_size:
        scale = max_size / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Checkpoint
# ---------------------------------------------------------------------------

def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_checkpoint(checkpoint):
    with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
        json.dump(checkpoint, f, ensure_ascii=False)


# ---------------------------------------------------------------------------
# OCR
# ---------------------------------------------------------------------------

def extract_text(client, image_bytes, title, description, location):
    """Send resized image bytes to Claude; return transcribed/translated text."""
    image_data = base64.standard_b64encode(image_bytes).decode("utf-8")

    context_parts = []
    if title:
        context_parts.append(f"Title: {title}")
    if description:
        context_parts.append(f"Description: {description}")
    if location:
        context_parts.append(f"Country: {', '.join(location)}")
    context = "\n".join(context_parts)

    prompt = f"""This is a scan of a historical financial document (bond, stock certificate, banknote, etc.).

Document context:
{context}

Please:
1. Read and transcribe ALL visible text in the image as accurately as possible.
2. If the text is in a language other than English, provide the original text first, then an English translation.
3. Note any visible signatures, stamps, serial numbers, or other markings.
4. If parts of the text are illegible or unclear, indicate this with [illegible] or [unclear].

Format your response as:
- If the document is in English: just provide the transcribed text.
- If in another language: provide the original text under "Original:" and the translation under "English translation:".
- End with any notable visual elements (stamps, seals, signatures) under "Notable markings:" if present.

Be concise but thorough. Focus on the actual text content."""

    for attempt in range(MAX_RETRIES):
        try:
            message = client.messages.create(
                model=MODEL,
                max_tokens=2048,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/jpeg",
                                    "data": image_data,
                                },
                            },
                            {"type": "text", "text": prompt},
                        ],
                    }
                ],
            )
            return message.content[0].text
        except anthropic.AuthenticationError:
            # Invalid API key — no point retrying ever
            raise
        except anthropic.BadRequestError as e:
            if "credit balance is too low" in str(e):
                print(f"\nError: Insufficient credits. Add credits at console.anthropic.com")
                raise
            if attempt < MAX_RETRIES - 1:
                wait = RETRY_DELAY * (2 ** attempt)
                print(f"    Bad request: {e}. Retrying in {wait}s...")
                time.sleep(wait)
            else:
                print(f"    Failed after {MAX_RETRIES} attempts: {e}")
                return None
        except anthropic.RateLimitError:
            wait = RETRY_DELAY * (2 ** attempt)
            print(f"    Rate limited, waiting {wait}s...")
            time.sleep(wait)
        except anthropic.APIError as e:
            if attempt < MAX_RETRIES - 1:
                wait = RETRY_DELAY * (2 ** attempt)
                print(f"    API error: {e}. Retrying in {wait}s...")
                time.sleep(wait)
            else:
                print(f"    Failed after {MAX_RETRIES} attempts: {e}")
                return None

    return None


# ---------------------------------------------------------------------------
# Work-item expansion
# ---------------------------------------------------------------------------

def collect_work_items(items):
    """
    Expand museum-data.json items into a flat list of (doc_id, meta) tuples.

    Items WITHOUT pages → one entry for the item itself.
    Items WITH pages   → one entry per sub-page (the container ID has no JPEG).
    meta contains title/description/location for OCR context.
    """
    work = []
    seen = set()

    for item in items:
        parent_meta = {
            "title": item.get("title", ""),
            "description": item.get("description", ""),
            "location": item.get("location", []),
        }
        pages = item.get("pages", [])

        if not pages:
            doc_id = item["id"]
            if doc_id not in seen:
                work.append((doc_id, parent_meta))
                seen.add(doc_id)
        else:
            for page in pages:
                page_id = page["id"]
                if page_id not in seen:
                    page_meta = {
                        "title": parent_meta["title"],
                        "description": page.get("description", parent_meta["description"]),
                        "location": parent_meta["location"],
                    }
                    work.append((page_id, page_meta))
                    seen.add(page_id)

    return work


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="OCR museum documents using Claude vision API (full-res originals)"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Maximum number of new documents to process (0 = all)",
    )
    args = parser.parse_args()

    # Fix Windows console encoding so Unicode titles print cleanly
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: ANTHROPIC_API_KEY environment variable not set.")
        print('Set it with: $env:ANTHROPIC_API_KEY = "sk-ant-..."')
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    # Build image index from full-res originals
    if not os.path.isdir(JPEG_DIR):
        print(f"Error: JPEG directory not found: {JPEG_DIR}")
        sys.exit(1)
    image_index = build_image_index(JPEG_DIR)

    # Ensure OCR output directory exists
    os.makedirs(OCR_DIR, exist_ok=True)

    # Load museum data
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        items = json.load(f)

    # Expand to flat work list (one entry per image)
    work_items = collect_work_items(items)
    total = len(work_items)
    print(f"Total images to OCR: {total}")

    # Load checkpoint
    checkpoint = load_checkpoint()
    already_done = len(checkpoint)
    print(f"Already processed: {already_done} (from checkpoint)")

    # Write .txt files for any checkpoint entries that don't have them yet
    migrated = 0
    for doc_id, text in checkpoint.items():
        txt_path = os.path.join(OCR_DIR, f"{doc_id}.txt")
        if text and not os.path.exists(txt_path):
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text)
            migrated += 1
    if migrated:
        print(f"Wrote {migrated} .txt files from existing checkpoint.")

    skipped = 0
    processed = 0
    errors = 0
    not_found = 0
    limit = args.limit if args.limit > 0 else float("inf")

    for i, (doc_id, meta) in enumerate(work_items):
        if processed >= limit:
            print(f"\nReached limit of {args.limit} new documents.")
            break

        # Skip if already processed
        if doc_id in checkpoint:
            skipped += 1
            continue

        # Locate image
        image_path = image_index.get(doc_id.lower())
        if not image_path:
            print(f"[{i+1}/{total}] {doc_id}: not found in JPEG directory, skipping")
            checkpoint[doc_id] = ""
            save_checkpoint(checkpoint)
            not_found += 1
            continue

        print(f"[{i+1}/{total}] Processing {doc_id}: {meta['title'][:70]}...")

        # Resize image in memory
        try:
            image_bytes = resize_image_bytes(image_path)
            w_info = f"{Image.open(image_path).size[0]}px → resized"
        except Exception as e:
            print(f"    Failed to load/resize image: {e}")
            checkpoint[doc_id] = ""
            save_checkpoint(checkpoint)
            errors += 1
            continue

        result = extract_text(
            client,
            image_bytes,
            meta["title"],
            meta["description"],
            meta["location"],
        )

        if result is not None:
            checkpoint[doc_id] = result
            # Write individual .txt file
            txt_path = os.path.join(OCR_DIR, f"{doc_id}.txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(result)
            processed += 1
            print(f"    Done ({len(result)} chars)")
        else:
            checkpoint[doc_id] = ""
            errors += 1
            print(f"    Error — saved empty transcription")

        save_checkpoint(checkpoint)
        time.sleep(0.5)  # gentle rate limiting

    # Write transcriptions back into museum-data.json
    # Multi-page items: concatenate page texts; single-page: direct lookup
    print(f"\nWriting transcriptions to {DATA_PATH}...")
    for item in items:
        pages = item.get("pages", [])
        if pages:
            parts = []
            for page in pages:
                text = checkpoint.get(page["id"], "")
                if text:
                    parts.append(text)
            item["transcription"] = "\n\n---\n\n".join(parts)
        else:
            item["transcription"] = checkpoint.get(item["id"], "")

    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(items, f, indent=2, ensure_ascii=False)

    print(f"\nComplete!")
    print(f"  Processed (new): {processed}")
    print(f"  Skipped (checkpoint): {skipped}")
    print(f"  Not found in JPEG dir: {not_found}")
    print(f"  Errors: {errors}")
    print(f"  Total images: {total}")


if __name__ == "__main__":
    main()
