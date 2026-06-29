import argparse
import csv
import json
import math
import os
import re
from datetime import datetime
from pathlib import Path


FORMULA_CHARS = set("=+-*/^_()[]{}<>×÷·∙√πρτηΩωμαβγθλΣΔPWVFstghmNkgJ")


def parse_bool(value):
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def parse_float(value, default=0.0):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return default


def load_font(size):
    from PIL import ImageFont

    for path in (
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                pass
    return ImageFont.load_default()


def normalize_result(raw):
    if raw is None:
        return [], 0.0
    elapsed = 0.0
    result = raw
    if isinstance(raw, tuple):
        result = raw[0]
        if len(raw) > 1:
            try:
                elapsed = float(raw[1])
            except (TypeError, ValueError):
                elapsed = 0.0
    if result is None:
        return [], elapsed

    items = []
    for item in result:
        if isinstance(item, dict):
            text = str(item.get("text", "")).strip()
            score = parse_float(item.get("score", item.get("confidence", 0.0)))
            box = item.get("box", item.get("points", []))
        elif isinstance(item, (list, tuple)) and len(item) >= 3:
            box = item[0]
            text = str(item[1]).strip()
            score = parse_float(item[2])
        else:
            continue
        if text:
            items.append({"text": text, "score": score, "box": box})
    return items, elapsed


def preprocess_image(input_path, output_path):
    from PIL import Image, ImageFilter, ImageOps

    image = Image.open(input_path)
    try:
        image = ImageOps.exif_transpose(image).convert("L")
        max_side = max(image.size)
        if max_side > 1600:
            scale = 1600.0 / max_side
            image = image.resize(
                (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
                Image.Resampling.LANCZOS,
            )
        elif max_side < 1200:
            scale = min(2.5, 1200.0 / max(1, max_side))
            image = image.resize(
                (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
                Image.Resampling.LANCZOS,
            )
        image = ImageOps.autocontrast(image)
        image = image.filter(ImageFilter.SHARPEN)
        image.save(output_path)
    finally:
        image.close()


def formula_signal(text):
    if not text:
        return 0
    hits = sum(1 for ch in text if ch in FORMULA_CHARS)
    latin = len(re.findall(r"[A-Za-zΑ-ω]", text))
    digits = len(re.findall(r"\d", text))
    compact = len(re.sub(r"\s+", "", text))
    if compact == 0:
        return 0
    score = round(100 * (hits + min(latin, 8) * 0.35 + min(digits, 8) * 0.25) / compact)
    return max(0, min(100, score))


def has_explicit_math_operator(text):
    if not text:
        return False
    return bool(re.search(r"[=<>±×÷*/^√_]", text))


def choose_best(original_items, processed_items):
    def quality(items):
        if not items:
            return (-1.0, 0, 0)
        text = " ".join(item["text"] for item in items)
        avg = sum(item["score"] for item in items) / max(1, len(items))
        return (avg, formula_signal(text), len(text))

    original_quality = quality(original_items)
    processed_quality = quality(processed_items)
    if processed_quality > original_quality:
        return "preprocessed", processed_items
    return "original", original_items


def write_contact_sheet(rows, output_path, max_items=80):
    from PIL import Image, ImageDraw

    items = [row for row in rows if row.get("Status") == "Success" and row.get("ImagePath")]
    items = items[:max_items]
    if not items:
        return ""

    tile_w, tile_h = 360, 260
    image_h = 150
    cols = 2 if len(items) <= 10 else 4
    rows_count = math.ceil(len(items) / cols)
    sheet = Image.new("RGB", (cols * tile_w, rows_count * tile_h), "white")
    draw = ImageDraw.Draw(sheet)
    font = load_font(14)
    small = load_font(11)

    for index, row in enumerate(items):
        col = index % cols
        grid_row = index // cols
        x = col * tile_w
        y = grid_row * tile_h
        draw.rectangle([x + 6, y + 6, x + tile_w - 6, y + tile_h - 6], outline=(210, 210, 210), fill=(248, 248, 248))
        try:
            image = Image.open(row["ImagePath"]).convert("RGB")
            image.thumbnail((tile_w - 24, image_h), Image.Resampling.LANCZOS)
            px = x + (tile_w - image.width) // 2
            py = y + 14 + (image_h - image.height) // 2
            sheet.paste(image, (px, py))
            image.close()
        except OSError:
            pass

        text = row.get("BestText", "")
        if len(text) > 54:
            text = text[:51] + "..."
        draw.text((x + 12, y + image_h + 20), f"{row.get('FormulaCandidateLevel')} score={row.get('FormulaImageScore')}", fill=(0, 0, 0), font=font)
        draw.text((x + 12, y + image_h + 42), f"OCR {row.get('MeanConfidence')} signal={row.get('FormulaTextSignal')}", fill=(80, 80, 80), font=small)
        draw.text((x + 12, y + image_h + 62), text, fill=(30, 30, 30), font=small)
        media = row.get("MediaPath", "")
        if len(media) > 58:
            media = media[:55] + "..."
        draw.text((x + 12, y + image_h + 82), media, fill=(80, 80, 80), font=small)

    sheet.save(output_path)
    return str(output_path)


def main():
    parser = argparse.ArgumentParser(description="Run RapidOCR over formula image candidates.")
    parser.add_argument("--input-csv", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--max-images", type=int, default=30)
    parser.add_argument("--min-score", type=float, default=45.0)
    parser.add_argument("--max-pixels", type=int, default=1500000)
    parser.add_argument("--ocr-original", action="store_true")
    parser.add_argument("--no-contact-sheet", action="store_true")
    args = parser.parse_args()

    from rapidocr_onnxruntime import RapidOCR

    input_csv = Path(args.input_csv).resolve()
    output_dir = Path(args.output_dir).resolve()
    preview_dir = output_dir / "preprocessed"
    output_dir.mkdir(parents=True, exist_ok=True)
    preview_dir.mkdir(parents=True, exist_ok=True)

    with input_csv.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))

    candidates = []
    for row in rows:
        score = parse_float(row.get("FormulaImageScore"))
        pixels = parse_float(row.get("Pixels"))
        image_path = row.get("FormulaExtractedPath") or row.get("ExtractedPath") or ""
        if not parse_bool(row.get("FormulaImageCandidate")):
            continue
        if score < args.min_score:
            continue
        if pixels > args.max_pixels:
            continue
        if not image_path or not Path(image_path).exists():
            continue
        row["_score"] = score
        row["_image_path"] = image_path
        candidates.append(row)
    candidates.sort(key=lambda item: item.get("_score", 0.0), reverse=True)
    selected = candidates[: max(1, args.max_images)]

    engine = RapidOCR()
    output_rows = []
    for index, row in enumerate(selected, start=1):
        image_path = Path(row["_image_path"])
        processed_path = preview_dir / f"ocr-{index:03d}.png"
        status = "Success"
        message = ""
        original_items = []
        processed_items = []
        original_elapsed = 0.0
        processed_elapsed = 0.0
        try:
            preprocess_image(image_path, processed_path)
            processed_items, processed_elapsed = normalize_result(engine(str(processed_path)))
            if args.ocr_original:
                original_items, original_elapsed = normalize_result(engine(str(image_path)))
        except Exception as exc:
            status = "Failed"
            message = str(exc)

        source, best_items = choose_best(original_items, processed_items)
        best_text = " ".join(item["text"] for item in best_items)
        mean_confidence = 0.0
        if best_items:
            mean_confidence = sum(item["score"] for item in best_items) / len(best_items)
        signal = formula_signal(best_text)
        if best_text and (has_explicit_math_operator(best_text) or signal >= 50):
            action = "ReviewOcrFormulaText"
        elif best_text and signal >= 12:
            action = "ReviewOcrPhysicsLabels"
        else:
            action = "ManualVisualReview"
        if status != "Success":
            action = "OcrFailedManualReview"

        output_rows.append(
            {
                "Deck": row.get("Deck", ""),
                "DeckPath": row.get("DeckPath", ""),
                "MediaPath": row.get("MediaPath", ""),
                "UsedOnSlides": row.get("UsedOnSlides", ""),
                "FormulaImageScore": row.get("FormulaImageScore", ""),
                "FormulaCandidateLevel": row.get("FormulaCandidateLevel", ""),
                "FormulaCandidateReason": row.get("FormulaCandidateReason", ""),
                "ImagePath": str(image_path),
                "PreprocessedPath": str(processed_path) if processed_path.exists() else "",
                "Status": status,
                "Message": message,
                "BestSource": source,
                "BestText": best_text,
                "MeanConfidence": f"{mean_confidence:.4f}",
                "FormulaTextSignal": str(signal),
                "OcrLineCount": str(len(best_items)),
                "OriginalOcrText": " ".join(item["text"] for item in original_items),
                "PreprocessedOcrText": " ".join(item["text"] for item in processed_items),
                "OriginalElapsed": f"{original_elapsed:.4f}",
                "PreprocessedElapsed": f"{processed_elapsed:.4f}",
                "SuggestedAction": action,
            }
        )

    csv_path = output_dir / "formula-image-ocr.csv"
    json_path = output_dir / "formula-image-ocr.json"
    manifest_path = output_dir / "formula-image-ocr-manifest.json"
    sheet_path = output_dir / "formula-image-ocr.contact-sheet.png"

    fields = [
        "Deck",
        "DeckPath",
        "MediaPath",
        "UsedOnSlides",
        "FormulaImageScore",
        "FormulaCandidateLevel",
        "FormulaCandidateReason",
        "ImagePath",
        "PreprocessedPath",
        "Status",
        "Message",
        "BestSource",
        "BestText",
        "MeanConfidence",
        "FormulaTextSignal",
        "OcrLineCount",
        "OriginalOcrText",
        "PreprocessedOcrText",
        "OriginalElapsed",
        "PreprocessedElapsed",
        "SuggestedAction",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(output_rows)

    with json_path.open("w", encoding="utf-8", newline="") as handle:
        json.dump(output_rows, handle, ensure_ascii=False, indent=2)

    contact_sheet = ""
    if not args.no_contact_sheet:
        contact_sheet = write_contact_sheet(output_rows, sheet_path)

    success_count = sum(1 for row in output_rows if row["Status"] == "Success")
    formula_text_count = sum(1 for row in output_rows if row["SuggestedAction"] == "ReviewOcrFormulaText")
    physics_label_count = sum(1 for row in output_rows if row["SuggestedAction"] == "ReviewOcrPhysicsLabels")
    manifest = {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "inputCsv": str(input_csv),
        "outputDir": str(output_dir),
        "maxPixels": args.max_pixels,
        "ocrOriginal": args.ocr_original,
        "candidateCount": len(candidates),
        "selectedCount": len(selected),
        "successCount": success_count,
        "reviewFormulaTextCount": formula_text_count,
        "reviewPhysicsLabelCount": physics_label_count,
        "csv": str(csv_path),
        "json": str(json_path),
        "contactSheet": contact_sheet,
        "note": "RapidOCR is a text OCR probe for review only; it is not a math OCR replacement writer.",
    }
    with manifest_path.open("w", encoding="utf-8", newline="") as handle:
        json.dump(manifest, handle, ensure_ascii=False, indent=2)

    print(f"Formula image OCR probe done: {output_dir}")
    print(f"Success: {success_count} / Selected: {len(selected)}")
    print(f"Review formula text: {formula_text_count}")
    print(f"Review physics labels: {physics_label_count}")


if __name__ == "__main__":
    main()
