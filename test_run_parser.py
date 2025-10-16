import fitz  # PyMuPDF
import re
import json
import csv
import os

def load_regions_from_json(json_path='selected_regions.json'):
    """
    Load the selected regions from the JSON file saved by the GUI.
    Returns dict: field_name -> bbox (x0, y0, x1, y1)
    """
    if not os.path.exists(json_path):
        print(f"Error: {json_path} not found. Run the GUI selector first!")
        return None
    try:
        with open(json_path, 'r') as f:
            regions = json.load(f)
        print(f"Loaded regions: {list(regions.keys())}")
        print("BBoxes (for debug):")
        for field, bbox in regions.items():
            x0, y0, x1, y1 = bbox
            print(f"  {field}: ({x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}) -- Valid? y0 < y1: {y0 < y1}, in bounds: {0 <= y0 <= y1 <= 792}")
        return regions
    except Exception as e:
        print(f"Error loading JSON: {e}")
        return None

def extract_text_from_bbox(page, bbox):
    """
    Extract text from the bbox clip on the page.
    Returns cleaned text (strip, join lines).
    """
    if not bbox:
        return ""
    crop_box = fitz.Rect(bbox)
    # Default behavior: use PyMuPDF's get_text on the clip
    text = page.get_text(clip=crop_box).strip()
    cleaned = re.sub(r'\s+', ' ', text)
    return cleaned


def extract_text_from_bbox_strict(page, bbox):
    """
    Strict extraction: only include words whose center point lies inside the bbox.
    This avoids capturing words that only barely clip into the selection.
    Returns a space-joined string of the selected words (preserves word order by sorting by y,x).
    """
    if not bbox:
        return ""
    rect = fitz.Rect(bbox)
    # get words: list of tuples (x0, y0, x1, y1, word, block_no, line_no, word_no)
    words = page.get_text("words")
    # filter words whose center is inside rect
    selected = []
    for w in words:
        x0, y0, x1, y1, word_text = w[0], w[1], w[2], w[3], w[4]
        cx = (x0 + x1) / 2.0
        cy = (y0 + y1) / 2.0
        if rect.contains(fitz.Point(cx, cy)):
            # store sorting keys to preserve reading order
            selected.append((y0, x0, word_text))
    # sort by top-to-bottom, left-to-right
    selected.sort()
    out = " ".join([w[2] for w in selected])
    # collapse runs of identical consecutive words (e.g., 'WORD WORD' -> 'WORD')
    if out:
        toks = out.split()
        # collapse exact consecutive duplicates
        import itertools
        toks = [k for k,_ in itertools.groupby(toks)]
        # if the sequence is exactly two identical halves (entire name duplicated), collapse
        if len(toks) % 2 == 0 and len(toks) > 1:
            half = toks[:len(toks)//2]
            if half == toks[len(toks)//2:]:
                toks = half
        out = ' '.join(toks)
    cleaned = re.sub(r'\s+', ' ', out).strip()
    return cleaned

def quick_test_parse_pdf(pdf_path, regions_json='selected_regions.json', output_csv='extraction_results.csv', strict=False):
    """
    Test parser: Load bboxes from JSON and parse first 100 pages.
    Extracts text from each region's bbox on every page.
    Outputs to CSV: Page, Field1, Field2, ...
    """
    regions = load_regions_from_json(regions_json)
    if not regions:
        return []
    
    doc = fitz.open(pdf_path)
    results = []
    
    max_pages = min(doc.page_count, len(doc))
    field_names = list(regions.keys())
    
    for page_num in range(max_pages):
        page = doc[page_num]
        row = {'Page': page_num + 1}
        
        for field in field_names:
            bbox = regions.get(field)
            if strict:
                extracted = extract_text_from_bbox_strict(page, bbox)
            else:
                extracted = extract_text_from_bbox(page, bbox)
            row[field] = extracted.strip()
        
        results.append(row)
    
    doc.close()
    
    # Write to CSV
    if results:
        try:
            with open(output_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Page'] + field_names)
                writer.writeheader()
                writer.writerows(results)
            print(f"\nResults saved to {output_csv} ({len(results)} rows)")
        except PermissionError:
            import datetime
            fallback = f"extraction_results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(fallback, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Page'] + field_names)
                writer.writeheader()
                writer.writerows(results)
            print(f"\nPermission denied writing {output_csv}. Saved to {fallback} instead ({len(results)} rows)")
    
    return results

# Example usage: Replace with your PDF path
if __name__ == "__main__":
    pdf_file = r"C:/Users/philliph/Desktop/Tools/2025FremontNOV.pdf"  # Update this
    # Set strict=True to only include words whose centers lie inside the bbox
    test_results = quick_test_parse_pdf(pdf_file, strict=True)
    
    if test_results:
        print("\nSample rows:")
        for res in test_results[:3]:  # First 3
            print(res)
    else:
        print("No results (check JSON regions).")