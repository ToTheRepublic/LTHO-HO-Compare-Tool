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
    text = page.get_text(clip=crop_box).strip()
    cleaned = re.sub(r'\s+', ' ', text)
    return cleaned

def quick_test_parse_pdf(pdf_path, regions_json='selected_regions.json', output_csv='extraction_results.csv'):
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
    
    max_pages = min(100, len(doc))
    field_names = list(regions.keys())
    
    for page_num in range(max_pages):
        page = doc[page_num]
        row = {'Page': page_num + 1}
        
        for field in field_names:
            bbox = regions.get(field)
            extracted = extract_text_from_bbox(page, bbox)
            row[field] = extracted.strip()
        
        results.append(row)
        print(f"Processed Page {page_num + 1}: { {k: v[:50] + '...' if len(v) > 50 else v for k, v in row.items() if k != 'Page'} }")
    
    doc.close()
    
    # Write to CSV
    if results:
        with open(output_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['Page'] + field_names)
            writer.writeheader()
            writer.writerows(results)
        print(f"\nResults saved to {output_csv} ({len(results)} rows)")
    
    return results

# Example usage: Replace with your PDF path
if __name__ == "__main__":
    pdf_file = r"C:/Users/philliph/Desktop/Tools/2025FremontNOV.pdf"  # Update this
    test_results = quick_test_parse_pdf(pdf_file)
    
    if test_results:
        print("\nSample rows:")
        for res in test_results[:3]:  # First 3
            print(res)
    else:
        print("No results (check JSON regions).")