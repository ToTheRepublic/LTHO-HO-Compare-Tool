import fitz  # PyMuPDF

pdf_file = r"C:/Users/philliph/Desktop/Tools/2025FremontNOV.pdf"  # Update this
doc = fitz.open(pdf_file)
page = doc[0]  # Page 2 (0-indexed)

# Search for known text instances (adjust if OCR varies)
searches = {
    'Account Number': 'R0040717',
    'Property Address': 'MAZET RD',
    'Parcel ID': '91253310013700',
    'Legal Description': 'TWP 2N RNG 5E SEC 33: PARCEL IN LOTS 3 & 4 WD 2024-1460137'  # Partial match for the line
}

print("Real bboxes for Page 1 (x0,y0,x1,y1):")
for field, term in searches.items():
    instances = page.search_for(term)
    if instances:
        bbox = instances[0]  # Take first match
        print(f"  {field}: {bbox}")
    else:
        print(f"  {field}: Not found (try partial term)")

doc.close()