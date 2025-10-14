import fitz  # PyMuPDF
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import io
import json
import os

class PDFRegionSelector:
    def __init__(self, pdf_path, root):
        self.root = root
        self.root.title("PDF Region Selector - Fixed Rotation Mapping")
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.page = self.doc[0]  # First page (assumes consistent layout)
        self.rot = self.page.rotation
        self.flip_y = (self.rot % 360 == 180)
        self.root.title(self.root.title() + f" (Rotation: {self.rot}° - Y flipped: {self.flip_y})")
        self.page_image = None
        self.canvas = None
        self.scale = 1.0
        self.selected_regions = {}
        
        self.fields = ['Account Number', 'Property Address', 'Parcel ID', 'Legal Description']
        self.current_field = None
        self.start_x = self.start_y = 0
        self.rect = None
        self.existing_rects = {}  # Canvas IDs for existing regions
        
        self.setup_ui()
        self.load_regions_if_exists()  # Auto-load JSON if available
        self.load_page_image()
    
    def setup_ui(self):
        control_frame = ttk.Frame(self.root)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        
        ttk.Label(control_frame, text="Select Field to Fine-Tune:").pack(pady=5)
        self.field_var = tk.StringVar()
        field_combo = ttk.Combobox(control_frame, textvariable=self.field_var, values=self.fields)
        field_combo.pack(pady=5)
        field_combo.bind('<<ComboboxSelected>>', self.on_field_selected)
        
        ttk.Button(control_frame, text="Start/Re-Select", command=self.start_selection).pack(pady=5)
        ttk.Button(control_frame, text="Clear Selected", command=self.clear_selected).pack(pady=5)
        ttk.Button(control_frame, text="Save Updated Regions", command=self.save_regions).pack(pady=5)
        ttk.Button(control_frame, text="Load from JSON", command=self.load_regions_if_exists).pack(pady=5)
        
        self.status_label = ttk.Label(control_frame, text="Load JSON to fine-tune existing regions.")
        self.status_label.pack(pady=10)

        # Right panel: Canvas for PDF
        self.canvas_frame = ttk.Frame(self.root)
        self.canvas_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.canvas = tk.Canvas(self.canvas_frame, bg='white')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
    
    def load_regions_if_exists(self):
        json_path = 'selected_regions.json'
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r') as f:
                    self.selected_regions = json.load(f)
                self.status_label.config(text=f"Loaded {len(self.selected_regions)} regions from JSON.")
                print(f"Loaded regions: {list(self.selected_regions.keys())}")
                print("Current BBoxes:")
                for field, bbox in self.selected_regions.items():
                    print(f"  {field}: {bbox}")
            except Exception as e:
                messagebox.showerror("Load Error", f"Failed to load JSON: {e}")
        else:
            self.status_label.config(text="No JSON found—select fields to create new regions.")
    
    def load_page_image(self):
        # Render upright pixmap (derotate for display)
        mat = fitz.Matrix(self.scale, self.scale).prerotate(-self.rot)  # Derotate explicitly
        pix = self.page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        self.page_image = Image.open(io.BytesIO(img_data))
        self.photo = ImageTk.PhotoImage(self.page_image)
        
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        
        # Clear old rects
        for rect_id in self.existing_rects.values():
            self.canvas.delete(rect_id)
        self.existing_rects = {}
        
        # Redraw existing regions
        self.redraw_regions()
    
    def on_field_selected(self, event=None):
        self.current_field = self.field_var.get()
        if self.current_field:
            if self.current_field in self.selected_regions:
                messagebox.showinfo("Fine-Tune", f"Existing region for '{self.current_field}' loaded. Click/drag to adjust.")
            else:
                messagebox.showinfo("New", f"Click and drag to select region for '{self.current_field}'.")
    
    def start_selection(self):
        if not self.field_var.get():
            messagebox.showwarning("Select Field", "Please select a field first.")
            return
        self.current_field = self.field_var.get()
        self.canvas.config(cursor="cross")
        # Remove existing rect for this field if present
        if self.current_field in self.existing_rects:
            self.canvas.delete(self.existing_rects[self.current_field])
            del self.existing_rects[self.current_field]
    
    def on_click(self, event):
        if not self.current_field:
            return
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)
    
    def on_drag(self, event):
        if not self.current_field or not self.rect:
            return
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)
    
    def on_release(self, event):
        if not self.current_field or not self.rect:
            return
        end_x = event.x
        end_y = event.y
        self.canvas.config(cursor="")
        
        # Canvas coords (upright space: y=0 top)
        canvas_x0 = min(self.start_x, end_x)
        canvas_y0_canvas = min(self.start_y, end_y)  # Top in canvas
        canvas_x1 = max(self.start_x, end_x)
        canvas_y1_canvas = max(self.start_y, end_y)  # Bottom in canvas
        
        # Scale to points
        pt_x0 = canvas_x0 / self.scale
        pt_y0_canvas = canvas_y0_canvas / self.scale  # From top
        pt_x1 = canvas_x1 / self.scale
        pt_y1_canvas = canvas_y1_canvas / self.scale
        
        # Transform to PDF space using page.transformation_matrix (handles rotation)
        trans_mat = self.page.transformation_matrix
        # Transform top-left and bottom-right points
        pt_tl = fitz.Point(pt_x0, pt_y0_canvas)
        pt_br = fitz.Point(pt_x1, pt_y1_canvas)
        pdf_tl = pt_tl * trans_mat
        pdf_br = pt_br * trans_mat
        
        # Bbox in PDF space (min/max after transform)
        pdf_x0 = min(pdf_tl.x, pdf_br.x)
        pdf_y0 = min(pdf_tl.y, pdf_br.y)
        pdf_x1 = max(pdf_tl.x, pdf_br.x)
        pdf_y1 = max(pdf_tl.y, pdf_br.y)
        
        bbox = (pdf_x0, pdf_y0, pdf_x1, pdf_y1)
        self.selected_regions[self.current_field] = bbox
        
        # Draw permanent rect (green dashed)
        self.canvas.delete(self.rect)
        canvas_rect = self.canvas.create_rectangle(canvas_x0, canvas_y0_canvas, canvas_x1, canvas_y1_canvas, outline='green', width=2, dash=(5, 5))
        self.existing_rects[self.current_field] = canvas_rect
        self.canvas.tag_raise(canvas_rect)
        
        messagebox.showinfo("Updated", f"Updated region for '{self.current_field}': {bbox}")
        self.current_field = None
        self.rect = None
    
    def clear_selected(self):
        if self.current_field and self.current_field in self.existing_rects:
            self.canvas.delete(self.existing_rects[self.current_field])
            del self.existing_rects[self.current_field]
            if self.current_field in self.selected_regions:
                del self.selected_regions[self.current_field]
            messagebox.showinfo("Cleared", f"Cleared '{self.current_field}'.")
            self.current_field = None
    
    def redraw_regions(self):
        trans_mat_inv = self.page.transformation_matrix.invert()  # Fixed: invert() not inverted()
        page_height = self.page.rect.height
        for field, bbox in self.selected_regions.items():
            x0, y0, x1, y1 = bbox
            # Transform bbox corners back to canvas space
            pt_bl = fitz.Point(x0, y0) * trans_mat_inv
            pt_tr = fitz.Point(x1, y1) * trans_mat_inv
            canvas_x0 = min(pt_bl.x, pt_tr.x) * self.scale
            canvas_y0_canvas = min(pt_bl.y, pt_tr.y) * self.scale  # Top
            canvas_x1 = max(pt_bl.x, pt_tr.x) * self.scale
            canvas_y1_canvas = max(pt_bl.y, pt_tr.y) * self.scale  # Bottom
            rect_id = self.canvas.create_rectangle(canvas_x0, canvas_y0_canvas, canvas_x1, canvas_y1_canvas, outline='blue', width=1, dash=(10, 5))
            self.existing_rects[field] = rect_id
            self.canvas.create_text((canvas_x0 + canvas_x1)/2, canvas_y0_canvas - 10, text=field, fill='blue')
    
    def save_regions(self):
        if not self.selected_regions:
            messagebox.showwarning("No Regions", "No regions to save.")
            return
        with open('selected_regions.json', 'w') as f:
            json.dump(self.selected_regions, f, indent=2)
        messagebox.showinfo("Saved", f"Updated regions saved to selected_regions.json:\n{json.dumps(self.selected_regions, indent=2)}")
    
    def close(self):
        self.doc.close()
        self.root.quit()

# Example usage: Replace with your PDF path
if __name__ == "__main__":
    pdf_file = r"C:/Users/philliph/Desktop/Tools/2025FremontNOV.pdf"  # Update this
    root = tk.Tk()
    app = PDFRegionSelector(pdf_file, root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()