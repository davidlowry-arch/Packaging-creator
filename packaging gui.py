import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pymupdf  
from PIL import Image, ImageTk
import os
import tempfile
import pandas as pd
import importlib.util
import subprocess
import shutil
import openpyxl

spec = importlib.util.spec_from_file_location("pc", "packaging creator.py")
pc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pc)

def get_col_letter(col_idx):
    letter = ""
    while col_idx >= 0:
        letter = chr(col_idx % 26 + 65) + letter
        col_idx = col_idx // 26 - 1
    return letter

class PackagingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Packaging Creator")
        self.root.geometry("1400x800")
        
        self.excel_path = r"C:\Users\dclow\Desktop\packaging info.xlsx"
        self.cell_colors = {}
        self.row_tags = {}
        self.load_excel_data()
        
        self.log_path = os.path.join(pc.ASSETS_DIR, "input_log.csv")
        self.debounce_timer = None
        
        self.current_font_size = 10
        self.style = ttk.Style()
        self.hide_edit_warning = tk.BooleanVar(value=False)
        self.labels_per_page = tk.IntVar(value=9)
        self.img_v_offset = 0.0  # Normalized offset between -1.0 and 1.0
        self.img_zoom = 1.0      # Zoom scale factor
        
        self.pdf_jobs = []            # ADD THIS: Stores queued label jobs
        self.total_queued_labels = 0  # ADD THIS: Tracks the total number of labels in the queue
        
        self.setup_ui()

    @staticmethod
    def extract_cell_hex(cell):
        if not cell or not cell.fill:
            return None
        fill = cell.fill
        if fill.fill_type is None:
            return None
        color = fill.fgColor or fill.start_color
        if not color:
            return None
            
        if color.type == 'rgb' and color.rgb:
            rgb_str = str(color.rgb).upper()
            if len(rgb_str) == 8:
                return f"#{rgb_str[2:]}"
            elif len(rgb_str) == 6:
                return f"#{rgb_str}"
                
        if color.type == 'indexed' and color.indexed is not None:
            try:
                from openpyxl.styles.colors import COLOR_INDEX
                if 0 <= color.indexed < len(COLOR_INDEX):
                    idx_str = str(COLOR_INDEX[color.indexed]).upper()
                    if len(idx_str) == 8:
                        return f"#{idx_str[2:]}"
                    elif len(idx_str) == 6:
                        return f"#{idx_str}"
            except Exception:
                pass
                
        if color.type == 'theme' and color.theme is not None:
            theme_defaults = {7: "#FFC000", 9: "#70AD47"}
            return theme_defaults.get(color.theme)
            
        return None

    @staticmethod
    def classify_color(hex_code):
        if not hex_code or not isinstance(hex_code, str) or not hex_code.startswith('#'):
            return None
        hex_clean = hex_code.lstrip('#')
        if len(hex_clean) != 6:
            return None
        try:
            r = int(hex_clean[0:2], 16)
            g = int(hex_clean[2:4], 16)
            b = int(hex_clean[4:6], 16)
        except ValueError:
            return None
            
        if (r > 245 and g > 245 and b > 245) or (r < 30 and g < 30 and b < 30):
            return None
        if g > r and g > b:
            return 'green'
        if r > 140 and g > 120 and (r > b + 15) and (g > b + 15):
            return 'yellow'
            
        return 'other'

    def load_excel_data(self):
        self.df = pd.read_excel(self.excel_path, header=None)
        self.df.fillna("", inplace=True)
        self.cell_colors.clear()
        self.row_tags.clear()
        
        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            ws = wb.active
            for r_idx, row in enumerate(ws.iter_rows(values_only=False)):
                row_green = False
                row_yellow = False
                for c_idx, cell in enumerate(row):
                    hex_color = self.extract_cell_hex(cell)
                    if hex_color:
                        self.cell_colors[(r_idx, c_idx)] = hex_color
                        cat = self.classify_color(hex_color)
                        if cat == 'green':
                            row_green = True
                        elif cat == 'yellow':
                            row_yellow = True
                if row_green and row_yellow:
                    self.row_tags[r_idx] = 'mixed'
                elif row_green:
                    self.row_tags[r_idx] = 'green'
                elif row_yellow:
                    self.row_tags[r_idx] = 'yellow'
                else:
                    self.row_tags[r_idx] = 'normal'
            wb.close()
        except Exception as e:
            print(f"Warning: Could not read cell colors from Excel: {e}")

    def setup_ui(self):
        left_frame = tk.Frame(self.root, padx=10, pady=5)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        right_frame = tk.Frame(self.root, padx=10, pady=5, bg="#e0e0e0")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 1 & 2. Top Buttons (Instructions & Log)
        top_btn_frame = tk.Frame(left_frame)
        top_btn_frame.pack(pady=(5, 12), fill=tk.X)
        
        self.btn_instructions = tk.Button(top_btn_frame, text="View Instructions", command=self.show_instructions)
        self.btn_instructions.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        
        self.btn_view_log = tk.Button(top_btn_frame, text="View Log", command=self.show_log)
        self.btn_view_log.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(2, 0))

        # 3. Experiment Input
        tk.Label(left_frame, text="Type here to experiment (Press Enter to log):").pack(anchor="w")
        self.entry_experiment = tk.Entry(left_frame, width=45)
        self.entry_experiment.pack(pady=(0, 8), anchor="w", fill=tk.X)
        self.entry_experiment.bind("<KeyRelease>", self.on_typing)
        self.entry_experiment.bind("<Return>", self.log_experiment) 

        # 4. Preview Window Header with Import Button
        preview_header_frame = tk.Frame(left_frame)
        preview_header_frame.pack(fill=tk.X, pady=(0, 4))
        
        tk.Label(preview_header_frame, text="Preview:").pack(side=tk.LEFT, anchor="w")
        self.btn_import = tk.Button(preview_header_frame, text="Import", command=self.import_image, padx=10)
        self.btn_import.pack(side=tk.RIGHT)

        # Display container defaults to 9-per-page (280x396 px scaled proportionally)
        self.preview_container = tk.Frame(left_frame, width=280, height=396, bg="gray")
        self.preview_container.pack_propagate(False) 
        self.preview_container.pack(pady=(0, 4), fill=tk.NONE, expand=False)

        self.lbl_preview = tk.Label(self.preview_container, text="Preview will appear here", bg="gray")
        self.lbl_preview.pack(fill=tk.BOTH, expand=True)

        self.btn_edit = tk.Button(left_frame, text="Edit in Inkscape", command=self.open_in_inkscape, state=tk.DISABLED)
        self.btn_edit.pack(pady=(0, 8), fill=tk.X)

        # 5. Create Packaging Input (Overhauled)
        create_frame = tk.Frame(left_frame)
        create_frame.pack(pady=(0, 8), anchor="w", fill=tk.X)

        tk.Label(create_frame, text="Add ").pack(side=tk.LEFT)
        
        # Limit to 1 or 2 digits by setting a small width
        self.entry_copies = tk.Entry(create_frame, width=3)
        self.entry_copies.pack(side=tk.LEFT)
        
        tk.Label(create_frame, text=" to pdf").pack(side=tk.LEFT)

        self.btn_add_to_pdf = tk.Button(create_frame, text="Add", command=self.add_to_pdf)
        self.btn_add_to_pdf.pack(side=tk.LEFT, padx=(8, 4))

        self.btn_generate_pdf = tk.Button(create_frame, text="Generate PDF", command=self.generate_queued_pdf)
        self.btn_generate_pdf.pack(side=tk.LEFT, padx=(4, 0))

        # Status label to show the user what is currently in their mental PDF representation
        self.lbl_queue_status = tk.Label(left_frame, text="Queued labels: 0", fg="blue")
        self.lbl_queue_status.pack(anchor="w", pady=(0, 8))

        # 6. Page Density Toggle (9 vs 12 per page)
        density_frame = tk.Frame(left_frame)
        density_frame.pack(fill=tk.X, pady=(0, 4))

        self.btn_9_per_page = tk.Button(
            density_frame, text="9 per page (Default)", relief=tk.SUNKEN, bg="#cce5ff",
            command=lambda: self.set_density(9)
        )
        self.btn_9_per_page.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.btn_12_per_page = tk.Button(
            density_frame, text="12 per page", relief=tk.RAISED, bg="#f0f0f0",
            command=lambda: self.set_density(12)
        )
        self.btn_12_per_page.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(4, 0))

        # 7. Manual Controls (Up, Down, Zoom In, Zoom Out, Reset)
        manual_controls_frame = tk.Frame(left_frame)
        manual_controls_frame.pack(fill=tk.X, pady=(0, 6))

        self.btn_nudge_up = tk.Button(manual_controls_frame, text="▲ Up", command=lambda: self.nudge_image(0.05), width=6)
        self.btn_nudge_up.pack(side=tk.LEFT, padx=1)

        self.btn_nudge_down = tk.Button(manual_controls_frame, text="▼ Down", command=lambda: self.nudge_image(-0.05), width=6)
        self.btn_nudge_down.pack(side=tk.LEFT, padx=1)

        self.btn_zoom_in_img = tk.Button(manual_controls_frame, text="+ Zoom", command=lambda: self.zoom_image(0.05), width=6)
        self.btn_zoom_in_img.pack(side=tk.LEFT, padx=1)

        self.btn_zoom_out_img = tk.Button(manual_controls_frame, text="- Zoom", command=lambda: self.zoom_image(-0.05), width=6)
        self.btn_zoom_out_img.pack(side=tk.LEFT, padx=1)

        self.btn_nudge_reset = tk.Button(manual_controls_frame, text="Reset", command=self.reset_transform, width=5)
        self.btn_nudge_reset.pack(side=tk.RIGHT, padx=1)

        # --- RIGHT PANEL UI (EXCEL DATA) ---
        top_right_bar = tk.Frame(right_frame, bg="#e0e0e0")
        top_right_bar.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(top_right_bar, text="Source Data Preview", bg="#e0e0e0", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        self.lbl_cell_info = tk.Label(top_right_bar, text="", bg="#e0e0e0", font=("Arial", 9, "italic"), fg="#333333")
        self.lbl_cell_info.pack(side=tk.LEFT, padx=12)

        btn_zoom_out = tk.Button(top_right_bar, text="Zoom Out (-)", command=self.zoom_out)
        btn_zoom_out.pack(side=tk.RIGHT, padx=2)
        btn_zoom_in = tk.Button(top_right_bar, text="Zoom In (+)", command=self.zoom_in)
        btn_zoom_in.pack(side=tk.RIGHT, padx=2)

        tree_frame = tk.Frame(right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        col_count = len(self.df.columns)
        excel_cols = [get_col_letter(i) for i in range(col_count)]
        tree_cols = ["Row"] + excel_cols
        
        self.style.theme_use("clam")
        self.style.configure("Treeview", background="#ffffff", foreground="#000000", fieldbackground="#ffffff")
        self.style.map("Treeview", background=[('selected', '#0078d7')], foreground=[('selected', '#ffffff')])
        
        self.tree = ttk.Treeview(tree_frame, columns=tree_cols, show="headings")
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree.heading("Row", text="")
        self.tree.column("Row", width=40, anchor="center", stretch=False)
        for col_name in excel_cols:
            self.tree.heading(col_name, text=col_name)
            self.tree.column(col_name, width=120, anchor="w")
            
        self.tree.tag_configure("green", background="#d4edda", foreground="#000000")
        self.tree.tag_configure("yellow", background="#fff3cd", foreground="#000000")
        self.tree.tag_configure("mixed", background="#d1ecf1", foreground="#000000")
        self.tree.tag_configure("normal", background="#ffffff", foreground="#000000")
            
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
            
        self.populate_tree()
        self.update_tree_font()

    def nudge_image(self, delta):
        self.img_v_offset = max(-1.0, min(1.0, self.img_v_offset + delta))
        self.update_preview()

    def zoom_image(self, delta):
        self.img_zoom = max(0.2, min(5.0, self.img_zoom + delta))
        self.update_preview()

    def reset_transform(self):
        self.img_v_offset = 0.0
        self.img_zoom = 1.0
        self.update_preview()

    def set_density(self, count):
        self.labels_per_page.set(count)
        if count == 9:
            self.btn_9_per_page.config(relief=tk.SUNKEN, bg="#cce5ff")
            self.btn_12_per_page.config(relief=tk.RAISED, bg="#f0f0f0")
            self.preview_container.config(height=396)  # True 1/9 A4 aspect ratio
        else:
            self.btn_9_per_page.config(relief=tk.RAISED, bg="#f0f0f0")
            self.btn_12_per_page.config(relief=tk.SUNKEN, bg="#cce5ff")
            self.preview_container.config(height=297)  # True 1/12 A4 aspect ratio
        self.update_preview()

    def import_image(self):
        file_path = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        file_path = os.path.normpath(file_path)
        self.reset_transform()
        self.entry_experiment.delete(0, tk.END)
        self.entry_experiment.insert(0, file_path)
        self.update_preview()

    def show_instructions(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("How to Use Packaging Creator")
        dlg.geometry("640x540")
        dlg.minsize(500, 400)
        dlg.transient(self.root)
        dlg.grab_set() 
        
        inst_frame = tk.Frame(dlg, padx=12, pady=12)
        inst_frame.pack(fill=tk.BOTH, expand=True)
        
        inst_text = tk.Text(inst_frame, wrap="word", font=("Segoe UI", 9), padx=8, pady=8, bg="#fafafa", relief=tk.SOLID, borderwidth=1)
        inst_scroll = tk.Scrollbar(inst_frame, command=inst_text.yview)
        inst_text.configure(yscrollcommand=inst_scroll.set)
        inst_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inst_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        instructions = (
            "WELCOME TO PACKAGING CREATOR\n"
            "--------------------------------------------------------------------------------\n"
            "This tool helps you quickly design, preview, and print 3x3 label grids on A4\n"
            "sheets using resources listed in your spreadsheet.\n\n"

            "1. BASIC COMMAND FORMAT\n"
            "Every command starts with the language name, followed by the spreadsheet\n"
            "row numbers you want to put on the label, separated by commas:\n\n"
            "    wolof, 18, 7\n\n"
            "• Language Title: The word before the first comma (e.g., 'wolof') is\n"
            "  printed across the top of the label in spaced-out capital letters.\n"
            "• Row Numbers: The numbers match the row numbers shown in the table\n"
            "  on the right side of the screen.\n"
            "• Print Quantity: Add a number in parentheses at the very end to specify\n"
            "  how many labels to print. If you leave this off, it defaults to 9\n"
            "  (one complete A4 sheet):\n"
            "      wolof, 18, 7 (18)      -> Prints 18 copies (2 full pages)\n\n"

            "2. CUSTOMIZING INDIVIDUAL ITEMS\n"
            "You can attach special symbols directly to any row number to adjust how\n"
            "that specific resource is displayed:\n\n"
            "• Custom Heading Fonts (add 'f's after the number):\n"
            "       21       -> Font 1: Rockwell-Bold (Default)\n"
            "       21f      -> Font 2: Amiri-Bold (Classic calligraphic serif)\n"
            "       21ff     -> Font 3: El Messiri-Bold (Sweeping reed-pen display)\n"
            "       21fff    -> Font 4: Cinzel Decorative-Bold (Architectural / ornate)\n"
            "       21ffff   -> Font 5: Noto Serif-Bold (Crisp modern serif)\n"
            " Changes the font used for that specific resource title only. Can be \n"
            " combined with other symbols (e.g., 21f- for font 2 without description).\n\n"
            "• Hide Description (add a single minus '-'):\n"
            "      18-\n"
            "  Keeps the resource title, image, and copyright notice, but drops\n"
            "  the descriptive text. Great when you're short on vertical space.\n\n"
            "• Hide Title & Description (add a double minus '--'):\n"
            "      121--\n"
            "  Shows only the resource picture and copyright credit. Perfect for\n"
            "  items like the JESUS Film logo where the title is already in the artwork.\n\n"
            "• Combine Rows Closely (link with '&'):\n"
            "      118&119\n"
            "  Groups multiple resources together as a single block: removes the\n"
            "  horizontal divider line between them, uses a slightly smaller title,\n"
            "  and places the copyright notice only at the end of the group.\n\n"
            "• Text-Only License (add '-cc'):\n"
            "      18-cc   or   18&-cc\n"
            "  Replaces the graphical Creative Commons logo badge with a neat,\n"
            "  compact line of text ('CC BY-NC-ND 4.0') to save layout space.\n\n"

            "3. WHOLE-LABEL ADJUSTMENTS (MODIFIERS)\n"
            "Add any of these extra codes as separate comma items anywhere in your entry:\n\n"
            "• Islamic Decorative Borders ('i', 'ii', 'iii', etc.):\n"
            "      wolof, 18, 7, i       -> Uses standard islamic_border.png\n"
            "      wolof, 18, 7, iii     -> Uses islamic_border3.jpg\n"
            "      wolof, 18, 7, iiiii   -> Uses islamic_border5.png\n"
            "  Applies the chosen Islamic border and automatically scales contents to 85%.\n\n"
            "• Manual Shrink ('s' + percentage):\n"
            "      s10   (shrinks by 10%)    |    s25   (shrinks by 25%)\n"
            "  Scales down all text, pictures, and spacing inside the label by that\n"
            "  percentage. Use this when a label is overflowing or looks too crowded.\n\n"
            "• Vertical Nudge ('l' + pixels):\n"
            "      l15   (moves down)        |    l-10  (moves up)\n"
            "  Slides the entire content block up or down to help you center\n"
            "  everything perfectly within the border.\n\n"
            "• License Badge Shift ('cc' + pixels):\n"
            "      cc10  (shifts right)      |    cc-10 (shifts left)\n"
            "  Nudges the bottom Creative Commons license icon left or right to\n"
            "  balance out the footer alignment.\n\n"

            "4. QUEUING MULTIPLE BATCHES AT ONCE\n"
            "You can queue different languages and quantities in a single print run\n"
            "by separating each job with a semicolon (;) :\n\n"
            "    wolof, 18-, 7, i, s10 (9); seereer, 217-, 213 (5)\n\n"

            "5. HELPFUL WORKFLOW TIPS\n"
            "• Experiment Field: As you type here, the preview box on the left will\n"
            "  automatically update within a moment. When you find a look you like,\n"
            "  press Enter in this box to save the line to your History Log.\n"
            "• Create Packaging Field: When your design is ready, type or paste the\n"
            "  command here and press Enter. It will create 'Final_Packaging.pdf'\n"
            "  in your project folder.\n"
            "• Editing Spreadsheet Data: Double-click any cell in the table on the\n"
            "  right to temporarily tweak text or titles for your current session.\n"
            "  These edits are safe and will never overwrite your Excel file on disk.\n"
            "• Table Row Colors: Rows tinted green or yellow indicate flagged categories\n"
            "  from your Excel sheet. Click any row to see which specific columns are colored.\n"
            "• Edit in Inkscape: Click this button under the preview to pop open the\n"
            "  current generated label in Inkscape for manual vector touch-ups.\n\n"

            "6. TO RESYNC WITH THE GOOGLE SHEET GREEN LIGHT LIST\n"
            "Easiest is to copy and paste the whole greenlight list into an excel file,\n"
            "delete columns B, E, F and G, then save over the existing packaging info file.\n"
            "You will need to then manually add any image filenames (esp. jesus_film.jpg)\n"
            "into column F.\n\n"

            "7. IMPORTING IMAGES: Click the 'Import' button above the preview window to browse\n"
            "  for a PNG or JPG file. Use the Up, Down, +Zoom, and -Zoom buttons to position\n"
            "  and scale your artwork manually. Toggling 9 vs 12 per page switches the label size\n"
            "  without cropping or re-scaling your image automatically."
        )
        
        inst_text.insert(tk.END, instructions)
        inst_text.see(1.0) 
        inst_text.config(state=tk.DISABLED) 
        
        btn_exit = tk.Button(dlg, text="Close Instructions", command=dlg.destroy, width=18, pady=3)
        btn_exit.pack(pady=(0, 10))

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        if column == '#1':
            return
            
        if not self.hide_edit_warning.get():
            self.show_edit_warning()

        x, y, width, height = self.tree.bbox(item_id, column)
        current_value = self.tree.set(item_id, column)
        
        entry = tk.Entry(self.tree, font=("Arial", self.current_font_size))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.focus_set()
        
        df_col_idx = int(column[1:]) - 2 
        entry.bind("<Return>", lambda e: self.save_edit(entry, item_id, column, df_col_idx))
        entry.bind("<FocusOut>", lambda e: self.save_edit(entry, item_id, column, df_col_idx))
        entry.bind("<Escape>", lambda e: entry.destroy())

    def save_edit(self, entry, item_id, column, df_col_idx):
        if not entry.winfo_exists():
            return 
        new_value = entry.get()
        self.tree.set(item_id, column, new_value)
        row_num_str = self.tree.set(item_id, '#1')
        df_row_idx = int(row_num_str) - 1
        self.df.iat[df_row_idx, df_col_idx] = new_value
        entry.destroy()
        self.update_preview()

    def show_edit_warning(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Notice")
        dlg.geometry("350x150")
        dlg.transient(self.root)
        dlg.grab_set() 
        lbl = tk.Label(dlg, text="Changes you make here are temporary and\nwill not be saved to the source file.", pady=15)
        lbl.pack()
        chk = tk.Checkbutton(dlg, text="Don't show this again", variable=self.hide_edit_warning)
        chk.pack(pady=5)
        btn = tk.Button(dlg, text="Got it", command=dlg.destroy, width=15)
        btn.pack(pady=5)
        self.root.wait_window(dlg) 

    def populate_tree(self):
        for index, row in self.df.iterrows():
            values = [index + 1] + list(row)
            row_tag = self.row_tags.get(index, "normal")
            self.tree.insert("", "end", values=values, tags=(row_tag,))

    def on_row_select(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            self.lbl_cell_info.config(text="")
            return
            
        item_id = selected_items[0]
        row_num_str = self.tree.set(item_id, '#1')
        try:
            row_idx = int(row_num_str) - 1
        except ValueError:
            return
            
        colored_cells = []
        col_count = len(self.df.columns)
        for c_idx in range(col_count):
            if (row_idx, c_idx) in self.cell_colors:
                hex_c = self.cell_colors[(row_idx, c_idx)]
                cat = self.classify_color(hex_c)
                col_letter = get_col_letter(c_idx)
                tag_desc = cat.capitalize() if cat else hex_c
                colored_cells.append(f"Col {col_letter}: {tag_desc}")
                
        if colored_cells:
            self.lbl_cell_info.config(text="•  " + " | ".join(colored_cells))
        else:
            self.lbl_cell_info.config(text="")

    def zoom_in(self):
        if self.current_font_size < 24:
            self.current_font_size += 2
            self.update_tree_font()

    def zoom_out(self):
        if self.current_font_size > 6:
            self.current_font_size -= 2
            self.update_tree_font()

    def update_tree_font(self):
        self.style.configure("Treeview", font=("Arial", self.current_font_size), rowheight=int(self.current_font_size * 2.5))
        self.style.configure("Treeview.Heading", font=("Arial", self.current_font_size, "bold"))

    def show_log(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("History Log")
        dlg.geometry("500x350")
        dlg.transient(self.root)
        
        log_frame = tk.Frame(dlg, padx=12, pady=12)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_text = tk.Text(log_frame, wrap="word", font=("Segoe UI", 9), bg="#fafafa", relief=tk.SOLID, borderwidth=1)
        log_scroll = tk.Scrollbar(log_frame, command=log_text.yview)
        log_text.configure(yscrollcommand=log_scroll.set)
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load the log file contents into the popup
        if os.path.exists(self.log_path):
            with open(self.log_path, 'r', encoding='utf-8') as f:
                log_text.insert(tk.END, f.read())
        
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)
        
        btn_exit = tk.Button(dlg, text="Close", command=dlg.destroy, width=15)
        btn_exit.pack(pady=(0, 10))

    def on_typing(self, event):
        if event.keysym == "Return":
            return 
        if self.debounce_timer:
            self.root.after_cancel(self.debounce_timer)
        self.debounce_timer = self.root.after(800, self.update_preview)

    def log_experiment(self, event=None):
        user_input = self.entry_experiment.get().strip()
        if not user_input:
            return
        try:
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.write(user_input + '\n')
        except Exception as log_error:
            print(f"Warning: Could not write to log: {log_error}")

    def update_preview(self):
        # NEW: Force a live reload of the creator script so changes apply instantly!
        global pc, spec
        try:
            spec.loader.exec_module(pc)
        except Exception as e:
            print(f"Warning: Could not hot-reload creator script: {e}")
            
        user_input = self.entry_experiment.get().strip()
        if not user_input:
            self.lbl_preview.config(image="", text="Preview will appear here")
            self.btn_edit.config(state=tk.DISABLED) 
            self.current_temp_pdf = None
            return
            
        try:
            rows = 4 if self.labels_per_page.get() == 12 else 3
            # Use underscores to ignore unused gutter values
            label_h, m_x, _, m_y, _ = pc.get_layout_geometry(rows)

            jobs = pc.parse_input_to_jobs(user_input, self.df, rows_per_page=rows)
            for job in jobs:
                job['count'] = 1 
                
            temp_pdf = os.path.join(tempfile.gettempdir(), "preview.pdf")
            pc.generate_pdf(
                jobs, 
                output_filename=temp_pdf, 
                rows_per_page=rows, 
                img_v_offset=self.img_v_offset,
                img_zoom=self.img_zoom
            )

            self.current_temp_pdf = temp_pdf
            
            doc = pymupdf.open(temp_pdf)
            page = doc.load_page(0)
            
            dpi = 95
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            dpi_scale = dpi / 72.0 
            
            # STRICT EDGE-TO-EDGE CROP: Removing all artificial padding so the 
            # preview window perfectly matches the physical label bounds
            left = m_x * dpi_scale
            top = m_y * dpi_scale
            right = (m_x + pc.LABEL_WIDTH) * dpi_scale
            bottom = (m_y + label_h) * dpi_scale
            
            img = img.crop((left, top, right, bottom))
            self.preview_img = ImageTk.PhotoImage(img)
            self.lbl_preview.config(image=self.preview_img, text="")
            self.btn_edit.config(state=tk.NORMAL)
            doc.close()
            
        except Exception as e:
            self.lbl_preview.config(image="", text=f"Error generating preview:\n{e}")
            self.btn_edit.config(state=tk.DISABLED)

    def open_in_inkscape(self):
        if not hasattr(self, 'current_temp_pdf') or not os.path.exists(self.current_temp_pdf):
            return
            
        config_file = os.path.join(pc.ASSETS_DIR, "inkscape_path.txt")
        inkscape_path = None
        
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    saved_path = f.read().strip()
                    if os.path.exists(saved_path):
                        inkscape_path = saved_path
            except Exception:
                pass
        
        if not inkscape_path:
            inkscape_path = shutil.which("inkscape")
            
        if not inkscape_path:
            common_paths = [
                r"C:\Program Files\WindowsApps\25415Inkscape.Inkscape_1.4.40.0_x64__9waqn51p1ttv2\VFS\ProgramFilesX64\Inkscape\bin\inkscape.exe",
                r"C:\Program Files\Inkscape\bin\inkscape.exe",
                r"C:\Program Files\Inkscape\inkscape.exe",
                r"C:\Program Files (x86)\Inkscape\bin\inkscape.exe",
                r"C:\Program Files (x86)\Inkscape\inkscape.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inkscape\bin\inkscape.exe")
            ]
            for path in common_paths:
                if os.path.exists(path):
                    inkscape_path = path
                    break

        if not inkscape_path:
            prompt = messagebox.askyesno(
                "Inkscape Not Found", 
                "Could not automatically find Inkscape.\n\nWould you like to manually locate inkscape.exe on your computer?"
            )
            if prompt:
                inkscape_path = filedialog.askopenfilename(
                    title="Locate inkscape.exe",
                    filetypes=[("Executable Files", "*.exe")]
                )
                if inkscape_path and os.path.exists(inkscape_path):
                    try:
                        with open(config_file, 'w', encoding='utf-8') as f:
                            f.write(inkscape_path)
                    except Exception:
                        pass
            else:
                return 

        if inkscape_path and os.path.exists(inkscape_path):
            try:
                if not os.path.exists(config_file):
                    try:
                        with open(config_file, 'w', encoding='utf-8') as f:
                            f.write(inkscape_path)
                    except Exception:
                        pass
                
                launch_cmd = f'"{inkscape_path}" "{self.current_temp_pdf}"'
                subprocess.Popen(launch_cmd, shell=True)
            except Exception as e:
                messagebox.showerror("Launch Error", f"Could not launch Inkscape:\n{e}")

    def add_to_pdf(self):
        # Grabs the label currently displayed in the window
        user_input = self.entry_experiment.get().strip()
        if not user_input:
            messagebox.showwarning("Warning", "No label to add. Please configure a label in the experiment box first.")
            return

        try:
            copies_str = self.entry_copies.get().strip()
            copies = int(copies_str)
            if copies < 1 or len(copies_str) > 2:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid 1 or 2 digit number for copies.")
            return

        try:
            rows = 4 if self.labels_per_page.get() == 12 else 3
            
            # Parse the current experimental input into jobs using the creator script
            jobs = pc.parse_input_to_jobs(user_input, self.df, rows_per_page=rows)
            
            for job in jobs:
                job['count'] = copies 
                # --- NEW: Snapshot the current zoom and nudge state for this specific job ---
                job['img_v_offset'] = self.img_v_offset
                job['img_zoom'] = self.img_zoom
                
                self.pdf_jobs.append(job)
                self.total_queued_labels += copies

            self.lbl_queue_status.config(text=f"Queued labels: {self.total_queued_labels}")
            self.entry_copies.delete(0, tk.END)
            
            # Log the addition
            try:
                with open(self.log_path, 'a', encoding='utf-8') as f:
                    f.write(f"QUEUED ({copies}x): " + user_input + '\n')
            except Exception as log_error:
                print(f"Warning: Could not write to log: {log_error}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add label to queue:\n{e}")

    def generate_queued_pdf(self):
        if not self.pdf_jobs:
            messagebox.showinfo("Info", "There are no labels queued to generate.")
            return
            
        try:
            rows = 4 if self.labels_per_page.get() == 12 else 3
            output_file = "Final_Packaging.pdf"
            
            # Generate the PDF with the accumulated jobs list
            pc.generate_pdf(
                self.pdf_jobs, 
                output_filename=output_file, 
                rows_per_page=rows, 
                img_v_offset=self.img_v_offset,
                img_zoom=self.img_zoom
            )
            
            messagebox.showinfo("Success", f"Generated {output_file} successfully with {self.total_queued_labels} total labels!")
            
            # Reset the queue after successful generation
            self.pdf_jobs = []
            self.total_queued_labels = 0
            self.lbl_queue_status.config(text="Queued labels: 0")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate queued packaging:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PackagingApp(root)
    root.mainloop()