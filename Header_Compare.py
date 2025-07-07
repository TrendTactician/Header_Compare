import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import difflib

# Define delimiters
DELIMITERS = {
    'Pipe (|)': '|',
    'Backslash (\\)': '\\',
    'Forward Slash (/)': '/',
    'Comma (,)': ',',
    'Colon (:)': ':',
    'Semicolon (;)': ';',
    'Tab (\\t)': '\t',
    'New Line (\\n)': '\n',
    'Space ( )': ' ',
    'Double Colon (::)': '::',
    'Tilde (~)': '~'
}

def clean_col_name(col):
    if col.startswith('\ufeff'):
        col = col[1:]
    return col.strip()

def get_header(file_path, delimiter):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            first_line = f.readline().strip('\r\n')
            if not first_line:
                return None, "Header missing"
            if delimiter is None:
                cols = first_line.split()
            else:
                try:
                    cols = first_line.split(delimiter)
                except Exception:
                    cols = first_line.split()  # Fallback to whitespace
            return cols, None
    except Exception as e:
        return None, str(e)

def find_files(folder):
    files = [f for f in os.listdir(folder)
             if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith(('.txt', '.csv'))]
    files.sort()
    return files

class HeaderCompareApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Header Comparison Tool")
        width = self.winfo_screenwidth()
        height = self.winfo_screenheight()
        self.state('zoomed')

        self.configure(bg="#f5f7fa")

        self.main_folder = tk.StringVar()
        self.comp_folder = tk.StringVar()
        self.report_file = tk.StringVar()
        self.delimiter_name = tk.StringVar(value='None')
        self.override_custom_delim = tk.StringVar(value='')  # For override custom delimiter input
        

        self.main_files = []
        self.comp_files = []

        self.match_vars = {}
        self.rows = {}
        self.delim_vars = {}
        self.delim_custom_vars = {}  # per-file custom delimiter entry variables
        self.original_delim_values = {}

        self.override_delimiter = None  # stores delimiter character or None
        self.override_custom_active = False

        self.setup_ui()
        self.setup_style()

    def setup_style(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background="#f5f7fa")
        style.configure('TLabel', font=("Segoe UI", 11), background="#f5f7fa")
        style.configure('Header.TLabel', font=("Segoe UI", 11, "bold"), background="#f5f7fa")
        style.configure('TButton', font=("Segoe UI", 10, "bold"), foreground="#fff", background="#0078d7", padding=6)
        style.map('TButton', background=[('active','#005a9e'), ('disabled','#a6a6a6')])
        style.configure('TCombobox', fieldbackground='#fff', background='#fff', foreground='#333', font=("Segoe UI",10))

    def show_dev_info(self):
        popup = tk.Toplevel(self)
        popup.title("Application Info")
        popup.geometry("450x390")
        popup.configure(bg="#f5f7fa")
        self.eval(f'tk::PlaceWindow {str(popup)} center')
        info = (
            "Version: 3.0\n"
            "Developer: MarkQCode\n"
            "Features:\n"
            "- Select and compare headers from files in two folders\n"
            "- Supports multiple delimiters and auto-detect\n"
            "- Per-file and Delimiter override\n"
            "- Auto-matching and manual-matching of files\n"
            "- Generate detailed comparison report\n"
            "- Save report to any location\n\n"
            "Inside Report:\n"
            "- Missing, Extra, Reordered columns\n"
            "- Detect case sensitivity (Lower/Upper Case)\n"
            "- leading/trailing spaces\n"
            "- List of extra files from both folder\n"
        )
        ttk.Label(popup, text=info, background="#f5f7fa", font=("Segoe UI",10), justify='left').pack(padx=15,pady=15, fill='both', expand=True)
        ttk.Button(popup, text="Close", command=popup.destroy).pack(pady=(0,15))

    def setup_ui(self):
        frm = ttk.Frame(self)
        frm.pack(padx=10, pady=10, fill='x')
        ttk.Label(frm, text="Select Main Folder:").grid(row=0, column=0, sticky='w')
        ttk.Entry(frm, textvariable=self.main_folder, width=70).grid(row=0, column=1, sticky='ew', padx=5)
        ttk.Button(frm, text="Browse", command=lambda:self.browse_folder(self.main_folder)).grid(row=0, column=2)

        ttk.Label(frm, text="Select Comparison Folder:").grid(row=1, column=0, sticky='w')
        ttk.Entry(frm, textvariable=self.comp_folder, width=70).grid(row=1, column=1, sticky='ew', padx=5)
        ttk.Button(frm, text="Browse", command=lambda:self.browse_folder(self.comp_folder)).grid(row=1, column=2)

        ttk.Label(frm, text="Save Report As:").grid(row=2, column=0, sticky='w')
        ttk.Entry(frm, textvariable=self.report_file, width=70).grid(row=2, column=1, sticky='ew', padx=5)
        ttk.Button(frm, text="Browse", command=self.browse_report).grid(row=2, column=2)

        # Global Delimiter Override: dropdown with Apply & Reset closely aligned
        ttk.Label(frm, text="Delimiter Override:").grid(row=3, column=0, sticky='w')

        # Use a frame to group dropdown + custom entry + buttons
        delim_frame = ttk.Frame(frm)
        delim_frame.grid(row=3, column=1, sticky='w')

        # Combobox for delimiter override
        self.override_cb = ttk.Combobox(
            delim_frame,
            values=['None', 'Custom'] + list(DELIMITERS.keys()),
            state='readonly',
            textvariable=self.delimiter_name,
            width=25
        )
        self.override_cb.grid(row=0, column=0, pady=2)
        self.override_cb.bind('<<ComboboxSelected>>', self.on_override_delim_change)

        # Custom delimiter entry, initially hidden
        self.override_custom_entry = ttk.Entry(delim_frame, textvariable=self.override_custom_delim, width=5)
        self.override_custom_entry.grid(row=0, column=1, padx=(5,0), pady=2)
        self.override_custom_entry.grid_remove()  # hide initially

        # Apply and Reset buttons next to entry
        ttk.Button(delim_frame, text="Apply", command=self.apply_override_delimiter).grid(row=0, column=2, padx=(5, 2), pady=2)
        ttk.Button(delim_frame, text="Reset", command=self.reset_override_delimiter).grid(row=0, column=3, pady=2)

        frm.columnconfigure(1, weight=1)



        self.files_frame = ttk.Frame(self)
        self.files_frame.pack(padx=10, pady=10, fill='both', expand=True)
        canvas = tk.Canvas(self.files_frame, bg="#f5f7fa", highlightthickness=0)
        sb = ttk.Scrollbar(self.files_frame, orient='vertical', command=canvas.yview)
        self.scrollable = ttk.Frame(canvas)
        self.scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0,0), window=self.scrollable, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        btnfrm = ttk.Frame(self)
        btnfrm.pack(padx=10, pady=10, fill='x')

        btn_load = ttk.Button(btnfrm, text="Load Files", command=self.load_files)
        btn_load.grid(row=0, column=0, sticky='w', padx=5)

        btn_info = ttk.Button(btnfrm, text="App Info", command=self.show_dev_info)
        btn_info.grid(row=0, column=2, padx=5)

        btn_compare = ttk.Button(btnfrm, text="Compare", command=self.compare_and_save_report)
        btn_compare.grid(row=0, column=4, sticky='e', padx=5)

        # Add empty columns 1 and 3 that expand, pushing buttons apart and centering the middle button
        btnfrm.columnconfigure(1, weight=1)
        btnfrm.columnconfigure(3, weight=1)
        self.on_override_delim_change()


    def browse_folder(self, var):
        popup = tk.Toplevel(self)
        popup.title("Select Option")
        popup.geometry("120x100")
        self.eval(f'tk::PlaceWindow {str(popup)} center')

        def choose_files():
            popup.destroy()
            f = filedialog.askopenfilename(multiple=True, filetypes=[("Text/CSV Files", "*.txt *.csv"), ("All Files", "*.*")])
            if f:
                folder = os.path.dirname(f[0])
                var.set(folder)

        def choose_folder():
            popup.destroy()
            fld = filedialog.askdirectory()
            if fld:
                var.set(fld)

        ttk.Button(popup, text="File", command=choose_files).pack(padx=10, pady=5, fill='x')
        ttk.Button(popup, text="FolderS", command=choose_folder).pack(padx=10, pady=5, fill='x')


    def browse_report(self):
        fn = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[("Text","*.txt"),("CSV","*.csv"),("All","*.*")])
        if fn: self.report_file.set(fn)

    def load_files(self):
        mf, cf = self.main_folder.get(), self.comp_folder.get()
        if not os.path.isdir(mf) or not os.path.isdir(cf):
            messagebox.showerror("Error", "Invalid folder paths!")
            return
        self.main_files = find_files(mf)
        self.comp_files = find_files(cf)
        for w in self.scrollable.winfo_children():
            w.destroy()
        self.match_vars.clear()
        self.delim_vars.clear()
        self.delim_custom_vars.clear()
        self.rows.clear()

        # Headings with Remove heading aligned above button, empty space reserved for custom delimiter entry
        headings = ["Main File", "Match Comp File", "Delimiter", "", "Remove"]
        for col, text in enumerate(headings):
            ttk.Label(self.scrollable, text=text, style='Header.TLabel').grid(row=0, column=col, padx=5, pady=2)

        for i, mfname in enumerate(self.main_files, start=1):
            ttk.Label(self.scrollable, text=mfname).grid(row=i, column=0, sticky='w', padx=5)

            mv = tk.StringVar()
            lo = [f.lower() for f in self.comp_files]
            matches = difflib.get_close_matches(mfname.lower(), lo, cutoff=0.6)
            mv.set(self.comp_files[lo.index(matches[0])] if matches else '')
            cb_match = ttk.Combobox(self.scrollable, values=['']+self.comp_files, textvariable=mv, state='readonly', width=40)
            cb_match.grid(row=i, column=1, sticky='ew', padx=5)

            self.match_vars[mfname] = mv

            # Determine best delimiter automatically based on improved heuristic
            auto_delim = self.auto_detect_delimiter(os.path.join(mf, mfname),
                                                    os.path.join(cf, mv.get()) if mv.get() else None)

            # Per-file delimiter combobox with Custom option + reserved space for entry
            delim_var = tk.StringVar()
            delim_values = ['None', 'Custom'] + list(DELIMITERS.keys())
            delim_var.set('None' if auto_delim is None else self.delim_name_from_char(auto_delim))
            cb_delim = ttk.Combobox(self.scrollable, values=delim_values, textvariable=delim_var, state='readonly', width=15)
            cb_delim.grid(row=i, column=2, sticky='w', padx=5)

            # Custom delimiter entry per file, hidden by default
            custom_delim_var = tk.StringVar()
            ent_custom = ttk.Entry(self.scrollable, textvariable=custom_delim_var, width=5)
            ent_custom.grid(row=i, column=3, sticky='w', padx=(0, 10))
            ent_custom.grid_remove()

            # Bind to show/hide custom entry on combobox change
            cb_delim.bind('<<ComboboxSelected>>', lambda e, v=delim_var, ent=ent_custom: self.toggle_custom_entry(v, ent))

            self.delim_vars[mfname] = delim_var
            self.delim_custom_vars[mfname] = custom_delim_var
            self.original_delim_values[mfname] = delim_var.get()

            btn_remove = ttk.Button(self.scrollable, text="Remove", command=lambda f=mfname: self.remove_file(f))
            btn_remove.grid(row=i, column=4, sticky='e', padx=5)

            self.rows[mfname] = (cb_delim, ent_custom)

        # On initial load, set visibility of custom entries if default delimiter is Custom
        for fname, delim_var in self.delim_vars.items():
            ent = self.rows[fname][1]
            self.toggle_custom_entry(delim_var, ent)

    def toggle_custom_entry(self, delim_var, entry_widget):
        if delim_var.get() == 'Custom':
            entry_widget.grid()
        else:
            entry_widget.grid_remove()

    def remove_file(self, fname):
        for widget in self.scrollable.grid_slaves():
            info = widget.grid_info()
            if info['row'] > 0 and self.scrollable.grid_slaves(row=info['row'], column=0)[0].cget('text') == fname:
                for w in self.scrollable.grid_slaves(row=info['row']):
                    w.destroy()
                break
        self.match_vars.pop(fname, None)
        self.delim_vars.pop(fname, None)
        self.delim_custom_vars.pop(fname, None)
        self.rows.pop(fname, None)
        self.original_delim_values.pop(fname, None)

    def apply_override_delimiter(self):
        sel = self.delimiter_name.get()
        if sel == 'Custom':
            val = self.override_custom_delim.get()
            if not val:
                messagebox.showerror("Error", "Please enter a custom delimiter.")
                return
            self.override_delimiter = val
            self.override_custom_active = True
        elif sel == 'None':
            self.override_delimiter = None
            self.override_custom_active = False
        else:
            self.override_delimiter = DELIMITERS.get(sel, sel)
            self.override_custom_active = False

    def reset_override_delimiter(self):
        self.delimiter_name.set('None')
        self.override_custom_delim.set('')
        self.override_delimiter = None
        self.override_custom_active = False
        # Reset all per-file delimiters to original and hide custom entry
        for fname, delim_var in self.delim_vars.items():
            delim_var.set(self.original_delim_values.get(fname, 'None'))
            self.toggle_custom_entry(delim_var, self.rows[fname][1])

    def on_override_delim_change(self, event=None):
        if self.delimiter_name.get() == 'Custom':
            self.override_custom_entry.grid()
        else:
            self.override_custom_entry.grid_remove()

    def delim_name_from_char(self, delim_char):
        for k, v in DELIMITERS.items():
            if v == delim_char:
                return k
        return None

    def delim_char_from_name(self, delim_name):
        if delim_name == 'None':
            return None
        if delim_name == 'Custom':
            return None  # handled separately
        return DELIMITERS.get(delim_name, None)

    def auto_detect_delimiter(self, main_file, comp_file):
        """
        Enhanced delimiter auto-detection:
        For main_file and comp_file (both file paths or None), try each delimiter plus None (whitespace).
        Pick the delimiter that results in the maximum column count intersection between the two headers.
        Return delimiter character or None if no good delimiter found.
        """
        candidates = [None] + list(DELIMITERS.values())
        best_delim = None
        best_score = -1

        # Read headers once outside loop
        main_header_lines = []
        comp_header_lines = []

        try:
            with open(main_file, 'r', encoding='utf-8') as f:
                main_header_lines.append(f.readline().strip('\r\n'))
        except:
            return None
        if comp_file and os.path.exists(comp_file):
            try:
                with open(comp_file, 'r', encoding='utf-8') as f:
                    comp_header_lines.append(f.readline().strip('\r\n'))
            except:
                comp_header_lines = []
        else:
            comp_header_lines = []

        # If comp_file not present or empty, consider only main file for scoring (number of columns)
        for delim in candidates:
            try:
                if delim is None:
                    main_cols = main_header_lines[0].split()
                    comp_cols = comp_header_lines[0].split() if comp_header_lines else []
                else:
                    main_cols = main_header_lines[0].split(delim)
                    comp_cols = comp_header_lines[0].split(delim) if comp_header_lines else []

                main_cols = [clean_col_name(c) for c in main_cols]
                comp_cols = [clean_col_name(c) for c in comp_cols]

                if not main_cols or (comp_header_lines and not comp_cols):
                    continue

                # Score by intersection count or column count for main file if no comp
                if comp_cols:
                    score = len(set(map(str.lower, main_cols)) & set(map(str.lower, comp_cols)))
                else:
                    score = len(main_cols)

                # Choose delimiter with highest score; tie break: more columns in main file preferred
                if score > best_score or (score == best_score and len(main_cols) > (len(main_header_lines[0].split(best_delim)) if best_delim else 0)):
                    best_score = score
                    best_delim = delim
            except Exception:
                continue
        return best_delim

    def get_effective_delimiter(self, fname):
        if self.override_delimiter is not None:
            return self.override_delimiter

        if fname not in self.delim_vars:
            return None  # Safe fallback if missing

        delim_name = self.delim_vars[fname].get()
        if delim_name == 'None':
            return None
        elif delim_name == 'Custom':
            val = self.delim_custom_vars.get(fname, tk.StringVar()).get()
            return val if val else None
        else:
            delim = DELIMITERS.get(delim_name)
            if delim is None:
                # Fallback: Treat as custom delimiter not in list
                self.delim_vars[fname].set("Custom")
                self.delim_custom_vars[fname].set(delim_name)
                return delim_name
            return delim

    def compare_headers(self, main_cols, comp_cols):
        mc = [clean_col_name(c) for c in main_cols]
        cc = [clean_col_name(c) for c in comp_cols]
        ml, cl = [c.lower() for c in mc], [c.lower() for c in cc]

        missing = [orig for orig, low in zip(main_cols, ml) if low not in cl]
        extra = [orig for orig, low in zip(comp_cols, cl) if low not in ml]
        reorder = [(main_cols[i], i+1, cl.index(low)+1) for i, low in enumerate(ml) if low in cl and i != cl.index(low)]
        lt_issues = []
        mlc, clc = set(), set()
        def sp(s): return (len(s)-len(s.lstrip())), (len(s)-len(s.rstrip()))
        for c in main_cols:
            l, t = sp(c)
            if l>0 or t>0:
                lt_issues.append((c,l,t,'Main')); mlc.add(c.strip().lower())
        for c in comp_cols:
            l, t = sp(c)
            if l>0 or t>0:
                lt_issues.append((c,l,t,'Comparison')); clc.add(c.strip().lower())

        case_diff = [(main_cols[i], comp_cols[cl.index(ml[i])]) for i in range(len(ml)) if ml[i] in cl and main_cols[i] != comp_cols[cl.index(ml[i])] and ml[i] not in mlc and ml[i] not in clc]

        exact = len(main_cols)==len(comp_cols) and all(mc[i].lstrip('\ufeff')==cc[i].lstrip('\ufeff') for i in range(len(mc))) and not lt_issues
        out = ["-"*80, f"Main File: {self.current_main_file}", f"Comparison File: {self.current_comp_file}", "-"*80]
        if exact:
            out.append("Headers match exactly.")
        if lt_issues:
            out.append("\nLeading/Trailing space issues:")
            for c,l,t,o in lt_issues:
                out.append(f" - [{o}] '{c}' Lead:{l} Trail:{t}")
        if not exact:
            if missing:
                out.append("\nMissing in Comparison:")
                out += [f" - {c}" for c in missing]
            if extra:
                out.append("\nExtra in Comparison:")
                out += [f" - {c}" for c in extra]
            if reorder:
                out.append("\nReordered columns:")
                out += [f" - '{c}' Main:{pm} Comp:{pc}" for c,pm,pc in reorder]
            if case_diff:
                out.append("\nCase differences:")
                out += [f" - Main:'{m}' vs Comp:'{c}'" for m,c in case_diff]
        out.append("="*80)
        out.append("")
        return "\n".join(out)

    def generate_report(self):
        mfld, cfl = self.main_folder.get(), self.comp_folder.get()
        unmatched_m, unmatched_c = set(self.main_files), set(self.comp_files)
        blocks = []

        for mf in self.main_files:
            cf = self.match_vars[mf].get()
            if not cf:
                continue
            self.current_main_file, self.current_comp_file = mf, cf

            if self.override_delimiter is not None:
                delim = self.override_delimiter
            else:
                sel = self.delim_vars[mf].get()
                if sel in DELIMITERS:
                    delim = DELIMITERS[sel]
                else:
                    delim = None

            mp = os.path.join(mfld, mf); cp = os.path.join(cfl, cf)
            mcols, me = get_header(mp, delim)
            ccols, ce = get_header(cp, delim)

            if me or ce or not mcols or not ccols:
                err = ["-"*80, f"Main File: {mf}", f"Comparison File: {cf}", "-"*80]
                if me or not mcols:
                    err.append(f"In Main: {me or 'Header missing'}")
                if ce or not ccols:
                    err.append(f"In Comparison: {ce or 'Header missing'}")
                err.append("="*80)
                err.append("")
                blocks.append("\n".join(err))
            else:
                blocks.append(self.compare_headers(mcols, ccols))

            unmatched_m.discard(mf)
            unmatched_c.discard(cf)

        if unmatched_m:
            blocks.append("="*80 + "\nUnmatched Main Files:\n" + "\n".join(f" - {f}" for f in sorted(unmatched_m)) + "\n" + "="*80 + "\n")
        if unmatched_c:
            blocks.append("="*80 + "\nUnmatched Comparison Files:\n" + "\n".join(f" - {f}" for f in sorted(unmatched_c)) + "\n" + "="*80 + "\n")

        return "\n".join(blocks)

    def compare_and_save_report(self):
        if not self.main_folder.get() or not self.comp_folder.get() or not self.report_file.get():
            messagebox.showwarning("Warning","Please select folders and report path.")
            return
        try:
            rpt = self.generate_report()
            with open(self.report_file.get(), 'w', encoding='utf-8') as f:
                f.write(rpt)
            messagebox.showinfo("Success","Report saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed saving report: {e}")

if __name__ == "__main__":
    app = HeaderCompareApp()
    app.mainloop()