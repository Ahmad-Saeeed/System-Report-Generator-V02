import os, re, sys, json, base64, webbrowser, subprocess, tempfile, shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from docx import Document

APP_DIR = Path(__file__).resolve().parent
ASSETS = APP_DIR / 'assets'
TEMPLATE = ASSETS / 'template.html'
LOGO = ASSETS / 'logo.png'

# build_email.py is bundled beside this app when packaged.
sys.path.insert(0, str(APP_DIR / 'scripts'))
from build_email import build_email_html, build_intro_text

STATUSES = ['PASS', 'FAIL', 'BLOCKED', 'IN PROGRESS', 'INVALID', 'NOT RUN']
STATUS_KEYS = ['pass', 'fail', 'blocked', 'inprogress', 'invalid', 'notrun']

# Built-in starting choices for the 4 report-detail dropdowns — used only to seed each
# field's option list the very first time the app runs. After that, the current list for
# every field (including these) lives in config.json under "field_options" and can be
# freely edited (added to, renamed, or deleted) via "Edit list…" — nothing is protected.
FIELD_BASE_VALUES = {
    'system':        ['TMS', 'MW CP4I13'],
    'environment':   ['Testing', 'IOT', 'E2E'],
    'report_type':   ['E2E', 'IOT'],
    'report_status': ['Completed', 'Not Completed'],
}
FIELD_LABELS = {'system': 'System', 'environment': 'Environment',
                'report_type': 'Report Type', 'report_status': 'Status'}


# ---- Tester name + settings: asked once, then remembered for every future run ----
def config_path():
    # %APPDATA% on Windows, ~/.QCReportGenerator elsewhere — survives exe updates/moves.
    base = Path(os.environ.get('APPDATA', '')) if sys.platform == 'win32' and os.environ.get('APPDATA') else Path.home()
    cfg_dir = base / 'QCReportGenerator'
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir / 'config.json'


def _legacy_config_paths():
    """Older config locations, checked as fallbacks so existing users aren't
    asked for their name again after the app was renamed/reorganised."""
    base = Path(os.environ.get('APPDATA', '')) if sys.platform == 'win32' and os.environ.get('APPDATA') else Path.home()
    return [
        base / 'TMSReportGenerator' / 'config.json',        # pre-rename "TMS Report Generator"
        base / 'TMS_Report_Generator' / 'settings.json',    # NewVersion.txt settings file
    ]


def load_config():
    p = config_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding='utf-8'))
        except Exception:
            return {}
    for legacy in _legacy_config_paths():
        if legacy.exists():
            try:
                data = json.loads(legacy.read_text(encoding='utf-8'))
                # Normalise the NewVersion "options"/"last" shape into the OldVersion shape.
                if 'last' in data and 'options' in data:
                    cfg = {}
                    if isinstance(data['last'].get('Tester'), str):
                        cfg['tester'] = data['last']['Tester']
                    cfg['custom_values'] = {k: list(v) for k, v in data.get('options', {}).items()}
                    cfg['last_selection'] = {k.lower().replace(' ', '_'): v
                                             for k, v in data.get('last', {}).items()
                                             if isinstance(v, str) and not k.endswith('_other')}
                    return cfg
                return data
            except Exception:
                return {}
    return {}


def save_config(cfg):
    config_path().write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding='utf-8')


def copy_html_to_clipboard(html_str):
    """Best-effort: put rich HTML on the Windows clipboard (the standard CF_HTML format)
    so pasting into Outlook/Gmail/Word renders the formatted content directly — the same
    result as opening the file, Ctrl+A, Ctrl+C. Silently does nothing if pywin32 isn't
    available (e.g. running outside Windows) or if anything about the clipboard fails;
    never raises, since this is a convenience on top of the file that's already written."""
    try:
        import win32clipboard
    except Exception:
        return False

    marker_template = (
        "Version:0.9\r\n"
        "StartHTML:{start_html:08d}\r\n"
        "EndHTML:{end_html:08d}\r\n"
        "StartFragment:{start_fragment:08d}\r\n"
        "EndFragment:{end_fragment:08d}\r\n"
    )
    prefix = "<html><body><!--StartFragment-->"
    suffix = "<!--EndFragment--></body></html>"

    header_len = len(marker_template.format(start_html=0, end_html=0, start_fragment=0, end_fragment=0).encode('utf-8'))
    start_html = header_len
    start_fragment = start_html + len(prefix.encode('utf-8'))
    end_fragment = start_fragment + len(html_str.encode('utf-8'))
    end_html = end_fragment + len(suffix.encode('utf-8'))
    cf_html = marker_template.format(start_html=start_html, end_html=end_html,
                                      start_fragment=start_fragment, end_fragment=end_fragment) + prefix + html_str + suffix

    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            cf_format = win32clipboard.RegisterClipboardFormat('HTML Format')
            win32clipboard.SetClipboardData(cf_format, cf_html.encode('utf-8'))
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, html_str)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        return False


def cairo_now():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo('Africa/Cairo'))
    except Exception:
        return datetime.now()


def extract_cr_number(filename):
    base = os.path.splitext(os.path.basename(filename))[0]
    base = re.sub(r'[\s_-]*TestReport\s*$', '', base, flags=re.IGNORECASE).rstrip(' -_')
    m = re.match(r'^CR[\s_\-#]?(\d.*)$', base, flags=re.IGNORECASE)
    return 'CR#' + m.group(1) if m else base


def count_test_cases(docx_path):
    doc = Document(docx_path)
    counts = {s: 0 for s in STATUSES}
    pattern = re.compile(r'\[(PASS|FAIL|BLOCKED|IN PROGRESS|INVALID|NOT RUN)\]', re.I)
    tagged = []
    for p in doc.paragraphs:
        text = p.text.strip()
        style = (p.style.name if p.style else '').lower()
        is_heading = style.startswith('heading') or re.match(r'^#{1,6}\s+', text)
        if not is_heading:
            continue
        m = pattern.search(text.replace('\\', ''))
        if m:
            s = m.group(1).upper()
            counts[s] += 1
            tagged.append(text)
    return counts, tagged


def make_report_html(system, cr_number, test_date, environment, report_type, report_status,
                     tester, counts, observations):
    html = TEMPLATE.read_text(encoding='utf-8')
    now = cairo_now()
    gen_date = now.strftime('%B %d, %Y at %I:%M %p')
    title = f'{report_type} Testing Progress Report'
    logo_b64 = base64.b64encode(LOGO.read_bytes()).decode('ascii') if LOGO.exists() else ''

    status_values = [counts.get(s, 0) for s in STATUSES]
    status_pattern = re.compile(r"const statusData = \[.*?\n\s*\];", re.S)
    status_block = """const statusData = [\n""" + ''.join([
        f"    {{ key:'{k}', label:'{label.title() if label != 'IN PROGRESS' else 'In Progress'}', icon:'{icon}', count:{n}, color:'{color}' }},\n"
        for k, label, icon, n, color in [
            ('pass','PASS','✅',status_values[0],'#22a35a'), ('fail','FAIL','❌',status_values[1],'#e0384c'),
            ('blocked','BLOCKED','⚠️',status_values[2],'#e0a52c'), ('inprogress','IN PROGRESS','🕐',status_values[3],'#3b8fe0'),
            ('invalid','INVALID','❓',status_values[4],'#a15fe0'), ('notrun','NOT RUN','⊖',status_values[5],'#8b93a3')]]) + '  ];'
    html = status_pattern.sub(lambda m: status_block, html, count=1)

    cr_rows = [
        {'k':'System','v':system}, {'k':'CR Number','v':cr_number}, {'k':'Tester','v':tester},
        {'k':'Test Date','v':test_date}, {'k':'Environment','v':environment}, {'k':'Status','v':report_status}]
    cr_block = "let crInfoRows = [\n" + ''.join([f"    {{ k:'{r['k']}', v:{r['v']!r} }},\n" for r in cr_rows]) + '  ];'
    html = re.sub(r"let crInfoRows = \[.*?\n\s*\];", cr_block, html, count=1, flags=re.S)
    obs_js = 'let observations = ' + repr(observations) + ';'
    html = re.sub(r"let observations = \[\];", obs_js, html, count=1)
    html = html.replace('<span id="reportTitleText">E2E Testing Progress Report</span>', f'<span id="reportTitleText">{title}</span>')
    html = html.replace('<span id="footerTitleText">E2E Testing Progress Report</span>', f'<span id="footerTitleText">{title}</span>')
    html = re.sub(r'<span id="genDate">.*?</span>', f'<span id="genDate">{gen_date}</span>', html, count=1)
    if logo_b64:
        html = re.sub(r'<img id="heroLogo" class="hero-logo" src="" alt="Company logo">',
                      f'<img id="heroLogo" class="hero-logo visible" src="data:image/png;base64,{logo_b64}" alt="Company logo">', html, count=1)
        html = html.replace('<span id="logoBtnLabel">Upload Logo</span>', '<span id="logoBtnLabel">Change Logo</span>')
        html = html.replace('id="removeLogoBtn" style="display:none;"', 'id="removeLogoBtn" style="display:inline-flex;"')

    # ---- Strip "Download Excel Template" / "Upload CR Data (Excel)" from the report-only output ----
    html = html.replace('    <button class="tb-btn" id="downloadTemplateBtn">⬇ Download Excel Template</button>\n', '')
    html = re.sub(r'    <label class="tb-btn" id="excelBtn">.*?</label>\n', '', html, count=1, flags=re.S)
    start_marker = '  // ---- Excel template download ----'
    end_marker = '  // ---- Copy as Outlook-compatible HTML for email'
    i, j = html.find(start_marker), html.find(end_marker)
    if i != -1 and j != -1 and j > i:
        html = html[:i] + html[j:]
    html = re.sub(r'<script src="https://cdnjs\.cloudflare\.com/ajax/libs/xlsx/[^"]+"></script>\n', '', html, count=1)
    return html, gen_date


def build_email(system, cr_number, test_date, environment, report_type, report_status,
                tester, counts, observations, gen_date):
    status_data = [
        {'key':'pass','icon':'✅','label':'Pass','count':counts['PASS'],'color':'#22a35a'},
        {'key':'fail','icon':'❌','label':'Fail','count':counts['FAIL'],'color':'#e0384c'},
        {'key':'blocked','icon':'⚠️','label':'Blocked','count':counts['BLOCKED'],'color':'#e0a52c'},
        {'key':'inprogress','icon':'🕐','label':'In Progress','count':counts['IN PROGRESS'],'color':'#3b8fe0'},
        {'key':'invalid','icon':'❓','label':'Invalid','count':counts['INVALID'],'color':'#a15fe0'},
        {'key':'notrun','icon':'⊖','label':'Not Run','count':counts['NOT RUN'],'color':'#8b93a3'}]
    rows = [('System',system),('CR Number',cr_number),('Tester',tester),('Test Date',test_date),
            ('Environment',environment),('Status',report_status)]
    total = sum(d['count'] for d in status_data)
    rate = round(counts['PASS']/total*100) if total else 0
    intro = build_intro_text(report_type, cr_number, rate)
    return build_email_html(f'{report_type} Testing Progress Report', gen_date, str(LOGO),
                            rows, status_data, intro_html=intro, observations=observations)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('QC Report Generator')
        try:
            icon = ASSETS / 'app_icon.ico'
            if icon.exists():
                self.iconbitmap(str(icon))
        except Exception:
            pass
        self.geometry('780x700')
        self.minsize(560, 420)
        self.resizable(True, True)
        self.docx = None
        self.counts = None
        self.observations = []
        cfg = load_config()
        self.last_selection = cfg.get('last_selection', {}) or {}
        self.field_options = self._load_field_options(cfg)
        self.tester = self._load_or_ask_tester_name(cfg)
        self.combo_widgets = {}
        self.field_vars = {}   # populated in _build()
        self._build()
        self.protocol('WM_DELETE_WINDOW', self._on_close)

    def _load_field_options(self, cfg):
        """The editable value list for each dropdown. Seeded once from FIELD_BASE_VALUES
        (or migrated from an older config's "custom_values" list) and from then on lives
        entirely in config.json — every entry, including the original built-ins, can be
        renamed or deleted from "Edit list…"."""
        stored = cfg.get('field_options')
        if stored:
            return {key: list(stored.get(key, FIELD_BASE_VALUES[key])) for key in FIELD_BASE_VALUES}
        # First run (or upgrading from a version that only had "custom_values"): start from
        # the built-ins plus whatever custom values had already been added.
        legacy_custom = cfg.get('custom_values', {}) or {}
        options = {}
        for key, base in FIELD_BASE_VALUES.items():
            base_lower = {b.lower() for b in base}
            extra = [c for c in legacy_custom.get(key, []) if c.lower() not in base_lower]
            options[key] = list(base) + extra
        cfg['field_options'] = options
        cfg.pop('custom_values', None)
        save_config(cfg)
        return options

    def _load_or_ask_tester_name(self, cfg):
        name = (cfg.get('tester') or '').strip()
        if name:
            return name
        self.withdraw()  # hide the empty main window while the dialog is up
        while True:
            name = simpledialog.askstring(
                'Tester Name',
                "Enter your name as it should appear on reports (e.g. 'Ahmed Saeed (ahmed.saiid)').\n"
                "You'll only be asked this once — it will be reused automatically from now on.",
                parent=self)
            if name and name.strip():
                name = name.strip()
                break
            messagebox.showwarning('Name required', 'Please enter a name to continue.')
        cfg['tester'] = name
        save_config(cfg)
        self.deiconify()
        return name

    def _on_close(self):
        try:
            self._persist_current_selection()
        except Exception:
            pass
        self.destroy()

    def _build(self):
        # ---- Scrollable body -----------------------------------------------------------
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True)
        canvas = tk.Canvas(container, highlightthickness=0)
        vscroll = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        vscroll.pack(side='right', fill='y')

        body = ttk.Frame(canvas)
        body_id = canvas.create_window((0, 0), window=body, anchor='nw')
        body.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(body_id, width=e.width))

        def _on_mousewheel(event):
            step = -1 if (event.delta > 0 or getattr(event, 'num', None) == 4) else 1
            canvas.yview_scroll(step, 'units')
        canvas.bind_all('<MouseWheel>', _on_mousewheel)   # Windows / macOS
        canvas.bind_all('<Button-4>', _on_mousewheel)      # Linux scroll up
        canvas.bind_all('<Button-5>', _on_mousewheel)      # Linux scroll down
        # -----------------------------------------------------------------------------

        pad = {'padx': 16, 'pady': 8}
        ttk.Label(body, text='QC Report Generator', font=('Segoe UI', 20, 'bold')).pack(anchor='w', **pad)
        ttk.Label(body, text='Choose a test-report DOCX → fill 4 fields → generate the report (with or without the Outlook email).').pack(anchor='w', padx=16)

        tester_frm = ttk.Frame(body); tester_frm.pack(fill='x', padx=16, pady=(4, 0))
        self.tester_lbl = ttk.Label(tester_frm, text=f'Tester: {self.tester}')
        self.tester_lbl.pack(side='left')
        ttk.Button(tester_frm, text='Change', command=self.change_tester_name).pack(side='left', padx=8)

        frm = ttk.Frame(body); frm.pack(fill='x', padx=16, pady=12)
        ttk.Button(frm, text='Select Test Report (.docx)', command=self.select_docx).grid(row=0, column=0, sticky='w')
        self.file_lbl = ttk.Label(frm, text='No file selected'); self.file_lbl.grid(row=0, column=1, padx=10, sticky='w')
        self.count_lbl = ttk.Label(body, text=''); self.count_lbl.pack(anchor='w', padx=16)

        form = ttk.LabelFrame(body, text='Report Details'); form.pack(fill='x', padx=16, pady=12)

        def _default(key):
            stored = self.last_selection.get(key)
            if stored:
                return stored
            opts = self.field_options.get(key) or FIELD_BASE_VALUES[key]
            return opts[0] if opts else 'Other'

        self.system        = tk.StringVar(value=_default('system'))
        self.environment   = tk.StringVar(value=_default('environment'))
        self.report_type   = tk.StringVar(value=_default('report_type'))
        self.report_status = tk.StringVar(value=_default('report_status'))

        self.other_system        = tk.StringVar()
        self.other_environment   = tk.StringVar()
        self.other_report_type   = tk.StringVar()
        self.other_status        = tk.StringVar()

        self.field_vars = {
            'system':        (self.system,        self.other_system),
            'environment':   (self.environment,   self.other_environment),
            'report_type':   (self.report_type,   self.other_report_type),
            'report_status': (self.report_status, self.other_status),
        }

        self._combo(form, 'System',       'system',        self.system,        0)
        self._combo(form, 'Environment',  'environment',   self.environment,   1)
        self._combo(form, 'Report Type',  'report_type',   self.report_type,   2)
        self._combo(form, 'Status',       'report_status', self.report_status, 3)

        ttk.Label(form, text='Custom values (used only when the matching field is Other):').grid(
            row=4, column=0, columnspan=3, sticky='w', padx=10, pady=(8, 2))
        self._other_entry(form, 'System:',       self.other_system,       5)
        self._other_entry(form, 'Environment:',  self.other_environment,  6)
        self._other_entry(form, 'Report Type:',  self.other_report_type,  7)
        self._other_entry(form, 'Status:',       self.other_status,       8)

        obs = ttk.LabelFrame(body, text='Observations (optional)'); obs.pack(fill='both', padx=16, pady=12)
        self.obs = tk.Text(obs, height=6, wrap='word'); self.obs.pack(fill='both', expand=True, padx=8, pady=8)

        out_frm = ttk.LabelFrame(body, text='Output'); out_frm.pack(fill='x', padx=16, pady=(0, 8))
        self.output_mode = tk.StringVar(value='both')
        ttk.Radiobutton(out_frm, text='Report only', variable=self.output_mode, value='report').pack(side='left', padx=10, pady=6)
        ttk.Radiobutton(out_frm, text='Report + Email', variable=self.output_mode, value='both').pack(side='left', padx=10, pady=6)

        btns = ttk.Frame(body); btns.pack(fill='x', padx=16, pady=8)
        ttk.Button(btns, text='Generate', command=self.generate).pack(side='left')
        ttk.Label(btns, text='© Ahmed Saeed', foreground='#888888').pack(side='right', padx=4)

    # -- dropdown option helpers -------------------------------------------------------
    def _combo_values(self, key):
        return list(self.field_options.get(key, FIELD_BASE_VALUES[key])) + ['Other']

    def _combo(self, parent, label, key, var, row):
        ttk.Label(parent, text=label + ':', width=18).grid(row=row, column=0, sticky='w', padx=10, pady=6)
        cb = ttk.Combobox(parent, textvariable=var, values=self._combo_values(key), state='readonly', width=30)
        cb.grid(row=row, column=1, sticky='w', padx=10, pady=6)
        cb.bind('<<ComboboxSelected>>', lambda e: self._persist_current_selection())
        self.combo_widgets[key] = cb
        ttk.Button(parent, text='Edit list…', width=10, command=lambda k=key: self._manage_list(k)).grid(
            row=row, column=2, sticky='w', padx=(0, 10), pady=6)
        return cb

    def _other_entry(self, parent, label, var, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', padx=10, pady=3)
        e = ttk.Entry(parent, textvariable=var, width=35)
        e.grid(row=row, column=1, columnspan=2, sticky='w', padx=10, pady=3)
        e.bind('<FocusOut>', lambda ev: self._persist_current_selection())

    # -- "Edit list…" per field: add, rename, or delete ANY value, built-in or custom -----
    def _manage_list(self, key):
        """Dialog to add, rename, or delete the values offered for one dropdown. Every
        entry can be changed, including the original built-ins — if the last remaining
        value is deleted, the field falls back to "Other"."""
        dlg = tk.Toplevel(self)
        dlg.title(f'Edit {FIELD_LABELS[key]} options')
        dlg.transient(self); dlg.grab_set()
        dlg.geometry('380x380')

        ttk.Label(dlg,
                  text=f'{FIELD_LABELS[key]} values ("Other" is always available and kept separately):',
                  wraplength=340, justify='left').pack(padx=10, pady=(10, 4), anchor='w')

        frame = ttk.Frame(dlg); frame.pack(fill='both', expand=True, padx=10, pady=4)
        lb = tk.Listbox(frame, exportselection=False)
        lb.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(frame, orient='vertical', command=lb.yview); sb.pack(side='right', fill='y')
        lb.config(yscrollcommand=sb.set)

        def refresh():
            lb.delete(0, 'end')
            for v in self.field_options.get(key, []):
                lb.insert('end', v)

        refresh()

        def add_item():
            val = simpledialog.askstring('Add value', f'New {FIELD_LABELS[key]} value:', parent=dlg)
            if val is None:
                return
            val = val.strip()
            if not val or val.lower() == 'other':
                messagebox.showwarning('Invalid value', '"Other" is reserved and empty values are not allowed.', parent=dlg)
                return
            lst = self.field_options.setdefault(key, [])
            if any(val.lower() == x.lower() for x in lst):
                messagebox.showwarning('Duplicate', 'That value already exists.', parent=dlg)
                return
            lst.append(val)
            refresh(); self._refresh_combo(key); self._persist_current_selection()

        def edit_item():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo('Rename value', 'Select a value first.', parent=dlg)
                return
            old = lb.get(sel[0])
            val = simpledialog.askstring('Rename value', f'Rename "{old}" to:', initialvalue=old, parent=dlg)
            if val is None:
                return
            val = val.strip()
            if not val or val.lower() == 'other':
                messagebox.showwarning('Invalid value', '"Other" is reserved and empty values are not allowed.', parent=dlg)
                return
            lst = self.field_options.get(key, [])
            if val.lower() != old.lower() and any(val.lower() == x.lower() for x in lst):
                messagebox.showwarning('Duplicate', 'That value already exists.', parent=dlg)
                return
            for i, x in enumerate(lst):
                if x == old:
                    lst[i] = val
                    break
            var = self.field_vars[key][0]
            if var.get() == old:
                var.set(val)
            refresh(); self._refresh_combo(key); self._persist_current_selection()

        def delete_item():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo('Delete value', 'Select a value first.', parent=dlg)
                return
            old = lb.get(sel[0])
            if not messagebox.askyesno('Delete value', f'Delete "{old}" from {FIELD_LABELS[key]} options?', parent=dlg):
                return
            self.field_options[key] = [x for x in self.field_options.get(key, []) if x != old]
            var = self.field_vars[key][0]
            if var.get() == old:
                remaining = self.field_options.get(key, [])
                var.set(remaining[0] if remaining else 'Other')
            refresh(); self._refresh_combo(key); self._persist_current_selection()

        btns = ttk.Frame(dlg); btns.pack(fill='x', padx=10, pady=(4, 10))
        ttk.Button(btns, text='Add…',    command=add_item).pack(side='left')
        ttk.Button(btns, text='Rename…', command=edit_item).pack(side='left', padx=6)
        ttk.Button(btns, text='Delete',  command=delete_item).pack(side='left')
        ttk.Button(btns, text='Close',   command=dlg.destroy).pack(side='right')

    def _refresh_combo(self, key):
        cb = self.combo_widgets.get(key)
        if cb is not None:
            cb.config(values=self._combo_values(key))

    # -- selection persistence ----------------------------------------------------------
    def val(self, var, custom):
        return custom.get().strip() if var.get() == 'Other' else var.get()

    def _persist_current_selection(self):
        """Remember the current System/Environment/Report Type/Status picks, and add any
        freshly-typed "Other" value to that field's dropdown so it doesn't need retyping."""
        cfg = load_config()
        last = cfg.get('last_selection', {}) or {}
        for key, (var, other_var) in self.field_vars.items():
            resolved = self.val(var, other_var)
            if not resolved:
                continue
            last[key] = resolved
            if var.get() == 'Other':
                lst = self.field_options.setdefault(key, [])
                if not any(resolved.lower() == x.lower() for x in lst):
                    lst.append(resolved)
        cfg['last_selection'] = last
        cfg['field_options'] = self.field_options
        cfg['tester'] = self.tester
        save_config(cfg)
        self.last_selection = last
        for key, (var, other_var) in self.field_vars.items():
            self._refresh_combo(key)
            if var.get() == 'Other':
                resolved = self.val(var, other_var)
                if resolved:
                    var.set(resolved)

    # -- existing behaviour ------------------------------------------------------------
    def select_docx(self):
        p = filedialog.askopenfilename(title='Select test-results Word document',
                                       filetypes=[('Word documents', '*.docx')])
        if not p:
            return
        try:
            counts, tagged = count_test_cases(p)
            self.docx, self.counts = p, counts
            self.file_lbl.config(text=Path(p).name)
            self.count_lbl.config(text='Found %d test cases: %d Pass, %d Fail, %d Blocked, %d In Progress, %d Invalid, %d Not Run.' % (
                len(tagged), counts['PASS'], counts['FAIL'], counts['BLOCKED'],
                counts['IN PROGRESS'], counts['INVALID'], counts['NOT RUN']))
        except Exception as e:
            messagebox.showerror('Could not read DOCX', str(e))

    def change_tester_name(self):
        name = simpledialog.askstring('Tester Name', 'Update the name used on reports:',
                                      initialvalue=self.tester, parent=self)
        if name and name.strip():
            self.tester = name.strip()
            self.tester_lbl.config(text=f'Tester: {self.tester}')
            cfg = load_config()
            cfg['tester'] = self.tester
            save_config(cfg)

    def generate(self):
        if not self.docx or not self.counts:
            messagebox.showwarning('Missing file', 'Select the test-report DOCX first.')
            return
        system = self.val(self.system, self.other_system)
        env = self.val(self.environment, self.other_environment)
        rtype = self.val(self.report_type, self.other_report_type)
        report_status = self.val(self.report_status, self.other_status)
        if not all([system, env, rtype, report_status]):
            messagebox.showwarning('Missing details', 'Please fill all report details.')
            return
        self._persist_current_selection()
        observations = [x.strip() for x in self.obs.get('1.0', 'end').splitlines() if x.strip()]
        cr = extract_cr_number(self.docx)
        cr_title = re.sub(r'\W+', '_', cr).strip('_')  # filesystem-safe version (CR#6703 -> CR_6703)
        now = cairo_now(); test_date = now.strftime('%d %b %Y')
        include_email = self.output_mode.get() == 'both'
        try:
            report_html, gen_date = make_report_html(system, cr, test_date, env, rtype,
                                                     report_status, self.tester,
                                                     self.counts, observations)
            out_dir = Path(self.docx).parent
            report_path = out_dir / f'QC_{cr_title}_Report.html'
            report_path.write_text(report_html, encoding='utf-8')
            created = [report_path]

            if include_email:
                # Report + Email produces one email-ready HTML file only.
                email_html = build_email(system, cr, test_date, env, rtype, report_status,
                                         self.tester, self.counts, observations, gen_date)
                report_path.unlink(missing_ok=True)
                report_path = out_dir / f'QC_{cr_title}_Report.html'
                report_path.write_text(email_html, encoding='utf-8')
                created = [report_path]

            subject = f'{system}-{cr}'
            if include_email:
                messagebox.showinfo('Done', f'Email-ready report created:\n{report_path}\n\nSubject:\n{subject}')
                # Best-effort only: copies the same email-ready HTML to the clipboard (Windows
                # "HTML Format") so it can be pasted straight into Outlook/Gmail with formatting
                # intact — skipped silently wherever that clipboard format isn't available.
                copy_html_to_clipboard(email_html)
            else:
                messagebox.showinfo('Done', f'Report created:\n{report_path}')

            for p in created:
                webbrowser.open(p.as_uri())
        except Exception as e:
            messagebox.showerror('Generation failed', str(e))


if __name__ == '__main__':
    App().mainloop()
