"""
 SPC GENERATOR
 ----------------------------------------------------------------------
 Logic:
  1. SCANS for tabs starting with "SPC_"
  2. EXTRACTS HEADER ROWS (1-7) to find Part #, Batch, Date, etc.
  3. CALCULATES Analysis (Stats + Rules + Cpk) OR Attribute Analysis
  4. GENERATES PDF with SECTION BREAKS per Tab.
  
 Version: 4.4.0 (Added Pass/Fail Attribute Support)
"""

# --- IMPORTS ---
import sys
import os
import glob
import time
import re
from datetime import datetime
from pathlib import Path

# --- CRITICAL FIX FOR THREADING ERROR ---
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

# Data Science
import pandas as pd
import numpy as np
from scipy.stats import norm

# Excel
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# PDF Reporting
from fpdf import FPDF

# TUI Imports
import questionary
from rich.console import Console
from rich.panel import Panel
from rich.tree import Tree

console = Console()

# --- CONSTANTS ---
TOOL_VERSION = "4.4.0"
INPUT_PREFIX = "SPC_"        
INSERT_START_ROW = 8   
HEADER_SEARCH_ROWS = 7 

# --- COLORS ---
COLOR_PURPLE_DARK = "7030A0"   
COLOR_PURPLE_LIGHT = "8064A2"  

# --- PDF ENGINE CLASS ---
class WhitePaperPDF(FPDF):
    def __init__(self, filename):
        super().__init__()
        self.report_filename = filename
        self.set_auto_page_break(auto=True, margin=15)

    def header(self):
        self.set_font('Arial', 'B', 14)
        self.set_text_color(50, 50, 50)
        self.cell(0, 10, 'Process Capability Engineering Report', 0, 1, 'L')
        self.set_draw_color(0, 0, 0)
        self.line(10, 20, 200, 20)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'File: {self.report_filename} | Page {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title, subtitle=None):
        self.set_font('Arial', 'B', 12)
        self.set_fill_color(230, 240, 255) # Light Blue
        self.set_text_color(0, 0, 0)
        self.cell(0, 8, f"  {title}", 0, 1, 'L', 1)
        if subtitle:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 6, f"  {subtitle}", 0, 1, 'L')
        self.ln(4)

    def add_section_header(self, sheet_name, metadata):
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_fill_color(112, 48, 160) # Purple
        self.set_text_color(255, 255, 255)
        self.cell(0, 12, f"  SECTION: {sheet_name}", 0, 1, 'L', 1)
        self.ln(5)
        
        if metadata:
            self.set_text_color(0, 0, 0)
            self.set_font('Arial', 'B', 12)
            self.cell(0, 8, "Run Information:", 0, 1, 'L')
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.set_fill_color(240, 240, 240)
            
            col_w_key = 50
            col_w_val = 130
            
            for key, val in metadata.items():
                self.set_font('Arial', 'B', 10)
                self.cell(col_w_key, 8, f"  {key}", 1, 0, 'L', 1)
                self.set_font('Arial', '', 10)
                self.cell(col_w_val, 8, f"  {val}", 1, 1, 'L')
            
            self.ln(10)

    def add_stat_table(self, metrics):
        self.set_font('Arial', 'B', 10)
        self.set_fill_color(240, 240, 240)
        self.set_text_color(0, 0, 0)
        
        col_w = 45
        self.cell(col_w, 7, "Metric", 1, 0, 'C', 1)
        self.cell(col_w, 7, "Value", 1, 1, 'C', 1)
        
        self.set_font('Arial', '', 10)
        for name, val in metrics:
            self.cell(col_w, 7, f"  {name}", 1)
            self.cell(col_w, 7, f"  {val}", 1)
            self.ln()
        self.ln(5)

    def body_text(self, txt):
        self.set_font('Times', '', 11)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 5, txt)
        self.ln()

# --- UTILITIES ---
def get_unique_filepath(filepath):
    if not os.path.exists(filepath): return filepath
    folder, filename = os.path.split(filepath)
    name, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(os.path.join(folder, f"{name}_{counter}{ext}")):
        counter += 1
    return os.path.join(folder, f"{name}_{counter}{ext}")

def wait_for_file_access(filepath):
    if not os.path.exists(filepath): return
    while True:
        try:
            with open(filepath, 'a'): break
        except IOError:
            console.print(f"[yellow]WAITING:[/yellow] File {os.path.basename(filepath)} is open. Close it to continue.")
            time.sleep(2)

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", str(filename))

def extract_sheet_metadata(ws):
    metadata = {}
    target_keys = ["PART", "BATCH", "DATE", "NOTE", "OPERATOR", "MACHINE", "ORDER", "LOT"]
    for row in ws.iter_rows(min_row=1, max_row=HEADER_SEARCH_ROWS, min_col=1, max_col=10):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val_str = str(cell.value).strip().rstrip(':')
                is_target = any(k in val_str.upper() for k in target_keys)
                if is_target:
                    next_col_idx = cell.column + 1
                    neighbor = ws.cell(row=cell.row, column=next_col_idx).value
                    if neighbor:
                        metadata[val_str] = str(neighbor)
    return metadata

def calculate_cpk(data, usl, lsl):
    if len(data) < 2: return 0, 0, 0, 0
    mean = np.mean(data)
    sigma = np.std(data, ddof=1)
    if sigma < 1e-9: sigma = 1e-9
    cpu = (usl - mean) / (3 * sigma)
    cpl = (mean - lsl) / (3 * sigma)
    cpk = min(cpu, cpl)
    cp = (usl - lsl) / (6 * sigma)
    return cp, cpk, mean, sigma

# --- EXCEL FORMATTING UTILS ---
def style_header_cell(cell):
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=COLOR_PURPLE_LIGHT, end_color=COLOR_PURPLE_LIGHT, fill_type="solid")
    cell.alignment = Alignment(horizontal="center")
    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def style_data_cell(cell, is_nominal=False, is_numeric=False, align_left=False):
    cell.alignment = Alignment(horizontal='left' if align_left else 'center')
    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    if is_nominal: cell.font = Font(bold=True)
    if is_numeric: cell.number_format = '0.0000'

def shift_existing_images(ws, insertion_row, count):
    if not hasattr(ws, '_images'): return
    for img in ws._images:
        try:
            if hasattr(img.anchor, '_from'):
                current_row = img.anchor._from.row
                if current_row >= (insertion_row - 1):
                    img.anchor._from.row += count
                    if hasattr(img.anchor, 'to'):
                        img.anchor.to.row += count
        except Exception: pass

# --- LEGEND GENERATOR ---
def write_rule_legend(ws, start_row):
    legend_data = [
        ("WECO Rule 1", "Any single point outside 3σ limit"),
        ("Trend", "6 consecutive points increasing or decreasing"),
        ("Attribute Fail", "Non-Numeric failure (Pass/Fail)")
    ]
    col_start = 10
    ws.column_dimensions['J'].width = 15
    ws.column_dimensions['K'].width = 55
    head_r = ws.cell(start_row, col_start, "RULE REFERENCE")
    head_r.font = Font(bold=True, size=14, color="FFFFFF")
    head_r.fill = PatternFill(start_color=COLOR_PURPLE_DARK, end_color=COLOR_PURPLE_DARK, fill_type="solid")
    r = start_row + 1
    style_header_cell(ws.cell(r, col_start, "Rule / Pattern"))
    style_header_cell(ws.cell(r, col_start+1, "Description"))
    r += 1
    for name, desc in legend_data:
        c1 = ws.cell(r, col_start, name)
        c2 = ws.cell(r, col_start+1, desc)
        style_data_cell(c1, is_nominal=True)
        style_data_cell(c2, align_left=True)
        r += 1

# --- SPC LOGIC ---
def check_spc_rules_full_scan(data, mean, std_dev):
    try:
        n = len(data)
        if n == 0: return "INSUFFICIENT DATA", False, []
        if n == 1: return f"LIMITED DATA (N=1)", False, [] 
        if std_dev == 0: return "NO VARIATION", False, []

        z = (data - mean) / std_dev
        
        r1_indices = np.where(np.abs(z) > 3)[0]
        if len(r1_indices) > 0: return f"WECO Rule 1 (Sample #{r1_indices[0]+1})", True, r1_indices

        if n >= 6:
            for i in range(n - 5):
                window = data[i:i+6]
                diffs = np.diff(window)
                if np.all(diffs > 0) or np.all(diffs < 0):
                    return f"Trend Detected (Samples {i+1}-{i+6})", True, range(i, i+6)

        return "STABLE", False, []
    except:
        return "CALC ERROR", False, []

# --- PLOTTING: NUMERIC ---
def create_bell_curve_plot(data, usl, lsl, nominal, mean, sigma, feature_name, output_dir, prefix):
    plt.figure(figsize=(10, 6))
    effective_sigma = max(sigma, (usl-lsl)/20) 
    spread = max((usl-lsl), (6*effective_sigma)) * 1.5
    x = np.linspace(mean - spread/2, mean + spread/2, 1000)
    y = norm.pdf(x, mean, sigma)
    
    plt.plot(x, y, color='blue', linewidth=2, label=f'Process (σ={sigma:.4f})')
    plt.fill_between(x, y, alpha=0.2, color='blue')
    plt.axvline(lsl, color='red', linestyle='--', linewidth=2, label=f'LSL ({lsl})')
    plt.axvline(usl, color='red', linestyle='--', linewidth=2, label=f'USL ({usl})')
    plt.axvline(nominal, color='black', linestyle=':', linewidth=1, label='Nominal')
    plt.axvline(mean, color='green', linestyle='-', linewidth=1, label=f'Mean ({mean:.4f})')
    plt.title(f'Capability Analysis: {feature_name}', fontsize=12, fontweight='bold')
    plt.xlabel('Measured Value')
    plt.ylabel('Probability Density')
    plt.legend(loc='best')
    plt.grid(True, alpha=0.3)
    
    filename = f"TEMP_BELL_{prefix}_{sanitize_filename(feature_name)}.png"
    save_path = os.path.join(output_dir, filename)
    plt.savefig(save_path, bbox_inches='tight', dpi=100)
    plt.close()
    return save_path

# --- PLOTTING: ATTRIBUTE ---
def create_attribute_bar_plot(pass_count, fail_count, feature_name, output_dir, prefix):
    plt.figure(figsize=(10, 6))
    
    labels = ['PASS', 'FAIL']
    values = [pass_count, fail_count]
    colors = ['green', 'red']
    
    bars = plt.bar(labels, values, color=colors, alpha=0.7, edgecolor='black')
    
    plt.title(f'Attribute Analysis: {feature_name}', fontsize=12, fontweight='bold')
    plt.ylabel('Count')
    plt.grid(axis='y', alpha=0.3)
    
    # Add counts on top of bars
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height,
                 f'{int(height)}', ha='center', va='bottom', fontsize=12, fontweight='bold')

    filename = f"TEMP_BAR_{prefix}_{sanitize_filename(feature_name)}.png"
    save_path = os.path.join(output_dir, filename)
    plt.savefig(save_path, bbox_inches='tight', dpi=100)
    plt.close()
    return save_path

def get_row_index_fuzzy(df, keywords):
    for idx_val in df.index:
        str_val = str(idx_val).upper()
        for k in keywords:
            if k.upper() in str_val:
                return idx_val
    return None

def process_single_file(filepath, output_dir):
    filename = os.path.basename(filepath)
    wait_for_file_access(filepath)
    
    date_str = datetime.now().strftime("%Y%m%d")
    suffix = filename[9:] if filename.startswith("SPC-DATA_") else filename
    target_filename = f"{date_str}_SPC-RESULTS_{suffix}"
    pdf_filename = f"{date_str}_REPORT_{suffix}.pdf"
    
    target_path = output_dir / target_filename
    output_path = get_unique_filepath(str(target_path))
    pdf_path = get_unique_filepath(str(output_dir / pdf_filename))
    
    tab_log = [] 

    try:
        wb_output = load_workbook(filepath)
    except Exception as e:
        return {"critical_error": f"Could not open Excel file: {str(e)}", "logs": []}

    xls = pd.ExcelFile(filepath)
    all_sheet_names = xls.sheet_names
    
    processed_tabs = []
    temp_files = [] 
    
    pdf = WhitePaperPDF(filename)
    pdf.add_page()
    pdf.chapter_title(f"Executive Summary")
    pdf.body_text(f"File: {filename}")
    pdf.body_text(f"Processed on: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    pdf.ln(5)
    
    pdf_summary_data = []

    for sheet_name in all_sheet_names:
        if not sheet_name.startswith(INPUT_PREFIX):
            tab_log.append({'name': sheet_name, 'status': 'SKIP', 'msg': f"Tab name must start with '{INPUT_PREFIX}'"})
            continue

        try:
            ws_meta = wb_output[sheet_name]
            metadata = extract_sheet_metadata(ws_meta)
            
            try:
                df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=0, skiprows=7)
            except Exception as e:
                tab_log.append({'name': sheet_name, 'status': 'ERR', 'msg': f"Pandas Read Error: {str(e)}"})
                continue

            if df_raw.shape[1] > 0 and df_raw.columns[0] not in df_raw.index.names:
                df_raw.set_index(df_raw.columns[0], inplace=True)

            nom_idx = get_row_index_fuzzy(df_raw, ["Nominal", "Nom", "Target"])
            usl_idx = get_row_index_fuzzy(df_raw, ["USL", "Upper", "High"])
            lsl_idx = get_row_index_fuzzy(df_raw, ["LSL", "Lower", "Low"])

            if nom_idx is None:
                tab_log.append({'name': sheet_name, 'status': 'SKIP', 'msg': "Could not find 'Nominal' row"})
                continue
            
            # --- FEATURE LOOP ---
            feature_cols = [c for c in df_raw.columns if "Unnamed" not in str(c)]
            metadata_keys = [nom_idx, usl_idx, lsl_idx]
            features_data = []
            
            for col in feature_cols:
                try:
                    # Raw data scan
                    raw_sample_indices = [x for x in df_raw.index if x not in metadata_keys]
                    raw_values = df_raw.loc[raw_sample_indices, col].tolist()
                    
                    # --- AUTO-DETECT TYPE: ATTRIBUTE VS NUMERIC ---
                    float_conversions = 0
                    valid_items = 0
                    for v in raw_values:
                        if str(v).strip() == "": continue
                        valid_items += 1
                        try:
                            f = float(v)
                            if not np.isnan(f): float_conversions += 1
                        except: pass
                    
                    is_attribute = False
                    if valid_items > 0 and (float_conversions / valid_items) < 0.5:
                        is_attribute = True

                    # --- PATH A: ATTRIBUTE (PASS/FAIL) ---
                    if is_attribute:
                        pass_kw = ["PASS", "OK", "GOOD", "ACCEPT", "ACC"]
                        fail_kw = ["FAIL", "NOK", "BAD", "REJECT", "F"]
                        
                        clean_samples = []
                        pass_count = 0
                        fail_count = 0
                        
                        for val in raw_values:
                            s_val = str(val).upper().strip()
                            if not s_val: continue
                            
                            # 1 = PASS, 0 = FAIL
                            if any(k in s_val for k in pass_kw):
                                clean_samples.append(1)
                                pass_count += 1
                            elif any(k in s_val for k in fail_kw):
                                clean_samples.append(0)
                                fail_count += 1
                        
                        if not clean_samples: continue

                        total = pass_count + fail_count
                        fail_rate = (fail_count / total) * 100 if total > 0 else 0
                        
                        features_data.append({
                            'type': 'attribute',
                            'name': col,
                            'data': np.array(clean_samples),
                            'pass_count': pass_count,
                            'fail_count': fail_count,
                            'fail_rate': fail_rate,
                            'nominal': "N/A", 'usl': "N/A", 'lsl': "N/A"
                        })
                        
                        status = "PASS" if fail_count == 0 else "FAIL"
                        pdf_summary_data.append([sheet_name, col, "N/A", status])

                    # --- PATH B: NUMERIC ---
                    else:
                        nom_val = df_raw.loc[nom_idx, col]
                        if pd.isna(nom_val) or str(nom_val).strip() == "": continue
                        
                        clean_samples = []
                        plot_split_locs = []
                        
                        for i, (label, val) in enumerate(zip(raw_sample_indices, raw_values)):
                            is_split = (str(label).upper().strip() == "SPLIT" or str(val).upper().strip() == "SPLIT")
                            if is_split:
                                plot_split_locs.append(len(clean_samples))
                                continue
                            try:
                                num = float(val)
                                if not np.isnan(num): clean_samples.append(num)
                            except: pass

                        # Tolerance Logic
                        nom_val_float = float(nom_val)
                        # We need USL/LSL for numeric
                        if usl_idx is None or lsl_idx is None: continue
                        
                        raw_usl = float(df_raw.loc[usl_idx, col])
                        raw_lsl = float(df_raw.loc[lsl_idx, col])

                        usl_val = raw_usl
                        lsl_val = raw_lsl

                        if raw_usl < nom_val_float:
                             usl_val = nom_val_float + abs(raw_usl)
                             lsl_val = nom_val_float - abs(raw_lsl)
                        
                        if usl_val == lsl_val: usl_val += 0.0001; lsl_val -= 0.0001

                        cp, cpk, mean_val, sigma_val = calculate_cpk(np.array(clean_samples), usl_val, lsl_val)
                        
                        features_data.append({
                            'type': 'numeric',
                            'name': col, 'nominal': nom_val_float,
                            'usl': usl_val, 'lsl': lsl_val,
                            'data': np.array(clean_samples),
                            'split_locs': plot_split_locs,
                            'cp': cp, 'cpk': cpk, 'mean': mean_val, 'sigma': sigma_val
                        })
                        pdf_summary_data.append([sheet_name, col, f"{cpk:.2f}", "PASS" if cpk >= 1.33 else "FAIL"])

                except: continue
            
            if not features_data: 
                tab_log.append({'name': sheet_name, 'status': 'SKIP', 'msg': "No valid data found"})
                continue

            # --- EXCEL OUTPUT ---
            ws = wb_output[sheet_name]
            processed_tabs.append(sheet_name)
            safe_name = sanitize_filename(sheet_name)[:20]

            ws['A1'] = "SPC ANALYSIS RESULTS"
            ws['A1'].fill = PatternFill(start_color=COLOR_PURPLE_DARK, end_color=COLOR_PURPLE_DARK, fill_type="solid")

            num_features = len(features_data)
            table_rows_needed = 2 + num_features + 1
            chart_rows_needed = num_features 
            total_insert_count = table_rows_needed + chart_rows_needed + 2

            shift_existing_images(ws, INSERT_START_ROW, total_insert_count)
            ws.insert_rows(INSERT_START_ROW, amount=total_insert_count)
            write_rule_legend(ws, 1)

            r = INSERT_START_ROW
            ws.cell(r,1,"ANALYSIS SUMMARY").font = Font(bold=True, size=14, color="FFFFFF")
            ws.cell(r,1).fill = PatternFill(start_color=COLOR_PURPLE_DARK, end_color=COLOR_PURPLE_DARK, fill_type="solid")
            r += 1
            
            headers = ["Feature", "Nominal", "LSL", "USL", "Mean / Rate", "Sigma", "Status", "Notes"]
            for i, h in enumerate(headers, 1): style_header_cell(ws.cell(r, i, h))
            
            f_pass, font_p = PatternFill(start_color="C6EFCE", fill_type="solid"), Font(color="006100", bold=True)
            f_warn, font_w = PatternFill(start_color="FFEB9C", fill_type="solid"), Font(color="9C6500", bold=True)
            f_fail, font_f = PatternFill(start_color="FFC7CE", fill_type="solid"), Font(color="9C0006", bold=True)
            
            r += 1 
            for feat in features_data:
                # Common Cells
                ws.cell(r, 1, str(feat['name'])).border = Border(bottom=Side(style='thin'))
                
                if feat['type'] == 'numeric':
                    # Numeric row writing
                    mean = feat['mean']
                    status, is_unstable, _ = check_spc_rules_full_scan(feat['data'], mean, feat['sigma'])
                    
                    row_vals = [feat['nominal'], feat['lsl'], feat['usl'], mean, feat['sigma'], status, f"Cpk: {feat['cpk']:.2f}"]
                    for c_idx, val in enumerate(row_vals, 2):
                        cell = ws.cell(r, c_idx, val)
                        style_data_cell(cell, is_numeric=(c_idx in [2,3,4,5,6]))
                        # Coloring
                        if c_idx == 7: # Status
                            if is_unstable: cell.fill, cell.font = f_fail, font_f
                            elif "LIMITED" in status: cell.fill, cell.font = f_warn, font_w
                            else: cell.fill, cell.font = f_pass, font_p
                
                else:
                    # Attribute row writing
                    ws.cell(r, 2, "N/A").alignment = Alignment(horizontal='center')
                    ws.cell(r, 3, "N/A").alignment = Alignment(horizontal='center')
                    ws.cell(r, 4, "N/A").alignment = Alignment(horizontal='center')
                    
                    # Col 5: Mean -> Failure Rate
                    c_rate = ws.cell(r, 5, f"{feat['fail_rate']:.1f}% Fail")
                    c_rate.alignment = Alignment(horizontal='center')
                    
                    ws.cell(r, 6, "N/A").alignment = Alignment(horizontal='center') # Sigma
                    
                    # Col 7: Status
                    status_text = "PASS" if feat['fail_count'] == 0 else "FAIL"
                    c_stat = ws.cell(r, 7, status_text)
                    c_stat.alignment = Alignment(horizontal='center')
                    c_stat.font = Font(bold=True)
                    if feat['fail_count'] > 0: c_stat.fill, c_stat.font = f_fail, font_f
                    else: c_stat.fill, c_stat.font = f_pass, font_p

                    # Col 8: Notes
                    ws.cell(r, 8, f"Count: {feat['pass_count']} Pass / {feat['fail_count']} Fail")

                r += 1

            r += 1 
            img_start_row = r
            
            pdf.add_section_header(sheet_name, metadata)

            for i, feat in enumerate(features_data):
                name = str(feat['name'])
                unique_prefix = f"{safe_name}_{i}"
                
                # --- CHART GENERATION ---
                img_path = ""
                
                if feat['type'] == 'numeric':
                    # Numeric PDF
                    pdf.add_page()
                    pdf.chapter_title(f"Feature: {name}", subtitle=f"Sheet: {sheet_name}")
                    bell_path = create_bell_curve_plot(feat['data'], feat['usl'], feat['lsl'], feat['nominal'], 
                                              feat['mean'], feat['sigma'], name, output_dir, f"BELL_{unique_prefix}")
                    temp_files.append(bell_path)
                    pdf.image(bell_path, x=15, w=170)
                    pdf.ln(5)
                    pdf.chapter_title("Statistical Interpretation")
                    status_text = "CAPABLE" if feat['cpk'] >= 1.33 else "NOT CAPABLE"
                    pdf.body_text(f"Process is {status_text} (Cpk {feat['cpk']:.2f}).")
                    pdf.ln(5)
                    metrics = [("Cpk", f"{feat['cpk']:.2f}"), ("Mean", f"{feat['mean']:.4f}"), ("Sigma", f"{feat['sigma']:.4f}")]
                    pdf.add_stat_table(metrics)

                    # Numeric Excel Chart
                    mean = feat['mean']
                    _, is_unstable, highlight_indices = check_spc_rules_full_scan(feat['data'], mean, feat['sigma'])
                    fig, ax = plt.subplots(figsize=(10, 5), dpi=100)
                    x_axis = np.arange(1, len(feat['data']) + 1)
                    
                    ax.axhline(mean, color='green')
                    ax.axhline(feat['usl'], color='red')
                    ax.axhline(feat['lsl'], color='red')
                    ax.plot(x_axis, feat['data'], marker='o', color='black', lw=1, ms=4)
                    
                    if is_unstable:
                        ax.plot(x_axis[highlight_indices], feat['data'][highlight_indices], 'o', color='red', ms=10, mfc='none', mew=2)

                    ax.set_title(f"{name} (Numeric)", fontweight='bold')
                    img_path = os.path.join(output_dir, f"TEMP_CHART_{unique_prefix}.png")
                    plt.savefig(img_path, bbox_inches='tight')
                    plt.close(fig)

                else:
                    # Attribute PDF
                    pdf.add_page()
                    pdf.chapter_title(f"Attribute: {name}", subtitle=f"Sheet: {sheet_name}")
                    bar_path = create_attribute_bar_plot(feat['pass_count'], feat['fail_count'], name, output_dir, f"BAR_{unique_prefix}")
                    temp_files.append(bar_path)
                    pdf.image(bar_path, x=15, w=170)
                    pdf.ln(5)
                    pdf.chapter_title("Attribute Summary")
                    pdf.body_text(f"Failure Rate: {feat['fail_rate']:.1f}%")
                    pdf.body_text(f"Total Samples: {len(feat['data'])}")
                    pdf.ln(5)
                    metrics = [("Pass Count", str(feat['pass_count'])), ("Fail Count", str(feat['fail_count'])), ("Fail Rate", f"{feat['fail_rate']:.1f}%")]
                    pdf.add_stat_table(metrics)

                    # Attribute Excel Chart (Step Chart)
                    fig, ax = plt.subplots(figsize=(10, 4), dpi=100)
                    x_axis = np.arange(1, len(feat['data']) + 1)
                    
                    # Plot 1s and 0s
                    ax.step(x_axis, feat['data'], where='mid', color='blue', lw=2)
                    ax.set_yticks([0, 1])
                    ax.set_yticklabels(['FAIL', 'PASS'])
                    ax.set_ylim(-0.2, 1.2)
                    
                    # Highlight Fails
                    fail_indices = np.where(feat['data'] == 0)[0]
                    if len(fail_indices) > 0:
                        ax.plot(x_axis[fail_indices], feat['data'][fail_indices], 'x', color='red', ms=10, mew=3)

                    ax.set_title(f"{name} (Attribute)", fontweight='bold')
                    img_path = os.path.join(output_dir, f"TEMP_CHART_{unique_prefix}.png")
                    plt.savefig(img_path, bbox_inches='tight')
                    plt.close(fig)

                temp_files.append(img_path)
                img = XLImage(img_path)
                img.width, img.height = 600, 300
                ws.add_image(img, f"B{img_start_row}")
                ws.row_dimensions[img_start_row].height = 225
                ws.cell(img_start_row, 1, name).font = Font(bold=True, size=12)
                img_start_row += 1

            # Styling original data rows (Gray out headers)
            grey_f = PatternFill(start_color="E7E6E6", fill_type="solid")
            for scan_row in range(total_insert_count + 1, ws.max_row + 1):
                val = str(ws.cell(scan_row, 1).value).strip()
                if any(k in val for k in ["Nominal", "USL", "LSL", "Target", "Upper", "Lower"]):
                    for col_idx in range(1, ws.max_column + 1):
                        ws.cell(scan_row, col_idx).fill = grey_f
            
            tab_log.append({'name': sheet_name, 'status': 'OK', 'msg': f"Processed {len(features_data)} features"})

        except Exception as e:
            tab_log.append({'name': sheet_name, 'status': 'ERR', 'msg': str(e)})
            continue

    if processed_tabs:
        wb_output.active = 0
        wb_output.save(output_path)
        pdf.output(pdf_path)
        for f in temp_files:
            try: os.remove(f)
            except: pass
    
    return {"processed": processed_tabs, "logs": tab_log, "ignored": []}

def main():
    os.system('cls' if os.name == 'nt' else 'clear')
    if os.name == 'nt': os.system(f'title SPC Tool v{TOOL_VERSION}')
    console.print(Panel.fit(r"""[bold cyan]SPC GENERATOR[/bold cyan]""", subtitle=f"v{TOOL_VERSION}"))

    cd = os.getcwd()
    files = [f for f in glob.glob(os.path.join(cd, "SPC-DATA_*.xlsx")) if not os.path.basename(f).startswith("~$")]

    if not files:
        console.print("[red]No input files found![/red]")
        questionary.press_any_key_to_continue().ask()
        return

    launcher_options = ["[ ▶ PROCESS ALL FILES ]", "[ ▶ SELECT MULTIPLE... ]", questionary.Separator()] + sorted([os.path.basename(f) for f in files])
    selected_action = questionary.select("Choose Input Action:", choices=launcher_options).ask()
    if not selected_action: return

    if selected_action == "[ ▶ PROCESS ALL FILES ]": files_to_process = files
    elif selected_action == "[ ▶ SELECT MULTIPLE... ]":
        selected = questionary.checkbox("Select:", choices=sorted([os.path.basename(f) for f in files])).ask()
        if not selected: return
        files_to_process = [f for f in files if os.path.basename(f) in selected]
    else: files_to_process = [f for f in files if os.path.basename(f) == selected_action]

    output_root = Path(cd) / "output"
    existing = []
    if output_root.exists(): existing = sorted([d.name for d in output_root.iterdir() if d.is_dir()])
    
    selected_opt = questionary.select("Select Output Folder:", choices=existing + ["< Create New Project >"]).ask()
    if selected_opt == "< Create New Project >":
        project_name = questionary.text("Enter New Project Name:", default="New_Project").ask()
    else: project_name = selected_opt
        
    output_dir = output_root / project_name
    output_dir.mkdir(parents=True, exist_ok=True)

    for fname in files_to_process:
        with console.status(f"Processing {os.path.basename(fname)}..."):
            result = process_single_file(fname, output_dir)
        
        file_tree = Tree(f"[bold cyan]{os.path.basename(fname)}[/bold cyan]")
        if "critical_error" in result: file_tree.add(f"[red]{result['critical_error']}")
        else:
            for log in result['logs']:
                color = "green" if log['status'] == 'OK' else "red"
                file_tree.add(f"[{color}]{log['name']}: {log['msg']}")
        console.print(file_tree)

if __name__ == "__main__":
    main()