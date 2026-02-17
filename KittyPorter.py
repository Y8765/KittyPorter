import pandas as pd
import os
import json
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

RISK_WEIGHTS = {'High': 50, 'Medium': 20, 'Low': 5, 'Passed': 0}
CRITICAL_KEYWORDS = [
    'LSA', 'Credential', 'WDigest', 'LSASS', 'SMB', 'NetBIOS', 'LLMNR', 
    'Signing', 'Spooler', 'Driver', 'Defender', 'Ntlm'
]
CRITICAL_BONUS = 50

def select_files_gui():
    root = tk.Tk()
    root.withdraw()
    print("Please select the Hardening Kitty REPORT CSV...")
    report = filedialog.askopenfilename(title="1. Select REPORT CSV", filetypes=[("CSV", "*.csv")])
    if not report: return None, None
    
    print("Please select the Hardening Kitty TEMPLATE CSV(s) (You can select multiple)...")
    templates = filedialog.askopenfilenames(title="2. Select TEMPLATE CSV(s)", filetypes=[("CSV", "*.csv")])
    return report, templates

def calculate_risk(row):
    if str(row.get('TestResult','')).lower() == 'passed': return 0
    severity = row.get('Severity', 'Low')
    score = RISK_WEIGHTS.get(severity, 5)
    text = f"{row.get('Category','')} {row.get('Description','')} {row.get('Name','')}"
    if any(k.lower() in text.lower() for k in CRITICAL_KEYWORDS): score += CRITICAL_BONUS
    return min(score, 100)

def generate_fix(row):
    """
    Generates the PowerShell command.
    CRITICAL: This MUST retain 'HKLM:' or 'HKCU:' (with colon).
    """
    path = row.get('RegistryPath')
    item = row.get('RegistryItem')
    val = row.get('RecommendedValue')
    if pd.isna(path) or pd.isna(item): return ""
    
    ps_path = str(path).replace('HKEY_LOCAL_MACHINE', 'HKLM:').replace('HKEY_CURRENT_USER', 'HKCU:')
    clean_val = str(val).replace('"', '')
    return f'Set-ItemProperty -Path "{ps_path}" -Name "{item}" -Value "{clean_val}" -Force'

def generate_excel(df, output_path, df_failed, df_passed):
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    wb = writer.book
    
    ws_dash = wb.add_worksheet('Dashboard')

    # --- עיצובים (Formats) ---
    
    # Passed: ירוק בהיר סטנדרטי (רקע בהיר, טקסט ירוק כהה)
    fmt_green = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    
    # Fixed: אותו ירוק בהיר בדיוק, אבל עם טקסט מודגש (Bold) כדי לסמן שזה תוקן
    fmt_fixed = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})
    
    fmt_yellow = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
    fmt_red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    fmt_grey = wb.add_format({'bg_color': '#D9D9D9', 'font_color': '#333333'})

    total_checks = len(df)
    
    preferred_order = [
        'CIS', 'Category', 'Description', 'Method', 'MethodArgument', 
        'RegistryPath', 'RegistryItem', 'Result', 'Recommended', 'Fix', 'Status', 'RiskScore'
    ]
    
    status_options = ['Fixed', 'Not Relevant', 'To Discuss', "Can't Fix/Exclude"]

    def clean_reg_col(path):
        if pd.isna(path): return path
        return str(path).replace('HKLM:', 'HKLM').replace('HKCU:', 'HKCU')

    # --- Action Items Sheet ---
    avail_cols = [c for c in preferred_order if c in df_failed.columns or c == 'Status']
    df_failed_export = df_failed[[c for c in avail_cols if c != 'Status']].copy()
    
    if 'RegistryPath' in df_failed_export.columns:
        df_failed_export['RegistryPath'] = df_failed_export['RegistryPath'].apply(clean_reg_col)

    df_failed_export['Status'] = '' 
    df_failed_export = df_failed_export[avail_cols] 
    
    df_failed_export.to_excel(writer, sheet_name='Action Items', index=False)
    ws_fail = writer.sheets['Action Items']
    ws_fail.set_tab_color('#C00000')
    ws_fail.freeze_panes(1, 0)
    
    fail_cols = df_failed_export.columns.tolist()
    status_idx_fail = fail_cols.index('Status')
    status_char_fail = xl_col_to_name(status_idx_fail) 
    cat_idx_fail = fail_cols.index('Category')
    cat_char_fail = xl_col_to_name(cat_idx_fail)        

    max_row_fail = len(df_failed_export) + 1
    max_col_fail = len(fail_cols) - 1
    last_col_char_fail = xl_col_to_name(max_col_fail)
    
    if len(df_failed_export) > 0:
        ws_fail.add_table(0, 0, max_row_fail-1, max_col_fail, {
            'columns': [{'header': c} for c in fail_cols],
            'style': 'TableStyleMedium9'
        })

    ws_fail.data_validation(1, status_idx_fail, max_row_fail-1, status_idx_fail, {'validate': 'list', 'source': status_options})
    
    full_rng_fail = f"A2:{last_col_char_fail}{max_row_fail}"
    
    # החלת העיצוב המתוקן (Fixed = ירוק בהיר מודגש)
    ws_fail.conditional_format(full_rng_fail, {'type': 'formula', 'criteria': f'=${status_char_fail}2="Fixed"', 'format': fmt_fixed})
    ws_fail.conditional_format(full_rng_fail, {'type': 'formula', 'criteria': f'=${status_char_fail}2="Not Relevant"', 'format': fmt_grey})
    ws_fail.conditional_format(full_rng_fail, {'type': 'formula', 'criteria': f'=${status_char_fail}2="To Discuss"', 'format': fmt_yellow})
    ws_fail.conditional_format(full_rng_fail, {'type': 'formula', 'criteria': f'=${status_char_fail}2="Can\'t Fix/Exclude"', 'format': fmt_red})

    ws_fail.set_column(f'{status_char_fail}:{status_char_fail}', 18) 
    ws_fail.set_column(f'{cat_char_fail}:{cat_char_fail}', 25) 
    
    # --- Passed Checks Sheet ---
    df_passed_export = df_passed[[c for c in avail_cols if c != 'Status']].copy()
    if 'RegistryPath' in df_passed_export.columns:
        df_passed_export['RegistryPath'] = df_passed_export['RegistryPath'].apply(clean_reg_col)
        
    df_passed_export['Status'] = 'Passed'
    df_passed_export = df_passed_export[avail_cols]
    
    df_passed_export.to_excel(writer, sheet_name='Passed Checks', index=False)
    ws_pass = writer.sheets['Passed Checks']
    ws_pass.set_tab_color('#00B050')
    ws_pass.freeze_panes(1, 0)
    
    pass_cols = df_passed_export.columns.tolist()
    status_idx_pass = pass_cols.index('Status')
    status_char_pass = xl_col_to_name(status_idx_pass)
    cat_idx_pass = pass_cols.index('Category')
    cat_char_pass = xl_col_to_name(cat_idx_pass)
    
    max_row_pass = len(df_passed_export) + 1
    max_col_pass = len(pass_cols) - 1
    last_col_char_pass = xl_col_to_name(max_col_pass)

    if len(df_passed_export) > 0:
        ws_pass.add_table(0, 0, len(df_passed_export), len(avail_cols), {
            'columns': [{'header': c} for c in pass_cols], 
            'style': 'TableStyleLight9'
        })

    ws_pass.data_validation(1, status_idx_pass, max_row_pass-1, status_idx_pass, {'validate': 'list', 'source': status_options})
    
    full_rng_pass = f"A2:{last_col_char_pass}{max_row_pass}"
    
    ws_pass.conditional_format(full_rng_pass, {'type': 'formula', 'criteria': f'=${status_char_pass}2="Passed"', 'format': fmt_green})
    # Fixed = ירוק בהיר מודגש
    ws_pass.conditional_format(full_rng_pass, {'type': 'formula', 'criteria': f'=${status_char_pass}2="Fixed"', 'format': fmt_fixed})
    ws_pass.conditional_format(full_rng_pass, {'type': 'formula', 'criteria': f'=${status_char_pass}2="Not Relevant"', 'format': fmt_grey})
    ws_pass.conditional_format(full_rng_pass, {'type': 'formula', 'criteria': f'=${status_char_pass}2="To Discuss"', 'format': fmt_yellow})
    ws_pass.conditional_format(full_rng_pass, {'type': 'formula', 'criteria': f'=${status_char_pass}2="Can\'t Fix/Exclude"', 'format': fmt_red})
    
    ws_pass.set_column(f'{status_char_pass}:{status_char_pass}', 18)

    # --- Notes Sheet ---
    ws_notes = wb.add_worksheet('Notes')
    ws_notes.set_tab_color('#FFC000') 
    
    notes_headers = ['Date', 'Author', 'Category/Control', 'Note/Comment']
    ws_notes.write_row('A1', notes_headers, wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white'}))
    ws_notes.set_column('A:A', 15)
    ws_notes.set_column('B:B', 20)
    ws_notes.set_column('C:C', 30)
    ws_notes.set_column('D:D', 60)
    ws_notes.add_table('A1:D20', {'columns': [{'header': c} for c in notes_headers], 'style': 'TableStyleMedium2'})

    # --- Stats Logic (Hidden Sheet) ---
    ws_stats = wb.add_worksheet('Stats')
    ws_stats.hide()
    
    pivot = df.pivot_table(index='Category', columns='TestResult', aggfunc='size', fill_value=0)
    if 'Passed' not in pivot.columns: pivot['Passed'] = 0
    if 'Failed' not in pivot.columns: pivot['Failed'] = 0
    
    pivot = pivot.sort_values('Failed', ascending=True)
    categories = pivot.index.tolist()
    init_pass_vals = pivot['Passed'].tolist()
    init_fail_vals = pivot['Failed'].tolist()
    
    ws_stats.write_row('A1', ['Overall Status', 'Count'])
    ws_stats.write_row('D1', ['Category', 'Live Failed', 'Live Discuss', 'Live Not Relevant', 'Live Fixed', 'Live Passed', 'Init Pass', 'Init Fail', 'Total', '% Compliant', 'Chart Label'])
    
    for i, cat in enumerate(categories):
        row = i + 1
        ws_stats.write(row, 3, cat) 
        ws_stats.write(row, 9, init_pass_vals[i]) # J
        ws_stats.write(row, 10, init_fail_vals[i]) # K
        
        # Live Not Relevant
        f_nr = f'=COUNTIFS(\'Action Items\'!{cat_char_fail}:{cat_char_fail}, D{row+1}, \'Action Items\'!{status_char_fail}:{status_char_fail}, "Not Relevant") + COUNTIFS(\'Passed Checks\'!{cat_char_pass}:{cat_char_pass}, D{row+1}, \'Passed Checks\'!{status_char_pass}:{status_char_pass}, "Not Relevant")'
        ws_stats.write_formula(row, 6, f_nr)

        # Live To Discuss
        f_disc = f'=COUNTIFS(\'Action Items\'!{cat_char_fail}:{cat_char_fail}, D{row+1}, \'Action Items\'!{status_char_fail}:{status_char_fail}, "To Discuss") + COUNTIFS(\'Passed Checks\'!{cat_char_pass}:{cat_char_pass}, D{row+1}, \'Passed Checks\'!{status_char_pass}:{status_char_pass}, "To Discuss")'
        ws_stats.write_formula(row, 5, f_disc)

        # Live Fixed (הפרדה לסטטוס עצמאי)
        f_fix = f'=COUNTIFS(\'Action Items\'!{cat_char_fail}:{cat_char_fail}, D{row+1}, \'Action Items\'!{status_char_fail}:{status_char_fail}, "Fixed") + COUNTIFS(\'Passed Checks\'!{cat_char_pass}:{cat_char_pass}, D{row+1}, \'Passed Checks\'!{status_char_pass}:{status_char_pass}, "Fixed")'
        ws_stats.write_formula(row, 7, f_fix)

        # Live Passed (מנכה את ה-Fixed כי הם נספרים בנפרד)
        f_pass = f'=J{row+1} - COUNTIFS(\'Passed Checks\'!{cat_char_pass}:{cat_char_pass}, D{row+1}, \'Passed Checks\'!{status_char_pass}:{status_char_pass}, "To Discuss") - COUNTIFS(\'Passed Checks\'!{cat_char_pass}:{cat_char_pass}, D{row+1}, \'Passed Checks\'!{status_char_pass}:{status_char_pass}, "Not Relevant") - COUNTIFS(\'Passed Checks\'!{cat_char_pass}:{cat_char_pass}, D{row+1}, \'Passed Checks\'!{status_char_pass}:{status_char_pass}, "Fixed")'
        ws_stats.write_formula(row, 8, f_pass)
        
        # Live Failed
        f_fail = f'=K{row+1} - COUNTIFS(\'Action Items\'!{cat_char_fail}:{cat_char_fail}, D{row+1}, \'Action Items\'!{status_char_fail}:{status_char_fail}, "Fixed") - COUNTIFS(\'Action Items\'!{cat_char_fail}:{cat_char_fail}, D{row+1}, \'Action Items\'!{status_char_fail}:{status_char_fail}, "To Discuss") - COUNTIFS(\'Action Items\'!{cat_char_fail}:{cat_char_fail}, D{row+1}, \'Action Items\'!{status_char_fail}:{status_char_fail}, "Not Relevant")'
        ws_stats.write_formula(row, 4, f_fail)
        
        # Total
        ws_stats.write_formula(row, 11, f'=E{row+1}+F{row+1}+G{row+1}+H{row+1}+I{row+1}')
        
        # % Compliant (Fixed + Passed) / (Total - Not Relevant)
        ws_stats.write_formula(row, 12, f'=IF((L{row+1}-G{row+1})=0, 0, (H{row+1}+I{row+1})/(L{row+1}-G{row+1}))')
        
        # Chart Label
        ws_stats.write_formula(row, 13, f'=D{row+1} & " (" & TEXT(M{row+1}, "0%") & ")"')

    # KPI Calculation
    count_pass_pc = f"COUNTIF('Passed Checks'!{status_char_pass}:{status_char_pass}, \"Passed\")"
    count_fix_ai  = f"COUNTIF('Action Items'!{status_char_fail}:{status_char_fail}, \"Fixed\")"
    count_fix_pc  = f"COUNTIF('Passed Checks'!{status_char_pass}:{status_char_pass}, \"Fixed\")"
    
    ws_stats.write('A2', 'Passed')
    ws_stats.write_formula('B2', f"={count_pass_pc}")
    
    ws_stats.write('A3', 'Fixed')
    ws_stats.write_formula('B3', f"={count_fix_ai} + {count_fix_pc}")

    count_disc_ai = f"COUNTIF('Action Items'!{status_char_fail}:{status_char_fail}, \"To Discuss\")"
    count_disc_pc = f"COUNTIF('Passed Checks'!{status_char_pass}:{status_char_pass}, \"To Discuss\")"
    ws_stats.write('A4', 'To Discuss') 
    ws_stats.write_formula('B4', f"={count_disc_ai} + {count_disc_pc}")
    
    count_nr_ai = f"COUNTIF('Action Items'!{status_char_fail}:{status_char_fail}, \"Not Relevant\")"
    count_nr_pc = f"COUNTIF('Passed Checks'!{status_char_pass}:{status_char_pass}, \"Not Relevant\")"
    ws_stats.write('A5', 'Not Relevant')
    ws_stats.write_formula('B5', f"={count_nr_ai} + {count_nr_pc}")
    
    ws_stats.write('A6', 'Failed')
    ws_stats.write_formula('B6', f"={total_checks} - B2 - B3 - B4 - B5")

    # --- Dashboard ---
    ws_dash.hide_gridlines(2)
    ws_dash.set_column('B:H', 20)
    
    fmt_title = wb.add_format({'bold': True, 'font_size': 24, 'font_color': '#203764'})
    fmt_head = wb.add_format({'bold': True, 'font_size': 12, 'color': 'white', 'bg_color': '#4472C4', 'align': 'center', 'border': 1})
    fmt_val = wb.add_format({'bold': True, 'font_size': 22, 'align': 'center', 'bg_color': '#f2f2f2', 'border': 1})
    fmt_pct = wb.add_format({'bold': True, 'font_size': 22, 'align': 'center', 'bg_color': '#f2f2f2', 'border': 1, 'num_format': '0.0%'})
    
    ws_dash.write('B2', 'Security Assessment Dashboard', fmt_title)
    
    ws_dash.write('B5', "Compliance Score", fmt_head)
    ws_dash.write('C5', "Total Controls", fmt_head)
    ws_dash.write('D5', "Passed Checks", fmt_head)
    ws_dash.write('E5', "Fixed Checks", fmt_head)
    ws_dash.write('F5', "Failed Checks", fmt_head)
    ws_dash.write('G5', "To Discuss", fmt_head)
    ws_dash.write('H5', "Not Relevant", fmt_head)

    # Values
    ws_dash.merge_range('B6:B7', '', fmt_pct)
    ws_dash.write_formula('B6', f"=(Stats!B2 + Stats!B3) / ({total_checks} - Stats!B5)", fmt_pct)
    
    ws_dash.merge_range('C6:C7', total_checks, fmt_val)
    
    ws_dash.merge_range('D6:D7', '', fmt_val)
    ws_dash.write_formula('D6', "=Stats!B2", fmt_val) # Passed
    
    ws_dash.merge_range('E6:E7', '', fmt_val)
    ws_dash.write_formula('E6', "=Stats!B3", fmt_val) # Fixed
    
    ws_dash.merge_range('F6:F7', '', fmt_val)
    ws_dash.write_formula('F6', "=Stats!B6", fmt_val) # Failed
    
    ws_dash.merge_range('G6:G7', '', fmt_val)
    ws_dash.write_formula('G6', "=Stats!B4", fmt_val) # Discuss
    
    ws_dash.merge_range('H6:H7', '', fmt_val)
    ws_dash.write_formula('H6', "=Stats!B5", fmt_val) # Not Relevant

    # --- Charts ---

    # 1. Doughnut Chart
    chart1 = wb.add_chart({'type': 'doughnut'})
    chart1.add_series({
        'categories': '=Stats!$A$2:$A$6', 
        'values':     '=Stats!$B$2:$B$6', 
        'points':     [
            {'fill': {'color': '#00B050'}}, # Passed (Green)
            {'fill': {'color': '#92D050'}}, # Fixed (Lime Green - בהיר ונעים)
            {'fill': {'color': '#FFC000'}}, # Discuss
            {'fill': {'color': '#D9D9D9'}}, # NR
            {'fill': {'color': '#C00000'}}  # Failed
        ]
    })
    chart1.set_title({'name': 'Status Overview (Live)'})
    ws_dash.insert_chart('B10', chart1)

    # 2. Stacked Bar Chart
    num_cats = len(categories)
    chart2 = wb.add_chart({'type': 'bar', 'subtype': 'stacked'})
    
    chart2.add_series({'name': 'Failed', 'categories': f'=Stats!$N$2:$N${num_cats+1}', 'values': f'=Stats!$E$2:$E${num_cats+1}', 'fill': {'color': '#C00000'}})
    chart2.add_series({'name': 'To Discuss', 'categories': f'=Stats!$N$2:$N${num_cats+1}', 'values': f'=Stats!$F$2:$F${num_cats+1}', 'fill': {'color': '#FFC000'}})
    chart2.add_series({'name': 'Not Relevant', 'categories': f'=Stats!$N$2:$N${num_cats+1}', 'values': f'=Stats!$G$2:$G${num_cats+1}', 'fill': {'color': '#D9D9D9'}})
    
    # הסדרה של Fixed בצבע Lime Green (בהיר)
    chart2.add_series({'name': 'Fixed', 'categories': f'=Stats!$N$2:$N${num_cats+1}', 'values': f'=Stats!$H$2:$H${num_cats+1}', 'fill': {'color': '#92D050'}})
    
    # הסדרה של Passed בירוק רגיל
    chart2.add_series({'name': 'Passed', 'categories': f'=Stats!$N$2:$N${num_cats+1}', 'values': f'=Stats!$I$2:$I${num_cats+1}', 'fill': {'color': '#00B050'}})
    
    chart2.set_title({'name': 'Category Analysis & Compliance % (Live)'})
    chart2.set_size({'width': 850, 'height': 500})
    chart2.set_x_axis({'name': 'Count of Checks'})
    chart2.set_y_axis({'reverse': True})
    ws_dash.insert_chart('F10', chart2)

    writer.close()
    print(f"✅ Excel Created with Notes & Clean Registry Paths: {output_path}")

def generate_html(df, output_path, score, total, passed, failed, categories):
    df_failed = df[df['TestResult'].str.contains('Failed', na=False)].sort_values(by=['RiskScore'], ascending=False)
    df_passed = df[df['TestResult'].str.contains('Passed', na=False)].sort_values(by=['Category'])
    sorted_cats = sorted([str(c) for c in categories if str(c) != 'nan'])

    def render_rows(dframe, is_action=True):
        rows = []
        for _, row in dframe.iterrows():
            risk_class = "risk-low"
            if is_action:
                if row['RiskScore'] == 100: risk_class = "risk-100"
                elif row['RiskScore'] >= 60: risk_class = "risk-high"
                elif row['RiskScore'] >= 40: risk_class = "risk-med"
            
            fix = str(row.get('Fix', '')).replace('"', '&quot;')
            desc = str(row.get('Description', row.get('Name', ''))).replace('"', '&quot;')
            cis = str(row.get('CIS', row.get('ID', ''))).replace('nan', '')
            
            curr_raw = str(row.get('Result', ''))
            exp_raw = str(row.get('Recommended', row.get('RecommendedValue', '')))
            
            curr = f'<div class="code-box">{curr_raw}</div>'
            exp = f'<div class="code-box">{exp_raw}</div>'
            
            reg_path = str(row.get('RegistryPath', ''))
            reg_item = str(row.get('RegistryItem', ''))
            reg_html = ""
            reg_btns = ""
            
            if reg_path and reg_path != "nan":
                # --- VISUAL FIX: Clean display path ---
                clean_short_path = reg_path.replace('HKLM:', 'HKLM').replace('HKCU:', 'HKCU')
                # Ensure full names are shortened if present
                clean_short_path = clean_short_path.replace('HKEY_LOCAL_MACHINE', 'HKLM').replace('HKEY_CURRENT_USER', 'HKCU')
                
                # --- FUNCTIONAL FIX: Build full path for Regedit Address Bar ---
                reg_addr = clean_short_path.replace('HKLM', r'Computer\HKEY_LOCAL_MACHINE')\
                                           .replace('HKCU', r'Computer\HKEY_CURRENT_USER')\
                                           .replace('HKCR', r'Computer\HKEY_CLASSES_ROOT')\
                                           .replace('HKU', r'Computer\HKEY_USERS')
                
                js_reg_addr = reg_addr.replace(chr(92), chr(92)+chr(92))

                reg_html = f"<div class='reg-box'><div><b>Key:</b> {clean_short_path}</div><div><b>Val:</b> {reg_item}</div></div>"
                
                # Only Copy Path Button Remains
                btn1 = f'''<button class="btn-icon" onclick="copyPathOnly('{js_reg_addr}')" title="Copy path to clipboard">📄 Copy Path</button>'''
                reg_btns = f"<div class='btn-group'>{btn1}</div>"

            data_attrs = f'''
                data-id="{row['ID']}" 
                data-cat="{row['Category']}" 
                data-desc="{desc}" 
                data-reg="{str(reg_html).replace('"', '&quot;')}"
                data-regbtns="{str(reg_btns).replace('"', '&quot;')}"
                data-curr="{str(curr).replace('"', '&quot;')}" 
                data-exp="{str(exp).replace('"', '&quot;')}" 
                data-fix="{fix}" 
                data-score="{row['RiskScore']}"
                data-cis="{cis}"
            '''

            rows.append(f'''
            <tr {data_attrs}>
                <td>{'<input type="checkbox" class="select-row">' if is_action else ''}</td>
                <td><span class="badge {risk_class if is_action else 'OK'}">{row['RiskScore'] if is_action else 'OK'}</span></td>
                <td style="font-weight:bold; white-space:nowrap;">{cis}</td>
                <td>{row['Category']}</td>
                <td><div class="desc-text">{desc}</div></td>
                <td>{reg_html} {reg_btns}</td>
                <td>{curr}</td>
                <td>{exp}</td>
            </tr>''')
        return "\n".join(rows)

    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Hardening App</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/fixedheader/3.3.2/css/fixedHeader.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/colreorder/1.6.2/css/colReorder.dataTables.min.css">
    
    <style>
        :root {{ --bg: #f4f6f8; --card: #fff; --text: #212b36; --r100: #7A0916; --rh: #D93025; --rm: #F9A825; --rl: #1877F2; --ok: #00AB55; }}
        body.dark {{ --bg: #121212; --card: #1E1E1E; --text: #E0E0E0; }}
        body.dark table {{ color: #E0E0E0 !important; }}
        body.dark thead th {{ background: #333 !important; color: white !important; }}
        body.dark tbody tr {{ background: #1E1E1E !important; }}
        body.dark tbody tr:hover {{ background: #2C2C2C !important; }}
        body.dark .reg-box, body.dark .code-box {{ background: #2c2c2c; border-color: #444; color: #ccc; }}
        body.dark select, body.dark input, body.dark .dataTables_filter input, body.dark .dataTables_length select {{ background: #333 !important; color: white !important; border: 1px solid #555 !important; }}
        body.dark .dataTables_wrapper {{ color: #E0E0E0 !important; }}
        body.dark .modal-content {{ background: #333 !important; color: #fff !important; }}
        body.dark #custom-modal {{ background-color: #2c2c2c !important; border-color: #444; color: #f0f0f0 !important; }}
        body.dark .modal-header {{ background-color: #1a1a1a; }}
        body {{ font-family: 'Inter', sans-serif; background: var(--bg); color: var(--text); padding: 20px; }}
        .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }}
        .btn {{ padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; margin-left: 5px; }}
        .btn-primary {{ background: #212b36; color: white; }}
        .btn-outline {{ border: 1px solid #919eab; background: transparent; color: var(--text); }}
        .kpi {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px; }}
        .card {{ background: var(--card); padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
        .val {{ font-size: 28px; font-weight: bold; }}
        .val-bad {{ color: var(--rh); }} .val-good {{ color: var(--ok); }}
        .tabs {{ display: flex; gap: 10px; border-bottom: 1px solid #ccc; margin-bottom: 15px; }}
        .tab {{ padding: 10px 20px; cursor: pointer; opacity: 0.6; font-weight: 600; }}
        .tab.active {{ border-bottom: 3px solid var(--text); opacity: 1; }}
        .tab-content {{ display: none; }} .tab-content.active {{ display: block; }}
        .badge {{ padding: 4px 8px; border-radius: 4px; color: white; font-size: 12px; font-weight: bold; }}
        .risk-100 {{ background: var(--r100); }} .risk-high {{ background: var(--rh); }} .risk-med {{ background: var(--rm); }} 
        .risk-low {{ background: var(--rl); }} .OK {{ background: var(--ok); }}
        .btn-group {{ display: flex; gap: 5px; margin-top: 5px; }}
        .btn-icon {{ flex: 1; background: #eee; border: 1px solid #ccc; padding: 4px 6px; border-radius: 4px; cursor: pointer; font-size: 11px; color: #333; }}
        .btn-restore {{ background: var(--rl); color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; }}
        .reg-box, .code-box {{ font-family: 'JetBrains Mono', monospace; font-size: 11px; background: #f8f9fa; padding: 6px; border-radius: 4px; border: 1px solid #ddd; word-break: break-all; max-height: 100px; overflow-y: auto; }}
        .controls {{ display: flex; justify-content: space-between; margin-bottom: 10px; }}
        .desc-text {{ font-size: 14px; line-height: 1.5; }}
        table.dataTable {{ width: 100% !important; border-collapse: collapse; table-layout: fixed; }}
        table.dataTable tbody td {{ white-space: normal !important; word-wrap: break-word; vertical-align: top; }}
        table.dataTable thead th:nth-child(1) {{ width: 2%; }}
        table.dataTable thead th:nth-child(2) {{ width: 4%; }}
        table.dataTable thead th:nth-child(3) {{ width: 6%; }}
        table.dataTable thead th:nth-child(4) {{ width: 10%; }}
        table.dataTable thead th:nth-child(5) {{ width: 28%; }}
        table.dataTable thead th:nth-child(6) {{ width: 30%; }}
        table.dataTable thead th:nth-child(7) {{ width: 10%; }}
        table.dataTable thead th:nth-child(8) {{ width: 10%; }}
        .resizer {{ position: absolute; top: 0; right: 0; bottom: 0; width: 5px; cursor: col-resize; user-select: none; background: transparent; z-index: 100; }}
        .resizer:hover {{ background: #007bff; opacity: 0.5; }}
        #custom-modal {{ display: none; position: fixed; z-index: 1000; left: 50%; top: 30%; width: 450px; background-color: #fff; box-shadow: 0 4px 8px rgba(0,0,0,0.2); border-radius: 8px; border: 1px solid #ccc; }}
        .modal-header {{ padding: 10px 15px; cursor: move; background-color: #212b36; color: white; border-top-left-radius: 8px; border-top-right-radius: 8px; font-weight: bold; display: flex; justify-content: space-between; }}
        .modal-body {{ padding: 20px; font-size: 14px; line-height: 1.5; color: var(--text); }}
        .modal-close {{ cursor: pointer; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="header" style="border-bottom: 2px solid rgba(0,0,0,0.05); padding-bottom: 15px; margin-bottom: 25px;">
        <h1 style="font-weight: 800; letter-spacing: -1px; background: linear-gradient(90deg, #212b36, #4472C4); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin: 0;">KittyPorter - Make Hardening Kitty Reports Great Again</h1>
        <div style="display: flex; align-items: center; gap: 8px;">
            <a href="https://downloads.cisecurity.org/#/" target="_blank" class="btn btn-outline" style="text-decoration: none; display: inline-flex; align-items: center; height: 38px; box-sizing: border-box; font-size: 13px; transition: all 0.2s;">📚 Download CIS Guides</a>
            <input type="file" id="load-file" style="display:none" accept=".json" onchange="loadFromFile(this)">
            <button class="btn btn-outline" style="height: 38px; display: inline-flex; align-items: center; font-size: 13px;" onclick="document.getElementById('load-file').click()">📂 Load Progress</button>
            <button class="btn btn-outline" style="height: 38px; display: inline-flex; align-items: center; font-size: 13px;" onclick="document.body.classList.toggle('dark')">🌗 Dark Mode</button>
            <button class="btn btn-primary" style="height: 38px; display: inline-flex; align-items: center; font-size: 13px;" onclick="saveFile()">💾 Save & Close</button>
        </div>
    </div>

    <div class="kpi">
        <div class="card"><div>Score</div><div class="val {'val-good' if score > 80 else 'val-bad'}" id="score-val">{score:.1f}%</div></div>
        <div class="card"><div>Pending</div><div class="val val-bad" id="cnt-p">{failed}</div></div>
        <div class="card"><div>Fixed</div><div class="val val-good" id="cnt-f">0</div></div>
        <div class="card"><div>Passed</div><div class="val val-good">{passed}</div></div>
    </div>

    <div class="tabs">
        <div class="tab active" onclick="tab('pending')">Pending</div>
        <div class="tab" onclick="tab('fixed')">Fixed ✅</div>
        <div class="tab" onclick="tab('passed')">Passed</div>
    </div>

    <div id="pending" class="tab-content active">
        <div class="controls">
            <button class="btn btn-primary" onclick="moveMarked()">✅ Move Marked to Fixed</button>
            <select onchange="applyFilter(this.value)"><option value="">All Categories</option>{''.join(f'<option value="{c}">{c}</option>' for c in sorted_cats)}</select>
        </div>
        <table id="tp" class="display">
            <thead><tr><th><input type="checkbox" class="select-all"></th><th>Score</th><th>CIS</th><th>Category</th><th>Description</th><th>Registry Info</th><th>Current Value</th><th>Expected Value</th></tr></thead>
            <tbody>{render_rows(df_failed, True)}</tbody>
        </table>
    </div>

    <div id="fixed" class="tab-content">
        <div class="controls"><div></div><select onchange="applyFilter(this.value)"><option value="">All Categories</option>{''.join(f'<option value="{c}">{c}</option>' for c in sorted_cats)}</select></div>
        <table id="tf" class="display">
            <thead><tr><th>Restore</th><th>Score</th><th>CIS</th><th>Category</th><th>Description</th><th>Registry Info</th><th>Current Value</th><th>Expected Value</th></tr></thead>
            <tbody></tbody>
        </table>
    </div>

    <div id="passed" class="tab-content">
        <div class="controls"><div></div><select onchange="applyFilter(this.value)"><option value="">All Categories</option>{''.join(f'<option value="{c}">{c}</option>' for c in sorted_cats)}</select></div>
        <table id="tpass" class="display">
            <thead><tr><th>-</th><th>Status</th><th>CIS</th><th>Category</th><th>Description</th><th>Registry Info</th><th>Current Value</th><th>Expected Value</th></tr></thead>
            <tbody>{render_rows(df_passed, False)}</tbody>
        </table>
    </div>

    <div id="custom-modal">
        <div class="modal-header" id="modal-drag-area">
            <span id="modal-title">Info</span>
            <span class="modal-close" onclick="closeModal()">✕</span>
        </div>
        <div class="modal-body" id="modal-msg"></div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/fixedheader/3.3.2/js/dataTables.fixedHeader.min.js"></script>
    <script src="https://cdn.datatables.net/colreorder/1.6.2/js/dataTables.colReorder.min.js"></script>
    
    <script>
        let tp, tf, tpass, fileHandle;
        
        $(document).ready(function() {{
            jQuery.extend( jQuery.fn.dataTable.ext.oSort, {{
                "cis-sort-asc": function ( a, b ) {{ return a.toString().localeCompare(b.toString(), undefined, {{ numeric: true, sensitivity: 'base' }}); }},
                "cis-sort-desc": function ( a, b ) {{ return b.toString().localeCompare(a.toString(), undefined, {{ numeric: true, sensitivity: 'base' }}); }}
            }});

            const cfg = {{ 
                "pageLength": -1, "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]], "order": [[ 2, "asc" ]], 
                "fixedHeader": true, "colReorder": true, "autoWidth": false, "columnDefs": [ {{ "type": "cis-sort", "targets": 2 }} ],
                "initComplete": function() {{ initCustomResize(this); }}
            }};
            
            tp = $('#tp').DataTable(cfg);
            tf = $('#tf').DataTable(cfg);
            tpass = $('#tpass').DataTable(cfg);
            
            setTimeout(loadFromStorage, 300);
            setupDraggableModal();
            updateScore(); // Initial Calc

            $('#tp thead').on('click', '.select-all', function() {{
                const isChecked = this.checked;
                $('input.select-row', tp.rows({{ 'search': 'applied' }}).nodes()).prop('checked', isChecked);
            }});
            $('#tp tbody').on('change', 'input.select-row', function(){{
                if(!this.checked) {{ $('.select-all').prop('checked', false); }}
            }});
        }});

        function initCustomResize(dt) {{
            const table = $(dt.api().table().node());
            const headers = table.find('thead th');
            headers.each(function() {{
                const th = $(this);
                if (th.find('.resizer').length === 0) {{
                    const resizer = $('<div class="resizer"></div>');
                    th.append(resizer);
                    createResizableColumn(th[0], resizer[0]);
                }}
            }});
        }}

        function createResizableColumn(th, resizer) {{
            let x = 0; let w = 0;
            const mouseDownHandler = function(e) {{
                e.stopPropagation(); x = e.clientX;
                const styles = window.getComputedStyle(th);
                w = parseInt(styles.width, 10);
                document.addEventListener('mousemove', mouseMoveHandler);
                document.addEventListener('mouseup', mouseUpHandler);
                resizer.classList.add('resizing');
            }};
            const mouseMoveHandler = function(e) {{ const dx = e.clientX - x; th.style.width = `${{w + dx}}px`; }};
            const mouseUpHandler = function() {{
                document.removeEventListener('mousemove', mouseMoveHandler);
                document.removeEventListener('mouseup', mouseUpHandler);
                resizer.classList.remove('resizing');
                $(window).trigger('resize'); 
            }};
            resizer.addEventListener('mousedown', mouseDownHandler);
        }}

        function tab(id) {{
            $('.tab-content').removeClass('active'); $('#'+id).addClass('active');
            $('.tab').removeClass('active'); $(event.target).addClass('active');
            tp.columns.adjust().draw(); tf.columns.adjust().draw(); tpass.columns.adjust().draw();
        }}
        
        function applyFilter(v) {{ tp.column(3).search(v).draw(); tf.column(3).search(v).draw(); tpass.column(3).search(v).draw(); }}

        function copyPathOnly(path) {{
            navigator.clipboard.writeText(path).then(() => {{
                showModal("Path Copied!", `The registry path has been copied to your clipboard.<br><br><b>Instructions:</b><br>1. Open Registry Editor (Regedit).<br>2. Paste the path into the address bar at the top.<br>3. Press Enter.`);
            }});
        }}

        function showModal(title, content) {{ $('#modal-title').text(title); $('#modal-msg').html(content); $('#custom-modal').fadeIn(200); }}
        function closeModal() {{ $('#custom-modal').fadeOut(200); }}

        function setupDraggableModal() {{
            const el = document.getElementById("custom-modal");
            const header = document.getElementById("modal-drag-area");
            let isDragging = false, offsetX, offsetY;
            header.onmousedown = function(e) {{
                isDragging = true; offsetX = e.clientX - el.offsetLeft; offsetY = e.clientY - el.offsetTop;
                document.onmousemove = function(e) {{ if (isDragging) {{ el.style.left = (e.clientX - offsetX) + 'px'; el.style.top = (e.clientY - offsetY) + 'px'; }} }};
                document.onmouseup = function() {{ isDragging = false; document.onmousemove = null; document.onmouseup = null; }};
            }};
        }}

        function getRisk(s) {{ if(s==100) return 'risk-100'; if(s>=60) return 'risk-high'; if(s>=40) return 'risk-med'; return 'risk-low'; }}
        
        // --- SCORE LOGIC ---
        function updateScore() {{
            const pending = tp.rows().count();
            const fixed = tf.rows().count();
            const passed = tpass.rows().count();
            const total = pending + fixed + passed;
            const compliant = fixed + passed;
            
            let pct = 0;
            if (total > 0) {{ pct = (compliant / total) * 100; }}
            
            $('#score-val').text(pct.toFixed(1) + '%');
            $('#score-val').removeClass('val-good val-bad');
            if (pct > 80) {{ $('#score-val').addClass('val-good'); }} else {{ $('#score-val').addClass('val-bad'); }}
            
            $('#cnt-p').text(pending);
            $('#cnt-f').text(fixed);
        }}

        function moveMarked() {{
            const rowsToRemove = [];
            const dataToMove = [];
            tp.rows().every(function() {{
                const $node = $(this.node());
                if ($node.find('input.select-row').is(':checked')) {{
                    const d = {{
                        id: $node.attr('data-id'), cis: $node.attr('data-cis'),
                        cat: $node.attr('data-cat'), desc: $node.attr('data-desc'),
                        reg: $node.attr('data-reg'), regbtns: $node.attr('data-regbtns'),
                        curr: $node.attr('data-curr'), exp: $node.attr('data-exp'),
                        score: $node.attr('data-score')
                    }};
                    dataToMove.push(d);
                    rowsToRemove.push(this.node());
                }}
            }});
            dataToMove.forEach(d => addFixed(d));
            tp.rows(rowsToRemove).remove();
            $('.select-all').prop('checked', false);
            tp.draw(); tf.draw(); 
            updateScore(); // Recalc Score
            saveToStorage();
        }}

        function restoreItem(btn) {{
            const row = $(btn).closest('tr');
            const d = {{
                id: row.attr('data-id'), cis: row.attr('data-cis'),
                cat: row.attr('data-cat'), desc: row.attr('data-desc'),
                reg: row.attr('data-reg'), regbtns: row.attr('data-regbtns'),
                curr: row.attr('data-curr'), exp: row.attr('data-exp'),
                score: row.attr('data-score')
            }};
            addPending(d);
            tf.row(row).remove().draw();
            tp.draw(); 
            updateScore(); 
            saveToStorage();
        }}

        function addFixed(d) {{
            if (!d || !d.id) return;
            const r = getRisk(d.score);
            const node = tf.row.add([
                `<button class="btn-restore" onclick="restoreItem(this)">↩️</button>`,
                `<span class="badge ${{r}}">${{d.score}}</span>`, d.cis, d.cat, d.desc, d.reg + ' ' + d.regbtns, d.curr, d.exp
            ]).node();
            $(node).attr('data-id', d.id).attr('data-cis', d.cis)
                   .attr('data-cat', d.cat).attr('data-desc', d.desc)
                   .attr('data-reg', d.reg).attr('data-regbtns', d.regbtns)
                   .attr('data-curr', d.curr).attr('data-exp', d.exp)
                   .attr('data-score', d.score);
        }}
        
        function addPending(d) {{
            if (!d || !d.id) return;
            const r = getRisk(d.score);
            const node = tp.row.add([
                `<input type="checkbox" class="select-row">`,
                `<span class="badge ${{r}}">${{d.score}}</span>`, d.cis, d.cat, d.desc, d.reg + ' ' + d.regbtns, d.curr, d.exp
            ]).node();
            $(node).attr('data-id', d.id).attr('data-cis', d.cis)
                   .attr('data-cat', d.cat).attr('data-desc', d.desc)
                   .attr('data-reg', d.reg).attr('data-regbtns', d.regbtns)
                   .attr('data-curr', d.curr).attr('data-exp', d.exp)
                   .attr('data-score', d.score);
        }}

        function saveToStorage() {{
            const data = [];
            tf.rows().every(function() {{ 
                const id = $(this.node()).attr('data-id');
                if (id && id !== "undefined") data.push(id); 
            }});
            localStorage.setItem('hk_progress', JSON.stringify(data));
        }}

        function loadFromStorage() {{
            const ids = JSON.parse(localStorage.getItem('hk_progress') || '[]');
            applyProgress(ids);
        }}
        
        function loadFromFile(input) {{
            const file = input.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {{
                try {{
                    const ids = JSON.parse(e.target.result);
                    if (Array.isArray(ids)) {{
                        applyProgress(ids);
                        alert("Progress loaded successfully! (" + ids.length + " items)");
                    }} else {{ alert("Invalid file format."); }}
                }} catch (err) {{ alert("Error parsing JSON: " + err); }}
            }};
            reader.readAsText(file);
            input.value = '';
        }}

        function applyProgress(ids) {{
            const nodesToRemove = [];
            tp.rows().every(function() {{
                const node = this.node();
                const id = String($(node).attr('data-id'));
                if (ids.includes(id)) {{
                    const d = {{
                        id: $(node).attr('data-id'), cis: $(node).attr('data-cis'),
                        cat: $(node).attr('data-cat'), desc: $(node).attr('data-desc'),
                        reg: $(node).attr('data-reg'), regbtns: $(node).attr('data-regbtns'),
                        curr: $(node).attr('data-curr'), exp: $(node).attr('data-exp'),
                        score: $(node).attr('data-score')
                    }};
                    addFixed(d);
                    nodesToRemove.push(node);
                }}
            }});
            if (nodesToRemove.length > 0) {{
                tp.rows(nodesToRemove).remove();
                tp.draw(); tf.draw(); 
                updateScore(); 
                saveToStorage();
            }}
        }}

        async function saveFile() {{
            saveToStorage(); 
            const data = [];
            tf.rows().every(function() {{ 
                const id = $(this.node()).attr('data-id');
                if (id && id !== "undefined") data.push(id);
            }});
            const str = JSON.stringify(data, null, 2); 
            try {{
                if (!fileHandle) {{
                    fileHandle = await window.showSaveFilePicker({{ suggestedName: 'progress.json', types: [{{ description: 'JSON', accept: {{ 'application/json': ['.json'] }} }}] }});
                }}
                const w = await fileHandle.createWritable(); await w.write(str); await w.close();
                alert('Saved successfully!');
            }} catch (e) {{ console.error(e); alert('Save cancelled.'); }}
        }}
    </script>
</body>
</html>
    """
    with open(output_path, "w", encoding="utf-8") as f: f.write(html)
    print(f"✅ HTML App Created: {output_path}")

def main():
    report, templates = select_files_gui()
    if not report: return
    
    df = pd.read_csv(report)
    if 'id' in df.columns: df.rename(columns={'id': 'ID'}, inplace=True)
    if 'ID' in df.columns: df['CIS'] = df['ID']
        
    df['ID'] = df['ID'].astype(str).str.strip()
    if 'TestResult' not in df.columns and 'Result' in df.columns:
         df['TestResult'] = df['TestResult'].apply(lambda x: 'Failed' if 'Failed' in str(x) else 'Passed')
    df['TestResult'] = df['TestResult'].astype(str).str.title()

    if templates:
        print(f"Processing {len(templates)} template files...")
        combined_tmpl = pd.DataFrame()
        for t_file in templates:
            try:
                t = pd.read_csv(t_file)
                if 'id' in t.columns: t.rename(columns={'id': 'ID'}, inplace=True)
                t['ID'] = t['ID'].astype(str).str.strip()
                
                if 'Name' in t.columns:
                    t['Description'] = t['Name']
                
                desired_cols = ['ID', 'Description', 'Method', 'MethodArgument', 'RegistryPath', 'RegistryItem', 'RecommendedValue']
                actual_cols = [c for c in desired_cols if c in t.columns]
                
                if combined_tmpl.empty: combined_tmpl = t[actual_cols]
                else: combined_tmpl = pd.concat([combined_tmpl, t[actual_cols]])
            except Exception as e: print(f"Error loading template {t_file}: {e}")
        
        if not combined_tmpl.empty:
            combined_tmpl = combined_tmpl.drop_duplicates(subset=['ID'], keep='last')
            df = pd.merge(df, combined_tmpl, on='ID', how='left', suffixes=('', '_tmpl'))
            
            if 'Description_tmpl' in df.columns:
                df['Description'] = df['Description_tmpl'].combine_first(df['Description'])
                df.drop(columns=['Description_tmpl'], inplace=True, errors='ignore')
            
            for col in ['Method', 'MethodArgument']:
                 if col + '_tmpl' in df.columns:
                     df[col] = df[col + '_tmpl']
                     df.drop(columns=[col + '_tmpl'], inplace=True)

    df['RiskScore'] = df.apply(calculate_risk, axis=1)
    df['Fix'] = df.apply(generate_fix, axis=1) if 'RegistryPath' in df.columns else ""
    
    base = os.path.splitext(report)[0]
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    
    df_failed = df[df['TestResult'].str.contains('Failed', na=False)].copy()
    df_passed = df[df['TestResult'].str.contains('Passed', na=False)].copy()

    generate_excel(df, f"{base}_Report_{ts}.xlsx", df_failed, df_passed)
    generate_html(df, f"{base}_App_{ts}.html", 0, len(df), len(df_passed), len(df_failed), df['Category'].unique())
    print("\n🎉 Full Suite Generated Successfully!")

if __name__ == "__main__":
    main()