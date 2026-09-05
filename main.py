#!/usr/bin/env python3
"""
ER-AI Offline DOCX Generator GUI
Reads Admission Order output from Gemini Gem clipboard/text
and generates Doctor Order, Inpatient H&P, and Informed Consent docx files.

Usage:
  python3 main.py

Or package as .exe:
  pyinstaller --onefile --windowed --add-data "templates:templates" --name "ER-AI_DOCX_Generator" main.py
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk

APP_VERSION = 'v1.2.0'

from tkinter import ttk, messagebox, filedialog

# --- Path Resolution (works for both script and PyInstaller .exe) ---
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")


# ============================================================
# UTILITY FUNCTIONS (ported from backend/main.py)
# ============================================================

def strip_markdown_tags(text: str) -> str:
    """Remove markdown code fences, bold markers, and HTML tags."""
    if not text:
        return ""
    text = re.sub(r'```[a-z]*', '', text, flags=re.I)
    text = re.sub(r'\*\*', '', text)
    text = re.sub(r'<[^>]+>', '', text)
    return text.strip()


def extract_tag(source: str, tag: str, max_len: int, default_val: str = "") -> str:
    """Extract [TAG]: value or TAG: value from source text."""
    m = re.search(rf'\[{tag}\][ \t]*[:=]?[ \t]*([^\n]+)', source, re.I)
    if not m:
        m = re.search(rf'{tag}[ \t]*[:=]?[ \t]*([^\n]+)', source, re.I)
    val = m.group(1).strip() if m else ""
    val = re.sub(r'\[.*?\]', '', val).strip()
    if not val:
        val = default_val
    val = val.replace('\n', ' ').replace('\r', '')
    val = re.sub(r'[\x00-\x1F\x7F]', '', val)
    return val[:max_len]


def extract_multiline_tag(source: str, tag: str, default_val: str = "") -> str:
    """Extract [TAG]...[/TAG] block content."""
    m = re.search(rf'\[{tag}\](.*?)\[/{tag}\]', source, re.DOTALL | re.I)
    if m:
        val = m.group(1).strip()
        return strip_markdown_tags(val)
    return default_val


def clean_pdx(raw_pdx: str) -> str:
    """Strip ICD-10 codes from diagnosis text."""
    if not raw_pdx:
        return ""
    cleaned = re.sub(r'^[A-Z]\d{2}(?:\.\d+)?\s*[-:]?\s*', '', raw_pdx, flags=re.I)
    cleaned = re.sub(r'\s*[\(\[]?[A-Z]\d{2}(?:\.\d+)?[\)\]]?', '', cleaned, flags=re.I)
    return cleaned.strip()[:60]


def extract_first_diagnosis(text: str, max_len: int = 100) -> str:
    if not text:
        return ""
    m = re.search(r'[A-Z]\d{2}(?:\.\d+)?\s*[-:]?\s*([^\n]+)', text, flags=re.I)
    if m:
        return m.group(1).strip()[:max_len]
    return text.strip()[:max_len]


def merge_docx_namespaces(orig_xml_str: str, generated_xml_bytes: bytes) -> bytes:
    """
    Restore original xmlns declarations to prevent MS Word 'unreadable content' errors.
    Python's ElementTree drops unused namespaces that Word requires.
    """
    orig_root_match = re.search(r'<w:document[^>]+>', orig_xml_str)
    gen_root_match = re.search(rb'<w:document[^>]+>', generated_xml_bytes)

    if orig_root_match and gen_root_match:
        orig_root_str = orig_root_match.group(0)
        gen_root_str = gen_root_match.group(0).decode('utf-8')

        gen_xmlns = re.findall(r'xmlns:[a-zA-Z0-9_]+="[^"]+"', gen_root_str)
        for xmlns in gen_xmlns:
            if xmlns not in orig_root_str:
                orig_root_str = orig_root_str.replace('>', f' {xmlns}>')

        return generated_xml_bytes.replace(gen_root_match.group(0), orig_root_str.encode('utf-8'))

    return generated_xml_bytes


def fill_textbox_template(template_path: str, replacements: dict, output_path: str):
    """
    Fill a .docx template that uses text boxes (txbxContent) as placeholders.
    Used by H&P and Consent templates.
    """
    with zipfile.ZipFile(template_path, 'r') as zin:
        xml_bytes = zin.read('word/document.xml')
        all_files = {name: zin.read(name) for name in zin.namelist() if name != 'word/document.xml'}

    xml_str_decoded = xml_bytes.decode('utf-8')
    ns_matches = re.findall(r'xmlns:([a-zA-Z0-9_]+)="([^"]+)"', xml_str_decoded)
    for prefix, uri in set(ns_matches):
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(xml_bytes)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for elem in root.iter():
        if elem.tag.endswith('txbxContent'):
            p_nodes = elem.findall('.//w:p', ns)
            full_t = ''.join([t.text for p in p_nodes for t in p.findall('.//w:t', ns) if t.text]).strip()
            if full_t in replacements:
                val = replacements[full_t]
                for p in p_nodes:
                    pPr = p.find('w:pPr', ns)
                    if pPr is None:
                        pPr = ET.SubElement(p, f'{{{ns["w"]}}}pPr')
                    spacing = pPr.find('w:spacing', ns)
                    if spacing is None:
                        spacing = ET.SubElement(pPr, f'{{{ns["w"]}}}spacing')
                    spacing.attrib[f'{{{ns["w"]}}}after'] = '0'
                    spacing.attrib[f'{{{ns["w"]}}}line'] = '240'
                    spacing.attrib[f'{{{ns["w"]}}}lineRule'] = 'auto'

                    all_t_nodes = p.findall('.//w:t', ns)
                    if all_t_nodes:
                        all_t_nodes[0].text = val
                        for ot in all_t_nodes[1:]:
                            ot.text = ''

                    r_nodes = p.findall('.//w:r', ns)
                    if r_nodes:
                        rPr = r_nodes[0].find('w:rPr', ns)
                        if rPr is None:
                            rPr = ET.SubElement(r_nodes[0], f'{{{ns["w"]}}}rPr')
                        sz = rPr.find('w:sz', ns)
                        if sz is None:
                            sz = ET.SubElement(rPr, f'{{{ns["w"]}}}sz')
                        sz.attrib[f'{{{ns["w"]}}}val'] = '20'
                        szCs = rPr.find('w:szCs', ns)
                        if szCs is None:
                            szCs = ET.SubElement(rPr, f'{{{ns["w"]}}}szCs')
                        szCs.attrib[f'{{{ns["w"]}}}val'] = '20'

    new_xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    new_xml_bytes = merge_docx_namespaces(xml_str_decoded, new_xml_bytes)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('word/document.xml', new_xml_bytes)
        for fname, fbytes in all_files.items():
            zout.writestr(fname, fbytes)


def fill_paragraph_template(template_path: str, replacements: dict, output_path: str):
    """
    Fill a .docx template that uses paragraph-level placeholders.
    Used by Doctor Order template. Supports multi-line content with <w:br/>.
    """
    with zipfile.ZipFile(template_path, 'r') as zin:
        xml_bytes = zin.read('word/document.xml')
        all_files = {name: zin.read(name) for name in zin.namelist() if name != 'word/document.xml'}

    xml_str_decoded = xml_bytes.decode('utf-8')
    ns_matches = re.findall(r'xmlns:([a-zA-Z0-9_]+)="([^"]+)"', xml_str_decoded)
    for prefix, uri in set(ns_matches):
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(xml_bytes)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for p in root.iter(f'{{{ns["w"]}}}p'):
        all_t_nodes = p.findall('./w:r/w:t', ns)
        if not all_t_nodes:
            continue

        full_text = ''.join(t.text for t in all_t_nodes if t.text)
        matched = False
        matched_val = ""
        for k, v in replacements.items():
            if k in full_text:
                full_text = full_text.replace(k, v)
                matched = True
                matched_val = v

        if matched:
            if '\n' in matched_val:
                for child in list(p):
                    if child.tag == f'{{{ns["w"]}}}r':
                        p.remove(child)
                lines = full_text.split('\n')
                for i, line in enumerate(lines):
                    r = ET.SubElement(p, f'{{{ns["w"]}}}r')
                    rPr = ET.SubElement(r, f'{{{ns["w"]}}}rPr')
                    rFonts = ET.SubElement(rPr, f'{{{ns["w"]}}}rFonts')
                    rFonts.attrib[f'{{{ns["w"]}}}ascii'] = 'TH SarabunPSK'
                    rFonts.attrib[f'{{{ns["w"]}}}hAnsi'] = 'TH SarabunPSK'
                    rFonts.attrib[f'{{{ns["w"]}}}cs'] = 'TH SarabunPSK'
                    sz = ET.SubElement(rPr, f'{{{ns["w"]}}}sz')
                    sz.attrib[f'{{{ns["w"]}}}val'] = '28'
                    szCs = ET.SubElement(rPr, f'{{{ns["w"]}}}szCs')
                    szCs.attrib[f'{{{ns["w"]}}}val'] = '28'
                    t = ET.SubElement(r, f'{{{ns["w"]}}}t')
                    t.text = line
                    if i < len(lines) - 1:
                        ET.SubElement(r, f'{{{ns["w"]}}}br')
            else:
                all_t_nodes[0].text = full_text
                for ot in all_t_nodes[1:]:
                    ot.text = ''

    new_xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    new_xml_bytes = merge_docx_namespaces(xml_str_decoded, new_xml_bytes)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('word/document.xml', new_xml_bytes)
        for fname, fbytes in all_files.items():
            zout.writestr(fname, fbytes)


# ============================================================
# DOCX GENERATORS
# ============================================================

def generate_doctor_order(raw_text: str, output_dir: str) -> str:
    """Generate Doctor Order .docx from Admission Order output."""
    template_path = os.path.join(TEMPLATES_DIR, "order_template.docx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    clean = strip_markdown_tags(raw_text)

    patient_name = extract_multiline_tag(clean, "Pt_Name", "Unregistered")
    if patient_name == "Unregistered":
        patient_name = extract_tag(clean, "Patient", 50, "Unregistered")

    age = extract_multiline_tag(clean, "Pt_Age", "-")
    if age == "-":
        age = extract_tag(clean, "Age", 3, "-")

    hn = extract_multiline_tag(clean, "Pt_HN", "Unregistered")
    if not hn or hn == "-":
        hn = extract_tag(clean, "HN", 20, "Unregistered")

    pl_order = extract_multiline_tag(clean, "PL_Order", "No Problem List")
    one_order = extract_multiline_tag(clean, "One_Order", "No One Day Order")
    cont_order = extract_multiline_tag(clean, "Cont_Order", "")

    replacements = {
        "{PL_Order}": pl_order,
        "{One_Order}": one_order,
        "{Cont_Order}": cont_order,
        "{Pt_Name}": patient_name,
        "{Pt_Age}": age,
        "{Pt_HN}": hn
    }

    safe_hn = re.sub(r'[^a-zA-Z0-9]', '_', hn) or "Guest"
    filename = f"Doctor_Order_{safe_hn}.docx"
    output_path = os.path.join(output_dir, filename)

    fill_paragraph_template(template_path, replacements, output_path)
    return output_path


def generate_inpatient_hp(raw_text: str, output_dir: str) -> str:
    """Generate Inpatient H&P .docx from Admission Order output."""
    template_path = os.path.join(TEMPLATES_DIR, "hp_template.docx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    clean = strip_markdown_tags(raw_text)

    # Patient demographics
    patient_name = extract_multiline_tag(clean, "Pt_Name", "Unregistered")
    if patient_name == "Unregistered":
        patient_name = extract_tag(clean, "Patient", 50, "Unregistered")

    age = extract_multiline_tag(clean, "Pt_Age", "-")
    if age == "-":
        age = extract_tag(clean, "Age", 3, "-")

    hn = extract_multiline_tag(clean, "Pt_HN", "Unregistered")
    if not hn or hn == "-":
        hn = extract_tag(clean, "HN", 20, "Unregistered")

    # Vitals - try [HP_Data] tags first, then fallback to raw input parsing
    bt = extract_tag(clean, "BT", 4, "-")
    pr = extract_tag(clean, "PR", 3, "-")
    rr = extract_tag(clean, "RR", 2, "-")
    bw = extract_tag(clean, "BW", 4, "-")

    bp_match = re.search(r'BP\s*[:=]?\s*(\d+)\s*/\s*(\d+)', clean, re.I)
    sbp_default = bp_match.group(1) if bp_match else "-"
    dbp_default = bp_match.group(2) if bp_match else "-"
    sbp = extract_tag(clean, "SBP", 3, sbp_default)
    dbp = extract_tag(clean, "DBP", 3, dbp_default)

    # Clinical data
    cc = extract_tag(clean, "CC", 125, "-")
    pi1 = extract_tag(clean, "PI1", 100, "-")
    pi2 = extract_tag(clean, "PI2", 130, "")
    ud = extract_tag(clean, "UD", 19, "ไม่มี")
    fda = extract_tag(clean, "FDA", 20, "ปฏิเสธแพ้ยา")

    # PE
    ga = extract_tag(clean, "GA", 45, "Good consciousness, not pale, no jaundice")
    heent = extract_tag(clean, "HEENT", 46, "Pharynx and tonsils not injected")
    cvs = extract_tag(clean, "CVS", 46, "Normal S1 S2, no murmur")
    rs = extract_tag(clean, "RS", 46, "Normal breath sound, no adventitious sound")
    abd = extract_tag(clean, "Abd", 46, "Soft, non-tender, normal bowel sound")
    ext = extract_tag(clean, "Ext", 46, "No deformities, no edema")
    cns = extract_tag(clean, "CNS", 46, "E4V5M6, motor power grade 5 all")

    # Assessment
    pl = extract_tag(clean, "PL", 50, "-")
    raw_pdx = extract_tag(clean, "PDx", 80, "Admitted for evaluation")
    pdx = clean_pdx(raw_pdx)
    plan = extract_tag(clean, "Plan", 60, "Admit ward for further management")

    replacements = {
        "CC": cc, "PI line 1": pi1, "PI line 2": pi2, "UD": ud, "FDA": fda,
        "BW": bw, "DBP": dbp, "SBP": sbp, "RR": rr, "PR": pr, "BT": bt,
        "GA": ga, "HEENT": heent, "CVS": cvs, "RS": rs, "Abd": abd, "Ext": ext,
        "CNS": cns, "PL": pl, "PDx": pdx, "Plan": plan,
        "{Pt_Name}": patient_name,
        "{Pt_Age}": age,
        "{Pt_HN}": hn
    }

    safe_hn = re.sub(r'[^a-zA-Z0-9]', '_', hn) or "Guest"
    filename = f"Inpatient_HP_{safe_hn}.docx"
    output_path = os.path.join(output_dir, filename)

    fill_textbox_template(template_path, replacements, output_path)
    return output_path


def generate_informed_consent(raw_text: str, output_dir: str) -> str:
    """Generate Informed Consent .docx from Admission Order output."""
    template_path = os.path.join(TEMPLATES_DIR, "infomed_consent_template.docx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    clean = strip_markdown_tags(raw_text)

    # Patient demographics
    patient_name = extract_multiline_tag(clean, "Pt_Name", "Unregistered")
    if patient_name == "Unregistered":
        patient_name = extract_tag(clean, "Patient", 50, "Unregistered")

    age = extract_multiline_tag(clean, "Pt_Age", "-")
    if age == "-":
        age = extract_tag(clean, "Age", 3, "-")

    hn = extract_multiline_tag(clean, "Pt_HN", "Unregistered")
    if not hn or hn == "-":
        hn = extract_tag(clean, "HN", 20, "Unregistered")

    # Diagnosis
    m_dx = re.search(r'\[Dx_Consent\][ \t]*[:=]?[ \t]*([^\n]+)', raw_text, re.I)
    if not m_dx:
        m_dx = re.search(r'Dx_Consent[ \t]*[:=]?[ \t]*([^\n]+)', raw_text, re.I)
    if m_dx:
        dx_consent = m_dx.group(1).strip()
        dx_consent = re.sub(r'\[.*?\]', '', dx_consent).strip()
    else:
        dx_consent = extract_first_diagnosis(clean, max_len=100)

    # Purpose checkboxes
    purpose_str = extract_tag(clean, "Purpose_Consent", 100, "Mx")
    purpose_list = [p.strip().upper() for p in purpose_str.split(",")]
    chk_dx = "✓" if "DX" in purpose_list else " "
    chk_sx = "✓" if "SX" in purpose_list else " "
    chk_ix = "✓" if "IX" in purpose_list else " "
    chk_mx = "✓" if "MX" in purpose_list or not purpose_list else " "
    chk_else_why = "✓" if "OTHER" in purpose_list else " "
    other_why = extract_tag(clean, "Other_Why_Consent", 40, "") if "OTHER" in purpose_list else ""

    # Treatment checkboxes
    tx_str = extract_tag(clean, "Tx_Consent", 100, "IV, Med")
    tx_list = [t.strip().upper() for t in tx_str.split(",")]
    chk_iv = "✓" if "IV" in tx_list else " "
    chk_med = "✓" if "MED" in tx_list else " "
    chk_nb = "✓" if "NB" in tx_list else " "
    chk_prc = "✓" if "PRC" in tx_list else " "
    chk_else_mx = "✓" if "OTHER" in tx_list else " "
    other_mx = extract_tag(clean, "Other_Mx_Consent", 40, "") if "OTHER" in tx_list else ""

    # Consent text fields
    pros = extract_tag(clean, "Pros_Consent", 90, "เพื่อควบคุมการติดเชื้อ บรรเทาอาการ และป้องกันภาวะแทรกซ้อนที่อาจเป็นอันตราย")
    cons = extract_tag(clean, "Cons_Consent", 90, "อาจมีอาการข้างเคียงจากยา เช่น ถ่ายเหลว คลื่นไส้ หรือแพ้ยา")
    alt = extract_tag(clean, "Alt_Consent", 90, "ปฏิเสธการรักษาและรับยากลับไปรับประทานที่บ้าน (อาจทำให้โรครุนแรงขึ้น)")
    outcome = extract_tag(clean, "Outcome_Consent", 90, "คาดว่าอาการจะค่อยๆ ดีขึ้น และสามารถกลับบ้านได้เมื่ออาการคงที่")
    duration = extract_tag(clean, "Duration_Consent", 30, "ประมาณ 3-5 วัน")

    # Risk checkboxes
    risk_str = extract_tag(clean, "Risk_Consent", 100, "")
    risk_list = [r.strip().upper() for r in risk_str.split(",")]
    chk_infect = "✓" if "INFECT" in risk_list else " "
    chk_fda = "✓" if "FDA" in risk_list else " "
    chk_prc_a = "✓" if "PRC_A" in risk_list else " "
    chk_else_comp = "✓" if "OTHER" in risk_list else " "
    other_comp = extract_tag(clean, "Other_Comp_Consent", 40, "") if "OTHER" in risk_list else ""

    replacements = {
        "Dx_Consent": dx_consent, "Dx": chk_dx, "Sx": chk_sx, "Ix": chk_ix,
        "Mx": chk_mx, "Else_why": chk_else_why, "Other_Why": other_why,
        "IV": chk_iv, "Med": chk_med, "NB": chk_nb, "PRC": chk_prc,
        "Else_Mx": chk_else_mx, "Other_Mx": other_mx, "Pros_Consent": pros,
        "Cons_Consent": cons, "Alt_Consent": alt, "Outcome_Consent": outcome,
        "Duration_Consent": duration, "Infect": chk_infect, "FDA": chk_fda,
        "PRC_A": chk_prc_a, "Else_comp": chk_else_comp, "Other_comp": other_comp,
        "{Pt_Name}": patient_name,
        "{Pt_Age}": age,
        "{Pt_HN}": hn
    }

    safe_hn = re.sub(r'[^a-zA-Z0-9]', '_', hn) or "Guest"
    filename = f"Informed_Consent_{safe_hn}.docx"
    output_path = os.path.join(output_dir, filename)

    fill_textbox_template(template_path, replacements, output_path)
    return output_path


import json
from datetime import datetime
import subprocess
import threading
from pathlib import Path

# ============================================================
# GUI APPLICATION
# ============================================================


class ManualEntryWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title(f"Manual IPD Entry Form {APP_VERSION}")
        self.geometry("900x750")
        self.configure(bg="#1a1a2e")
        
        # --- Variables ---
        self.v_name = tk.StringVar()
        self.v_age = tk.StringVar()
        self.v_hn = tk.StringVar()
        
        # HxPE Vars
        self.v_cc = tk.StringVar()
        self.v_pi1 = tk.StringVar()
        self.v_pi2 = tk.StringVar()
        self.v_ud = tk.StringVar(value="ไม่มี")
        self.v_fda = tk.StringVar(value="ปฏิเสธแพ้ยา")
        self.v_bw = tk.StringVar()
        self.v_bt = tk.StringVar()
        self.v_pr = tk.StringVar()
        self.v_rr = tk.StringVar()
        self.v_sbp = tk.StringVar()
        self.v_dbp = tk.StringVar()
        self.v_ga = tk.StringVar(value="alert, good consciousness")
        self.v_heent = tk.StringVar(value="not pale conjunctiva, anicteric sclera, no LN")
        self.v_cvs = tk.StringVar(value="CRT<2sec, pulse RA 2+")
        self.v_rs = tk.StringVar(value="clear equal BS both")
        self.v_abd = tk.StringVar(value="soft, not tender, no rebound, no guarding")
        self.v_ext = tk.StringVar(value="no edema, no rash")
        self.v_cns = tk.StringVar(value="grossly intact")
        self.v_pl = tk.StringVar()
        self.v_pdx = tk.StringVar()
        self.v_plan = tk.StringVar(value="Admit for management")
        
        # Consent Vars
        self.v_dx_consent = tk.StringVar()
        self.cb_purp_mx = tk.BooleanVar(value=True)
        self.cb_purp_ix = tk.BooleanVar(value=True)
        self.cb_purp_sx = tk.BooleanVar()
        self.cb_purp_dx = tk.BooleanVar()
        self.cb_tx_iv = tk.BooleanVar(value=True)
        self.cb_tx_med = tk.BooleanVar(value=True)
        self.cb_tx_nb = tk.BooleanVar()
        self.cb_tx_prc = tk.BooleanVar()
        self.v_pros = tk.StringVar(value="เพื่อให้ได้รับการรักษาที่เหมาะสมและปลอดภัย")
        self.v_cons = tk.StringVar(value="อาจเกิดอาการแพ้ยา ติดเชื้อ หรือภาวะแทรกซ้อนอื่นๆ")
        self.v_alt = tk.StringVar(value="ปฏิเสธการรักษาและกลับบ้าน (อาจมีความเสี่ยง)")
        self.v_outcome = tk.StringVar(value="คาดว่าอาการดีขึ้นและกลับบ้านได้")
        self.v_duration = tk.StringVar(value="ประมาณ 3-5 วัน")
        self.cb_risk_infect = tk.BooleanVar(value=True)
        self.cb_risk_fda = tk.BooleanVar(value=True)
        self.cb_risk_prc_a = tk.BooleanVar()
        self.cb_purp_else = tk.BooleanVar()
        self.v_other_why = tk.StringVar()
        self.cb_tx_else = tk.BooleanVar()
        self.v_other_mx = tk.StringVar()
        self.cb_risk_else = tk.BooleanVar()
        self.v_other_comp = tk.StringVar()
        
        self.build_ui()
        
    def build_ui(self):
        # Top Header (Global)
        top_frame = tk.Frame(self, bg="#1a1a2e")
        top_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Label(top_frame, text="Patient Name:", bg="#1a1a2e", fg="white").grid(row=0, column=0, sticky="w", padx=5)
        tk.Entry(top_frame, textvariable=self.v_name, width=30).grid(row=0, column=1, padx=5)
        
        tk.Label(top_frame, text="Age:", bg="#1a1a2e", fg="white").grid(row=0, column=2, sticky="w", padx=5)
        tk.Entry(top_frame, textvariable=self.v_age, width=10).grid(row=0, column=3, padx=5)
        
        tk.Label(top_frame, text="HN:", bg="#1a1a2e", fg="white").grid(row=0, column=4, sticky="w", padx=5)
        tk.Entry(top_frame, textvariable=self.v_hn, width=20).grid(row=0, column=5, padx=5)
        
        # Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab_hxpe = tk.Frame(self.notebook, bg="#21262d")
        self.tab_consent = tk.Frame(self.notebook, bg="#21262d")
        self.tab_order = tk.Frame(self.notebook, bg="#21262d")
        
        self.notebook.add(self.tab_hxpe, text="Hx & PE")
        self.notebook.add(self.tab_consent, text="Consent")
        self.notebook.add(self.tab_order, text="Orders")
        
        self.build_hxpe_tab()
        self.build_consent_tab()
        self.build_order_tab()
        
        # Bottom Buttons
        bottom_frame = tk.Frame(self, bg="#1a1a2e")
        bottom_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Button(bottom_frame, text="Push to Main Window & Generate", command=self.generate_docs,
                  bg="#238636", fg="white", font=("Arial", 11, "bold")).pack(side="right")
        tk.Button(bottom_frame, text="Close", command=self.destroy).pack(side="left")

    def build_hxpe_tab(self):
        f = self.tab_hxpe
        
        # Left Col
        tk.Label(f, text="CC:", bg="#21262d", fg="white").grid(row=0, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_cc, width=40).grid(row=0, column=1)
        
        tk.Label(f, text="PI 1:", bg="#21262d", fg="white").grid(row=1, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_pi1, width=40).grid(row=1, column=1)
        
        tk.Label(f, text="PI 2:", bg="#21262d", fg="white").grid(row=2, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_pi2, width=40).grid(row=2, column=1)
        
        tk.Label(f, text="U/D:", bg="#21262d", fg="white").grid(row=3, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_ud, width=40).grid(row=3, column=1)
        
        tk.Label(f, text="FDA:", bg="#21262d", fg="white").grid(row=4, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_fda, width=40).grid(row=4, column=1)
        
        tk.Label(f, text="Problem List:", bg="#21262d", fg="white").grid(row=5, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_pl, width=40).grid(row=5, column=1)
        
        tk.Label(f, text="Primary Dx:", bg="#21262d", fg="white").grid(row=6, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_pdx, width=40).grid(row=6, column=1)
        
        tk.Label(f, text="Plan:", bg="#21262d", fg="white").grid(row=7, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_plan, width=40).grid(row=7, column=1)
        
        # Right Col (Vitals & PE)
        tk.Label(f, text="BW:", bg="#21262d", fg="white").grid(row=0, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_bw, width=15).grid(row=0, column=3)
        tk.Label(f, text="BT:", bg="#21262d", fg="white").grid(row=1, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_bt, width=15).grid(row=1, column=3)
        tk.Label(f, text="PR:", bg="#21262d", fg="white").grid(row=2, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_pr, width=15).grid(row=2, column=3)
        tk.Label(f, text="RR:", bg="#21262d", fg="white").grid(row=3, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_rr, width=15).grid(row=3, column=3)
        tk.Label(f, text="SBP:", bg="#21262d", fg="white").grid(row=4, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_sbp, width=15).grid(row=4, column=3)
        tk.Label(f, text="DBP:", bg="#21262d", fg="white").grid(row=5, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_dbp, width=15).grid(row=5, column=3)
        
        tk.Label(f, text="GA:", bg="#21262d", fg="white").grid(row=6, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_ga, width=25).grid(row=6, column=3)
        tk.Label(f, text="HEENT:", bg="#21262d", fg="white").grid(row=7, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_heent, width=25).grid(row=7, column=3)
        tk.Label(f, text="CVS:", bg="#21262d", fg="white").grid(row=8, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_cvs, width=25).grid(row=8, column=3)
        tk.Label(f, text="RS:", bg="#21262d", fg="white").grid(row=9, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_rs, width=25).grid(row=9, column=3)
        tk.Label(f, text="Abd:", bg="#21262d", fg="white").grid(row=10, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_abd, width=25).grid(row=10, column=3)
        tk.Label(f, text="Ext:", bg="#21262d", fg="white").grid(row=11, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_ext, width=25).grid(row=11, column=3)
        tk.Label(f, text="CNS:", bg="#21262d", fg="white").grid(row=12, column=2, sticky="w", padx=(30,5))
        tk.Entry(f, textvariable=self.v_cns, width=25).grid(row=12, column=3)

    def build_consent_tab(self):
        f = self.tab_consent
        
        tk.Label(f, text="Diagnosis:", bg="#21262d", fg="white").grid(row=0, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_dx_consent, width=40).grid(row=0, column=1, columnspan=3, sticky="w")
        
        # Purpose
        tk.Label(f, text="Purpose:", bg="#21262d", fg="white").grid(row=1, column=0, sticky="w", pady=5, padx=10)
        f_purp = tk.Frame(f, bg="#21262d")
        f_purp.grid(row=1, column=1, columnspan=3, sticky="w")
        tk.Checkbutton(f_purp, text="Mx", variable=self.cb_purp_mx, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_purp, text="Ix", variable=self.cb_purp_ix, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_purp, text="Sx", variable=self.cb_purp_sx, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_purp, text="Dx", variable=self.cb_purp_dx, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_purp, text="Other", variable=self.cb_purp_else, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Entry(f_purp, textvariable=self.v_other_why, width=15).pack(side="left", padx=5)
        
        # Tx
        tk.Label(f, text="Treatment:", bg="#21262d", fg="white").grid(row=2, column=0, sticky="w", pady=5, padx=10)
        f_tx = tk.Frame(f, bg="#21262d")
        f_tx.grid(row=2, column=1, columnspan=3, sticky="w")
        tk.Checkbutton(f_tx, text="IV", variable=self.cb_tx_iv, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_tx, text="Med", variable=self.cb_tx_med, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_tx, text="NB", variable=self.cb_tx_nb, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_tx, text="PRC", variable=self.cb_tx_prc, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_tx, text="Other", variable=self.cb_tx_else, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Entry(f_tx, textvariable=self.v_other_mx, width=15).pack(side="left", padx=5)
        
        # Risks
        tk.Label(f, text="Risks:", bg="#21262d", fg="white").grid(row=3, column=0, sticky="w", pady=5, padx=10)
        f_risk = tk.Frame(f, bg="#21262d")
        f_risk.grid(row=3, column=1, columnspan=3, sticky="w")
        tk.Checkbutton(f_risk, text="Infect", variable=self.cb_risk_infect, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_risk, text="FDA", variable=self.cb_risk_fda, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_risk, text="PRC_A", variable=self.cb_risk_prc_a, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Checkbutton(f_risk, text="Other", variable=self.cb_risk_else, bg="#21262d", fg="white", selectcolor="#1a1a2e").pack(side="left")
        tk.Entry(f_risk, textvariable=self.v_other_comp, width=15).pack(side="left", padx=5)
        
        tk.Label(f, text="Pros:", bg="#21262d", fg="white").grid(row=4, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_pros, width=50).grid(row=4, column=1, columnspan=3, sticky="w")
        
        tk.Label(f, text="Cons:", bg="#21262d", fg="white").grid(row=5, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_cons, width=50).grid(row=5, column=1, columnspan=3, sticky="w")
        
        tk.Label(f, text="Alternatives:", bg="#21262d", fg="white").grid(row=6, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_alt, width=50).grid(row=6, column=1, columnspan=3, sticky="w")
        
        tk.Label(f, text="Outcome:", bg="#21262d", fg="white").grid(row=7, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_outcome, width=50).grid(row=7, column=1, columnspan=3, sticky="w")
        
        tk.Label(f, text="Duration:", bg="#21262d", fg="white").grid(row=8, column=0, sticky="w", pady=5, padx=10)
        tk.Entry(f, textvariable=self.v_duration, width=50).grid(row=8, column=1, columnspan=3, sticky="w")
        
    def build_order_tab(self):
        f = self.tab_order
        
        tk.Label(f, text="Problem List:", bg="#21262d", fg="white").pack(anchor="w", pady=(5,0), padx=10)
        self.t_pl = tk.Text(f, height=5, width=80)
        self.t_pl.pack(pady=5, padx=10)
        
        tk.Label(f, text="One Day Order:", bg="#21262d", fg="white").pack(anchor="w", pady=(5,0), padx=10)
        self.t_one = tk.Text(f, height=5, width=80)
        self.t_one.pack(pady=5, padx=10)
        
        tk.Label(f, text="Continue Order:", bg="#21262d", fg="white").pack(anchor="w", pady=(5,0), padx=10)
        self.t_cont = tk.Text(f, height=5, width=80)
        self.t_cont.pack(pady=5, padx=10)

    def generate_docs(self):
        out = []
        
        out.append(f"[Pt_Name]\n{self.v_name.get()}\n[/Pt_Name]\n")
        out.append(f"[Pt_Age]\n{self.v_age.get()}\n[/Pt_Age]\n")
        out.append(f"[Pt_HN]\n{self.v_hn.get()}\n[/Pt_HN]\n")
        
        out.append(f"[PL_Order]\n{self.t_pl.get('1.0', tk.END).strip()}\n[/PL_Order]\n")
        out.append(f"[One_Order]\n{self.t_one.get('1.0', tk.END).strip()}\n[/One_Order]\n")
        out.append(f"[Cont_Order]\n{self.t_cont.get('1.0', tk.END).strip()}\n[/Cont_Order]\n")
        
        out.append("[HP_Data]")
        out.append(f"[CC]: {self.v_cc.get()}")
        out.append(f"[PI1]: {self.v_pi1.get()}")
        out.append(f"[PI2]: {self.v_pi2.get()}")
        out.append(f"[UD]: {self.v_ud.get()}")
        out.append(f"[FDA]: {self.v_fda.get()}")
        out.append(f"[BW]: {self.v_bw.get()}")
        out.append(f"[BT]: {self.v_bt.get()}")
        out.append(f"[PR]: {self.v_pr.get()}")
        out.append(f"[RR]: {self.v_rr.get()}")
        out.append(f"[SBP]: {self.v_sbp.get()}")
        out.append(f"[DBP]: {self.v_dbp.get()}")
        out.append(f"[GA]: {self.v_ga.get()}")
        out.append(f"[HEENT]: {self.v_heent.get()}")
        out.append(f"[CVS]: {self.v_cvs.get()}")
        out.append(f"[RS]: {self.v_rs.get()}")
        out.append(f"[Abd]: {self.v_abd.get()}")
        out.append(f"[Ext]: {self.v_ext.get()}")
        out.append(f"[CNS]: {self.v_cns.get()}")
        out.append(f"[PL]: {self.v_pl.get()}")
        out.append(f"[PDx]: {self.v_pdx.get()}")
        out.append(f"[Plan]: {self.v_plan.get()}")
        out.append("[/HP_Data]\n")
        
        purps = []
        if self.cb_purp_mx.get(): purps.append("Mx")
        if self.cb_purp_ix.get(): purps.append("Ix")
        if self.cb_purp_sx.get(): purps.append("Sx")
        if self.cb_purp_dx.get(): purps.append("Dx")
        if self.cb_purp_else.get(): purps.append("Other")
        
        txs = []
        if self.cb_tx_iv.get(): txs.append("IV")
        if self.cb_tx_med.get(): txs.append("Med")
        if self.cb_tx_nb.get(): txs.append("NB")
        if self.cb_tx_prc.get(): txs.append("PRC")
        if self.cb_tx_else.get(): txs.append("Other")
        
        risks = []
        if self.cb_risk_infect.get(): risks.append("Infect")
        if self.cb_risk_fda.get(): risks.append("FDA")
        if self.cb_risk_prc_a.get(): risks.append("PRC_A")
        if self.cb_risk_else.get(): risks.append("Other")
        
        out.append("[Consent_Data]")
        out.append(f"[Dx_Consent]: {self.v_dx_consent.get() or self.v_pdx.get()}")
        out.append(f"[Purpose_Consent]: {','.join(purps)}")
        out.append(f"[Tx_Consent]: {','.join(txs)}")
        out.append(f"[Pros_Consent]: {self.v_pros.get()}")
        out.append(f"[Cons_Consent]: {self.v_cons.get()}")
        out.append(f"[Alt_Consent]: {self.v_alt.get()}")
        out.append(f"[Outcome_Consent]: {self.v_outcome.get()}")
        out.append(f"[Duration_Consent]: {self.v_duration.get()}")
        out.append(f"[Risk_Consent]: {','.join(risks)}")
        out.append(f"[Other_Why_Consent]: {self.v_other_why.get()}")
        out.append(f"[Other_Mx_Consent]: {self.v_other_mx.get()}")
        out.append(f"[Other_Comp_Consent]: {self.v_other_comp.get()}")
        out.append("[/Consent_Data]")
        
        compiled_text = "\n".join(out)
        self.parent.txt_input.delete("1.0", tk.END)
        self.parent.txt_input.insert("1.0", compiled_text)
        
        # Trigger generation on parent
        self.parent.run_generate()


class DocxGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"IPD Document Generator {APP_VERSION}")
        self.geometry("750x650")
        self.configure(bg="#1a1a2e")

        self.config_file = os.path.expanduser("~/.er_ai_docx_config.json")
        self.base_save_dir = self.load_config()
        self.last_generated_dir = None

        subtitle = tk.Label(self, text="Paste details according to Admission Order template below, then click Generate.",
                            font=("Arial", 10), bg="#1a1a2e", fg="#888")
        subtitle.pack(pady=(12, 8))

        # Text input area
        text_frame = tk.Frame(self, bg="#1a1a2e")
        text_frame.pack(fill="both", expand=True, padx=14, pady=4)

        self.txt_input = tk.Text(text_frame, wrap="word", font=("Consolas", 10),
                                 bg="#0d1117", fg="#c9d1d9", insertbackground="#58a6ff",
                                 selectbackground="#264f78", relief="flat", bd=0,
                                 padx=10, pady=10)
        self.txt_input.pack(fill="both", expand=True)

        # Button row 1 (Input controls)
        btn_frame1 = tk.Frame(self, bg="#1a1a2e")
        btn_frame1.pack(fill="x", padx=14, pady=(8, 4))

        btn_paste = tk.Button(btn_frame1, text="Paste from Clipboard", command=self.paste_clipboard,
                              bg="#21262d", fg="#c9d1d9", font=("Arial", 10, "bold"),
                              activebackground="#30363d", activeforeground="#fff",
                              relief="flat", padx=12, pady=6, cursor="hand2")
        btn_paste.pack(side="left", padx=4)
        btn_manual = tk.Button(btn_frame1, text="Manual Entry Form", command=self.open_manual_form,
                               bg="#2ea043", fg="#ffffff", font=("Arial", 10, "bold"),
                               relief="flat", padx=12, pady=6, cursor="hand2")
        btn_manual.pack(side="left", padx=4)


        btn_clear = tk.Button(btn_frame1, text="Clear", command=self.clear_input,
                              bg="#21262d", fg="#c9d1d9", font=("Arial", 10),
                              activebackground="#30363d", activeforeground="#fff",
                              relief="flat", padx=12, pady=6, cursor="hand2")
        btn_clear.pack(side="left", padx=4)

        self.btn_change_dir = tk.Button(btn_frame1, text="Set Default Save Folder", command=self.change_save_dir,
                                        bg="#21262d", fg="#c9d1d9", font=("Arial", 10),
                                        activebackground="#30363d", activeforeground="#fff",
                                        relief="flat", padx=12, pady=6, cursor="hand2")
        self.btn_change_dir.pack(side="right", padx=4)

        # Button row 2 (Action controls)
        btn_frame2 = tk.Frame(self, bg="#1a1a2e")
        btn_frame2.pack(fill="x", padx=14, pady=(4, 8))

        self.btn_open_folder = tk.Button(btn_frame2, text="Open Last Folder", command=self.open_last_folder,
                                         bg="#d29922", fg="#ffffff", font=("Arial", 11, "bold"),
                                         activebackground="#e3b341", activeforeground="#fff",
                                         relief="flat", padx=16, pady=6, cursor="hand2", state="disabled")
        self.btn_open_folder.pack(side="left", padx=4)

        btn_template = tk.Button(btn_frame2, text="View Template Format", command=self.show_template_guide,
                                 bg="#21262d", fg="#c9d1d9", font=("Arial", 10),
                                 activebackground="#30363d", activeforeground="#fff",
                                 relief="flat", padx=12, pady=6, cursor="hand2")
        btn_template.pack(side="left", padx=4)

        btn_gen = tk.Button(btn_frame2, text="Generate Documents", command=self.run_generate,
                            bg="#238636", fg="#ffffff", font=("Arial", 11, "bold"),
                            activebackground="#2ea043", activeforeground="#fff",
                            relief="flat", padx=16, pady=6, cursor="hand2")
        btn_gen.pack(side="right", padx=4)

        # Status bar
        self.status_var = tk.StringVar(value=f" Ready, Paste output and click Generate  |  {APP_VERSION}")
        status_bar = tk.Label(self, textvariable=self.status_var, font=("Arial", 9),
                              bg="#161b22", fg="#8b949e", anchor="w", padx=10, pady=4)
        status_bar.pack(fill="x", side="bottom")

        # Check for updates in background
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def show_template_guide(self):
        guide_win = tk.Toplevel(self)
        guide_win.title("IPD Document Template Guide")
        guide_win.geometry("600x650")
        guide_win.configure(bg="#1a1a2e")

        lbl = tk.Label(guide_win, text="How to format data for this generator:",
                       font=("Arial", 12, "bold"), bg="#1a1a2e", fg="#e0e0e0")
        lbl.pack(pady=(12, 4))

        desc = tk.Label(guide_win, text="Fill this template manually and paste it into the main window.",
                        font=("Arial", 10), bg="#1a1a2e", fg="#888")
        desc.pack(pady=(0, 8))

        txt = tk.Text(guide_win, wrap="word", font=("Consolas", 9),
                      bg="#0d1117", fg="#c9d1d9", insertbackground="#58a6ff",
                      relief="flat", bd=0, padx=10, pady=10)
        txt.pack(fill="both", expand=True, padx=14, pady=4)

        template_text = """[Pt_Name]
Name Surname
[/Pt_Name]

[Pt_Age]
50
[/Pt_Age]

[Pt_HN]
1234567
[/Pt_HN]

[PL_Order]
- Problem 1
[/PL_Order]

[One_Order]
- Record V/S
[/One_Order]

[Cont_Order]
- Paracetamol PRN
[/Cont_Order]

[HP_Data]
[CC]: Chief Complaint
[PI1]: Present Illness 1
[PI2]: Present Illness 2
[UD]: ไม่มี
[FDA]: ปฏิเสธแพ้ยา
[BW]: 60
[BT]: 37.0
[PR]: 80
[RR]: 20
[SBP]: 120
[DBP]: 80
[GA]: Good consciousness
[HEENT]: Normal
[CVS]: Normal
[RS]: Normal
[Abd]: Normal
[Ext]: Normal
[CNS]: Normal
[PL]: Problem List
[PDx]: Primary Diagnosis
[Plan]: Admit for management
[/HP_Data]

[Consent_Data]
[Dx_Consent]: Primary Diagnosis
[Purpose_Consent]: Mx, Ix
[Tx_Consent]: IV, Med
[Pros_Consent]: เพื่อการรักษา
[Cons_Consent]: อาจแพ้ยา
[Alt_Consent]: ปฏิเสธการรักษา
[Outcome_Consent]: อาการดีขึ้น
[Duration_Consent]: 3-5 วัน
[Risk_Consent]: Infect, FDA
[/Consent_Data]"""
        
        txt.insert("1.0", template_text)

        def copy_tmpl():
            self.clipboard_clear()
            self.clipboard_append(txt.get("1.0", tk.END).strip())
            messagebox.showinfo("Copied", "Template copied to clipboard!", parent=guide_win)

        btn_copy = tk.Button(guide_win, text="Copy Template to Clipboard", command=copy_tmpl,
                             bg="#2ea043", fg="#ffffff", font=("Arial", 10, "bold"),
                             relief="flat", padx=12, pady=6, cursor="hand2")
        btn_copy.pack(pady=10)


    def open_manual_form(self):
        ManualEntryWindow(self)

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get("base_save_dir")
        except Exception:
            pass
        return None

    def save_config(self, path):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({"base_save_dir": path}, f)
        except Exception:
            pass

    def change_save_dir(self):
        initial = self.base_save_dir if self.base_save_dir else os.path.expanduser("~/Desktop")
        new_dir = filedialog.askdirectory(title="Select Default Base Folder", initialdir=initial)
        if new_dir:
            self.base_save_dir = new_dir
            self.save_config(new_dir)
            self.status_var.set(f"Default save folder updated: {new_dir}")
            messagebox.showinfo("Folder Saved", f"Default save folder is now:\n{new_dir}")

    def open_last_folder(self):
        if not self.last_generated_dir or not os.path.exists(self.last_generated_dir):
            return
        
        # Cross-platform folder opening
        if os.name == 'nt': # Windows
            os.startfile(self.last_generated_dir)
        elif sys.platform == 'darwin': # Mac
            subprocess.Popen(['open', self.last_generated_dir])
        else: # Linux
            subprocess.Popen(['xdg-open', self.last_generated_dir])


    def check_for_updates(self):
        try:
            import urllib.request
            import json
            
            url = "https://api.github.com/repos/newnavapol/offline_docx/releases/latest"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode())
            
            latest_version = data.get("tag_name", "")
            if latest_version and latest_version != APP_VERSION:
                assets = data.get("assets", [])
                download_url = None
                filename = None
                
                target = ".exe" if os.name == 'nt' else "_Mac.zip"
                    
                for asset in assets:
                    if asset.get("name", "").endswith(target):
                        download_url = asset.get("browser_download_url")
                        filename = asset.get("name")
                        break
                
                if download_url:
                    self.after(0, lambda: self.prompt_update(latest_version, download_url, filename))
        except Exception as e:
            print(f"Update check failed: {e}")

    def prompt_update(self, latest_version, download_url, filename):
        if messagebox.askyesno("Update Available", f"Version {latest_version} is available!\n\nWould you like to download it to your Downloads folder now?"):
            self.download_update(download_url, filename)

    def download_update(self, url, filename):
        self.status_var.set(f" Downloading {filename}...")
        self.update_idletasks()
        
        def do_download():
            try:
                import urllib.request
                import os
                
                downloads_dir = str(Path.home() / "Downloads")
                save_path = os.path.join(downloads_dir, filename)
                
                urllib.request.urlretrieve(url, save_path)
                
                self.after(0, lambda: self.finish_update(save_path, downloads_dir))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Download Error", f"Failed to download: {e}"))
                self.after(0, lambda: self.status_var.set(f" Ready, Paste output and click Generate  |  {APP_VERSION}"))
                
        threading.Thread(target=do_download, daemon=True).start()

    def finish_update(self, save_path, downloads_dir):
        self.status_var.set(f" Download complete!  |  {APP_VERSION}")
        messagebox.showinfo("Download Complete", f"Saved to:\n{save_path}\n\nPlease close this app and open the new version.")
        
        if os.name == 'nt':
            os.startfile(downloads_dir)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', downloads_dir])
        else:
            subprocess.Popen(['xdg-open', downloads_dir])

    def paste_clipboard(self):
        try:
            text = self.clipboard_get()
            self.txt_input.delete("1.0", tk.END)
            self.txt_input.insert("1.0", text)
            self.status_var.set(f"Pasted {len(text)} characters from clipboard.")
        except Exception:
            messagebox.showwarning("Clipboard", "Clipboard is empty or contains non-text data.")

    def clear_input(self):
        self.txt_input.delete("1.0", tk.END)
        self.status_var.set("Cleared.")

    def run_generate(self):
        text = self.txt_input.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Input Required", "Please paste the Admission Order text before generating.")
            return

        # Ensure base save directory is selected
        if not self.base_save_dir or not os.path.exists(self.base_save_dir):
            messagebox.showinfo("Select Save Location", "Please select a base folder to save your DOCX files.\n\nYou only need to do this once.")
            new_dir = filedialog.askdirectory(title="Select Default Base Folder", initialdir=os.path.expanduser("~/Desktop"))
            if not new_dir:
                return
            self.base_save_dir = new_dir
            self.save_config(new_dir)

        # Extract HN
        hn = extract_multiline_tag(text, "Pt_HN", "Unregistered")
        if hn == "Unregistered" or not hn:
            hn = extract_tag(text, "HN", 20, "Unregistered")
        safe_hn = re.sub(r'[^a-zA-Z0-9]', '_', hn)

        # Extract Primary Dx
        dx = extract_first_diagnosis(text, max_len=100)
        if not dx:
            dx = "Admit"
        safe_dx = re.sub(r'[^a-zA-Z0-9\-\s]', '_', clean_pdx(dx)).strip().replace(' ', '_')

        # Build dynamic subfolder: YYYY/MM/DD/HN_Dx
        now = datetime.now()
        year = now.strftime("%Y")
        month = now.strftime("%m")
        day = now.strftime("%d")
        folder_name = f"{safe_hn}_{safe_dx}"

        out_dir = os.path.join(self.base_save_dir, year, month, day, folder_name)
        os.makedirs(out_dir, exist_ok=True)
        self.last_generated_dir = out_dir

        generated = []
        errors = []

        # Generate Doctor Order
        try:
            path = generate_doctor_order(text, out_dir)
            generated.append(os.path.basename(path))
        except Exception as e:
            errors.append(f"Doctor Order: {e}")

        # Generate Inpatient H&P
        try:
            path = generate_inpatient_hp(text, out_dir)
            generated.append(os.path.basename(path))
        except Exception as e:
            errors.append(f"Inpatient H&P: {e}")

        # Generate Informed Consent
        try:
            path = generate_informed_consent(text, out_dir)
            generated.append(os.path.basename(path))
        except Exception as e:
            errors.append(f"Informed Consent: {e}")

        # Show results and enable open folder button
        if generated:
            self.btn_open_folder.config(state="normal")
            
            file_list = "\n".join(f"  - {f}" for f in generated)
            error_list = "\n".join(f"  [ERROR] {e}" for e in errors) if errors else ""
            msg = f"Generated {len(generated)} file(s) in:\n{out_dir}\n\n{file_list}"
            if error_list:
                msg += f"\n\nErrors:\n{error_list}"
            
            self.status_var.set(f"Saved to {year}/{month}/{day}/{folder_name}")
            
            # Automatically open the folder
            self.open_last_folder()
        else:
            messagebox.showerror("Error", f"Failed to generate any files.\n\n" + "\n".join(errors))
            self.status_var.set("Generation failed.")


if __name__ == "__main__":
    app = DocxGeneratorApp()
    app.mainloop()
