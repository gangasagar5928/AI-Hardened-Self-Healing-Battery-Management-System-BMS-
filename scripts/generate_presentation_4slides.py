import os
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# 1. Copy Blueprint Image to Assets
ASSETS_DIR = r"c:\Users\mksin\Desktop\AI hardened BMS\assets"
os.makedirs(ASSETS_DIR, exist_ok=True)
SRC_IMAGE = r"C:\Users\mksin\.gemini\antigravity\brain\0c7977e6-c5f6-4522-8deb-7765a9c616f0\.user_uploaded\media_1785876165433.jpg"
BLUEPRINT_IMG_PATH = os.path.join(ASSETS_DIR, "blueprint_3d.png")
shutil.copy(SRC_IMAGE, BLUEPRINT_IMG_PATH)
print(f"Copied blueprint image to: {BLUEPRINT_IMG_PATH}")

# 2. Color Palette (Sleek Dark Tech Theme)
BG_COLOR      = RGBColor(10, 17, 30)      # Deep Dark Navy (#0A111E)
CARD_BG       = RGBColor(18, 28, 48)      # Card Background (#121C30)
CARD_BORDER   = RGBColor(30, 58, 95)      # Card Border (#1E3A5F)
CYAN_ACCENT   = RGBColor(0, 210, 255)     # Electric Cyan (#00D2FF)
GREEN_ACCENT  = RGBColor(16, 233, 122)    # Cyber Green (#10E97A)
WHITE         = RGBColor(255, 255, 255)   # Pure White
MUTED_TEXT    = RGBColor(180, 195, 215)   # Muted Blue-Gray (#B4C3D7)
RED_ALERT     = RGBColor(255, 75, 75)     # Alert Red (#FF4B4B)

# 3. Helper Functions
def set_bg(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

def add_header(slide, title_text, category_text="CYBER-HARDENED BMS"):
    # Header container
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.13), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p_cat = tf.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = CYAN_ACCENT
    
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE
    p_title.space_before = Pt(2)

def add_card(slide, left, top, width, height, border_color=CARD_BORDER, bg_color=CARD_BG):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1.5)
    return shape

# 4. Initialize 16:9 Widescreen Presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1: TITLE & PROBLEM STATEMENT
# ─────────────────────────────────────────────────────────────────────────────
s1 = prs.slides.add_slide(blank_layout)
set_bg(s1)

# Main Title Card
add_card(s1, Inches(0.6), Inches(0.6), Inches(12.13), Inches(2.2), border_color=CYAN_ACCENT)
tb1 = s1.shapes.add_textbox(Inches(0.9), Inches(0.8), Inches(11.5), Inches(1.8))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "Self-Healing Cyber-Hardened Battery Management System (BMS)"
p1.font.size = Pt(26); p1.font.bold = True; p1.font.color.rgb = WHITE

p2 = tf1.add_paragraph()
p2.text = "Dual-Core TinyML Intrusion Detection, Adaptive EKF Covariance Scaling & 7-Layer Defense for Indian EVs"
p2.font.size = Pt(14); p2.font.color.rgb = CYAN_ACCENT; p2.space_before = Pt(6)

p3 = tf1.add_paragraph()
p3.text = "GCET EEE Department | Production-Grade Open Source Architecture | Target Cost: ₹2,100 – ₹2,400"
p3.font.size = Pt(11); p3.font.color.rgb = MUTED_TEXT; p3.space_before = Pt(8)

# Problem Statement Header
tb_prob_hdr = s1.shapes.add_textbox(Inches(0.6), Inches(3.0), Inches(12.13), Inches(0.4))
tf_ph = tb_prob_hdr.text_frame
p_ph = tf_ph.paragraphs[0]
p_ph.text = "REAL-WORLD EV SECURITY & TRUST PROBLEM STATEMENT IN INDIA"
p_ph.font.size = Pt(13); p_ph.font.bold = True; p_ph.font.color.rgb = GREEN_ACCENT

# 4 Problem Cards
problems = [
    ("1. Unauthenticated CAN/BLE Buses", "EV 2W/3W buses lack encryption. Attackers flood DoS & spoof false telemetry, causing >18.4% EKF SoC estimation divergence.", RED_ALERT),
    ("2. Zero Resale Trust Infrastructure", "No tamper-evident battery history exists in India. Buyers cannot verify real SoH or abuse history, depressing EV resale value.", CYAN_ACCENT),
    ("3. Battery Pack Theft & Cell Harvesting", "Battery costs 40-50% of EV price. Rising whole-pack theft and unauthorized cell-bypassing with jumper wires.", GREEN_ACCENT),
    ("4. Uncoordinated Grid Stress", "Peak-hour EV charging overloads tier-2/3 city power distribution grids without smart demand-response scheduling.", MUTED_TEXT)
]

for idx, (title, desc, col) in enumerate(problems):
    col_idx = idx % 2
    row_idx = idx // 2
    left = Inches(0.6 + col_idx * 6.16)
    top = Inches(3.5 + row_idx * 1.8)
    
    add_card(s1, left, top, Inches(5.97), Inches(1.6), border_color=col)
    tb_card = s1.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), Inches(5.57), Inches(1.3))
    tf_c = tb_card.text_frame
    tf_c.word_wrap = True
    
    pt = tf_c.paragraphs[0]
    pt.text = title; pt.font.size = Pt(13); pt.font.bold = True; pt.font.color.rgb = col
    
    pd = tf_c.add_paragraph()
    pd.text = desc; pd.font.size = Pt(10.5); pd.font.color.rgb = MUTED_TEXT; pd.space_before = Pt(4)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2: SYSTEM BLUEPRINT ARCHITECTURE (IMAGE SLIDE)
# ─────────────────────────────────────────────────────────────────────────────
s2 = prs.slides.add_slide(blank_layout)
set_bg(s2)
add_header(s2, "3D Hardware Blueprint & Cyber-Hardened System Architecture")

# Embed Blueprint Image (Left Container)
img_left = Inches(0.6)
img_top = Inches(1.4)
img_width = Inches(8.5)
s2.shapes.add_picture(BLUEPRINT_IMG_PATH, img_left, img_top, width=img_width)

# Right Highlights Card Container
right_card = add_card(s2, Inches(9.3), Inches(1.4), Inches(3.43), Inches(5.6), border_color=CYAN_ACCENT)
tb_specs = s2.shapes.add_textbox(Inches(9.5), Inches(1.6), Inches(3.03), Inches(5.2))
tf_s = tb_specs.text_frame
tf_s.word_wrap = True

ps_hdr = tf_s.paragraphs[0]
ps_hdr.text = "HARDWARE CALLOUTS"; ps_hdr.font.size = Pt(14); ps_hdr.font.bold = True; ps_hdr.font.color.rgb = CYAN_ACCENT

callouts = [
    ("🔋 Battery Pack", "4S1P 14.8V Modular (Scalable to 16S 57.6V Daisy-Chain)"),
    ("🛡️ BQ76920 AFE", "TI Cell Balance & Analog Front-End IC"),
    ("🧠 ESP32 Master", "Dual-Core 240MHz (Core 0 ML / Core 1 EKF)"),
    ("⚡ SSR Cutoff", "High-Side Optocoupler P-FET Driver (GPIO 17)"),
    ("⚡ CAN Bus", "Dual SN65HVD230 Transceivers (500 kbps)"),
    ("🔴 Attacker Node", "ESP32 Attack Injector (Modes 1-9)"),
    ("☁️ TCU Gateway", "MQTT Cloud Telematics & Resale Passport")
]

for title, desc in callouts:
    p_t = tf_s.add_paragraph()
    p_t.text = title; p_t.font.size = Pt(11); p_t.font.bold = True; p_t.font.color.rgb = GREEN_ACCENT; p_t.space_before = Pt(6)
    p_d = tf_s.add_paragraph()
    p_d.text = desc; p_d.font.size = Pt(9.5); p_d.font.color.rgb = MUTED_TEXT

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3: IMPLEMENTATION & LOW-COST BOM MATRIX
# ─────────────────────────────────────────────────────────────────────────────
s3 = prs.slides.add_slide(blank_layout)
set_bg(s3)
add_header(s3, "Hardware Implementation & Low-Cost BOM Matrix (Target ₹2,100 – ₹2,400)")

# Left Card: Dual Core Execution Model
add_card(s3, Inches(0.6), Inches(1.4), Inches(5.9), Inches(5.6), border_color=CYAN_ACCENT)
tb_core = s3.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.5), Inches(5.2))
tf_core = tb_core.text_frame
tf_core.word_wrap = True

pc_hdr = tf_core.paragraphs[0]
pc_hdr.text = "DUAL-CORE FREERTOS EXECUTION MODEL"; pc_hdr.font.size = Pt(13); pc_hdr.font.bold = True; pc_hdr.font.color.rgb = CYAN_ACCENT

cores = [
    ("Core 0: Security & AI Engine (240 MHz)", "• TinyML Random Forest IDS (<0.35 ms C++ decision tree)\n• Layer 3 UDS (ISO 14229) 0x7E0 SecurityAccess Inspector\n• Layer 1 BLE ECDH + AES-128 Authentication\n• Federated Learning Model Delta Aggregator (0x188)", GREEN_ACCENT),
    ("Core 1: Deterministic Control & EKF (240 MHz)", "• BQ76920 Cell Voltage & Current Sampling (100 ms)\n• Layer 4 Adaptive EKF State Estimation (2RC ECM)\n• Layer 5 Self-Healing Engine & Passive Cell Balancing\n• Layer 6 High-Side SSR Cutoff Driver (GPIO 17 trip <1.2 ms)", CYAN_ACCENT)
]

for title, desc, col in cores:
    pt = tf_core.add_paragraph()
    pt.text = title; pt.font.size = Pt(11.5); pt.font.bold = True; pt.font.color.rgb = col; pt.space_before = Pt(10)
    pd = tf_core.add_paragraph()
    pd.text = desc; pd.font.size = Pt(10); pd.font.color.rgb = MUTED_TEXT; pd.space_before = Pt(4)

# Right Card: Low-Cost BOM Matrix Table
add_card(s3, Inches(6.8), Inches(1.4), Inches(5.93), Inches(5.6), border_color=GREEN_ACCENT)
tb_bom_hdr = s3.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.5), Inches(0.5))
tf_bh = tb_bom_hdr.text_frame
pbh = tf_bh.paragraphs[0]
pbh.text = "LOW-COST BOM OPTIMIZATION MATRIX"; pbh.font.size = Pt(13); pbh.font.bold = True; pbh.font.color.rgb = GREEN_ACCENT

# Add Table inside Right Container
table_shape = s3.shapes.add_table(7, 3, Inches(7.0), Inches(2.2), Inches(5.53), Inches(4.5))
table = table_shape.table
table.columns[0].width = Inches(2.3)
table.columns[1].width = Inches(2.03)
table.columns[2].width = Inches(1.2)

bom_data = [
    ("Component Category", "Optimized Choice", "Price (INR)"),
    ("Microcontrollers", "1x ESP32-S3 / 2x Bare ICs", "₹690"),
    ("CAN Transceivers", "2x VP230 / TJA1051 ICs", "₹240"),
    ("Battery Pack", "4S1P Reclaimed 18650 Pack", "₹450"),
    ("SSR Cutoff Driver", "Optocoupler + P-FET Switch", "₹90"),
    ("AFE & Sensors", "TI BQ76920 IC + Shunt", "₹800"),
    ("TOTAL BENCHMARK", "Complete Hardware Setup", "₹2,100 – ₹2,400")
]

for row_idx, row in enumerate(bom_data):
    for col_idx, text in enumerate(row):
        cell = table.cell(row_idx, col_idx)
        cell.text = text
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if col_idx < 2 else PP_ALIGN.RIGHT
        p.font.size = Pt(9.5)
        if row_idx == 0:
            p.font.bold = True
            p.font.color.rgb = WHITE
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(30, 58, 95)
        elif row_idx == 6:
            p.font.bold = True
            p.font.color.rgb = GREEN_ACCENT
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(20, 40, 65)
        else:
            p.font.color.rgb = MUTED_TEXT
            cell.fill.solid(); cell.fill.fore_color.rgb = CARD_BG

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4: SYSTEM WORKING & 7-LAYER SECURITY ARCHITECTURE
# ─────────────────────────────────────────────────────────────────────────────
s4 = prs.slides.add_slide(blank_layout)
set_bg(s4)
add_header(s4, "System Working & 7-Layer Defense Stack")

layers = [
    ("L1: BLE Peripheral Auth", "ECDH SECP256R1 + AES-128-GCM + 64-nonce replay cache.", CYAN_ACCENT),
    ("L2: HMAC Command Auth", "HMAC-SHA256 frame signature verification over CAN payloads.", GREEN_ACCENT),
    ("L3: TinyML IDS & UDS Inspector", "Zero-GPU Random Forest C++ decision tree (<0.35ms) + UDS 0x7E0 filter.", CYAN_ACCENT),
    ("L4: Adaptive EKF Scaling", "R_eff = R_base * e^(10*S). Under attack (S->1), K->0 isolates state.", GREEN_ACCENT),
    ("L5: Self-Healing Engine", "5-stage threshold: Normal -> Red Charge -> Disable Bal -> Limp -> Isolate.", CYAN_ACCENT),
    ("L6: High-Side SSR Isolation", "Automated GPIO 17 P-FET gate driver cutoff in <1.2ms during S > 0.90.", RED_ALERT),
    ("L7: Secure Cloud Gateway", "Secure MQTT telemetry & digital twin logging to Node-RED / Grafana.", GREEN_ACCENT)
]

for idx, (title, desc, col) in enumerate(layers):
    top = Inches(1.4 + idx * 0.82)
    add_card(s4, Inches(0.6), top, Inches(12.13), Inches(0.72), border_color=col)
    tb_l = s4.shapes.add_textbox(Inches(0.8), top + Inches(0.08), Inches(11.7), Inches(0.56))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    
    p = tf_l.paragraphs[0]
    r1 = p.add_run(); r1.text = title + " — "; r1.font.bold = True; r1.font.size = Pt(11.5); r1.font.color.rgb = col
    r2 = p.add_run(); r2.text = desc; r2.font.size = Pt(10.5); r2.font.color.rgb = MUTED_TEXT

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5: ADVANCED HIGH-IMPACT NOVELTIES & VALIDATION
# ─────────────────────────────────────────────────────────────────────────────
s5 = prs.slides.add_slide(blank_layout)
set_bg(s5)
add_header(s5, "High-Impact Zero-Cost Novelties & Experimental Results")

novelties = [
    ("📄 Digital Battery Health Passport", "Cryptographically signed SHA-256 digest (0x190) over SoH, cycle count, and thermal events for used-EV resale trust.", CYAN_ACCENT),
    ("🔒 Battery Theft & Tamper Detection", "Layer 3 IDS detects abrupt pack removal (V<4V) and physical cell-bypass jumpering (Delta V > 3.5V).", GREEN_ACCENT),
    ("⚡ Grid-Aware Adaptive Charging", "50% charge current throttling during peak grid stress signals (0x198) without extra hardware.", CYAN_ACCENT),
    ("🤖 Federated Learning Fleet Intelligence", "Local gradient delta sharing (0x188) across deployed BMS units for privacy-preserving fleet updates.", GREEN_ACCENT)
]

# Left 4 Cards
for idx, (title, desc, col) in enumerate(novelties):
    top = Inches(1.4 + idx * 1.35)
    add_card(s5, Inches(0.6), top, Inches(6.8), Inches(1.22), border_color=col)
    tb_n = s5.shapes.add_textbox(Inches(0.8), top + Inches(0.1), Inches(6.4), Inches(1.0))
    tf_n = tb_n.text_frame; tf_n.word_wrap = True
    pt = tf_n.paragraphs[0]; pt.text = title; pt.font.size = Pt(12); pt.font.bold = True; pt.font.color.rgb = col
    pd = tf_n.add_paragraph(); pd.text = desc; pd.font.size = Pt(9.5); pd.font.color.rgb = MUTED_TEXT; pd.space_before = Pt(3)

# Right Card: Experimental Validation Benchmarks
add_card(s5, Inches(7.6), Inches(1.4), Inches(5.13), Inches(5.6), border_color=GREEN_ACCENT)
tb_bench = s5.shapes.add_textbox(Inches(7.8), Inches(1.6), Inches(4.73), Inches(5.2))
tf_b = tb_bench.text_frame; tf_b.word_wrap = True

pb_hdr = tf_b.paragraphs[0]
pb_hdr.text = "EXPERIMENTAL BENCHMARK RESULTS"; pb_hdr.font.size = Pt(13); pb_hdr.font.bold = True; pb_hdr.font.color.rgb = GREEN_ACCENT

benchmarks = [
    ("• SoC Error Under Attack:", "<1.4% (vs >18.4% in unprotected baselines)"),
    ("• TinyML IDS Accuracy:", "98.1% across 6 attack classes (AUC = 0.994)"),
    ("• Inference Latency:", "<0.35 ms on ESP32 Core 0 (Zero GPU overhead)"),
    ("• SSR Isolation Time:", "<1.2 ms hardware trip on GPIO 17 (S > 0.90)"),
    ("• Memory Footprint:", "38.4 KB SRAM footprint (<7.5% ESP32 memory)"),
    ("• Re-convergence Time:", "<300 ms post-attack recovery (3 EKF cycles)")
]

for title, val in benchmarks:
    pt = tf_b.add_paragraph()
    pt.text = title; pt.font.size = Pt(11); pt.font.bold = True; pt.font.color.rgb = CYAN_ACCENT; pt.space_before = Pt(8)
    pd = tf_b.add_paragraph()
    pd.text = val; pd.font.size = Pt(10); pd.font.color.rgb = WHITE; pd.space_before = Pt(2)

# Save Presentation
PPTX_PATH = r"c:\Users\mksin\Desktop\AI hardened BMS\Cyber_Hardened_BMS_Short_Presentation.pptx"
prs.save(PPTX_PATH)
print(f"SUCCESSFULLY GENERATED 5-SLIDE PPTX AT: {PPTX_PATH}")
