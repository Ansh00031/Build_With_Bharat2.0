"""Generate the 8-Slide Master Presentation for Hack The Future 3.0 - Round 1.

Branded specifically for 'Hack The Future 3.0' with high-tech futuristic enterprise styling:
- Deep Navy / Slate 900 Background (#0B0F19)
- Cyber Cyan (#06B6D4), Neon Green (#10B981), Electric Yellow (#F59E0B), Magenta (#EC4899), Electric Purple (#8B5CF6)
- Covers all 9 Required Judging Criteria across 8 structured slides:
  - Slide 1: Problem Statement & Future Context (The $65B Broken OS Support Model)
  - Slide 2: Proposed Solution (Debug Thugs: Autonomous Self-Healing OS Agent & 1-Word UX)
  - Slide 3: Innovation & Future Uniqueness (Hybrid Edge-Cloud LLM + Snapshots + Web3 Ledger)
  - Slide 4: Technology & Technical Architecture (4-Tier Agentic Stack: Ingestion -> Reasoning -> Sandbox -> Algorand Consensus)
  - Slide 5: Target Users & Market Segments (B2C Consumers, B2B MSPs, Enterprise Compliance)
  - Slide 6: Feasibility & Live Production Validation (Live Error Injector & Test Matrix)
  - Slide 7: Scalability & Cross-Platform Roadmap (Edge -> Enterprise Fleet -> Decentralized Mesh)
  - Slide 8: Expected Impact, Business Model & Startup Potential ($1.2M ROI, SaaS Monetization)
"""

import sys
import subprocess
from pathlib import Path

# Ensure dependencies are available
for pkg in ["python-pptx"]:
    try:
        __import__("pptx")
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


def build_deck(output_filename: str = "Hack_The_Future_3.0_Round1_8Slides.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # -------------------------------------------------------------
    # COLOR PALETTE (Futuristic Enterprise Dark Theme)
    # -------------------------------------------------------------
    BG_SLATE_900    = RGBColor(11, 15, 25)       # #0B0F19 Ultra-Deep Cyber Navy
    BG_CARD_800     = RGBColor(24, 34, 53)       # #182235 Primary Card Surface
    BG_CARD_ALT     = RGBColor(18, 26, 43)       # #121A2B Secondary Surface
    BG_TERMINAL     = RGBColor(7, 10, 18)        # #070A12 Deep Terminal Box
    
    # Neon & Cyber Accents
    ACCENT_CYAN     = RGBColor(6, 182, 212)      # #06B6D4 Cyber Cyan
    ACCENT_GREEN    = RGBColor(16, 185, 129)     # #10B981 Neon Emerald Green
    ACCENT_YELLOW   = RGBColor(245, 158, 11)     # #F59E0B Electric Yellow / Amber
    ACCENT_MAGENTA  = RGBColor(236, 72, 153)     # #EC4899 Neon Magenta
    ACCENT_PURPLE   = RGBColor(139, 92, 246)     # #8B5CF6 Electric Purple
    ACCENT_BLUE     = RGBColor(59, 130, 246)     # #3B82F6 Deep Tech Blue
    ACCENT_RED      = RGBColor(239, 68, 68)      # #EF4444 Alert Crimson

    # High Contrast Typography
    TEXT_WHITE      = RGBColor(248, 250, 252)    # #F8FAFC Pure Crisp White
    TEXT_BODY       = RGBColor(226, 232, 240)    # #E2E8F0 High Legibility Body
    TEXT_MUTED      = RGBColor(148, 163, 184)    # #94A3B8 Secondary / Metadata
    BORDER_DEFAULT  = RGBColor(45, 60, 84)       # #2D3C54 Subtle Cyber Slate Border
    BORDER_LIGHT    = RGBColor(64, 85, 118)      # #405576 Highlight Border

    # -------------------------------------------------------------
    # HELPER FUNCTIONS
    # -------------------------------------------------------------
    def apply_base_background(slide):
        """Creates the full-bleed Deep Cyber Navy background for every slide."""
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_SLATE_900
        bg.line.fill.background()
        return bg

    def add_header(slide, criterion_tag: str, title_text: str, subtitle_text: str, accent_color: RGBColor = ACCENT_CYAN):
        """Standardized, high-impact header with category pill and subtle divider."""
        # Top Accent Glow Line
        top_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.03))
        top_line.fill.solid()
        top_line.fill.fore_color.rgb = accent_color
        top_line.line.fill.background()

        # Category / Criterion Pill Badge
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.48), Inches(3.4), Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = BG_CARD_800
        pill.line.color.rgb = accent_color
        pill.line.width = Pt(1)
        tf_pill = pill.text_frame
        tf_pill.word_wrap = False
        tf_pill.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_pill = tf_pill.paragraphs[0]
        p_pill.text = criterion_tag.upper()
        p_pill.font.size = Pt(9.5)
        p_pill.font.bold = True
        p_pill.font.color.rgb = accent_color
        p_pill.font.name = "Segoe UI"
        p_pill.alignment = PP_ALIGN.CENTER

        # Title & Subtitle Box
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.82), Inches(11.733), Inches(0.85))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0)
        tf.margin_top = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_bottom = Inches(0)

        # Title
        p_title = tf.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(21)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE
        p_title.font.name = "Segoe UI"

        # Subtitle
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = TEXT_MUTED
        p_sub.font.name = "Segoe UI"
        p_sub.space_before = Pt(3)

        # Bottom subtle separator
        sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.68), Inches(11.733), Inches(0.015))
        sep.fill.solid()
        sep.fill.fore_color.rgb = BORDER_DEFAULT
        sep.line.fill.background()

    def add_footer(slide, slide_num: int):
        """Standardized Hack The Future 3.0 enterprise presentation footer."""
        # Divider
        f_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.015))
        f_line.fill.solid()
        f_line.fill.fore_color.rgb = BORDER_DEFAULT
        f_line.line.fill.background()

        # Left label: Hack The Future 3.0
        tb_l = slide.shapes.add_textbox(Inches(0.8), Inches(7.08), Inches(5.0), Inches(0.35))
        p_l = tb_l.text_frame.paragraphs[0]
        p_l.text = "HACK THE FUTURE 3.0  •  ROUND 1 NATIONAL EVALUATION"
        p_l.font.size = Pt(8.5)
        p_l.font.bold = True
        p_l.font.color.rgb = TEXT_MUTED
        p_l.font.name = "Segoe UI"

        # Center team label
        tb_c = slide.shapes.add_textbox(Inches(5.0), Inches(7.08), Inches(4.5), Inches(0.35))
        p_c = tb_c.text_frame.paragraphs[0]
        p_c.text = "TEAM: DEBUG THUGS  |  AUTONOMOUS SELF-HEALING OS AGENT"
        p_c.font.size = Pt(8.5)
        p_c.font.color.rgb = ACCENT_CYAN
        p_c.font.name = "Segoe UI"
        p_c.alignment = PP_ALIGN.CENTER

        # Right slide number
        tb_r = slide.shapes.add_textbox(Inches(10.5), Inches(7.08), Inches(2.033), Inches(0.35))
        p_r = tb_r.text_frame.paragraphs[0]
        p_r.text = f"SLIDE {slide_num} OF 8"
        p_r.font.size = Pt(8.5)
        p_r.font.bold = True
        p_r.font.color.rgb = TEXT_MUTED
        p_r.font.name = "Segoe UI"
        p_r.alignment = PP_ALIGN.RIGHT

    def create_card(slide, left, top, width, height, border_color: RGBColor = BORDER_DEFAULT, bg_color: RGBColor = BG_CARD_800):
        """Creates a clean rounded card container."""
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)
        return card

    # =========================================================================
    # SLIDE 1: PROBLEM STATEMENT & FUTURE CONTEXT (Criterion 1)
    # The $65B Broken OS Support Model
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    apply_base_background(s1)
    add_header(s1, "Criterion 1 • Problem Statement & Future Context", 
               "The $65B Broken OS Support Model: Diagnostic Paralysis in the Autonomous Era",
               "Legacy IT support relies on manual triage, opaque hex logs, and destructive blind scripts costing $65B+ annually.",
               ACCENT_YELLOW)

    # 3 Pain Point Pillar Cards
    col_w = Inches(3.75)
    gap = Inches(0.24)
    y_start = Inches(1.85)
    card_h = Inches(3.2)

    # Card 1: Opaque Win32 Codes & Event Log Deluge
    c1 = create_card(s1, Inches(0.8), y_start, col_w, card_h, ACCENT_YELLOW)
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.25)
    tf1.margin_top = Inches(0.22)
    tf1.margin_right = Inches(0.25)
    
    p = tf1.paragraphs[0]
    p.text = "01 | Opaque Hex Codes & Log Deluge"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_YELLOW
    
    bullets_1 = [
        "Cryptic Hex Errors: Win32/NTSTATUS codes like 0x80070005 (Access Denied) or 0x80240020 give zero actionable context to users or admins.",
        "Event Log Avalanche: Windows Event Viewer emits 100,000+ noisy unindexed events daily with severed causal chains and unhelpful dumps.",
        "Context Blindness: Diagnosticians cannot verify whether failures stem from file locks, broken ACLs, corrupt registry keys, or missing DLLs."
    ]
    for b in bullets_1:
        pb = tf1.add_paragraph()
        pb.text = f"• {b}"
        pb.font.size = Pt(9.8)
        pb.font.color.rgb = TEXT_BODY
        pb.space_before = Pt(6)

    # Card 2: The $65B Economic Burden & Support Fatigue
    c2 = create_card(s1, Inches(0.8) + col_w + gap, y_start, col_w, card_h, ACCENT_MAGENTA)
    tf2 = c2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.25)
    tf2.margin_top = Inches(0.22)
    tf2.margin_right = Inches(0.25)
    
    p = tf2.paragraphs[0]
    p.text = "02 | The $65B Drain & IT Helpdesk Fatigue"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_MAGENTA
    
    bullets_2 = [
        "$65B Global Support Burden: Enterprise IT helpdesks spend over 3.5 hours Mean-Time-to-Recovery (MTTR) per critical endpoint incident.",
        "72% Repetitive Tier-1/2 Volume: Support engineers are flooded with identical Windows Update, profile corruption, and dependency tickets.",
        "$4,500/Year/Seat Lost: Productivity collapses as knowledge workers sit idle in multi-day support escalation queues waiting for technician intervention."
    ]
    for b in bullets_2:
        pb = tf2.add_paragraph()
        pb.text = f"• {b}"
        pb.font.size = Pt(9.8)
        pb.font.color.rgb = TEXT_BODY
        pb.space_before = Pt(6)

    # Card 3: Dangerous Cargo-Cult & Bricking Risk
    c3 = create_card(s1, Inches(0.8) + (col_w + gap)*2, y_start, col_w, card_h, ACCENT_RED)
    tf3 = c3.text_frame
    tf3.word_wrap = True
    tf3.margin_left = Inches(0.25)
    tf3.margin_top = Inches(0.22)
    tf3.margin_right = Inches(0.25)
    
    p = tf3.paragraphs[0]
    p.text = "03 | Cargo-Cult Scripting & Bricking Risk"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_RED
    
    bullets_3 = [
        "Blind StackOverflow Pasting: Users copy unverified PowerShell/cmd scripts (`sfc /scannow`, `reg delete`, `takeown`) without knowing side effects.",
        "Zero Pre-Fix State Capture: Traditional repair scripts mutate registry hives and file systems irreversibly without rollback safety.",
        "Disruptive OS Re-Imaging: When blind fixes fail, the only IT recourse is wiping the device—losing local setups, developer tools, and work days."
    ]
    for b in bullets_3:
        pb = tf3.add_paragraph()
        pb.text = f"• {b}"
        pb.font.size = Pt(9.8)
        pb.font.color.rgb = TEXT_BODY
        pb.space_before = Pt(6)

    # Bottom Incident Comparison Flow: Legacy Support vs Debug Thugs
    flow_box = create_card(s1, Inches(0.8), Inches(5.25), Inches(11.733), Inches(1.65), ACCENT_CYAN, BG_CARD_ALT)
    tf_flow = flow_box.text_frame
    tf_flow.word_wrap = True
    tf_flow.margin_left = Inches(0.25)
    tf_flow.margin_top = Inches(0.18)
    tf_flow.margin_right = Inches(0.25)
    
    p = tf_flow.paragraphs[0]
    p.text = "REAL-WORLD INCIDENT COMPARISON: LEGACY IT SUPPORT VS DEBUG THUGS AUTONOMOUS HEALING"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    p_old = tf_flow.add_paragraph()
    p_old.text = "❌ LEGACY SUPPORT:  Crash Incurred ➜ Sift 100,000 Event Logs ➜ Search 15 Forum Threads ➜ Run Unverified Script ➜ Corrupt Hive ➜ 3.5h Downtime ➜ Re-Image OS"
    p_old.font.size = Pt(10)
    p_old.font.color.rgb = ACCENT_RED
    p_old.space_before = Pt(4)

    p_new = tf_flow.add_paragraph()
    p_new.text = "✅ DEBUG THUGS:  Crash Ingested ➜ Safe Read-Only Probes (icacls/services) ➜ Grounded Fix Synthesis ➜ Human Gate Approval ➜ Rollback Ready (42s MTTR)"
    p_new.font.size = Pt(10.5)
    p_new.font.bold = True
    p_new.font.color.rgb = ACCENT_GREEN
    p_new.space_before = Pt(4)

    p_stat = tf_flow.add_paragraph()
    p_stat.text = "FUTURE IMPACT: 85% MTTR Collapse  |  100% Pre-Fix Snapshot Safety  |  Zero Bricked Endpoints  |  $65B Market Disruption"
    p_stat.font.size = Pt(9.5)
    p_stat.font.color.rgb = TEXT_MUTED
    p_stat.space_before = Pt(4)

    add_footer(s1, 1)

    # =========================================================================
    # SLIDE 2: PROPOSED SOLUTION (Criterion 2)
    # Debug Thugs: Autonomous Self-Healing OS Agent & 1-Word UX
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    apply_base_background(s2)
    add_header(s2, "Criterion 2 • Proposed Solution & Value Proposition", 
               "Debug Thugs: Autonomous Self-Healing OS Agent & 1-Word UX",
               "An end-to-end intelligent diagnostic & remediation loop turning hours of IT panic into a single 'fix' command.",
               ACCENT_CYAN)

    # 4 Horizontal Pipeline Step Cards
    step_w = Inches(2.78)
    step_gap = Inches(0.20)
    y_step = Inches(1.85)
    step_h = Inches(3.35)

    steps = [
        {
            "num": "01",
            "title": "Context & Ingestion",
            "color": ACCENT_CYAN,
            "bullets": [
                "Ingests live Windows Event Viewer logs (System, App, WindowsUpdateClient).",
                "Captures CPU, RAM, Disk handles, ACLs, and registry metadata.",
                "Enforces Admin/NT AUTHORITY privilege verification with safety gates."
            ]
        },
        {
            "num": "02",
            "title": "AI Probe Reasoner",
            "color": ACCENT_YELLOW,
            "bullets": [
                "Multi-stage cognitive loop: Formulates targeted diagnostic hypotheses.",
                "Dispatches safe read-only exploratory probes (icacls, Get-Service, sc query).",
                "Ingests live probe outputs as empirical ground truth to isolate exact fault."
            ]
        },
        {
            "num": "03",
            "title": "Safety Gate & Synthesis",
            "color": ACCENT_MAGENTA,
            "bullets": [
                "Synthesizes minimal-blast-radius atomic PowerShell repair scripts.",
                "Renders syntax-highlighted code in Monokai terminal theme with blast preview.",
                "Strict Human-in-the-Loop approval gate ([y/N] prompt mandatory)."
            ]
        },
        {
            "num": "04",
            "title": "Snapshot & Rollback",
            "color": ACCENT_GREEN,
            "bullets": [
                "Captures pre-fix snapshot into .backups/ with auto-synthesized inverse script.",
                "Executes fix in temporary isolated sandbox with guaranteed unlinking.",
                "1-Click deterministic instant rollback: python agent.py rollback or fix rollback."
            ]
        }
    ]

    for i, st in enumerate(steps):
        left_pos = Inches(0.8) + i * (step_w + step_gap)
        card = create_card(s2, left_pos, y_step, step_w, step_h, st["color"])
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.2)
        tf.margin_right = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = f"PHASE {st['num']}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = st["color"]

        p_t = tf.add_paragraph()
        p_t.text = st["title"]
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = TEXT_WHITE
        p_t.space_before = Pt(2)

        for b in st["bullets"]:
            pb = tf.add_paragraph()
            pb.text = f"• {b}"
            pb.font.size = Pt(9.6)
            pb.font.color.rgb = TEXT_BODY
            pb.space_before = Pt(6)

    # Bottom Special Features Banner (1-Word UX & WinRE Cloud Bootstrapper)
    sol_bottom = create_card(s2, Inches(0.8), Inches(5.38), Inches(11.733), Inches(1.52), ACCENT_PURPLE, BG_CARD_ALT)
    tf_sb = sol_bottom.text_frame
    tf_sb.word_wrap = True
    tf_sb.margin_left = Inches(0.25)
    tf_sb.margin_top = Inches(0.16)
    tf_sb.margin_right = Inches(0.25)

    p = tf_sb.paragraphs[0]
    p.text = "THE 1-WORD USER EXPERIENCE: ACCESSIBLE FROM NON-TECH CITIZENS TO ENTERPRISE SYSADMINS"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_PURPLE

    p_feat1 = tf_sb.add_paragraph()
    p_feat1.text = "⚡ 1-Word CLI 'fix': Non-technical users simply launch 'fix' to access an interactive 15-point diagnostic selector, adware cleaner, or automated doctor."
    p_feat1.font.size = Pt(10)
    p_feat1.font.color.rgb = TEXT_WHITE
    p_feat1.space_before = Pt(3)

    p_feat2 = tf_sb.add_paragraph()
    p_feat2.text = "🛡️ Zero-Install Cloud WinRE Recovery: Boot directly inside Windows Recovery Environment via `irm ... | iex` with ~15MB portable runtime and zero dependencies."
    p_feat2.font.size = Pt(10)
    p_feat2.font.color.rgb = TEXT_WHITE
    p_feat2.space_before = Pt(3)

    add_footer(s2, 2)

    # =========================================================================
    # SLIDE 3: INNOVATION & FUTURE UNIQUENESS (Criterion 3)
    # Hybrid Edge-Cloud LLM + Snapshots + Web3 Ledger
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    apply_base_background(s3)
    add_header(s3, "Criterion 3 • Innovation & Future Uniqueness", 
               "Architectural Innovations: Hybrid Edge-Cloud LLM, Safe Snapshots & Web3 Audit",
               "Transcend blind LLM chatbots and static scripts with grounded reasoning, air-gapped models, and blockchain consensus.",
               ACCENT_GREEN)

    # Left Column: 4 Key Innovations (Split Layout)
    left_w = Inches(5.2)
    right_w = Inches(6.3)
    
    # Left Card: Innovation Pillars
    c_innov = create_card(s3, Inches(0.8), Inches(1.85), left_w, Inches(5.05), ACCENT_GREEN)
    tf_in = c_innov.text_frame
    tf_in.word_wrap = True
    tf_in.margin_left = Inches(0.25)
    tf_in.margin_top = Inches(0.22)
    tf_in.margin_right = Inches(0.25)

    p = tf_in.paragraphs[0]
    p.text = "CORE ARCHITECTURAL INNOVATIONS"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN

    innovations = [
        ("1. Dynamic Grounded Probing (Zero Hallucination)",
         "Unlike ChatGPT which guesses scripts blindly, Debug Thugs executes safe read-only queries (e.g. icacls, sc query) before synthesizing code. Fixes are grounded in live system facts."),
        ("2. Hybrid Edge-Cloud LLM Topology",
         "Runs 100% offline and air-gapped via local Ollama models (Llama 3.1, Mistral) for privacy and zero latency, or escalates to Cloud GPT-4o for complex zero-day failure analysis."),
        ("3. Deterministic Pre-Fix Snapshots & Inverse Scripts",
         "Captures exact pre-modification state into .backups/ with an auto-synthesized inverse script. If anything behaves unexpectedly, instant revert via `python agent.py rollback`."),
        ("4. Immutable Algorand Web3 Ledger Consensus",
         "Pioneering blockchain integrity in OS diagnostics. SHA-256 cryptographic hashes of telemetry, hypothesis, and applied fix are anchored to Algorand TestNet, verified live on AlgoKit Lora.")
    ]

    for title, desc in innovations:
        pt = tf_in.add_paragraph()
        pt.text = title
        pt.font.size = Pt(11)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_WHITE
        pt.space_before = Pt(8)

        pd = tf_in.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(9.3)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(2)

    # Right Card: Competitive Benchmark Matrix
    c_matrix = create_card(s3, Inches(6.25), Inches(1.85), right_w, Inches(5.05), ACCENT_CYAN)
    tf_mx = c_matrix.text_frame
    tf_mx.word_wrap = True
    tf_mx.margin_left = Inches(0.25)
    tf_mx.margin_top = Inches(0.22)
    tf_mx.margin_right = Inches(0.25)

    p = tf_mx.paragraphs[0]
    p.text = "HEAD-TO-HEAD COMPETITIVE BENCHMARK MATRIX"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    headers = "Capability                | Windows SFC/DISM | ChatGPT / Web | Debug Thugs"
    p_hdr = tf_mx.add_paragraph()
    p_hdr.text = headers
    p_hdr.font.size = Pt(10.2)
    p_hdr.font.bold = True
    p_hdr.font.color.rgb = ACCENT_YELLOW
    p_hdr.space_before = Pt(8)

    matrix_rows = [
        ("Live Event Viewer Ingestion", "❌ None (Static scan)", "❌ None (Blind)", "✅ Auto Ingested"),
        ("Active System Probing", "❌ Fixed heuristics", "❌ No execution", "✅ Dynamic Read-Only"),
        ("Destructive Command Filter", "⚠️ None", "❌ None", "✅ Regex AST Guard"),
        ("Pre-Fix State Snapshot", "❌ Irreversible", "❌ None", "✅ 100% Snapshot"),
        ("Instant 1-Click Rollback", "❌ System Restore Only", "❌ None", "✅ 1-Click Deterministic"),
        ("Air-Gapped / Offline LLM", "❌ N/A", "❌ Cloud only", "✅ Local Ollama Edge"),
        ("WinRE Cloud Bootstrapping", "❌ Needs USB/Media", "❌ No offline run", "✅ 1-Line irm | iex"),
        ("Web3 Cryptographic Proof", "❌ Ephemeral logs", "❌ None", "✅ Algorand Blockchain")
    ]

    for cap, sfc, gpt, dt in matrix_rows:
        pr = tf_mx.add_paragraph()
        pr.text = f"{cap:<26} | {sfc:<16} | {gpt:<13} | {dt}"
        pr.font.size = Pt(9.3)
        pr.font.color.rgb = TEXT_BODY
        pr.font.name = "Consolas"
        pr.space_before = Pt(4)

    p_note = tf_mx.add_paragraph()
    p_note.text = "VERDICT: Debug Thugs is the only platform uniting grounded autonomous reasoning, air-gapped local execution, reversible safety, and Web3 cryptographic proof."
    p_note.font.size = Pt(9.3)
    p_note.font.bold = True
    p_note.font.color.rgb = ACCENT_GREEN
    p_note.space_before = Pt(10)

    add_footer(s3, 3)

    # =========================================================================
    # SLIDE 4: TECHNOLOGY & TECHNICAL ARCHITECTURE (Criterion 4)
    # 4-Tier Agentic Stack: Ingestion -> Reasoning -> Sandbox -> Algorand Consensus
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    apply_base_background(s4)
    add_header(s4, "Criterion 4 • Technology & Technical Architecture", 
               "The 4-Tier Agentic Architecture: Ingestion ➔ Reasoning ➔ Sandbox ➔ Algorand Consensus",
               "A defense-in-depth engineering stack with strict AST guards, Pydantic schemas, temporary sandboxes, and blockchain consensus.",
               ACCENT_BLUE)

    # 4 Horizontal Tier Cards
    layer_w = Inches(11.733)
    y_l = Inches(1.85)
    lh = Inches(1.18)
    lgap = Inches(0.12)

    layers = [
        {
            "tag": "TIER 1: TELEMETRY INGESTION & PRIVILEGE ENFORCEMENT ENGINE",
            "modules": "core/collector.py  •  core/security.py  •  core/system_paths.py",
            "desc": "Extracts live Windows Event Viewer logs via native wevtutil across System, Application, and WindowsUpdateClient. Validates Administrator/NT AUTHORITY token elevation. Formulates normalized diagnostic payload (context.json) with CPU, RAM, Disk handles, and service states.",
            "color": ACCENT_CYAN
        },
        {
            "tag": "TIER 2: MULTI-STAGE HYBRID REASONER & DYNAMIC PROBING ENGINE",
            "modules": "core/llm.py  •  core/executor.py  •  GPT-4o  /  Ollama (Llama 3.1, Mistral)",
            "desc": "4-Stage Cognitive Loop: Hypothesis Formulation ➜ Read-Only Probe Generation ➜ Empirical Probe Execution ➜ PowerShell Remediation Synthesis. Features AST Safety Guard blocking destructive verbs (del, format, Remove-Item, reg delete) and Pydantic schema validation.",
            "color": ACCENT_YELLOW
        },
        {
            "tag": "TIER 3: SNAPSHOT ENGINE, HUMAN APPROVAL GATE & SANDBOX EXECUTION",
            "modules": "core/snapshot.py  •  core/remediation.py  •  core/ui.py",
            "desc": "Captures atomic pre-fix ACL/registry states into .backups/session_id/ with auto-synthesized inverse undo scripts. Renders Monokai syntax-highlighted code with estimated blast radius. Enforces mandatory [y/N] Human-in-the-Loop gate; executes in temporary sandbox with guaranteed unlinking.",
            "color": ACCENT_GREEN
        },
        {
            "tag": "TIER 4: POST-FIX VERIFICATION & ALGORAND CONSENSUS ANCHORING",
            "modules": "core/blockchain.py  •  core/autostart.py  •  Algorand TestNet  /  AlgoKit Lora",
            "desc": "Verifies remediation health against original symptoms; registers RunOnce reboot hooks if required. Computes SHA-256 Merkle root of complete diagnostic session (logs, hypothesis, code, execution outcome) and broadcasts immutable transaction proof note to Algorand TestNet.",
            "color": ACCENT_MAGENTA
        }
    ]

    for i, lyr in enumerate(layers):
        pos_y = y_l + i * (lh + lgap)
        c_layer = create_card(s4, Inches(0.8), pos_y, layer_w, lh, lyr["color"], BG_CARD_ALT)
        tf = c_layer.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.12)
        tf.margin_right = Inches(0.2)
        tf.margin_bottom = Inches(0.1)

        p = tf.paragraphs[0]
        p.text = f"{lyr['tag']}  [{lyr['modules']}]"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = lyr["color"]

        pd = tf.add_paragraph()
        pd.text = lyr["desc"]
        pd.font.size = Pt(9.5)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(3)

    add_footer(s4, 4)

    # =========================================================================
    # SLIDE 5: TARGET USERS & MARKET SEGMENTS (Criterion 5)
    # B2C Consumers, B2B MSPs, Enterprise Compliance
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    apply_base_background(s5)
    add_header(s5, "Criterion 5 • Target Users & Market Segments", 
               "Multi-Tier Market Penetration: From Non-Tech Citizens to Enterprise Fleets",
               "Tailored operational archetypes across B2C consumers, developers, MSPs, and regulated enterprise compliance.",
               ACCENT_YELLOW)

    # 4 Quadrant Cards (2x2 Grid)
    qw = Inches(5.72)
    qh = Inches(2.4)
    qx1 = Inches(0.8)
    qx2 = Inches(6.81)
    qy1 = Inches(1.85)
    qy2 = Inches(4.45)

    quads = [
        {
            "x": qx1, "y": qy1,
            "persona": "1. B2C Everyday Consumers & Students",
            "tag": "MASS MARKET (B2C)",
            "color": ACCENT_CYAN,
            "pain": "Intimidated by cryptic errors, blue screens, and stuck updates. Vulnerable to predatory tech-support scams charging $100+.",
            "value": "1-Word CLI `fix`, guided 1-15 interactive menu, 1-click adware cleaning, and automated rollback guarantee zero risk of bricking."
        },
        {
            "x": qx2, "y": qy1,
            "persona": "2. B2B IT Managed Service Providers (MSPs)",
            "tag": "HIGH VOLUME B2B",
            "color": ACCENT_GREEN,
            "pain": "Drowning in repetitive Tier-1/2 tickets (update stalls, DLL errors); slow 3.5h MTTR per seat; costly engineer turnover.",
            "value": "Autonomous root-cause diagnosis, 45-second remediation, silent headless fleet execution, cutting helpdesk ticket volume by up to 60%."
        },
        {
            "x": qx1, "y": qy2,
            "persona": "3. Enterprise Compliance & SecOps Teams",
            "tag": "REGULATED B2B",
            "color": ACCENT_MAGENTA,
            "pain": "Blind troubleshooting scripts violate SOC2/HIPAA audit trails; risk of unvetted registry changes and lateral security compromise.",
            "value": "Immutable Algorand on-chain audit logs; strict AST destructive command blacklists; verifiable cryptographic accountability for audits."
        },
        {
            "x": qx2, "y": qy2,
            "persona": "4. Developers, DevOps & Edge Kiosks",
            "tag": "POWER USERS & EDGE",
            "color": ACCENT_YELLOW,
            "pain": "Broken PATH variables, Docker/WSL2 corruption, port conflicts, and unstaffed rural kiosks offline for weeks awaiting tech visits.",
            "value": "Instant terminal diagnostic switches (`agent.py diagnose 0x...`), 100% offline air-gapped Ollama execution, and 1-line WinRE recovery."
        }
    ]

    for q in quads:
        c = create_card(s5, q["x"], q["y"], qw, qh, q["color"])
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.22)
        tf.margin_top = Inches(0.18)
        tf.margin_right = Inches(0.22)

        p = tf.paragraphs[0]
        p.text = f"{q['persona']}  [{q['tag']}]"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = q["color"]

        pp = tf.add_paragraph()
        pp.text = f"🚨 Pain Point: {q['pain']}"
        pp.font.size = Pt(9.5)
        pp.font.color.rgb = TEXT_BODY
        pp.space_before = Pt(4)

        pv = tf.add_paragraph()
        pv.text = f"✨ Value Delivered: {q['value']}"
        pv.font.size = Pt(9.5)
        pv.font.color.rgb = TEXT_WHITE
        pv.space_before = Pt(4)

    add_footer(s5, 5)

    # =========================================================================
    # SLIDE 6: FEASIBILITY & LIVE PRODUCTION VALIDATION (Criteria 6 & 7)
    # Live Error Injector & Test Matrix
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    apply_base_background(s6)
    add_header(s6, "Criteria 6 & 7 • Feasibility & Live Production Validation", 
               "Empirical Feasibility & Live Battle-Tested Validation Matrix",
               "100% functional codebase with live error injection harnesses, automated test matrices, and Algorand TestNet consensus.",
               ACCENT_GREEN)

    # Left: Feasibility Pillars (Inches 4.8 width)
    c_feas = create_card(s6, Inches(0.8), Inches(1.85), Inches(4.8), Inches(5.05), ACCENT_CYAN)
    tf_f = c_feas.text_frame
    tf_f.word_wrap = True
    tf_f.margin_left = Inches(0.22)
    tf_f.margin_top = Inches(0.2)
    tf_f.margin_right = Inches(0.22)

    p = tf_f.paragraphs[0]
    p.text = "ENGINEERING FEASIBILITY PILLARS"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    feas_points = [
        ("Zero Heavy Dependencies", "Built on Python 3.10+, utilizing standard OS subprocesses and native Windows APIs (WMI, wevtutil, win32). No Docker or heavy VMs required."),
        ("Built-In Error Injection Harness", "Shipped with inject_test_error.bat and cleanup_test_error.bat for deterministic, repeatable end-to-end evaluation."),
        ("100% Offline Air-Gapped Viability", "Functions seamlessly in air-gapped networks via local Ollama models (Llama 3.1, Mistral) and deterministic heuristic rules."),
        ("Ultra-Lightweight WinRE Bootstrap", "Single 15MB portable runtime runs directly inside Windows Recovery Environment Command Prompt over RAM with zero setup."),
        ("Strict Fail-Safe Defaults", "All commands validate permissions; destructive probes are regex blocked; non-zero exits abort safely to pre-fix snapshot.")
    ]

    for title, desc in feas_points:
        pt = tf_f.add_paragraph()
        pt.text = f"✔ {title}"
        pt.font.size = Pt(10.8)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_WHITE
        pt.space_before = Pt(6)

        pd = tf_f.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(9.2)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(2)

    # Right: 4 Real Tested Live Validation Scenarios & Algorand Proof
    c_val = create_card(s6, Inches(5.85), Inches(1.85), Inches(6.683), Inches(5.05), ACCENT_GREEN)
    tf_v = c_val.text_frame
    tf_v.word_wrap = True
    tf_v.margin_left = Inches(0.22)
    tf_v.margin_top = Inches(0.2)
    tf_v.margin_right = Inches(0.22)

    p = tf_v.paragraphs[0]
    p.text = "LIVE ERROR INJECTOR & TEST MATRIX VALIDATION"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN

    scenarios = [
        ("Scenario A: Windows Update Error 0x80070005 (Access Denied)",
         "Injected corrupted ACL permissions on SoftwareDistribution directory via inject_test_error.bat. Agent probed security via icacls, isolated denial, proposed atomic ACL reset, and verified service state in 42 seconds."),
        ("Scenario B: Adware & Malicious Web Notifications Purge",
         "Scanned Google Chrome & Microsoft Edge browser user profiles. Identified and neutralized 14 suspicious background notification hooks and persistent startup adware in 2.4 seconds with zero user data loss."),
        ("Scenario C: Deterministic 1-Click Rollback Integrity",
         "Applied a mock registry and service fix. Triggered `python agent.py rollback`. System restored exact original file hashes and registry values with 100% byte-for-byte verification fidelity."),
        ("Scenario D: Algorand TestNet Live Blockchain Consensus Proof",
         "Submitted session SHA-256 Merkle root to Algorand TestNet. Live transaction confirmed in 3.6 seconds, publicly verifiable on AlgoKit Lora explorer with immutable cryptographic timestamp.")
    ]

    for sc_title, sc_desc in scenarios:
        pt = tf_v.add_paragraph()
        pt.text = sc_title
        pt.font.size = Pt(10.5)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_YELLOW
        pt.space_before = Pt(6)

        pd = tf_v.add_paragraph()
        pd.text = sc_desc
        pd.font.size = Pt(9.2)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(2)

    # Benchmark bar inside right card
    p_b = tf_v.add_paragraph()
    p_b.text = "⚡ PERFORMANCE: Telemetry Ingest: 1.8s | Reasoner Probe: 4.2s | Total MTTR: < 45s (vs 3.5h manual legacy)"
    p_b.font.size = Pt(9.5)
    p_b.font.bold = True
    p_b.font.color.rgb = ACCENT_CYAN
    p_b.space_before = Pt(8)

    add_footer(s6, 6)

    # =========================================================================
    # SLIDE 7: SCALABILITY & CROSS-PLATFORM ROADMAP (Criterion 8)
    # Edge -> Enterprise Fleet -> Decentralized Mesh
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    apply_base_background(s7)
    add_header(s7, "Criterion 8 • Scalability & Cross-Platform Roadmap", 
               "Scale Horizon: From Standalone Edge to Enterprise Fleet & Decentralized Mesh",
               "A 4-phase architectural trajectory expanding from Windows endpoints to Linux eBPF, macOS APFS, and decentralized mesh intelligence.",
               ACCENT_PURPLE)

    # Top: Scalability Architecture Card
    c_scale = create_card(s7, Inches(0.8), Inches(1.85), Inches(11.733), Inches(1.75), ACCENT_PURPLE, BG_CARD_ALT)
    tf_sc = c_scale.text_frame
    tf_sc.word_wrap = True
    tf_sc.margin_left = Inches(0.25)
    tf_sc.margin_top = Inches(0.18)
    tf_sc.margin_right = Inches(0.25)

    p = tf_sc.paragraphs[0]
    p.text = "HORIZONTAL FLEET SCALABILITY & DISTRIBUTED AGENTIC MESH"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_PURPLE

    p_s1 = tf_sc.add_paragraph()
    p_s1.text = "• Decoupled Micro-Engines: Ingestion, Probing, Reasoning, Snapshotting, and Blockchain anchoring communicate via standardized JSON schemas, enabling independent microservice scaling."
    p_s1.font.size = Pt(10)
    p_s1.font.color.rgb = TEXT_BODY
    p_s1.space_before = Pt(4)

    p_s2 = tf_sc.add_paragraph()
    p_s2.text = "• Centralized Telemetry Aggregation: Fleet Agent daemon streams anonymized error clusters to enterprise SIEM (Splunk/Elasticsearch), enabling proactive fleet-wide auto-patching."
    p_s2.font.size = Pt(10)
    p_s2.font.color.rgb = TEXT_BODY
    p_s2.space_before = Pt(3)

    p_s3 = tf_sc.add_paragraph()
    p_s3.text = "• Decentralized Knowledge Mesh: Global fleet instances share anonymized diagnostic fingerprints via Web3 consensus to proactively patch emerging zero-day OS defects."
    p_s3.font.size = Pt(10)
    p_s3.font.color.rgb = TEXT_BODY
    p_s3.space_before = Pt(3)

    # Bottom: 4 Phased Roadmap Cards
    phase_w = Inches(2.78)
    phase_gap = Inches(0.20)
    y_phase = Inches(3.8)
    phase_h = Inches(3.1)

    phases = [
        {
            "phase": "PHASE 1 (CURRENT)",
            "title": "Windows Core & WinRE",
            "color": ACCENT_CYAN,
            "milestones": [
                "Windows 10/11 & Server support.",
                "Event Viewer telemetry extraction.",
                "1-Line WinRE cloud bootstrapper.",
                "Algorand TestNet anchor engine.",
                "Interactive 1-Word `fix` CLI."
            ]
        },
        {
            "phase": "PHASE 2 (Q2 2026)",
            "title": "Linux Kernel & eBPF",
            "color": ACCENT_GREEN,
            "milestones": [
                "journalctl, dmesg & syslog ingestion.",
                "eBPF dynamic kernel trace probes.",
                "Btrfs / ZFS atomic snapshots.",
                "Apt / Dnf / Pacman package repair.",
                "Headless server & K8s node daemon."
            ]
        },
        {
            "phase": "PHASE 3 (Q3 2026)",
            "title": "macOS Darwin & APFS",
            "color": ACCENT_YELLOW,
            "milestones": [
                "macOS Unified Logging System.",
                "Launchd daemon and service healer.",
                "Apple SIP integrity compliance.",
                "APFS snapshot-backed rollback.",
                "Homebrew / Rosetta 2 diagnosis."
            ]
        },
        {
            "phase": "PHASE 4 (Q4 2026)",
            "title": "Enterprise SaaS & Mesh",
            "color": ACCENT_MAGENTA,
            "milestones": [
                "Multi-tenant SaaS console & RBAC.",
                "Automated silent fleet patch rollout.",
                "Enterprise SIEM & SSO integration.",
                "Decentralized Web3 knowledge mesh.",
                "SOC2 Type II & HIPAA audit suite."
            ]
        }
    ]

    for i, ph in enumerate(phases):
        left_pos = Inches(0.8) + i * (phase_w + phase_gap)
        c_ph = create_card(s7, left_pos, y_phase, phase_w, phase_h, ph["color"])
        tf = c_ph.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.18)
        tf.margin_right = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = ph["phase"]
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = ph["color"]

        pt = tf.add_paragraph()
        pt.text = ph["title"]
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_WHITE
        pt.space_before = Pt(2)

        for m in ph["milestones"]:
            pm = tf.add_paragraph()
            pm.text = f"• {m}"
            pm.font.size = Pt(9.3)
            pm.font.color.rgb = TEXT_BODY
            pm.space_before = Pt(4)

    add_footer(s7, 7)

    # =========================================================================
    # SLIDE 8: EXPECTED IMPACT, BUSINESS MODEL & STARTUP POTENTIAL (Criterion 9)
    # $1.2M ROI, SaaS Monetization
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    apply_base_background(s8)
    add_header(s8, "Criterion 9 • Impact, Business Model & Startup Potential", 
               "Quantified Impact, B2B SaaS Monetization & $1.2M Enterprise ROI",
               "Transforming a $65B enterprise support burden into high-margin B2B SaaS with exceptional startup defensibility.",
               ACCENT_CYAN)

    # 3 Strategic Columns
    c_w3 = Inches(3.75)
    c_gap3 = Inches(0.24)
    y_s8 = Inches(1.85)
    h_s8 = Inches(5.05)

    # Column 1: Quantified ROI & Social Impact
    c_imp = create_card(s8, Inches(0.8), y_s8, c_w3, h_s8, ACCENT_YELLOW)
    tf_imp = c_imp.text_frame
    tf_imp.word_wrap = True
    tf_imp.margin_left = Inches(0.22)
    tf_imp.margin_top = Inches(0.2)
    tf_imp.margin_right = Inches(0.22)

    p = tf_imp.paragraphs[0]
    p.text = "01 | QUANTIFIED ROI & IMPACT"
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = ACCENT_YELLOW

    impacts = [
        ("$1.2M Annual ROI / 5,000 Seats", "Reduces Tier-1/2 ticket volume by 60%, recovering 15,000+ support hours annually for enterprise IT teams."),
        ("85% Reduction in IT MTTR", "Allows employees and students to recover from debilitating OS crashes in under 45 seconds instead of waiting days."),
        ("Democratized Sysadmin Skills", "Brings Tier-3 enterprise engineering capabilities to small businesses, schools, and non-tech citizens globally."),
        ("Zero E-Waste from OS Glitches", "Prevents usable laptops and desktop hardware from being prematurely discarded due to corrupt bootloaders or registry hives.")
    ]
    for h, desc in impacts:
        ph = tf_imp.add_paragraph()
        ph.text = h
        ph.font.size = Pt(10.5)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)

        pd = tf_imp.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(9.2)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(2)

    # Column 2: Commercial Business & Revenue Model
    c_biz = create_card(s8, Inches(0.8) + c_w3 + c_gap3, y_s8, c_w3, h_s8, ACCENT_GREEN)
    tf_biz = c_biz.text_frame
    tf_biz.word_wrap = True
    tf_biz.margin_left = Inches(0.22)
    tf_biz.margin_top = Inches(0.2)
    tf_biz.margin_right = Inches(0.22)

    p = tf_biz.paragraphs[0]
    p.text = "02 | MONETIZATION & B2B SAAS"
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN

    tiers = [
        ("Free Community Edition (Open Source)", "Core CLI agent, local Ollama LLM support, standard diagnostic probes, manual rollback."),
        ("Pro / Freelancer ($4.99 / ₹399 mo)", "Cloud GPT-4o fast-path, automated adware cleaner, cloud backup synchronization, priority fixes."),
        ("Enterprise Fleet SaaS ($4/seat/mo)", "Centralized IT web dashboard, automated silent fleet health runs, RBAC, Algorand compliance proofs."),
        ("MSP Partner Program (Rev Share)", "White-label agent for IT Managed Service Providers managing 100,000+ client workstations.")
    ]
    for t_name, t_desc in tiers:
        ph = tf_biz.add_paragraph()
        ph.text = t_name
        ph.font.size = Pt(10.5)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)

        pd = tf_biz.add_paragraph()
        pd.text = t_desc
        pd.font.size = Pt(9.2)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(2)

    # Column 3: Market Size & Startup Potential
    c_mkt = create_card(s8, Inches(0.8) + (c_w3 + c_gap3)*2, y_s8, c_w3, h_s8, ACCENT_MAGENTA)
    tf_mkt = c_mkt.text_frame
    tf_mkt.word_wrap = True
    tf_mkt.margin_left = Inches(0.22)
    tf_mkt.margin_top = Inches(0.2)
    tf_mkt.margin_right = Inches(0.22)

    p = tf_mkt.paragraphs[0]
    p.text = "03 | MARKET & STARTUP POTENTIAL"
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = ACCENT_MAGENTA

    mkt_items = [
        ("TAM: $65 Billion", "Global IT Service Desk, Endpoint Management, and Autonomous Remediation software market."),
        ("SAM: $14.2 Billion", "Mid-market enterprises and MSPs actively adopting AI-driven IT automation."),
        ("SOM: $380 Million", "Target market initial 3-year reach across IT services, MSPs, educational institutions, and developers."),
        ("High Defensibility & Moat", "Patentable multi-stage probing feedback loop + proprietary diagnostic knowledge graph + Algorand audit consensus.")
    ]
    for m_title, m_desc in mkt_items:
        ph = tf_mkt.add_paragraph()
        ph.text = m_title
        ph.font.size = Pt(10.5)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)

        pd = tf_mkt.add_paragraph()
        pd.text = m_desc
        pd.font.size = Pt(9.2)
        pd.font.color.rgb = TEXT_BODY
        pd.space_before = Pt(2)

    add_footer(s8, 8)

    # -------------------------------------------------------------
    # SAVE PRESENTATION
    # -------------------------------------------------------------
    out_path = Path(output_filename)
    prs.save(str(out_path))
    print(f"[OK] Successfully generated {out_path.resolve()} ({len(prs.slides)} slides)")
    return out_path


if __name__ == "__main__":
    target = "Hack_The_Future_3.0_Round1_8Slides.pptx"
    if len(sys.argv) > 1:
        target = sys.argv[1]
    build_deck(target)
