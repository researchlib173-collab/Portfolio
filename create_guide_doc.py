import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

doc = Document()

# Page Margins
for section in doc.sections:
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

# Brand Color Palette
C_DARK = RGBColor(14, 15, 23)
C_ORANGE = RGBColor(219, 114, 44)
C_BLUE = RGBColor(30, 64, 175)
C_GRAY = RGBColor(71, 85, 105)
C_TEXT = RGBColor(30, 41, 59)
C_LIGHT_BG = "F8FAFC"
C_HEADER_BG = "DB722C"

# Bookmark Generator
def add_bookmark(paragraph, bookmark_name):
    tag = paragraph._p
    bm_start = parse_xml(r'<w:bookmarkStart {} w:id="{}" w:name="{}"/>'.format(
        nsdecls('w'), abs(hash(bookmark_name)) % 100000, bookmark_name
    ))
    bm_end = parse_xml(r'<w:bookmarkEnd {} w:id="{}"/>'.format(
        nsdecls('w'), abs(hash(bookmark_name)) % 100000
    ))
    tag.append(bm_start)
    tag.append(bm_end)

# Clickable Hyperlink Generator
def add_bookmark_link(paragraph, text, bookmark_name, font_size=9, color="DB722C", bold=True):
    hyperlink = parse_xml(r'<w:hyperlink {} w:anchor="{}" />'.format(
        nsdecls('w'), bookmark_name
    ))
    run = parse_xml(r'<w:r {}><w:rPr><w:rFont w:ascii="Consolas" w:hAnsi="Consolas"/><w:color w:val="{}"/><w:sz w:val="{}"/><w:b/></w:rPr><w:t>{}</w:t></w:r>'.format(
        nsdecls('w'), color, int(font_size * 2), text
    ))
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

# Code Block Styler
def add_code_block(doc, code_text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(7.1)
    
    shading = parse_xml(r'<w:shd {} w:fill="0A0A10"/>'.format(nsdecls('w')))
    cell._tc.get_or_add_tcPr().append(shading)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        r'<w:tcBorders {}><w:top w:val="single" w:sz="4" w:space="0" w:color="222222"/>'
        r'<w:left w:val="single" w:sz="16" w:space="0" w:color="DB722C"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="222222"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="222222"/></w:tcBorders>'.format(nsdecls('w'))
    )
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = Pt(12.5)
    
    run = p.add_run(code_text)
    run.font.name = 'Consolas'
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(226, 232, 240)
    
    sp_p = doc.add_paragraph()
    sp_p.paragraph_format.space_before = Pt(2)
    sp_p.paragraph_format.space_after = Pt(4)

# Section Spec Table (Key-Value)
def add_spec_table(doc, specs):
    table = doc.add_table(rows=len(specs), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [Inches(1.8), Inches(5.3)]
    
    for row_idx, (k, v) in enumerate(specs):
        row = table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = widths[0], widths[1]
        
        shd0 = parse_xml(r'<w:shd {} w:fill="1E293B"/>'.format(nsdecls('w')))
        shd1 = parse_xml(r'<w:shd {} w:fill="0F172A"/>'.format(nsdecls('w')))
        c0._tc.get_or_add_tcPr().append(shd0)
        c1._tc.get_or_add_tcPr().append(shd1)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(3)
        p0.paragraph_format.space_after = Pt(3)
        r0 = p0.add_run(k)
        r0.font.name = 'Consolas'
        r0.font.size = Pt(8.5)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(255, 255, 255)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(3)
        p1.paragraph_format.space_after = Pt(3)
        r1 = p1.add_run(v)
        r1.font.name = 'Consolas'
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = RGBColor(203, 213, 225)

    sp_p = doc.add_paragraph()
    sp_p.paragraph_format.space_before = Pt(2)
    sp_p.paragraph_format.space_after = Pt(4)

# Multi-Column Table (for Mapping Tables)
def add_mapping_table(doc, headers, rows_data, col_widths):
    table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Header Row
    for i, h in enumerate(headers):
        c = table.rows[0].cells[i]
        c.width = col_widths[i]
        shd = parse_xml(r'<w:shd {} w:fill="1E293B"/>'.format(nsdecls('w')))
        c._tc.get_or_add_tcPr().append(shd)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(3)
        r = p.add_run(h)
        r.font.name = 'Consolas'; r.font.bold = True; r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    # Data Rows
    for r_idx, row_items in enumerate(rows_data):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_items):
            c = row.cells[c_idx]
            c.width = col_widths[c_idx]
            bg_color = "F1F5F9" if r_idx % 2 == 0 else "FFFFFF"
            shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), bg_color))
            c._tc.get_or_add_tcPr().append(shd)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Consolas'; r.font.size = Pt(8)
            if c_idx == 0:
                r.font.bold = True
                r.font.color.rgb = C_DARK
            elif c_idx == 1:
                r.font.color.rgb = C_ORANGE
            else:
                r.font.color.rgb = C_TEXT
                
    sp_p = doc.add_paragraph()
    sp_p.paragraph_format.space_before = Pt(2)
    sp_p.paragraph_format.space_after = Pt(4)

# ==========================================
# COVER / TITLE
# ==========================================
title_p = doc.add_paragraph()
title_run = title_p.add_run("PORTFOLIO MASTER CUSTOMIZATION MANUAL")
title_run.font.name = 'Consolas'
title_run.font.size = Pt(22)
title_run.font.bold = True
title_run.font.color.rgb = C_DARK

sub_p = doc.add_paragraph()
sub_run = sub_p.add_run(
    "Interactive One-Click Navigation Index\n"
    "Complete Blueprint for Code Architecture, Animations, and Internal Content Mapping\n"
    "Structured by: *Page* -> Customisation guide for -> *Sections* -> *Title/subtitles* -> *Matter* -> *Content Mapping & Modification Guide* -> *Animations in that Page*"
)
sub_run.font.name = 'Consolas'
sub_run.font.size = Pt(10)
sub_run.font.color.rgb = C_ORANGE
sub_p.paragraph_format.space_after = Pt(12)

# ==========================================
# MASTER INDEX TABLE (CLICKABLE)
# ==========================================
idx_h = doc.add_paragraph()
r_idx = idx_h.add_run("CLICKABLE MODIFICATION INDEX (CLICK [GO] TO JUMP TO THAT PAGE)")
r_idx.font.name = 'Consolas'
r_idx.font.size = Pt(11)
r_idx.font.bold = True
r_idx.font.color.rgb = C_DARK
idx_h.paragraph_format.space_after = Pt(6)

index_table = doc.add_table(rows=1, cols=4)
index_table.alignment = WD_TABLE_ALIGNMENT.CENTER
index_table.autofit = False

headers = ["Jump Link", "*Page*", "*Sections*", "Engine & Transitions"]
col_w = [Inches(1.8), Inches(1.6), Inches(1.9), Inches(1.8)]

for i, h in enumerate(headers):
    c = index_table.rows[0].cells[i]
    c.width = col_w[i]
    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), C_HEADER_BG))
    c._tc.get_or_add_tcPr().append(shd)
    hp = c.paragraphs[0]
    hrun = hp.add_run(h)
    hrun.font.name = 'Consolas'
    hrun.font.bold = True
    hrun.font.size = Pt(9)
    hrun.font.color.rgb = RGBColor(255, 255, 255)

index_entries = [
    ("@page-home-about", "index.html", "About Me & Abstract Teasers", "Direct Image Overlay, Drop Shadows, Hover Glow"),
    ("@page-skills", "skills.html", "Skills Strips, Office Tools, Search Queries", "GSAP 3D Tilt Accordion, Tilted Collage"),
    ("@page-interests", "interests.html", "Math, Vintage Lectures, Retro, WW2, Photography", "Three.js WebGL Helix, Polaroid Depth Stack"),
    ("@page-achievements", "achievements.html", "Course Credentials & Certifications", "CSS 3D CollectionSurfer, Lift Parallax, Modal"),
    ("@page-feedback", "feedback.html", "Peer Reviews & Note Form", "Componentry 3-Panel Split, Staggered 3D Flip"),
    ("@page-profiles", "profiles-contact.html", "Coding Profiles & Social Presence", "React-Bits BounceCards, FlowingMenu Marquee"),
    ("@page-journey", "journey.html", "Milestones, Breakthroughs & Future Focus", "Timeline Sections, Subpage Back Nav")
]

for tag, page_file, sections_list, anim_summary in index_entries:
    row = index_table.add_row()
    c0, c1, c2, c3 = row.cells[0], row.cells[1], row.cells[2], row.cells[3]
    c0.width, c1.width, c2.width, c3.width = col_w[0], col_w[1], col_w[2], col_w[3]
    for c in [c0, c1, c2, c3]:
        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), C_LIGHT_BG))
        c._tc.get_or_add_tcPr().append(shd)
        
    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(3); p0.paragraph_format.space_after = Pt(3)
    add_bookmark_link(p0, f"[GO] {tag}", tag.replace("@", ""), font_size=8.5)
    
    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_before = Pt(3); p1.paragraph_format.space_after = Pt(3)
    r1 = p1.add_run(page_file)
    r1.font.name = 'Consolas'; r1.font.size = Pt(8.5); r1.font.bold = True; r1.font.color.rgb = C_DARK
    
    p2 = c2.paragraphs[0]
    p2.paragraph_format.space_before = Pt(3); p2.paragraph_format.space_after = Pt(3)
    r2 = p2.add_run(sections_list)
    r2.font.name = 'Consolas'; r2.font.size = Pt(8.5); r2.font.color.rgb = C_TEXT
    
    p3 = c3.paragraphs[0]
    p3.paragraph_format.space_before = Pt(3); p3.paragraph_format.space_after = Pt(3)
    r3 = p3.add_run(anim_summary)
    r3.font.name = 'Consolas'; r3.font.size = Pt(8); r3.font.color.rgb = C_ORANGE

doc.add_page_break()

# =========================================================================
# HELPER: SECTION BUILDER WITH EXACT REQUESTED STRUCTURE + CONTENT MAPPING
# =========================================================================
def build_customization_page(doc, tag, page_name, guide_desc, sections_text, title_sub_text, matter_text, mapping_headers, mapping_rows, mapping_widths, mapping_notes, anim_specs, code_blocks):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(2)
    add_bookmark(h, tag.replace("@", ""))
    
    r_tag = h.add_run(f"{tag}\n")
    r_tag.font.name = 'Consolas'; r_tag.font.size = Pt(11); r_tag.font.bold = True; r_tag.font.color.rgb = C_ORANGE
    
    r_pg = h.add_run(f"*Page*: {page_name}")
    r_pg.font.name = 'Consolas'; r_pg.font.size = Pt(16); r_pg.font.bold = True; r_pg.font.color.rgb = C_DARK
    
    # Customisation guide for
    p_cg = doc.add_paragraph()
    p_cg.paragraph_format.space_before = Pt(2); p_cg.paragraph_format.space_after = Pt(4)
    r_cg_label = p_cg.add_run("Customisation guide for: ")
    r_cg_label.font.name = 'Consolas'; r_cg_label.font.size = Pt(10); r_cg_label.font.bold = True; r_cg_label.font.color.rgb = C_ORANGE
    r_cg_val = p_cg.add_run(guide_desc)
    r_cg_val.font.name = 'Consolas'; r_cg_val.font.size = Pt(9.5); r_cg_val.font.color.rgb = C_TEXT
    
    # 1. *Sections*
    p_sec = doc.add_paragraph()
    p_sec.paragraph_format.space_before = Pt(4); p_sec.paragraph_format.space_after = Pt(2)
    r_s_label = p_sec.add_run("*Sections*:\n")
    r_s_label.font.name = 'Consolas'; r_s_label.font.size = Pt(10.5); r_s_label.font.bold = True; r_s_label.font.color.rgb = C_ORANGE
    r_s_val = p_sec.add_run(sections_text)
    r_s_val.font.name = 'Consolas'; r_s_val.font.size = Pt(9.5); r_s_val.font.color.rgb = C_TEXT
    
    # 2. *Title/subtitles*
    p_ts = doc.add_paragraph()
    p_ts.paragraph_format.space_before = Pt(4); p_ts.paragraph_format.space_after = Pt(2)
    r_ts_label = p_ts.add_run("*Title/subtitles*:\n")
    r_ts_label.font.name = 'Consolas'; r_ts_label.font.size = Pt(10.5); r_ts_label.font.bold = True; r_ts_label.font.color.rgb = C_ORANGE
    r_ts_val = p_ts.add_run(title_sub_text)
    r_ts_val.font.name = 'Consolas'; r_ts_val.font.size = Pt(9.5); r_ts_val.font.color.rgb = C_TEXT
    
    # 3. *Matter*
    p_mat = doc.add_paragraph()
    p_mat.paragraph_format.space_before = Pt(4); p_mat.paragraph_format.space_after = Pt(4)
    r_m_label = p_mat.add_run("*Matter*:\n")
    r_m_label.font.name = 'Consolas'; r_m_label.font.size = Pt(10.5); r_m_label.font.bold = True; r_m_label.font.color.rgb = C_ORANGE
    r_m_val = p_mat.add_run(matter_text)
    r_m_val.font.name = 'Consolas'; r_m_val.font.size = Pt(9.5); r_m_val.font.color.rgb = C_TEXT
    
    # 4. Content Mapping & Modification Guide (The guide to map and change internal content)
    p_map = doc.add_paragraph()
    p_map.paragraph_format.space_before = Pt(6); p_map.paragraph_format.space_after = Pt(2)
    r_map_label = p_map.add_run("*Guide to Map & Change Internal Content*:\n")
    r_map_label.font.name = 'Consolas'; r_map_label.font.size = Pt(10.5); r_map_label.font.bold = True; r_map_label.font.color.rgb = C_ORANGE
    
    if mapping_notes:
        p_mn = doc.add_paragraph()
        p_mn.paragraph_format.space_before = Pt(1); p_mn.paragraph_format.space_after = Pt(4)
        r_mn = p_mn.add_run(mapping_notes)
        r_mn.font.name = 'Consolas'; r_mn.font.size = Pt(8.5); r_mn.font.italic = True; r_mn.font.color.rgb = C_GRAY
        
    add_mapping_table(doc, mapping_headers, mapping_rows, mapping_widths)
    
    # 5. *Animations in that Page*
    p_an = doc.add_paragraph()
    p_an.paragraph_format.space_before = Pt(6); p_an.paragraph_format.space_after = Pt(4)
    r_an_label = p_an.add_run("*Animations in that Page*:\n")
    r_an_label.font.name = 'Consolas'; r_an_label.font.size = Pt(10.5); r_an_label.font.bold = True; r_an_label.font.color.rgb = C_ORANGE
    
    add_spec_table(doc, anim_specs)
    
    # 6. Exact Code Modification Blocks
    for code_title, code_snip in code_blocks:
        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(6); p_c.paragraph_format.space_after = Pt(2)
        r_c = p_c.add_run(f"Code Modification Block: {code_title}")
        r_c.font.name = 'Consolas'; r_c.font.size = Pt(9.5); r_c.font.bold = True; r_c.font.color.rgb = C_DARK
        add_code_block(doc, code_snip)
        
    doc.add_page_break()

# =========================================================================
# PAGE 1: @page-home-about (index.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-home-about",
    page_name="index.html (Main Landing Hub)",
    guide_desc="Modifying Hero bio text, spacing, drop shadow contrast, teaser card titles, descriptions, and redirection links.",
    sections_text="""1. #about - About Me Hero Section (Direct image overlay with text-shadow contrast)
2. #skills - Technical Skills (Abstract Teaser Card -> skills.html)
3. #interests - Interests & Passions (Abstract Teaser Card -> interests.html)
4. #journey - My Journey (Abstract Teaser Card -> journey.html)
5. #achievements - Achievements & Honors (Abstract Teaser Card -> achievements.html)
6. #feedback - Feedback & Testimonials (Abstract Teaser Card -> feedback.html)
7. #contact - Profiles & Contact (Abstract Teaser Card -> profiles-contact.html)""",
    title_sub_text="""- Section 1 Title: 'Welcome to My World' (.about-title)
- Section 1 Subquote: '“Simplicity is prerequisite for reliability...”' (.about-subquote in italic Courier New)
- Section 2-7 Titles: 'Technical Skills', 'Interests & Passions', 'My Journey', 'Achievements & Honors', 'Feedback & Testimonials', 'Profiles & Contact'
- Navbar Links: 'About Me', 'Skills', 'Interests', 'Journey', 'Achievements', 'Feedback', 'Profiles & Contact'""",
    matter_text="""The hero section features your personal developer summary directly over your background image without any glassmorphic borders or blurring. All teaser cards feature clear overview summaries with glowing orange redirect buttons linking to their respective detailed subpages.""",
    mapping_headers=["Target Element", "File & Location", "Current Code / Text Content", "How to Map & Modify"],
    mapping_rows=[
        ("Hero Title", "index.html (L33)", "Welcome to My World", "Change text inside <h1 class=\"about-title\">...</h1>"),
        ("Hero Subquote", "index.html (L34)", "“Simplicity is prerequisite for reliability...”", "Update quote inside <p class=\"about-subquote\">...<p>"),
        ("Hero Bio Paragraph", "index.html (L35-37)", "I am a developer who believes in first principles...", "Replace bio inside <p class=\"about-paragraph\">...<p>"),
        ("Background Photo", "styles.css (L28)", "url('me.jpeg')", "Replace me.jpeg in root folder or change filename in CSS"),
        ("Lateral Padding", "styles.css (L24)", "padding: 100px 8vw;", "Change 8vw to adjust distance from screen left/right edges"),
        ("Vertical Margins", "styles.css (L36, L45)", "margin-bottom: 16px; / 26px;", "Tune spacing between Title, Subquote, and Paragraph"),
        ("Teaser Card 1 (Skills)", "index.html (L45-52)", "Title: 'Technical Skills', Link: 'skills.html'", "Edit summary <p> and href inside <a class=\"abstract-btn\">"),
        ("Teaser Card 2 (Interests)", "index.html (L58-65)", "Title: 'Interests & Passions', Link: 'interests.html'", "Edit summary <p> and href inside <a class=\"abstract-btn\">"),
        ("Teaser Card 3 (Journey)", "index.html (L71-78)", "Title: 'My Journey', Link: 'journey.html'", "Edit summary <p> and href inside <a class=\"abstract-btn\">"),
        ("Teaser Card 4 (Achievements)", "index.html (L84-91)", "Title: 'Achievements & Honors', Link: 'achievements.html'", "Edit summary <p> and href inside <a class=\"abstract-btn\">"),
        ("Teaser Card 5 (Feedback)", "index.html (L97-104)", "Title: 'Feedback & Testimonials', Link: 'feedback.html'", "Edit summary <p> and href inside <a class=\"abstract-btn\">"),
        ("Teaser Card 6 (Contact)", "index.html (L110-117)", "Title: 'Profiles & Contact', Link: 'profiles-contact.html'", "Edit summary <p> and href inside <a class=\"abstract-btn\">")
    ],
    mapping_widths=[Inches(1.5), Inches(1.4), Inches(2.2), Inches(2.0)],
    mapping_notes="All hero bio copy and teaser cards can be updated directly in index.html. Padding and typography styles are located in styles.css.",
    anim_specs=[
        ("Direct Image Stacking", "me.jpeg rendered as background-size: cover; background-position: center; no-repeat"),
        ("Legibility Drop Shadow", "text-shadow: 0 3px 18px rgba(0,0,0,0.95), 0 1px 4px rgba(0,0,0,0.85) separates white text from photo"),
        ("Sticky Navigation", "Navbar stays pinned to top with position: sticky; top: 0; backdrop-filter: blur(10px)"),
        ("Button Hover Physics", ".abstract-btn lifts with translateY(-3px) and box-shadow: 0 0 20px rgba(219, 114, 44, 0.6)"),
        ("Smooth Anchor Scroll", "html { scroll-behavior: smooth; } smoothly animates jumps between navbar anchor IDs")
    ],
    code_blocks=[
        ("styles.css (Lines 20-80) - About Me Spacing & Layout",
"""/* Adjust lateral padding, title margins, and line heights in styles.css */
.about-hero-section {
    padding: 100px 8vw;   /* [Top/Bottom: 100px, Left/Right: 8vw from edge] */
    justify-content: flex-start; /* Pinned to left */
}

.about-title {
    margin-bottom: 16px;  /* Vertical gap between Title & Subquote */
    font-size: clamp(2.4rem, 4.2vw, 3.4rem);
}

.about-subquote {
    margin-bottom: 26px;  /* Vertical gap between Subquote & Paragraph */
    color: #db722c;       /* Accent orange */
    line-height: 1.55;
}

.about-paragraph {
    line-height: 1.8;     /* Vertical distance between text lines */
    max-width: 720px;     /* Line length limit for comfortable reading */
}""")
    ]
)

# =========================================================================
# PAGE 2: @page-skills (skills.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-skills",
    page_name="skills.html (Technical Stack & Deep Dives)",
    guide_desc="Adding or reordering programming skill panels, changing images, modifying Office design cards, and editing Google search operator badges.",
    sections_text="""1. Section 1: AccordionGallery (Componentry 3D Interactive Tilting Strips)
2. Section 2: MS Word, PowerPoint & Draw.io Showcase (Typography + 3-Card Collage)
3. Section 3: Google Search Operators & Query Playground (Sherlock Holmes Technique)""",
    title_sub_text="""- Section 1 Header: 'Technical Proficiency & Languages' (Tag: 'CORE LANGUAGES • ARCHITECTURE • TOOLING')
- Section 2 Header: 'MS word, Power Point and draw.io.' (Tag: 'DESIGN & VISUALISATION')
- Section 3 Header: 'Google Search Queries' (Tag: 'Sherlock Holmes in the town of Google')
- Interactive Search Chips: 'site:', 'filetype:pdf', 'intitle:', 'inurl:', '\"exact match\"', '-exclude'""",
    matter_text="""Showcases low-level system languages, data pipelines, presentation authoring, and precision web search operators. Individual skill panels display active title tags, glowing indicator bars, and customized tool descriptions.""",
    mapping_headers=["Target Item / Field", "File & Location", "Property / Array Key", "How to Map & Modify"],
    mapping_rows=[
        ("Skill 1: C Programming", "skills.html (L138-143)", "{ image: 'images/c.jpg', label: 'C Programming', link: '#' }", "Update label string, replace c.jpg in images folder"),
        ("Skill 2: Python", "skills.html (L144-149)", "{ image: 'images/python.jpg', label: 'Python', link: '#' }", "Update label string, replace python.jpg"),
        ("Skill 3: Java", "skills.html (L150-155)", "{ image: 'images/java-bg.jpg', label: 'Java', link: '#' }", "Update label string, replace java-bg.jpg"),
        ("Skill 4: C++ & Algos", "skills.html (L156-161)", "{ image: '...', label: 'C++ & Algorithms', link: '#' }", "Replace URL with local path or custom image"),
        ("Skill 5: JavaScript & Web", "skills.html (L162-167)", "{ image: '...', label: 'JavaScript & Web', link: '#' }", "Replace image and link target"),
        ("Skill 6: Databases & SQL", "skills.html (L168-173)", "{ image: '...', label: 'Databases & SQL', link: '#' }", "Add PostgreSQL / MySQL image and label"),
        ("Skill 7: Git & Linux", "skills.html (L174-179)", "{ image: '...', label: 'Git & Linux', link: '#' }", "Add terminal / devops image"),
        ("Adding New Skills", "skills.html (L137)", "items.push({ image: '...', label: '...' })", "Simply append another JSON object to the items array"),
        ("Office Collage Cards", "skills.html (L68-88)", "images/word-sample.jpg, ppt-sample.jpg, drawio-sample.jpg", "Replace images and edit .card-caption text"),
        ("Search Operator Chips", "skills.html (L96-102)", "<span class=\"badge\">site:</span>", "Add or modify chip tags in the HTML badge container")
    ],
    mapping_widths=[Inches(1.6), Inches(1.4), Inches(2.2), Inches(1.9)],
    mapping_notes="The items array in skills.html controls Section 1. All images can be local files (e.g. images/my-skill.png) or remote URLs.",
    anim_specs=[
        ("GSAP 3D Interactive Tilt", "Non-active panels rotate in 3D: rotateY: +8deg (left) / -8deg (right). Active panel: 0deg"),
        ("Grayscale to Color Blend", "Inactive panels: filter: grayscale(1) with --ag-dim: 0.35; Active: grayscale(0) vibrant color"),
        ("Parallax Counter-Drift", "Media moves with shift = drift * parallax * mediaSize * 0.06 as you hover across panels"),
        ("Glowing Bar Indicator", "GSAP staggered tween (stagger: 0.06s) animating opacity: 1 and translateX(-14px -> 0px)"),
        ("3-Card Collage Tilt", "Screenshot cards tilted at -3deg, +4deg, -2deg with hover scale(1.05) and z-index elevation")
    ],
    code_blocks=[
        ("skills.html (Lines 135-185) - Skills Array & GSAP Tuning",
"""// Update skill items, local image paths, and links in skills.html:
const items = [
    { image: "images/c.jpg", label: "C Programming", link: "#" },
    { image: "images/python.jpg", label: "Python", link: "#" },
    { image: "images/java-bg.jpg", label: "Java", link: "#" },
    { image: "https://picsum.photos/id/1039/900/1200", label: "C++ & Algorithms", link: "#" },
    { image: "https://picsum.photos/id/1043/900/1200", label: "JavaScript & Web", link: "#" },
    { image: "https://picsum.photos/id/1044/900/1200", label: "Databases & SQL", link: "#" },
    { image: "https://picsum.photos/id/1018/900/1200", label: "Git & Linux", link: "#" }
];

// Physical animation parameters:
const tilt = 8;          // 3D Y-axis tilt in degrees
const expandRatio = 0.52;// Portion of gallery width taken by active card (52%)
const duration = 0.6;    // Tween speed in seconds
const ease = 'power3.out'; // Smoothing curve""")
    ]
)

# =========================================================================
# PAGE 3: @page-interests (interests.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-interests",
    page_name="interests.html (Passions, Physics & Photography)",
    guide_desc="Mapping Vintage Lecture videos, speakers, YouTube redirects, space topics, WW2 historical items, and Polaroid photography slides.",
    sections_text="""1. Section 1 (#math): Mandelbrot Fractal & Pure Mathematics
2. Section 2 (#vintage-lectures): Vintage Lectures 3D WebGL Spiral Helix
3. Section 3 (#space-tech): Space Exploration & Retro Tech
4. Section 4 (#research): First Principles Systems & Architecture Papers
5. Section 5 (#ww2-tech): Wartime Technology, Signals & Enigma Decryption
6. Section 6 (#photography): Sticky Scroll Stacking Cards Polaroid Gallery""",
    title_sub_text="""- Section 1: '1. Math' (Tag: 'pure-mathematics / discrete-structures / logic')
- Section 2: '2. Vintage Lectures' (Tag: 'classic-archives / historic-pedagogy / timeless-minds')
- Section 3: '3. Space and Retro tech' (Tag: 'spacecraft-avionics / apollo-guidance / retro-hardware')
- Section 4: '4. Research' (Tag: 'foundational-papers / rfc-specs / systems-architecture')
- Section 5: '5. WW 2 Tech' (Tag: 'signals-intelligence / cryptographic-warfare / wartime-engineering')
- Section 6: '6. Photography' (Badge: '[SNAPSEED TUNED] • 5 FRAMES')""",
    matter_text="""Deep dive into intellectual passions outside software development. Features a custom Mandelbrot set backdrop, Feynman/Lewin historic lectures, cold-war electronics, and a film photography portfolio with Snapseed darkroom presets.""",
    mapping_headers=["Content Field", "File & Location", "Current Code / Value", "How to Map & Modify"],
    mapping_rows=[
        ("Feynman Lecture", "interests.html (L154-160)", "title, speaker, src: feynman.jpg, url: youtube.com", "Change lecture title, speaker, thumbnail image, or video link"),
        ("Walter Lewin Lecture", "interests.html (L161-166)", "title, speaker, src: lewin.jpg, url: youtube.com", "Change lecture title, speaker, thumbnail image, or video link"),
        ("Eric Grimson Lecture", "interests.html (L167-172)", "title, speaker, src: Eric.png, url: youtube.com", "Change lecture title, speaker, thumbnail image, or video link"),
        ("Anant Agarwal Lecture", "interests.html (L173-178)", "title, speaker, src: pranantag.png, url: youtube.com", "Change lecture title, speaker, thumbnail image, or video link"),
        ("Oppenheimer Lecture", "interests.html (L179-184)", "title, speaker, src: oppenheimer.jpg, url: youtube.com", "Change lecture title, speaker, thumbnail image, or video link"),
        ("Adding New Lectures", "interests.html (L153)", "lectures = [ ... ]", "Append new lecture JSON object; helix adapts automatically"),
        ("Photo 1: Misty Alps", "interests.html (L518-521)", "{ title: 'Misty Alps', src: 'unsplash...' }", "Change caption in title, replace src with your camera photo"),
        ("Photo 2: Sunlit Grove", "interests.html (L522-525)", "{ title: 'Sunlit Grove', src: 'unsplash...' }", "Change caption in title, replace src with your camera photo"),
        ("Photo 3: Shore", "interests.html (L526-529)", "{ title: 'Turquoise Shore', src: 'unsplash...' }", "Change caption in title, replace src with your camera photo"),
        ("Photo 4: Mountain Pass", "interests.html (L530-533)", "{ title: 'Mountain Pass', src: 'unsplash...' }", "Change caption in title, replace src with your camera photo"),
        ("Photo 5: Rolling Hills", "interests.html (L534-537)", "{ title: 'Rolling Hills', src: 'unsplash...' }", "Change caption in title, replace src with your camera photo"),
        ("Snapseed Darkroom Pills", "interests.html (L134-139)", "Ambiance +18, Curves, Film Grain L02", "Edit pill badge text inside <span class=\"snapseed-pill\">")
    ],
    mapping_widths=[Inches(1.6), Inches(1.4), Inches(2.2), Inches(1.9)],
    mapping_notes="Vintage lectures are loaded in the Three.js helix engine. Photography cards are defined in the cardsData array in the bottom script.",
    anim_specs=[
        ("Three.js WebGL Helix", "PerspectiveCamera (fov: 44deg) rendering cards along a 3D spiral ladder"),
        ("Custom GLSL Shaders", "Vertex: cosine curve mesh bending; Fragment: 10-tap Gaussian blur + luminance boost"),
        ("Anti-Overlap Helix Math", "Radius: 3.40, Gap: 1.35, Angle Step: 0.95 rad. Arc distance (3.23) > card width (1.85)"),
        ("4800px Pinned Runway", "Screen stays pinned to Section 2 until all 8 lectures are viewed via scroll progress"),
        ("Polaroid Stacking Physics", "Sequential scale-down: restingScale = max(0.56, 1 - (total-index-1)*0.095) with tilts"),
        ("Snapseed Darkroom Box", "Bottom-left glassmorphic panel with backdrop-filter: blur(24px) saturate(160%)")
    ],
    code_blocks=[
        ("interests.html (Lines 150-195) - Vintage Lectures Dataset",
"""// Update lectures, speakers, and YouTube links in interests.html:
const lectures = [
    {
        title: "Richard Feynman - The Character of Physical Law (1964)",
        speaker: "Richard Feynman",
        src: "images/lectures/feynman.jpg",
        url: "https://www.youtube.com/watch?v=j3mhkYbznBk"
    },
    {
        title: "Walter Lewin - 8.01 MIT Classical Mechanics",
        speaker: "Walter Lewin",
        src: "images/lectures/lewin.jpg",
        url: "https://www.youtube.com/watch?v=4a0FbQdH3dY"
    }
    // Add additional lectures here...
];"""),
        ("interests.html (Lines 515-545) - Photography Cards Dataset",
"""// Update photography collection in interests.html:
const cardsData = [
    { title: "Misty Alps", src: "https://images.unsplash.com/photo-1476514525535-07fb3b4ae5f1?w=1200&q=85" },
    { title: "Sunlit Grove", src: "https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=1200&q=85" },
    { title: "Turquoise Shore", src: "https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=1200&q=85" },
    { title: "Mountain Pass", src: "https://images.unsplash.com/photo-1482938289607-e9573fc25ebb?w=1200&q=85" },
    { title: "Rolling Hills", src: "https://images.unsplash.com/photo-1501854140801-50d01698950b?w=1200&q=85" }
];
const TILT_PATTERN = [-1.25, 0.85, -0.65, 1.35, -0.9];""")
    ]
)

# =========================================================================
# PAGE 4: @page-achievements (achievements.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-achievements",
    page_name="achievements.html (Honors & Credentials)",
    guide_desc="Adding and editing certificate courses, issuing organizations, dates, descriptions, and certificate image files.",
    sections_text="""1. Section 1 (#cert-surfer): 3D Certificate CollectionSurfer Stage
2. Lightbox Modal (#cert-modal): High-Resolution A4 Zoom Viewport""",
    title_sub_text="""- Top Overlay: 'COURSE CREDENTIALS' / 'COLLECTION (8)' (.surfer-heading-main)
- Focus Indicator: 'CERTIFICATE 01 / 08' (.focus-badge)
- Focus Title: Dynamically syncs to centered certificate name (.focus-title)
- Focus Issuer: 'Coursera / DeepLearning.AI • 2024' (.focus-issuer)
- Scroll Hint: 'scroll to surf ↓ (click to zoom)' (.surfer-scroll-hint)""",
    matter_text="""Presents 8 certified accomplishments rendered in true A4 landscape proportions (420px x 297px, 1.414:1 aspect ratio). The active certificate lifts 80px higher and synchronizes real-time credential metadata on the left overlay.""",
    mapping_headers=["Certificate #", "File & Location", "Issuer / Org Key", "How to Map & Modify"],
    mapping_rows=[
        ("Cert 1: Python & Data", "achievements.html (L78-85)", "Coursera / DeepLearning.AI (2024)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 2: Java Full Stack", "achievements.html (L86-93)", "Oracle Academy (2023)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 3: Systems & OS", "achievements.html (L94-101)", "edX / Harvard CS50 (2023)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 4: Algorithms", "achievements.html (L102-109)", "Stanford Online (2023)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 5: Web & DBs", "achievements.html (L110-117)", "freeCodeCamp / MongoDB (2024)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 6: Cloud & AWS", "achievements.html (L118-125)", "Amazon Web Services (2024)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 7: Deep Learning", "achievements.html (L126-133)", "DeepLearning.AI / TF (2023)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Cert 8: Git & DevOps", "achievements.html (L134-141)", "Linux Foundation / GitHub (2023)", "Edit title, issuer, date, desc, image path in certificates array"),
        ("Adding More Certs", "achievements.html (L75)", "certificates.push({ ... })", "Add new object; surfer counter and track auto-update"),
        ("Certificate Images", "images/certificates/", "cert1.jpg, cert2.jpg, ...", "Place scans in images/certificates/ or reference web URLs")
    ],
    mapping_widths=[Inches(1.6), Inches(1.5), Inches(2.1), Inches(1.9)],
    mapping_notes="All certificate credentials and descriptions are mapped in the certificates array in achievements.html. The collection counter updates automatically.",
    anim_specs=[
        ("CSS 3D Perspective Scene", "perspective: 1800px; transform-style: preserve-3d; perspective-origin: 50% 50%"),
        ("Parallax Elevation Lift", "Active focused certificate lifts translateY(-80px) and scales to 1.08x with glowing border"),
        ("Z-Index Card Layering", "Inactive cards sit at z-index: 10, active lifts to z-index: 50 with smooth 0.4s transitions"),
        ("Live Synchronized Details", "Title, issuer, date, and description update dynamically as cards cross the focus center"),
        ("A4 Lightbox Modal Zoom", "Clicking any card smoothly triggers #cert-modal with backdrop-filter: blur(20px)")
    ],
    code_blocks=[
        ("achievements.html (Lines 75-145) - Certificates Collection",
"""// Update certificate credentials in achievements.html:
const certificates = [
    {
        title: "Advanced Python & Data Architecture",
        issuer: "Coursera / DeepLearning.AI",
        date: "2024",
        desc: "Comprehensive mastery of data structures, automation pipelines, and core programming paradigms.",
        image: "images/certificates/cert1.jpg"
    },
    {
        title: "Java Full Stack Development",
        issuer: "Oracle Academy",
        date: "2023",
        desc: "Object-oriented design, multithreaded backend architecture, and design patterns.",
        image: "images/certificates/cert2.jpg"
    }
    // Add additional credentials here...
];

// Physical Surfer Variables:
const stepX = 260;         // Horizontal spacing between certificates
const stepY = -80;         // Elevation step per card
const stepZ = -300;        // Depth recession in 3D space""")
    ]
)

# =========================================================================
# PAGE 5: @page-feedback (feedback.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-feedback",
    page_name="feedback.html (Peer Reviews & Testimonials)",
    guide_desc="Modifying peer review quotes, authors, roles, badge titles, background card colors, and review form submission text.",
    sections_text="""1. Section 1 (#split-scroll-section): Componentry ScrollSplitCard 3-Panel Stage
2. Section 2 (#leave-feedback): Peer Review & Community Submission Form""",
    title_sub_text="""- Intro Tag: 'PEER REVIEWS • ARCHIVE' (.split-tag)
- Intro Heading: 'Words From Friends & Teammates' (.split-heading)
- Dynamic Phase Badge:
  - Phase 1: 'PHASE 1 / 3 • UNITED HERO IMAGE'
  - Phase 2: 'PHASE 2 / 3 • CARDS SEPARATING & FLIPPING'
  - Phase 3: 'PHASE 3 / 3 • ALL 3 PEER REVIEWS REVEALED'
- Form Title: 'Leave a Thought or Review'""",
    matter_text="""Features testimonials from collaborators and classmates. Begins as a single hero photograph, splits apart horizontally as the user scrolls down, and executes staggered 3D flips revealing three distinct styled review cards.""",
    mapping_headers=["Card Panel", "File & Location", "Current Text / Properties", "How to Map & Modify"],
    mapping_rows=[
        ("Hero Split Image", "feedback.html (L38, L61, L82)", "background-image in .split-img-slice", "Replace hero photo in CSS; front faces slice it into 3 parts"),
        ("Panel 1 (Left Card)", "feedback.html (L44-55)", "Author: Arun Kumar, Role: Project Partner", "Edit title, quote paragraph, author name, role, background-color"),
        ("Panel 2 (Center Card)", "feedback.html (L66-74)", "Author: Devin Reed, Role: Hackathon Teammate", "Edit title, quote paragraph, author name, role, background-color"),
        ("Panel 3 (Right Card)", "feedback.html (L86-97)", "Author: Siddharth Rao, Role: Study Group Peer", "Edit title, quote paragraph, author name, role, background-color"),
        ("Card Accent Colors", "feedback.html (L45, L66, L86)", "Panel 1: #161822 | Panel 2: #db722c | Panel 3: #0e0f14", "Change inline style background-color for each card back"),
        ("Review Form Heading", "feedback.html (L115)", "Leave a Thought or Review", "Change <h2> heading in section #leave-feedback"),
        ("Form Submission Logic", "feedback.html (L119)", "onsubmit=\"event.preventDefault(); alert('...');\"", "Connect to a Google Form, Formspree, or backend endpoint")
    ],
    mapping_widths=[Inches(1.6), Inches(1.5), Inches(2.1), Inches(1.9)],
    mapping_notes="Each panel contains a front face (slice of hero image) and a back face (review card). The back face HTML holds the friend quote, name, and role.",
    anim_specs=[
        ("3-Phase Scroll Runway", "3800px pinned runway with sticky viewport tracking scroll progress from 0.0 to 1.0"),
        ("Single Image Splitting", "Front faces use width: 300% and offsets [0%, -100%, -200%] to slice 1 hero image"),
        ("Lateral Separation", "Left panel shifts -38px, right shifts +38px to create crisp separation gaps"),
        ("Staggered 3D Y-Axis Flip", "Quadratic easing flip: Panel 0 (0.22->0.54), Panel 1 (0.38->0.70), Panel 2 (0.54->0.86)"),
        ("3D Card Preserved Depth", "transform-style: preserve-3d with backface-visibility: hidden for seamless front/back flips")
    ],
    code_blocks=[
        ("feedback.html (Lines 44-75) - Sliced Reviews Markup",
"""<!-- feedback.html: Update review text, authors, roles, and background colors -->
<!-- PANEL 1: Charcoal Card -->
<div class="split-face split-face-back" style="background-color: #161822;">
    <div class="panel-review-badge">COLLABORATOR</div>
    <h3 class="panel-review-title">Rapid Problem Solver</h3>
    <p class="panel-review-desc">"Always the first to dive into complex algorithmic blockers..."</p>
    <span class="panel-author-name">Arun Kumar</span>
    <span class="panel-author-role">Project Partner</span>
</div>

<!-- PANEL 2: Accent Orange Card -->
<div class="split-face split-face-back" style="background-color: #db722c; color: #000000;">
    <div class="panel-review-badge">CREATIVE DRIVE</div>
    <h3 class="panel-review-title">Vintage & Tech Enthusiast</h3>
    <p class="panel-review-desc">"His fascination with retro computing brings a totally unique perspective..."</p>
    <span class="panel-author-name">Devin Reed</span>
    <span class="panel-author-role">Hackathon Teammate</span>
</div>""")
    ]
)

# =========================================================================
# PAGE 6: @page-profiles (profiles-contact.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-profiles",
    page_name="profiles-contact.html (Coding & Social Profiles)",
    guide_desc="Mapping coding platform logos, URLs, usernames, ratings, and social ribbon marquee links (LinkedIn, Instagram, Pinterest, GitHub).",
    sections_text="""1. Section 1 (#coding-profiles): React-Bits BounceCards Developer Profiles
2. Section 2 (#social-presence): React-Bits FlowingMenu Kinetic Ribbon Marquees
3. Section 3 (#direct-contact): Get In Touch / Email Dispatcher""",
    title_sub_text="""- Section 1 Header: 'Coding & Developer Profiles' (Tag: 'DEV PLATFORMS • ALGORITHMS • OPEN SOURCE')
- Focus Details Bar: 'GITHUB', 'GitHub / @username', 'Repositories, open source contributions...'
- Section 2 Header: 'Professional & Social Presence' (Tag: 'DIGITAL FOOTPRINT • SOCIAL • COLLABORATION')
- Ribbons: 'LinkedIn', 'Instagram', 'Pinterest', 'GitHub'
- Section 3 Header: 'Get In Touch' (Button: 'Send an Email →')""",
    matter_text="""Consolidates GitHub, LeetCode, Codeforces, HackerRank, and Kaggle profiles on Page 1, followed by direction-aware kinetic ribbon bands for LinkedIn, Instagram, Pinterest, and GitHub on Page 2.""",
    mapping_headers=["Platform / Target", "File & Location", "Current Code / Value", "How to Map & Modify"],
    mapping_rows=[
        ("Profile 1: GitHub", "profiles-contact.html (L89-98)", "url: 'https://github.com', image: '...'", "Put your GitHub URL, username handle, and avatar/logo path"),
        ("Profile 2: LeetCode", "profiles-contact.html (L99-108)", "url: 'https://leetcode.com', image: '...'", "Put your LeetCode URL, contest handle, and stats"),
        ("Profile 3: Codeforces", "profiles-contact.html (L109-117)", "url: 'https://codeforces.com', image: '...'", "Put your Codeforces profile link and rating handle"),
        ("Profile 4: HackerRank", "profiles-contact.html (L118-126)", "url: 'https://hackerrank.com', image: '...'", "Put your HackerRank profile URL and star badges summary"),
        ("Profile 5: Kaggle", "profiles-contact.html (L127-135)", "url: 'https://kaggle.com', image: '...'", "Put your Kaggle profile URL and notebook count"),
        ("Ribbon 1: LinkedIn", "profiles-contact.html (L317-322)", "text: 'LinkedIn', link: 'https://linkedin.com'", "Update your personal LinkedIn profile URL and preview image"),
        ("Ribbon 2: Instagram", "profiles-contact.html (L323-328)", "text: 'Instagram', link: 'https://instagram.com'", "Update your Instagram profile URL and preview photo"),
        ("Ribbon 3: Pinterest", "profiles-contact.html (L329-334)", "text: 'Pinterest', link: 'https://pinterest.com'", "Update your Pinterest profile URL and board preview image"),
        ("Ribbon 4: GitHub Ribbon", "profiles-contact.html (L335-340)", "text: 'GitHub', link: 'https://github.com'", "Update your GitHub link and organization banner"),
        ("Email Contact Button", "profiles-contact.html (L70)", "href=\"mailto:your.email@example.com\"", "Replace placeholder with your real contact email address")
    ],
    mapping_widths=[Inches(1.6), Inches(1.5), Inches(2.1), Inches(1.9)],
    mapping_notes="BounceCards profiles are mapped in profileCards (L89-135). FlowingMenu ribbons are mapped in menuItems (L316-341).",
    anim_specs=[
        ("GSAP Elastic Pop-In", "Entrance bounce using elastic.out(1, 0.7) with staggered card pop-in (0.06s delay)"),
        ("Hover Straighten & Magnify", "Active card scales to 1.06x, straightens to 0deg rotation, and elevates to z-index: 50"),
        ("Directional Push Physics", "Neighboring cards slide away sideways by pushDistance * 0.35 on hover"),
        ("Direction-Aware Ribbons", "Calculates closest edge (top vs bottom) on mouse entry/exit to slide ribbons directionally"),
        ("Continuous Ribbon Marquee", "Infinite seamless GSAP loop (duration: 15s) streaming monospace text and capsule logos")
    ],
    code_blocks=[
        ("profiles-contact.html (Lines 89-135) - Developer Profiles Array",
"""// Update profile cards, URLs, and logos in profiles-contact.html:
const profileCards = [
    {
        title: "GitHub",
        handle: "GitHub / @your-username",
        desc: "Open-source projects, system utilities, and repositories.",
        url: "https://github.com/your-username",
        image: "images/logos/github.png",
        badge: "OPEN SOURCE"
    },
    {
        title: "LeetCode",
        handle: "LeetCode / Algorithmic Mastery",
        desc: "Data structures, graph algorithms, dynamic programming.",
        url: "https://leetcode.com/your-username",
        image: "images/logos/leetcode.png",
        badge: "PROBLEM SOLVER"
    }
];

// Physical parameters:
const pushDistance = 160; // Pixels neighboring cards push away
const hoverScale = 1.06;  // Hovered card magnification factor"""),
        ("profiles-contact.html (Lines 315-350) - Social Presence Ribbons Array",
"""// Update social ribbon links and capsule images in profiles-contact.html:
const menuItems = [
    { text: "LinkedIn",  link: "https://linkedin.com/in/your-profile", image: "https://images.unsplash.com/photo-1611944212129-29977ae1398c?w=600&h=300&fit=crop" },
    { text: "Instagram", link: "https://instagram.com/your-handle",    image: "https://images.unsplash.com/photo-1611162617213-7d7a39e9b1d7?w=600&h=300&fit=crop" },
    { text: "Pinterest", link: "https://pinterest.com/your-board",     image: "https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=600&h=300&fit=crop" },
    { text: "GitHub",    link: "https://github.com/your-username",     image: "https://images.unsplash.com/photo-1618401471353-b98afee0b2eb?w=600&h=300&fit=crop" }
];
const speed = 15; // Loop duration in seconds""")
    ]
)

# =========================================================================
# PAGE 7: @page-journey (journey.html)
# =========================================================================
build_customization_page(
    doc=doc,
    tag="@page-journey",
    page_name="journey.html (Career Milestones & Chronology)",
    guide_desc="Mapping career milestone dates, education timeline, project retrospective descriptions, and future vision roadmaps.",
    sections_text="""1. Section 1: The Beginning & Education (Early curiosity, academic milestones)
2. Section 2: Projects & Key Milestones (Breakthroughs, hackathons, open source)
3. Section 3: Where I Am Heading (Current focus, target technologies, long-term goals)""",
    title_sub_text="""- Section 1 Title: 'Page 1: The Beginning & Education'
- Section 2 Title: 'Page 2: Projects & Key Milestones'
- Section 3 Title: 'Page 3: Where I Am Heading'
- Subpage Header: 'My Journey' with Back Button: '← Back to Home' linking to index.html#journey""",
    matter_text="""Chronological timeline of your growth as an engineer. Template sections ready to be expanded with dates, project retrospective cards, and interactive milestone roadmaps.""",
    mapping_headers=["Timeline Phase", "File & Location", "Section Heading & Target", "How to Map & Modify"],
    mapping_rows=[
        ("Phase 1: Education", "journey.html (L20-24)", "<h2>Page 1: The Beginning & Education</h2>", "Add high school / university dates, degree, foundational courses"),
        ("Phase 2: Milestones", "journey.html (L26-30)", "<h2>Page 2: Projects & Key Milestones</h2>", "Add hackathon wins, major GitHub releases, production software launches"),
        ("Phase 3: Future", "journey.html (L32-36)", "<h2>Page 3: Where I Am Heading</h2>", "Add future learning goals, intended domains, engineering ambitions"),
        ("Back to Home Nav", "journey.html (L12)", "<a href=\"index.html#journey\" class=\"back-btn\">", "Controls redirect destination back to main portfolio page"),
        ("Adding Extra Sections", "journey.html (L37)", "<section class=\"subpage-section\">...</section>", "Copy and paste section template block to add Phase 4, Phase 5, etc.")
    ],
    mapping_widths=[Inches(1.6), Inches(1.5), Inches(2.1), Inches(1.9)],
    mapping_notes="Each phase is contained within a <section class=\"subpage-section\"> block in journey.html. Add paragraphs, lists, or timeline cards inside.",
    anim_specs=[
        ("Subpage Header Docking", "Sticky top navigation header with position: sticky; top: 0; z-index: 1000"),
        ("Full-Height Scroll Snapping", "Each section spans min-height: 100vh with alternating dark tones (#0d0d0d / #050505)"),
        ("Accent Hover Lighting", "Back button (.back-btn) expands with scale(1.05) and vibrant orange accent fill (#db722c)")
    ],
    code_blocks=[
        ("journey.html (Lines 18-35) - Journey Sections Structure",
"""<!-- In journey.html: Customize your career story and educational background -->
<section class="subpage-section">
    <h2>Page 1: The Beginning & Education</h2>
    <p>Describe your early introduction to computing, first programs, and university foundations.</p>
</section>

<section class="subpage-section">
    <h2>Page 2: Projects & Key Milestones</h2>
    <p>Detail your biggest engineering milestones, hackathon wins, and production builds.</p>
</section>

<section class="subpage-section">
    <h2>Page 3: Where I Am Heading</h2>
    <p>Outline your target engineering domains, research areas, and career goals.</p>
</section>""")
    ]
)

doc.save("Portfolio_Customization_Guide.docx")
print("Successfully generated Portfolio_Customization_Guide.docx with requested structure!")
