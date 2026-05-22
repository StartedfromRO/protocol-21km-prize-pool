import collections
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    # Set slide dimensions to 16:9 widescreen
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Colors
    bg_dark = RGBColor(18, 22, 28)       # Very dark navy/charcoal
    card_bg = RGBColor(30, 37, 48)       # Slightly lighter card background
    text_white = RGBColor(245, 247, 250)
    text_muted = RGBColor(160, 174, 192)
    accent_green = RGBColor(52, 211, 153) # Mint green accent
    accent_blue = RGBColor(96, 165, 250)  # Ice blue accent
    border_color = RGBColor(45, 55, 72)
    table_header_bg = RGBColor(45, 55, 72)

    def apply_slide_bg(slide):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bg_dark

    def add_title(slide, text, subtitle_text=None):
        title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.5), Inches(11.83), Inches(1.2))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_bottom = tf.margin_right = 0
        
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = 'Georgia'
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = text_white
        
        if subtitle_text:
            p2 = tf.add_paragraph()
            p2.text = subtitle_text
            p2.font.name = 'Arial'
            p2.font.size = Pt(14)
            p2.font.color.rgb = text_muted
            p2.space_before = Pt(6)

    # -------------------------------------------------------------
    # SLIDE 1: Title Slide
    # -------------------------------------------------------------
    slide_layout = prs.slide_layouts[6] # Blank layout
    slide1 = prs.slides.add_slide(slide_layout)
    apply_slide_bg(slide1)

    # Decorative visual element (accent block)
    shape = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(2.2), Inches(0.15), Inches(3.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = accent_green
    shape.line.fill.background()

    # Title box
    title_box = slide1.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(11.0), Inches(3.5))
    tf = title_box.text_frame
    tf.word_wrap = True
    
    p_main = tf.paragraphs[0]
    p_main.text = "2025 MALAYSIA RUNNING EVENT"
    p_main.font.name = 'Arial'
    p_main.font.size = Pt(16)
    p_main.font.bold = True
    p_main.font.color.rgb = accent_green
    p_main.space_after = Pt(12)

    p_title = tf.add_paragraph()
    p_title.text = "Official Cash Prize Structure\nRecommendation"
    p_title.font.name = 'Georgia'
    p_title.font.size = Pt(44)
    p_title.font.bold = True
    p_title.font.color.rgb = text_white
    p_title.space_after = Pt(16)

    p_sub = tf.add_paragraph()
    p_sub.text = "Comparative proposal for Option A (Premium 10 Ranks) and Option B (Budget 5 Ranks)\nDesigned for 8,000 Participants"
    p_sub.font.name = 'Arial'
    p_sub.font.size = Pt(16)
    p_sub.font.color.rgb = text_muted

    # -------------------------------------------------------------
    # SLIDE 2: Parameter Overview
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(slide_layout)
    apply_slide_bg(slide2)
    add_title(slide2, "Event Core Parameters", "Rules and baseline benchmarks for the proposed models")

    # Add 3 Cards for Key Metrics
    card_width = Inches(3.6)
    card_height = Inches(4.2)
    y_pos = Inches(2.0)
    
    cards_data = [
        {
            "title": "21KM INDIVIDUAL",
            "desc": "Baseline Anchor Category",
            "points": [
                "1st Place Open: RM 2,000",
                "1st Place Veteran: RM 1,500",
                "Age Groups: Men (18-39, 40+), Women (18-39, 40+)",
                "Standard Malaysian decay curve applied to subsequent ranks."
            ],
            "accent": accent_green
        },
        {
            "title": "21KM GROUP RELAY",
            "desc": "Special Team Category (4 Runners)",
            "points": [
                "1st Place Team: RM 1,000 (RM 250/runner)",
                "Awards restricted to Top 3 Teams only",
                "Relay exclusively organized for the 21KM distance",
                "Fosters team spirit and athletic club engagement."
            ],
            "accent": accent_blue
        },
        {
            "title": "10KM INDIVIDUAL",
            "desc": "Short Distance Competitive",
            "points": [
                "1st Place Open: RM 1,000",
                "1st Place Veteran: RM 800",
                "Prizes scaled at 50% of the 21KM baseline",
                "Keeps categories balanced while rewarding fast runners."
            ],
            "accent": text_white
        }
    ]

    for i, c in enumerate(cards_data):
        x_pos = Inches(0.75 + i * 4.1)
        # Card Background
        card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x_pos, y_pos, card_width, card_height)
        card.fill.solid()
        card.fill.fore_color.rgb = card_bg
        card.line.color.rgb = border_color
        
        # Text Frame
        tb = slide2.shapes.add_textbox(x_pos + Inches(0.2), y_pos + Inches(0.2), card_width - Inches(0.4), card_height - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        
        # Card Title
        p_title = tf.paragraphs[0]
        p_title.text = c["title"]
        p_title.font.name = 'Arial'
        p_title.font.size = Pt(16)
        p_title.font.bold = True
        p_title.font.color.rgb = c["accent"]
        p_title.space_after = Pt(4)
        
        # Card Desc
        p_desc = tf.add_paragraph()
        p_desc.text = c["desc"]
        p_desc.font.name = 'Arial'
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = text_muted
        p_desc.space_after = Pt(20)
        
        # Bullet Points
        for pt in c["points"]:
            p_pt = tf.add_paragraph()
            p_pt.text = "•  " + pt
            p_pt.font.name = 'Arial'
            p_pt.font.size = Pt(12)
            p_pt.font.color.rgb = text_white
            p_pt.space_after = Pt(10)

    # -------------------------------------------------------------
    # SLIDE 3: Option A: Premium Structure
    # -------------------------------------------------------------
    slide3 = prs.slides.add_slide(slide_layout)
    apply_slide_bg(slide3)
    add_title(slide3, "Option A: Premium Cash Structure (1st to 10th)", "Distributes cash down to the 10th position for all individual runs")

    # Table layout
    rows = 6
    cols = 7
    left = Inches(0.75)
    top = Inches(2.0)
    width = Inches(11.83)
    height = Inches(4.5)

    table_shape = slide3.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # Column widths
    table.columns[0].width = Inches(2.83)
    for col_idx in range(1, 6):
        table.columns[col_idx].width = Inches(1.5)
    table.columns[6].width = Inches(1.5)

    headers = ["Category", "1st", "2nd", "3rd", "4th", "5th", "6th-10th (each)"]
    data_a = [
        ["21KM Men/Women Open", "RM 2,000", "RM 1,200", "RM 800", "RM 500", "RM 300", "RM 150"],
        ["21KM Men/Women Veteran", "RM 1,500", "RM 900", "RM 600", "RM 400", "RM 250", "RM 120"],
        ["21KM Group Relay (Team)", "RM 1,000", "RM 600", "RM 400", "—", "—", "—"],
        ["10KM Men/Women Open", "RM 1,000", "RM 600", "RM 400", "RM 250", "RM 150", "RM 80"],
        ["10KM Men/Women Veteran", "RM 800", "RM 500", "RM 300", "RM 200", "RM 120", "RM 60"]
    ]

    # Populate Headers
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = table_header_bg
        p = cell.text_frame.paragraphs[0]
        p.text = header
        p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT
        p.font.name = 'Arial'
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_white

    # Populate Data
    for row_idx, row_data in enumerate(data_a):
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = card_bg
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT
            p.font.name = 'Arial'
            p.font.size = Pt(12)
            p.font.color.rgb = accent_green if (col_idx == 1 and val != "—") else text_white
            if col_idx == 0:
                p.font.bold = True

    # Total Box
    total_box = slide3.shapes.add_textbox(Inches(0.75), Inches(6.6), Inches(11.83), Inches(0.5))
    p_tot = total_box.text_frame.paragraphs[0]
    p_tot.text = "Total Cash Prize Budget Required (Option A): RM 31,640"
    p_tot.font.name = 'Arial'
    p_tot.font.size = Pt(14)
    p_tot.font.bold = True
    p_tot.font.color.rgb = accent_green

    # -------------------------------------------------------------
    # SLIDE 4: Option B: Budget-Saver Structure
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(slide_layout)
    apply_slide_bg(slide4)
    add_title(slide4, "Option B: Budget-Saver Structure (1st to 5th)", "Focuses cash prizes on the podium, rewarding 6th-10th with trophies only")

    table_shape_b = slide4.shapes.add_table(rows, cols, left, top, width, height)
    table_b = table_shape_b.table

    # Column widths
    table_b.columns[0].width = Inches(2.83)
    for col_idx in range(1, 6):
        table_b.columns[col_idx].width = Inches(1.5)
    table_b.columns[6].width = Inches(1.5)

    headers_b = ["Category", "1st", "2nd", "3rd", "4th", "5th", "6th-10th"]
    data_b = [
        ["21KM Men/Women Open", "RM 2,000", "RM 1,200", "RM 800", "RM 500", "RM 300", "Trophy Only"],
        ["21KM Men/Women Veteran", "RM 1,500", "RM 900", "RM 600", "RM 400", "RM 250", "Trophy Only"],
        ["21KM Group Relay (Team)", "RM 1,000", "RM 600", "RM 400", "—", "—", "—"],
        ["10KM Men/Women Open", "RM 1,000", "RM 600", "RM 400", "RM 250", "RM 150", "Trophy Only"],
        ["10KM Men/Women Veteran", "RM 800", "RM 500", "RM 300", "RM 200", "RM 120", "Trophy Only"]
    ]

    # Populate Headers
    for col_idx, header in enumerate(headers_b):
        cell = table_b.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = table_header_bg
        p = cell.text_frame.paragraphs[0]
        p.text = header
        p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT
        p.font.name = 'Arial'
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = text_white

    # Populate Data
    for row_idx, row_data in enumerate(data_b):
        for col_idx, val in enumerate(row_data):
            cell = table_b.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = card_bg
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT
            p.font.name = 'Arial'
            p.font.size = Pt(12)
            p.font.color.rgb = accent_green if (col_idx == 1 and val != "—") else (text_muted if val == "Trophy Only" else text_white)
            if col_idx == 0:
                p.font.bold = True

    # Total Box
    total_box_b = slide4.shapes.add_textbox(Inches(0.75), Inches(6.6), Inches(11.83), Inches(0.5))
    p_tot_b = total_box_b.text_frame.paragraphs[0]
    p_tot_b.text = "Total Cash Prize Budget Required (Option B): RM 25,440"
    p_tot_b.font.name = 'Arial'
    p_tot_b.font.size = Pt(14)
    p_tot_b.font.bold = True
    p_tot_b.font.color.rgb = accent_blue

    # -------------------------------------------------------------
    # SLIDE 5: Budget Comparison & Recommendation
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(slide_layout)
    apply_slide_bg(slide5)
    add_title(slide5, "Comparison & Recommendation", "Which structure matches your organizational strategy?")

    # Budgets Cards
    w = Inches(5.6)
    h = Inches(2.2)
    
    # Card Option A
    card_a = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(2.0), w, h)
    card_a.fill.solid()
    card_a.fill.fore_color.rgb = card_bg
    card_a.line.color.rgb = border_color
    
    tb_a = slide5.shapes.add_textbox(Inches(0.95), Inches(2.1), w - Inches(0.4), h - Inches(0.2))
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True
    p = tf_a.paragraphs[0]
    p.text = "OPTION A — PREMIUM STRUCTURE"
    p.font.name = 'Arial'
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = accent_green
    
    p = tf_a.add_paragraph()
    p.text = "Total Budget: RM 31,640"
    p.font.name = 'Arial'
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = text_white
    p.space_before = Pt(8)
    
    p = tf_a.add_paragraph()
    p.text = "• Maximizes elite & running club registrations by rewarding more depth.\n• Standard format for tier-1 national events (e.g. PBIM, Langkawi)."
    p.font.name = 'Arial'
    p.font.size = Pt(11)
    p.font.color.rgb = text_muted
    p.space_before = Pt(8)

    # Card Option B
    card_b = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.98), Inches(2.0), w, h)
    card_b.fill.solid()
    card_b.fill.fore_color.rgb = card_bg
    card_b.line.color.rgb = border_color
    
    tb_b = slide5.shapes.add_textbox(Inches(7.18), Inches(2.1), w - Inches(0.4), h - Inches(0.2))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p = tf_b.paragraphs[0]
    p.text = "OPTION B — BUDGET-SAVER"
    p.font.name = 'Arial'
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = accent_blue
    
    p = tf_b.add_paragraph()
    p.text = "Total Budget: RM 25,440"
    p.font.name = 'Arial'
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = text_white
    p.space_before = Pt(8)
    
    p = tf_b.add_paragraph()
    p.text = "• Saves RM 6,200 (19.6% reduction in total prize pool).\n• Maintains top-tier status while rewarding 6th-10th with high-quality trophies."
    p.font.name = 'Arial'
    p.font.size = Pt(11)
    p.font.color.rgb = text_muted
    p.space_before = Pt(8)

    # Recommendation Box
    rec_shape = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(4.6), Inches(11.83), Inches(2.0))
    rec_shape.fill.solid()
    rec_shape.fill.fore_color.rgb = card_bg
    rec_shape.line.color.rgb = border_color
    
    tb_rec = slide5.shapes.add_textbox(Inches(0.95), Inches(4.7), Inches(11.43), Inches(1.8))
    tf_rec = tb_rec.text_frame
    tf_rec.word_wrap = True
    
    p = tf_rec.paragraphs[0]
    p.text = "Our Recommendation:"
    p.font.name = 'Arial'
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = accent_green
    p.space_after = Pt(8)
    
    p = tf_rec.add_paragraph()
    p.text = "For an event targeting 8,000 participants, Option A is highly recommended. The additional RM 6,200 required to pay cash down to 10th place is easily offset by the incremental ticket revenue from competitive running clubs. By advertising cash prizes up to 10th place, you establish the event as a premier competitive milestone on the national calendar."
    p.font.name = 'Arial'
    p.font.size = Pt(13)
    p.font.color.rgb = text_white
    p.line_spacing = 1.3

    prs.save("/Users/shiban/.gemini/antigravity/scratch/prizes_presentation.pptx")
    print("Presentation created successfully at /Users/shiban/.gemini/antigravity/scratch/prizes_presentation.pptx")

if __name__ == '__main__':
    create_presentation()
