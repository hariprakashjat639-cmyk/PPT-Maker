from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_presentation(questions_data, output_filename="exam_model_paper.pptx"):
    # Presentation initialize karein (16:9 Widescreen)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Blank slide layout lein
    blank_layout = prs.slide_layouts[6]
    
    # Colors Definition
    BG_COLOR = RGBColor(245, 247, 250)         # Light Gray / Clean Background
    PRIMARY_COLOR = RGBColor(26, 54, 93)       # Dark Navy Blue
    ACCENT_COLOR = RGBColor(221, 107, 32)      # Orange / Highlight
    TEXT_COLOR = RGBColor(45, 55, 72)          # Dark Slate
    BOX_BG = RGBColor(255, 255, 255)           # White for options
    SPECIAL_BG = RGBColor(254, 235, 226)       # Soft Peach for Special Type 1
    SPECIAL_BG_2 = RGBColor(235, 248, 255)     # Soft Blue for Special Type 2
    SPECIAL_BG_3 = RGBColor(254, 243, 199)     # Soft Yellow for Special Type 3

    for idx, q in enumerate(questions_data):
        slide = prs.slides.add_slide(blank_layout)
        
        # 1. Background Shape
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background() # No border
        
        # 2. Header / Top Banner (Question Number & Title)
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.73), Inches(0.8))
        tf_h = header_box.text_frame
        tf_h.word_wrap = True
        p_h = tf_h.paragraphs[0]
        p_h.text = f"Prashan {idx + 1} / {len(questions_data)}"
        p_h.font.size = Pt(20)
        p_h.font.bold = True
        p_h.font.color.rgb = ACCENT_COLOR
        
        q_type = q.get("type", "standard") # Default standard MCQ
        
        # ==========================================
        # CONDITIONAL RENDERING FOR 3 SPECIAL TYPES
        # ==========================================
        
        if q_type == "special_type_1":
            # --- SPECIAL TYPE 1 LAYOUT (e.g., Statement / Assertion based) ---
            # Background highlight for special type 1
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.73), Inches(5.4))
            panel.fill.solid()
            panel.fill.fore_color.rgb = SPECIAL_BG
            panel.line.color.rgb = ACCENT_COLOR
            panel.line.width = Pt(1.5)
            
            # Question Text
            q_box = slide.shapes.add_textbox(Inches(1.1), Inches(1.6), Inches(11.13), Inches(1.5))
            tf_q = q_box.text_frame
            tf_q.word_wrap = True
            p_q = tf_q.paragraphs[0]
            p_q.text = f"[Vishisht Prashan Type 1]: {q['question']}"
            p_q.font.size = Pt(22)
            p_q.font.bold = True
            p_q.font.color.rgb = PRIMARY_COLOR
            
            # Options in special layout (Vertical stacked cards)
            options = q.get("options", [])
            top_pos = 3.3
            for opt_idx, option in enumerate(options):
                opt_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(top_pos), Inches(11.13), Inches(0.8))
                opt_box.fill.solid()
                opt_box.fill.fore_color.rgb = BOX_BG
                opt_box.line.color.rgb = RGBColor(226, 232, 240)
                
                tf_opt = opt_box.text_frame
                tf_opt.word_wrap = True
                p_opt = tf_opt.paragraphs[0]
                p_opt.text = f"({chr(65+opt_idx)})  {option}"
                p_opt.font.size = Pt(18)
                p_opt.font.color.rgb = TEXT_COLOR
                top_pos += 0.95

        elif q_type == "special_type_2":
            # --- SPECIAL TYPE 2 LAYOUT (e.g., Match the following / Column A & B) ---
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.73), Inches(5.4))
            panel.fill.solid()
            panel.fill.fore_color.rgb = SPECIAL_BG_2
            panel.line.color.rgb = PRIMARY_COLOR
            panel.line.width = Pt(1.5)
            
            q_box = slide.shapes.add_textbox(Inches(1.1), Inches(1.6), Inches(11.13), Inches(1.2))
            tf_q = q_box.text_frame
            tf_q.word_wrap = True
            p_q = tf_q.paragraphs[0]
            p_q.text = f"[Milan Karein Type 2]: {q['question']}"
            p_q.font.size = Pt(22)
            p_q.font.bold = True
            p_q.font.color.rgb = PRIMARY_COLOR
            
            # Split view for columns or specific text
            col_box = slide.shapes.add_textbox(Inches(1.1), Inches(3.0), Inches(11.13), Inches(3.5))
            tf_col = col_box.text_frame
            tf_col.word_wrap = True
            for line in q.get("details", []):
                p_line = tf_col.add_paragraph()
                p_line.text = line
                p_line.font.size = Pt(18)
                p_line.font.color.rgb = TEXT_COLOR
                p_line.space_after = Pt(10)

        elif q_type == "special_type_3":
            # --- SPECIAL TYPE 3 LAYOUT (e.g., Fact / Explanation / One-liner Highlight) ---
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.73), Inches(5.4))
            panel.fill.solid()
            panel.fill.fore_color.rgb = SPECIAL_BG_3
            panel.line.color.rgb = RGBColor(217, 119, 6)
            panel.line.width = Pt(1.5)
            
            q_box = slide.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.13), Inches(4.5))
            tf_q = q_box.text_frame
            tf_q.word_wrap = True
            
            p_q = tf_q.paragraphs[0]
            p_q.text = f"[Important Fact Type 3]"
            p_q.font.size = Pt(18)
            p_q.font.bold = True
            p_q.font.color.rgb = RGBColor(180, 83, 9)
            
            p_q2 = tf_q.add_paragraph()
            p_q2.text = q['question']
            p_q2.font.size = Pt(24)
            p_q2.font.bold = True
            p_q2.font.color.rgb = PRIMARY_COLOR
            p_q2.space_before = Pt(15)

        else:
            # --- STANDARD MCQ LAYOUT (Purana Style Safe & Intact) ---
            q_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.73), Inches(1.5))
            tf_q = q_box.text_frame
            tf_q.word_wrap = True
            p_q = tf_q.paragraphs[0]
            p_q.text = q['question']
            p_q.font.size = Pt(22)
            p_q.font.bold = True
            p_q.font.color.rgb = PRIMARY_COLOR
            
            # Standard Options Grid (2x2 layout)
            options = q.get("options", [])
            coords = [
                (Inches(0.8), Inches(3.2)),   # Option A
                (Inches(6.8), Inches(3.2)),   # Option B
                (Inches(0.8), Inches(5.0)),   # Option C
                (Inches(6.8), Inches(5.0))    # Option D
            ]
            
            for opt_idx, option in enumerate(options):
                if opt_idx < len(coords):
                    x, y = coords[opt_idx]
                    opt_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(1.5))
                    opt_box.fill.solid()
                    opt_box.fill.fore_color.rgb = BOX_BG
                    opt_box.line.color.rgb = RGBColor(226, 232, 240)
                    opt_box.line.width = Pt(1)
                    
                    tf_opt = opt_box.text_frame
                    tf_opt.word_wrap = True
                    p_opt = tf_opt.paragraphs[0]
                    p_opt.text = f"({chr(65+opt_idx)})  {option}"
                    p_opt.font.size = Pt(18)
                    p_opt.font.color.rgb = TEXT_COLOR

    # Presentation Save Karein
    prs.save(output_filename)
    print(f"Presentation successfully saved as '{output_filename}'")

# Example Data Structure with Standard + 3 Special Types
sample_questions = [
    {
        "type": "standard",
        "question": "Rajasthan High Court ki sthapna kab hui thi?",
        "options": ["1949", "1950", "1956", "1952"]
    },
    {
        "type": "special_type_1",
        "question": "Kathan (A): Rajasthan me sabse zyada jile kis sambhag me hain? Karan (R): Jodhpur aur Jaipur me 7-7 jile hain.",
        "options": ["A aur R dono sahi hain, R, A ki sahi vyakhya hai", "A aur R dono sahi hain par R sahi vyakhya nahi hai", "A sahi hai par R galat hai", "A galat hai par R sahi hai"]
    },
    {
        "type": "special_type_2",
        "question": "Nimnlikhit ka sahi milan karein (Aadiwasi mele aur sthan):",
        "details": ["1. Beneshwar Mela -> Dungarpur", "2. Kaila Devi Mela -> Karauli", "3. Ramdevra Mela -> Jaisalmer"]
    },
    {
        "type": "special_type_3",
        "question": "Rajasthan ka rajkiya pashu (Chinkara) vanyajiv ki shreni me kis varsh ghoshit kiya gaya tha? (Mahatvapurna Tathy)",
    }
]

# Run code to generate PPT
if __name__ == "__main__":
    create_presentation(sample_questions)
