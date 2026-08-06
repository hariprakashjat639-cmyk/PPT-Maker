import streamlit as st
import re
import io
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# पेज सेटिंग्स
st.set_page_config(page_title="Offline Model Paper PPT Maker", page_icon="📊", layout="centered")
st.markdown("""
    <style>
    [data-testid="stHeader"] {display: none;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Offline PDF/Photo to PPT Maker (No API Key)")
st.write("बिना किसी API Key के, आपके पुराने स्मार्ट PPT स्टाइल के साथ!")

# PPT स्लाइड साइज
slide_format = st.selectbox(
    "PPT Slide Size चुनें",
    ["16:9 (Widescreen)", "20:9 (Cinematic)", "4:3 (Standard)"]
)

uploaded_file = st.file_uploader("PDF या Photo अपलोड करें", type=["pdf", "png", "jpg", "jpeg"])

# --- फाइल से टेक्स्ट निकालने का ऑफलाइन तरीका (OCR) ---
def extract_text_offline(file_bytes, file_name):
    extracted_text = ""
    try:
        if file_name.endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page in doc:
                # पहले डायरेक्ट टेक्स्ट निकालने की कोशिश
                text = page.get_text("text")
                if len(text.strip()) > 20:
                    extracted_text += text + "\n"
                else:
                    # अगर टेक्स्ट नहीं मिला तो फोटो खींचकर OCR करें
                    pix = page.get_pixmap(dpi=200)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    # हिंदी और अंग्रेजी दोनों पढ़ने के लिए lang='hin+eng'
                    extracted_text += pytesseract.image_to_string(img, lang='hin+eng') + "\n"
        else:
            img = Image.open(io.BytesIO(file_bytes))
            extracted_text += pytesseract.image_to_string(img, lang='hin+eng')
    except Exception as e:
        st.error(f"टेक्स्ट निकालने में एरर: {e}")
    return extracted_text

# --- टेक्स्ट को प्रश्न, विकल्प और उत्तर में तोड़ने का स्मार्ट लॉजिक ---
def parse_raw_text(text):
    questions = []
    # यह लॉजिक "प्र.", "Q.", "1." जैसे शब्दों को ढूँढकर प्रश्नों को अलग करता है
    raw_blocks = re.split(r'\n(?=[Qप्र]\.?\s*\d+|\d+\.)', text)
    
    for block in raw_blocks:
        if not block.strip(): continue
        
        q_data = {'question': '', 'options': [], 'answer': '', 'explanation': []}
        lines = block.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line: continue
            
            # विकल्प ढूँढने का लॉजिक (A, B, C, D या 1, 2, 3, 4)
            if re.match(r'^[\(]?[A-Da-d1-4][\)\.]\s+', line) or line.startswith('O'):
                q_data['options'].append(line)
            # उत्तर ढूँढने का लॉजिक
            elif "उत्तर" in line or "Ans" in line:
                q_data['answer'] = line
            # व्याख्या ढूँढने का लॉजिक
            elif "व्याख्या" in line or "Exp" in line:
                q_data['explanation'].append(line)
            # बाकी सब प्रश्न है
            else:
                if not q_data['options']:  # अगर विकल्प अभी तक नहीं आए हैं, तो वह प्रश्न का हिस्सा है
                    q_data['question'] += " " + line
                    
        # अगर विकल्प नहीं मिले, तो डमी विकल्प डाल दें (ताकि PPT की डिज़ाइन खराब न हो)
        if q_data['question'] and not q_data['options']:
            q_data['options'] = ["(1) विकल्प 1", "(2) विकल्प 2", "(3) विकल्प 3", "(4) विकल्प 4"]
            
        if q_data['question']:
            questions.append(q_data)
            
    return questions

if uploaded_file is not None:
    if st.button("🚀 PPT Generate करें"):
        with st.spinner("सिस्टम फाइल को पढ़ रहा है और आपकी स्टाइलिश PPT बना रहा है..."):
            file_bytes = uploaded_file.getvalue()
            file_name = uploaded_file.name.lower()
            
            # 1. टेक्स्ट निकालें
            raw_text = extract_text_offline(file_bytes, file_name)
            
            # 2. टेक्स्ट को स्ट्रक्चर करें
            if raw_text:
                parsed_questions = parse_raw_text(raw_text)
                
                prs = Presentation()
                
                # --- आपकी ओरिजिनल PPT डिज़ाइन और साइज़ सेटिंग्स ---
                if slide_format == "20:9 (Cinematic)":
                    prs.slide_width = Inches(13.333)
                    prs.slide_height = Inches(6.0)
                    q_font_size = Pt(32)
                    opt_font_size = Pt(30)
                    ans_font_size = Pt(23)
                    exp_font_size = Pt(30)
                    card_width = Inches(12.333)
                    card_height = Inches(5.0)
                    box_width = Inches(11.733)
                    q_box_width = Inches(12.3)
                    opt_box_width = Inches(12.0)
                    exp_box_height = Inches(3.2)
                    opt_left = Inches(0.9)
                    opt_top = Inches(2.5)
                    opt_space_before = Pt(2)
                elif slide_format == "16:9 (Widescreen)":
                    prs.slide_width = Inches(13.333)
                    prs.slide_height = Inches(7.5)
                    q_font_size = Pt(40)
                    opt_font_size = Pt(40)
                    ans_font_size = Pt(28)
                    exp_font_size = Pt(30)
                    card_width = Inches(11.733)
                    card_height = Inches(6.4)
                    box_width = Inches(11.133)
                    q_box_width = Inches(12.3)
                    opt_box_width = Inches(12.0)
                    exp_box_height = Inches(4.5)
                    opt_left = Inches(0.6)
                    opt_top = Inches(3.1)
                    opt_space_before = Pt(7)
                elif slide_format == "4:3 (Standard)":
                    prs.slide_width = Inches(10)
                    prs.slide_height = Inches(7.5)
                    q_font_size = Pt(30)
                    opt_font_size = Pt(28)
                    ans_font_size = Pt(24)
                    exp_font_size = Pt(24)
                    card_width = Inches(8.8)
                    card_height = Inches(6.4)
                    box_width = Inches(8.2)
                    q_box_width = Inches(9.0)
                    opt_box_width = Inches(8.8)
                    exp_box_height = Inches(4.5)
                    opt_left = Inches(0.5)
                    opt_top = Inches(2.8)
                    opt_space_before = Pt(6)

                blank_layout = prs.slide_layouts[6]

                for q in parsed_questions:
                    # ==========================================
                    # SLIDE 1: Question + Options
                    # ==========================================
                    slide1 = prs.slides.add_slide(blank_layout)

                    q_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.4), q_box_width, Inches(2.5))
                    tf1 = q_box.text_frame
                    tf1.word_wrap = True
                    p1 = tf1.paragraphs[0]
                    p1.text = q['question']
                    p1.font.name = 'Nirmala UI'
                    p1.font.size = q_font_size
                    p1.font.bold = True
                    p1.font.color.rgb = RGBColor(255, 0, 0)
                    p1.line_spacing = 1.3

                    opt_box = slide1.shapes.add_textbox(opt_left, opt_top, opt_box_width, Inches(4.3))
                    tf_opt = opt_box.text_frame
                    tf_opt.word_wrap = True

                    for opt_idx, opt in enumerate(q['options']):
                        if opt_idx == 0:
                            p = tf_opt.paragraphs[0]
                        else:
                            p = tf_opt.add_paragraph()
                            p.space_before = opt_space_before
                        
                        p.text = opt
                        p.font.name = 'Nirmala UI'
                        p.font.size = opt_font_size
                        p.font.bold = True
                        p.font.color.rgb = RGBColor(0, 0, 0)
                        p.line_spacing = 1.4

                    # ==========================================
                    # SLIDE 2: Answer + Explanation (स्मार्ट नियम: केवल तभी बनेगी जब डेटा हो)
                    # ==========================================
                    has_answer_data = bool(q['answer'].strip() or q['explanation'])
                    
                    if has_answer_data:
                        slide2 = prs.slides.add_slide(blank_layout)

                        card = slide2.shapes.add_shape(
                            MSO_SHAPE.ROUNDED_RECTANGLE, 
                            Inches(0.8), Inches(0.5), card_width, card_height
                        )
                        card.fill.solid()
                        card.fill.fore_color.rgb = RGBColor(248, 250, 252)
                        card.line.color.rgb = RGBColor(203, 213, 225)
                        card.line.width = Pt(1.5)

                        ans_banner = slide2.shapes.add_shape(
                            MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(1.1), Inches(0.8), box_width, Inches(1.1)
                        )
                        ans_banner.fill.solid()
                        ans_banner.fill.fore_color.rgb = RGBColor(22, 163, 74) 
                        ans_banner.line.color.rgb = RGBColor(22, 163, 74)

                        tf_ans = ans_banner.text_frame
                        tf_ans.word_wrap = True
                        p_ans = tf_ans.paragraphs[0]
                        p_ans.text = q['answer'] if q['answer'] else "उत्तर उपलब्ध नहीं है"
                        p_ans.font.name = 'Nirmala UI'
                        p_ans.font.size = ans_font_size
                        p_ans.font.bold = True
                        p_ans.font.color.rgb = RGBColor(255, 255, 255)

                        exp_box = slide2.shapes.add_textbox(Inches(1.1), Inches(2.0), box_width, exp_box_height)
                        tf_exp = exp_box.text_frame
                        tf_exp.word_wrap = True

                        p_exp_title = tf_exp.paragraphs[0]
                        p_exp_title.text = "व्याख्या:"
                        p_exp_title.font.name = 'Nirmala UI'
                        p_exp_title.font.size = Pt(26)
                        p_exp_title.font.bold = True
                        p_exp_title.font.color.rgb = RGBColor(30, 41, 59)

                        expl_lines = q['explanation'] if q['explanation'] else ["व्याख्या उपलब्ध नहीं है।"]

                        for exp_line in expl_lines:
                            p_exp = tf_exp.add_paragraph()
                            p_exp.space_before = Pt(6)
                            p_exp.line_spacing = 1.25
                            p_exp.text = f"• {exp_line}" if not exp_line.startswith('•') else exp_line
                            p_exp.font.name = 'Nirmala UI'
                            p_exp.font.size = exp_font_size
                            p_exp.font.color.rgb = RGBColor(0, 0, 0)

                ppt_stream = io.BytesIO()
                prs.save(ppt_stream)
                ppt_stream.seek(0)

                st.success("🎉 आपकी स्टाइलिश PPT बिल्कुल तैयार है!")
                st.download_button(
                    label="📥 PPT Download करें",
                    data=ppt_stream,
                    file_name="Offline_Model_Paper.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.warning("फाइल से कोई टेक्स्ट नहीं मिल पाया। कृपया साफ फोटो या पीडीएफ अपलोड करें।")
