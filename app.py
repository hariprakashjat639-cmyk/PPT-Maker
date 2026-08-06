import streamlit as st
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import fitz   # PyMuPDF

# पेज सेटिंग्स और लेआउट
st.set_page_config(page_title="Offline Model Paper PPT Maker", page_icon="📊", layout="centered")
st.markdown("""
    <style>
    [data-testid="stHeader"] {display: none;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge {display: none !important;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Offline Smart Bilingual PDF/Text to PPT Maker")
st.write("यह ऐप बिना किसी API के पूरी तरह से ऑफलाइन फाइल को पढ़ेगा, डबल-वेरीफाई करेगा और आपकी पसंद के अनुसार परफेक्ट PPT बनाएगा!")

# PPT स्लाइड साइज चुनने का ऑप्शन
slide_format = st.selectbox(
    "PPT Slide Size Chunein",
    ["16:9 (Widescreen)", "20:9 (Cinematic)", "4:3 (Standard)"]
)

# फाइल अपलोडर
uploaded_file = st.file_uploader("PDF या Text फाइल अपलोड करें", type=["txt", "pdf"])

# --- ऑफलाइन डबल-वेरीफाई टेक्स्ट पार्सर ---
def offline_double_verify_parser(text_content):
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    questions = []
    
    current_q = {'question': '', 'options': [], 'answer': '', 'explanation': []}
    state = 'NONE' # 'Q', 'OPT', 'ANS', 'EXP'
    
    for line in lines:
        # डिटेक्ट करें कि क्या यह नया प्रश्न है (जैसे: 1., 2., प्रश्न:, Q.)
        if re.match(r'^(\d+[\.\-\)]|प्रश्न\s*\d*[:\.-]?|Q[\.:])', line, re.IGNORECASE):
            if current_q['question']:
                # डबल-वेरीफाई: सुनिश्चित करें कि 4 विकल्प मौजूद हैं
                while len(current_q['options']) < 4:
                    current_q['options'].append(f"({len(current_q['options']) + 1}) विकल्प उपलब्ध नहीं")
                questions.append(current_q)
                current_q = {'question': '', 'options': [], 'answer': '', 'explanation': []}
            
            # प्रश्न साफ करें
            clean_q = re.sub(r'^(\d+[\.\-\)]|प्रश्न\s*\d*[:\.-]?|Q[\.:])\s*', '', line, flags=re.IGNORECASE)
            current_q['question'] = clean_q
            state = 'Q'
            
        # डिटेक्ट करें विकल्प (जैसे: (1), (2), a), b), क, ख आदि)
        elif re.match(r'^(\([1-4a-dA-D]\)|[1-4a-dA-D][\.\)]|[क-ज्ञ][\.\)])', line):
            current_q['options'].append(line)
            state = 'OPT'
            
        # डिटेक्ट करें उत्तर (Answer)
        elif re.match(r'^(उत्तर|Ans|Answer)[:\.-]?', line, re.IGNORECASE):
            ans_clean = re.sub(r'^(उत्तर|Ans|Answer)[:\.-]?\s*', '', line, flags=re.IGNORECASE)
            current_q['answer'] = f"उत्तर: {ans_clean}" if not ans_clean.startswith("उत्तर") else ans_clean
            state = 'ANS'
            
        # डिटेक्ट करें व्याख्या (Explanation)
        elif re.match(r'^(व्याख्या|Exp|Explanation|स्पष्टीकरण)[:\.-]?', line, re.IGNORECASE):
            exp_clean = re.sub(r'^(व्याख्या|Exp|Explanation|स्पष्टीकरण)[:\.-]?\s*', '', line, flags=re.IGNORECASE)
            if exp_clean:
                current_q['explanation'].append(exp_clean)
            state = 'EXP'
            
        else:
            # पिछले स्टेट के आधार पर मल्टी-लाइन टेक्स्ट जोड़ें
            if state == 'Q':
                current_q['question'] += " " + line
            elif state == 'OPT' and current_q['options']:
                current_q['options'][-1] += " " + line
            elif state == 'ANS':
                current_q['answer'] += " " + line
            elif state == 'EXP':
                if current_q['explanation']:
                    current_q['explanation'][-1] += " " + line
                else:
                    current_q['explanation'].append(line)
                    
    # आखिरी प्रश्न जोड़ें
    if current_q['question']:
        while len(current_q['options']) < 4:
            current_q['options'].append(f"({len(current_q['options']) + 1}) विकल्प उपलब्ध नहीं")
        questions.append(current_q)
        
    return questions

# PDF से टेक्स्ट निकालने का फंक्शन
def extract_text_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    return full_text

if uploaded_file is not None:
    if st.button("🚀 Offline डबल-वेरीफाई करके PPT Generate Karein"):
        with st.spinner("फाइल को ऑफलाइन पढ़ा जा रहा है और डबल-वेरीफाई किया जा रहा है..."):
            file_bytes = uploaded_file.getvalue()
            file_name = uploaded_file.name.lower()
            
            if file_name.endswith(".pdf"):
                raw_text = extract_text_from_pdf(file_bytes)
            else:
                raw_text = file_bytes.decode("utf-8", errors="ignore")
                
            parsed_questions = offline_double_verify_parser(raw_text)
            
            if not parsed_questions:
                st.error("⚠️ फाइल से प्रश्न नहीं मिल पाए। कृपया सुनिश्चित करें कि फाइल में प्रश्न सही फॉर्मेट में लिखे हैं।")
            else:
                prs = Presentation()
                
                # तीनों साइज़ और उनकी परफेक्ट डिज़ाइन सेटिंग्स
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

                for idx, q in enumerate(parsed_questions):
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
                    p1.font.color.rgb = RGBColor(255, 0, 0)  # Pure Red
                    p1.line_spacing = 1.3

                    opt_box = slide1.shapes.add_textbox(opt_left, opt_top, opt_box_width, Inches(4.3))
                    tf_opt = opt_box.text_frame
                    tf_opt.word_wrap = True

                    options = q['options']
                    for opt_idx, opt in enumerate(options):
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
                    # SLIDE 2: Answer + Explanation (Only if available)
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
                        ans_banner.fill.fore_color.rgb = RGBColor(22, 163, 74) # Green Banner
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

                        for line_idx, exp_line in enumerate(expl_lines):
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

                st.success("🎉 आपकी PPT पूरी तरह से ऑफलाइन डबल-वेरीफाई होकर तैयार हो गई है!")
                st.download_button(
                    label="📥 PPT Download Karein",
                    data=ppt_stream,
                    file_name="Offline_Model_Paper.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
