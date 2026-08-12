import streamlit as st
import re
import io
import os
from PIL import Image

# PowerPoint Generation
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# PDF Generation (fpdf2)
try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

# Document & Image Processing Libraries
try:
    import fitz  # PyMuPDF for PDF
except ImportError:
    fitz = None

try:
    from docx import Document  # python-docx for Word files
except ImportError:
    Document = None

try:
    import pytesseract  # Offline OCR for Images
except ImportError:
    pytesseract = None

# Streamlit पेज सेटिंग्स
st.set_page_config(
    page_title="Master Offline Model Paper PPT & PDF Maker",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
    [data-testid="stHeader"] {display: none;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge {display: none !important;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Master Offline Multi-Format PPT & PDF Maker (2026)")
st.write("फाइल अपलोड करें या सीधे टेक्स्ट पेस्ट करें। AI पार्सर की मदद से प्रोफेशनल PPT और PDF तैयार करें!")

# Session State इनिशियलाइजेशन
if "parsed_questions" not in st.session_state:
    st.session_state.parsed_questions = []

# --- साईडबार कॉन्फ़िगरेशन ---
st.sidebar.header("⚙️ सेटिंग्स एवं विकल्प")

slide_format = st.sidebar.selectbox(
    "PPT & PDF Slide Size चुनें:",
    ["16:9 (Widescreen)", "20:9 (Cinematic)", "4:3 (Standard)"]
)

input_choice = st.sidebar.radio(
    "डेटा इनपुट का तरीका चुनें:",
    ["📁 File Upload (.txt, .pdf, .docx, Image)", "✍️ Direct Text Paste"]
)

# --- टेक्स्ट पार्सिंग फंक्शन ---
def double_verify_and_parse(text):
    questions = []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    current_q = None
    q_pattern = re.compile(r'^(प्रश्न|\bQ\b|\bQ\d+|\d+[\.\)]|\bQuestion\b)', re.IGNORECASE)
    opt_pattern = re.compile(r'^(\([A-Da-dअ-द1-4]\)|[A-Da-dअ-द1-4][\.\)]|\b[A-Da-dअ-द][\)])')
    ans_pattern = re.compile(r'^(उत्तर|Answer|Ans|सही उत्तर)[\s:\-]', re.IGNORECASE)
    exp_pattern = re.compile(r'^(व्याख्या|Explanation|Exp|स्पष्टीकरण)[\s:\-]', re.IGNORECASE)

    for line in lines:
        if q_pattern.match(line):
            if current_q and current_q['question']:
                while len(current_q['options']) < 4:
                    current_q['options'].append(f"({chr(65+len(current_q['options']))}) विकल्प उपलब्ध नहीं")
                questions.append(current_q)
            current_q = {'question': line, 'options': [], 'answer': '', 'explanation': []}
            continue
            
        if current_q is None:
            current_q = {'question': line, 'options': [], 'answer': '', 'explanation': []}
            continue

        if opt_pattern.match(line):
            current_q['options'].append(line)
        elif ans_pattern.match(line):
            current_q['answer'] = line
        elif exp_pattern.match(line):
            exp_text = exp_pattern.sub('', line).strip()
            if exp_text:
                current_q['explanation'].append(exp_text)
        else:
            if current_q['explanation']:
                if len(current_q['explanation']) < 3:
                    current_q['explanation'].append(line)
            elif current_q['answer']:
                current_q['answer'] += " " + line
            elif current_q['options']:
                if len(current_q['options']) < 4:
                    current_q['options'].append(line)
                else:
                    current_q['explanation'].append(line)
            else:
                current_q['question'] += " " + line

    if current_q and current_q['question']:
        while len(current_q['options']) < 4:
            current_q['options'].append(f"({chr(65+len(current_q['options']))}) विकल्प उपलब्ध नहीं")
        questions.append(current_q)
        
    return questions

# --- PDF शेप हेल्प फंक्शन ---
def draw_safe_rect(pdf_obj, x, y, w, h, style='FD', r=0.15):
    """FPDF में बिना किसी एरर के बॉक्स ड्रॉ करने का सेफ फंक्शन"""
    if hasattr(pdf_obj, 'rounded_rect'):
        try:
            pdf_obj.rounded_rect(x, y, w, h, r=r, style=style)
            return
        except Exception:
            pass
    pdf_obj.rect(x, y, w, h, style=style)

# --- इनपुट सेक्शन ---
raw_text = ""

col_input, col_preview = st.columns([1, 1])

with col_input:
    st.subheader("📥 इनपुट डेटा")
    if input_choice == "📁 File Upload (.txt, .pdf, .docx, Image)":
        uploaded_file = st.file_uploader(
            "फाइल चुनें (.txt, .pdf, .docx, .png, .jpg, .jpeg)", 
            type=["txt", "pdf", "docx", "png", "jpg", "jpeg"]
        )
        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            try:
                if file_extension == "txt":
                    raw_text = file_bytes.decode("utf-8", errors="ignore")
                elif file_extension == "pdf":
                    if fitz is None:
                        st.error("⚠️ PyMuPDF (fitz) इंस्टॉल करें: `pip install PyMuPDF`")
                    else:
                        doc = fitz.open(stream=file_bytes, filetype="pdf")
                        for page in doc:
                            raw_text += page.get_text() + "\n"
                elif file_extension == "docx":
                    if Document is None:
                        st.error("⚠️ python-docx इंस्टॉल करें: `pip install python-docx`")
                    else:
                        doc = Document(io.BytesIO(file_bytes))
                        for para in doc.paragraphs:
                            raw_text += para.text + "\n"
                elif file_extension in ["png", "jpg", "jpeg"]:
                    if pytesseract is None:
                        st.error("⚠️ pytesseract इंस्टॉल करें: `pip install pytesseract`")
                    else:
                        image = Image.open(io.BytesIO(file_bytes))
                        raw_text = pytesseract.image_to_string(image, lang='hin+eng')
            except Exception as e:
                st.error(f"फाइल पढ़ने में त्रुटि: {e}")
    else:
        raw_text = st.text_area(
            "यहाँ प्रश्न, विकल्प, उत्तर और व्याख्या पेस्ट करें:",
            height=300,
            placeholder="प्रश्न 1: राजस्थान High Court 4th Grade Exam 2026 की मुख्य पीठ कहाँ स्थित है?\n(A) जयपुर\n(B) जोधपुर\n(C) अजमेर\n(D) बीकानेर\nउत्तर: (B) जोधपुर\nव्याख्या: मुख्य पीठ जोधपुर में तथा खंडपीठ जयपुर में स्थित है।"
        )

    if st.button("🔍 डेटा पार्स करें"):
        if raw_text.strip():
            st.session_state.parsed_questions = double_verify_and_parse(raw_text)
            st.success(f"कुल {len(st.session_state.parsed_questions)} प्रश्न सफलतापूर्वक पार्स किए गए!")
        else:
            st.warning("⚠️ कृपया पहले कोई फाइल अपलोड करें या टेक्स्ट लिखें!")

# --- पूर्वावलोकन एवं संपादन ---
with col_preview:
    st.subheader(f"📋 प्रश्न पूर्वावलोकन ({len(st.session_state.parsed_questions)})")
    if st.session_state.parsed_questions:
        for idx, q in enumerate(st.session_state.parsed_questions):
            with st.expander(f"प्र. {idx+1}: {q['question'][:50]}...", expanded=(idx == 0)):
                st.write(f"**प्रश्न:** {q['question']}")
                for opt in q['options']:
                    st.write(opt)
                if q['answer']:
                    st.success(f"**{q['answer']}**")
                if q['explanation']:
                    st.info(f"**व्याख्या:** {' '.join(q['explanation'])}")
                
                if st.button(f"🗑️ हटाएँ", key=f"del_{idx}"):
                    st.session_state.parsed_questions.pop(idx)
                    st.rerun()
    else:
        st.info("यहाँ पार्स किए गए प्रश्नों का पूर्वावलोकन दिखाई देगा।")

st.divider()

# --- PPTX एवं PDF जनरेटर बटन ---
if st.button("🚀 Master PPT & PDF जनरेट करें", type="primary", use_container_width=True):
    if not st.session_state.parsed_questions:
        st.warning("⚠️ जनरेट करने के लिए कोई प्रश्न उपलब्ध नहीं है! पहले डेटा पार्स करें।")
    else:
        parsed_questions = st.session_state.parsed_questions
        
        # PPTX स्ट्रक्चर तैयार करना
        prs = Presentation()
        
        if slide_format == "20:9 (Cinematic)":
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(6.0)
            q_font_size, opt_font_size, ans_font_size, exp_font_size = Pt(32), Pt(30), Pt(23), Pt(30)
            card_width, card_height = Inches(12.333), Inches(5.0)
            box_width, q_box_width, opt_box_width = Inches(11.733), Inches(12.3), Inches(12.0)
            exp_box_height, opt_left, opt_top, opt_space_before = Inches(3.2), Inches(0.9), Inches(2.5), Pt(2)

        elif slide_format == "16:9 (Widescreen)":
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            q_font_size, opt_font_size, ans_font_size, exp_font_size = Pt(40), Pt(40), Pt(28), Pt(32)
            card_width, card_height = Inches(11.733), Inches(6.4)
            box_width, q_box_width, opt_box_width = Inches(11.133), Inches(12.3), Inches(12.0)
            exp_box_height, opt_left, opt_top, opt_space_before = Inches(4.5), Inches(0.8), Inches(3.0), Pt(8)

        else:  # 4:3 Standard
            prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
            q_font_size, opt_font_size, ans_font_size, exp_font_size = Pt(30), Pt(28), Pt(24), Pt(24)
            card_width, card_height = Inches(8.8), Inches(6.4)
            box_width, q_box_width, opt_box_width = Inches(8.2), Inches(9.0), Inches(8.8)
            exp_box_height, opt_left, opt_top, opt_space_before = Inches(4.5), Inches(0.5), Inches(2.8), Pt(6)

        blank_layout = prs.slide_layouts[6]

        for q in parsed_questions:
            # SLIDE 1: Question + Options
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
                p = tf_opt.paragraphs[0] if opt_idx == 0 else tf_opt.add_paragraph()
                if opt_idx > 0:
                    p.space_before = opt_space_before
                p.text = opt
                p.font.name = 'Nirmala UI'
                p.font.size = opt_font_size
                p.font.bold = True
                p.font.color.rgb = RGBColor(0, 0, 0)

            # SLIDE 2: Answer + Explanation Card
            slide2 = prs.slides.add_slide(blank_layout)
            card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.5), card_width, card_height)
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(248, 250, 252)
            card.line.color.rgb = RGBColor(203, 213, 225)

            ans_banner = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(0.8), box_width, Inches(1.1))
            ans_banner.fill.solid()
            ans_banner.fill.fore_color.rgb = RGBColor(22, 163, 74)

            tf_ans = ans_banner.text_frame
            tf_ans.word_wrap = True
            p_ans = tf_ans.paragraphs[0]
            p_ans.text = q['answer'] if q['answer'] else "उत्तर उपलब्ध नहीं"
            p_ans.font.name = 'Nirmala UI'
            p_ans.font.size = ans_font_size
            p_ans.font.bold = True
            p_ans.font.color.rgb = RGBColor(255, 255, 255)

            exp_box = slide2.shapes.add_textbox(Inches(1.1), Inches(2.1), box_width, exp_box_height)
            tf_exp = exp_box.text_frame
            tf_exp.word_wrap = True
            p_exp_title = tf_exp.paragraphs[0]
            p_exp_title.text = "व्याख्या:"
            p_exp_title.font.name = 'Nirmala UI'
            p_exp_title.font.size = Pt(26)
            p_exp_title.font.bold = True

            expl_lines = q['explanation'][:3] if q['explanation'] else ["कोई व्याख्या दर्ज नहीं है।"]
            for exp_line in expl_lines:
                p_exp = tf_exp.add_paragraph()
                p_exp.space_before = Pt(6)
                p_exp.text = f"• {exp_line}" if not exp_line.startswith('•') else exp_line
                p_exp.font.name = 'Nirmala UI'
                p_exp.font.size = exp_font_size

        # Save PPT Stream
        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)

        # PDF Generation (Matching Exact PPT Landscape Slide Layout, Styling & Dimensions)
        pdf_bytes = b""
        if FPDF is not None:
            try:
                # PPT की चुनी गई साइज़ और थीम के अनुसार PDF पैरामीटर्स (Inches में)
                if slide_format == "20:9 (Cinematic)":
                    pdf_w, pdf_h = 13.333, 6.0
                    p_q_font, p_opt_font, p_ans_font, p_exp_font = 32, 30, 23, 30
                    p_card_w, p_card_h = 12.333, 5.0
                    p_box_w, p_q_box_w, p_opt_box_w = 11.733, 12.3, 12.0
                    p_opt_l, p_opt_t = 0.9, 2.5
                elif slide_format == "16:9 (Widescreen)":
                    pdf_w, pdf_h = 13.333, 7.5
                    p_q_font, p_opt_font, p_ans_font, p_exp_font = 40, 40, 28, 32
                    p_card_w, p_card_h = 11.733, 6.4
                    p_box_w, p_q_box_w, p_opt_box_w = 11.133, 12.3, 12.0
                    p_opt_l, p_opt_t = 0.8, 3.0
                else:  # 4:3 Standard
                    pdf_w, pdf_h = 10.0, 7.5
                    p_q_font, p_opt_font, p_ans_font, p_exp_font = 30, 28, 24, 24
                    p_card_w, p_card_h = 8.8, 6.4
                    p_box_w, p_q_box_w, p_opt_box_w = 8.2, 9.0, 8.8
                    p_opt_l, p_opt_t = 0.5, 2.8

                # PPT स्लाइड डाइमेंशन के अनुसार PDF बनाएँ
                pdf = FPDF(unit='in', format=(pdf_w, pdf_h))
                pdf.set_auto_page_break(auto=False)

                font_path = "Nirmala.ttf" if os.path.exists("Nirmala.ttf") else ("nirmala.ttf" if os.path.exists("nirmala.ttf") else None)
                font_name = "Helvetica"
                if font_path:
                    pdf.add_font("Nirmala", style="", fname=font_path)
                    pdf.add_font("Nirmala", style="B", fname=font_path)
                    font_name = "Nirmala"

                for q in parsed_questions:
                    # ==================== PDF SLIDE 1: Question + Options ====================
                    pdf.add_page()
                    
                    # Question Text (Red, Bold, Exactly like PPT)
                    pdf.set_font(font_name, style="B" if font_path else "B", size=p_q_font)
                    pdf.set_text_color(255, 0, 0)
                    pdf.set_xy(0.5, 0.4)
                    pdf.multi_cell(w=p_q_box_w, h=(p_q_font/72.0)*1.25, text=q['question'], border=0, align='L')
                    
                    # Options Text (Black, Bold, Dynamic Overlap Prevention)
                    pdf.set_font(font_name, style="B" if font_path else "B", size=p_opt_font)
                    pdf.set_text_color(0, 0, 0)
                    opt_start_y = max(p_opt_t, pdf.get_y() + 0.15)
                    
                    for opt in q['options']:
                        pdf.set_xy(p_opt_l, opt_start_y)
                        pdf.multi_cell(w=p_opt_box_w, h=(p_opt_font/72.0)*1.25, text=opt, border=0, align='L')
                        opt_start_y = pdf.get_y() + 0.08

                    # ==================== PDF SLIDE 2: Answer + Explanation Card ====================
                    pdf.add_page()

                    # Card Background (Soft Slate Color)
                    pdf.set_fill_color(248, 250, 252)
                    pdf.set_draw_color(203, 213, 225)
                    draw_safe_rect(pdf, 0.8, 0.5, p_card_w, p_card_h, style='FD', r=0.2)

                    # Green Answer Banner
                    pdf.set_fill_color(22, 163, 74)
                    pdf.set_draw_color(22, 163, 74)
                    draw_safe_rect(pdf, 1.1, 0.8, p_box_w, 1.1, style='F', r=0.15)

                    # Answer Text (White, Bold)
                    pdf.set_font(font_name, style="B" if font_path else "B", size=p_ans_font)
                    pdf.set_text_color(255, 255, 255)
                    pdf.set_xy(1.2, 0.95)
                    ans_text = q['answer'] if q['answer'] else "उत्तर उपलब्ध नहीं"
                    pdf.multi_cell(w=p_box_w - 0.2, h=(p_ans_font/72.0)*1.2, text=ans_text, border=0, align='L')

                    # Explanation Heading ("व्याख्या:")
                    pdf.set_font(font_name, style="B" if font_path else "B", size=26)
                    pdf.set_text_color(0, 0, 0)
                    pdf.set_xy(1.1, 2.1)
                    pdf.multi_cell(w=p_box_w, h=(26/72.0)*1.2, text="व्याख्या:", border=0, align='L')

                    # Explanation Bullet Points
                    pdf.set_font(font_name, style="" if font_path else "", size=p_exp_font)
                    expl_lines = q['explanation'][:3] if q['explanation'] else ["कोई व्याख्या दर्ज नहीं है।"]
                    curr_exp_y = 2.6
                    for exp_line in expl_lines:
                        formatted_line = f"• {exp_line}" if not exp_line.startswith('•') else exp_line
                        pdf.set_xy(1.1, curr_exp_y)
                        pdf.multi_cell(w=p_box_w, h=(p_exp_font/72.0)*1.25, text=formatted_line, border=0, align='L')
                        curr_exp_y = pdf.get_y() + 0.08

                pdf_bytes = bytes(pdf.output())
            except Exception as e:
                st.error(f"PDF जनरेट करने में त्रुटि: {e}")

        st.success("🎉 आपकी PPTX और PDF सफलतापूर्वक तैयार हो गई हैं!")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="📥 PPTX डाउनलोड करें",
                data=ppt_stream,
                file_name="Master_Model_Paper_2026.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
        with c2:
            if pdf_bytes:
                st.download_button(
                    label="📄 PDF डाउनलोड करें",
                    data=pdf_bytes,
                    file_name="Master_Model_Paper_2026.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
