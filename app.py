import streamlit as st
import re
import io
import os
import subprocess
import tempfile
from PIL import Image

# PowerPoint Generation
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn  # XML लेवल पर हिंदी/देवनागरी फॉन्ट सेट करने के लिए

# Document & Image Processing
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

# Streamlit Page Config
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
st.write("फाइल अपलोड करें या सीधे टेक्स्ट पेस्ट करें। यह आपके पुराने सटीक स्टाइल वाली PPT और 100% सेम लेआउट वाली PDF तैयार करेगा!")

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
    ["📁 File Upload Karein (.txt, .pdf, .docx, Image)", "✍️ Direct Text Paste Karein"]
)

# --- देवनागरी/हिंदी फॉन्ट फिक्सर फंक्शन ---
def set_hindi_font(paragraph, font_name="Nirmala UI"):
    """PPTX के Complex Script (Devanagari) को फोर्स करने के लिए विशेष फंक्शन"""
    paragraph.font.name = font_name
    for run in paragraph.runs:
        run.font.name = font_name
        rPr = run._r.get_or_add_rPr()
        rPr.set(qn('a:cs'), font_name) # Complex Script (Hindi/Devanagari)
        rPr.set(qn('a:ea'), font_name) # East Asian

# --- टेक्स्ट पार्सिंग फंक्शन (आपके पुराने लॉजिक के अनुसार) ---
def double_verify_and_parse(text):
    questions = []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    current_q = None
    q_pattern = re.compile(r'^(प्रश्न|\bQ\b|\bQ\d+|\d+[\.\)]|\bQuestion\b)', re.IGNORECASE)
    opt_pattern = re.compile(r'^(\([A-Da-dअ-द1-4]\)|[A-Da-dअ-द1-4][\.\)]|\b[A-Da-dअ-द][\)])')
    ans_pattern = re.compile(r'^(उत्तर|Answer|Ans|सही उत्तर)[\s:\-]', re.IGNORECASE)
    exp_pattern = re.compile(r'^(व्याख्या|Explanation|Exp|स्पष्टीकरण)[\s:\-]', re.IGNORECASE)

    for line in lines:
        if q_pattern.match(line) or (current_q and line.startswith('प्रश्न')):
            if current_q and current_q['question']:
                while len(current_q['options']) < 4:
                    current_q['options'].append(f"({len(current_q['options']) + 1}) विकल्प उपलब्ध नहीं")
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
            current_q['options'].append(f"({len(current_q['options']) + 1}) विकल्प उपलब्ध नहीं")
        questions.append(current_q)
        
    return questions

# --- Cloud Safe PPT to PDF Conversion ---
def convert_pptx_to_pdf_cloud(pptx_path, output_dir):
    try:
        cmd = ['soffice', '--headless', '--convert-to', 'pdf', pptx_path, '--outdir', output_dir]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except Exception as e:
        st.error(f"⚠️ PDF कन्वर्जन में समस्या: {e}")
        return False

# --- इनपुट सेक्शन ---
raw_text = ""
col_input, col_preview = st.columns([1, 1])

with col_input:
    st.subheader("📥 इनपुट डेटा")
    if input_choice == "📁 File Upload Karein (.txt, .pdf, .docx, Image)":
        uploaded_file = st.file_uploader(
            "File Upload Karein (.txt, .pdf, .docx, .png, .jpg, .jpeg)", 
            type=["txt", "pdf", "docx", "png", "jpg", "jpeg"]
        )
        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            file_extension = uploaded_file.name.split('.')[-1].lower()
            try:
                if file_extension == "txt":
                    raw_text = file_bytes.decode("utf-8", errors="ignore")
                elif file_extension == "pdf":
                    if fitz:
                        doc = fitz.open(stream=file_bytes, filetype="pdf")
                        for page in doc:
                            raw_text += page.get_text() + "\n"
                elif file_extension == "docx":
                    if Document:
                        doc = Document(io.BytesIO(file_bytes))
                        for para in doc.paragraphs:
                            raw_text += para.text + "\n"
                elif file_extension in ["png", "jpg", "jpeg"]:
                    if pytesseract:
                        image = Image.open(io.BytesIO(file_bytes))
                        raw_text = pytesseract.image_to_string(image, lang='hin+eng')
            except Exception as e:
                st.error(f"File reading error: {e}")
    else:
        raw_text = st.text_area(
            "यहाँ प्रश्न, विकल्प, उत्तर और व्याख्या पेस्ट करें:",
            height=300,
            placeholder="प्रश्न 1: भारत की राजधानी क्या है?\n(A) मुंबई\n(B) दिल्ली\n(C) कोलकाता\n(D) चेन्नई\nउत्तर: (B) दिल्ली\nव्याख्या: दिल्ली भारत की राष्ट्रीय राजधानी है।"
        )

    if st.button("🔍 डेटा पार्स करें"):
        if raw_text.strip():
            st.session_state.parsed_questions = double_verify_and_parse(raw_text)
            st.success(f"कुल {len(st.session_state.parsed_questions)} प्रश्न पार्स हुए!")
        else:
            st.warning("⚠️ कृपया टेक्स्ट पेस्ट करें या फाइल अपलोड करें!")

with col_preview:
    st.subheader(f"📋 पूर्वावलोकन ({len(st.session_state.parsed_questions)})")
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

st.divider()

# --- Generator ---
if st.button("🚀 Master PPT & PDF जनरेट करें", type="primary", use_container_width=True):
    if not st.session_state.parsed_questions:
        st.warning("⚠️ कोई प्रश्न पार्स नहीं हुआ है!")
    else:
        with st.spinner("आपकी पुरानी सटीक स्टाइल में PPT और PDF जनरेट की जा रही है..."):
            parsed_questions = st.session_state.parsed_questions
            
            prs = Presentation()
            
            # --- आपके पुरानी फाइल वाली EXACT डायमेंशन और साइज़ ---
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
                exp_font_size = Pt(32)
                card_width = Inches(11.733)
                card_height = Inches(6.4)
                box_width = Inches(11.133)
                q_box_width = Inches(12.3)
                opt_box_width = Inches(12.0)
                exp_box_height = Inches(4.5)
                opt_left = Inches(0.8)
                opt_top = Inches(3.0)
                opt_space_before = Pt(8)

            else:  # 4:3 (Standard)
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
                # SLIDE 1: Question + Options (EXACT OLD STYLE)
                # ==========================================
                slide1 = prs.slides.add_slide(blank_layout)

                q_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.4), q_box_width, Inches(2.5))
                tf1 = q_box.text_frame
                tf1.word_wrap = True
                p1 = tf1.paragraphs[0]
                p1.text = q['question']
                set_hindi_font(p1, 'Nirmala UI')  # Devanagari Font Fix
                p1.font.size = q_font_size
                p1.font.bold = True
                p1.font.color.rgb = RGBColor(255, 0, 0)
                p1.line_spacing = 1.3

                opt_box = slide1.shapes.add_textbox(opt_left, opt_top, opt_box_width, Inches(4.3))
                tf_opt = opt_box.text_frame
                tf_opt.word_wrap = True

                options = q['options'] if q['options'] else ["(A) विकल्प 1", "(B) विकल्प 2", "(C) विकल्प 3", "(D) विकल्प 4"]
                for opt_idx, opt in enumerate(options):
                    if opt_idx == 0:
                        p = tf_opt.paragraphs[0]
                    else:
                        p = tf_opt.add_paragraph()
                        p.space_before = opt_space_before
                    
                    p.text = opt
                    set_hindi_font(p, 'Nirmala UI')  # Devanagari Font Fix
                    p.font.size = opt_font_size
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    p.line_spacing = 1.4

                # ==========================================
                # SLIDE 2: Answer + Explanation (EXACT OLD STYLE)
                # ==========================================
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
                p_ans.text = q['answer'] if q['answer'] else "उत्तर: (सही विकल्प का नाम)"
                set_hindi_font(p_ans, 'Nirmala UI')  # Devanagari Font Fix
                p_ans.font.size = ans_font_size
                p_ans.font.bold = True
                p_ans.font.color.rgb = RGBColor(255, 255, 255)

                exp_box = slide2.shapes.add_textbox(Inches(1.1), Inches(2.1), box_width, exp_box_height)
                tf_exp = exp_box.text_frame
                tf_exp.word_wrap = True

                p_exp_title = tf_exp.paragraphs[0]
                p_exp_title.text = "व्याख्या:"
                set_hindi_font(p_exp_title, 'Nirmala UI')  # Devanagari Font Fix
                p_exp_title.font.size = Pt(26)
                p_exp_title.font.bold = True
                p_exp_title.font.color.rgb = RGBColor(30, 41, 59)

                expl_lines = q['explanation'][:3] if q['explanation'] else ["महत्वपूर्ण व्याख्या बिंदु यहाँ आएंगे।"]

                for line_idx, exp_line in enumerate(expl_lines):
                    p_exp = tf_exp.add_paragraph()
                    p_exp.space_before = Pt(6)
                    p_exp.line_spacing = 1.25
                    
                    p_exp.text = f"• {exp_line}" if not exp_line.startswith('•') else exp_line
                    set_hindi_font(p_exp, 'Nirmala UI')  # Devanagari Font Fix
                    p_exp.font.size = exp_font_size
                    p_exp.font.color.rgb = RGBColor(0, 0, 0)

            # Save PPT Stream
            ppt_stream = io.BytesIO()
            prs.save(ppt_stream)
            ppt_stream.seek(0)

            # Cloud Convert PPT to PDF
            pdf_bytes = b""
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_pptx_path = os.path.join(temp_dir, "temp.pptx")
                temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
                
                prs.save(temp_pptx_path)
                
                if convert_pptx_to_pdf_cloud(temp_pptx_path, temp_dir):
                    if os.path.exists(temp_pdf_path):
                        with open(temp_pdf_path, "rb") as f:
                            pdf_bytes = f.read()

        st.success("🎉 आपकी पुरानी स्टाइल वाली PPTX और 100% सेम डिज़ाइन वाली PDF तैयार हैं!")

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
            else:
                st.warning("⚠️ PDF कन्वर्जन फेल हुआ। कृपया सुनिश्चित करें कि 'packages.txt' में libreoffice मौजूद है।")
