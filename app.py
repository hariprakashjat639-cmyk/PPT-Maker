import streamlit as st
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
from PIL import Image

# सुरक्षित रूप से वैकल्पिक लाइब्रेरीज को इम्पोर्ट करना ताकि ऐप क्रैश न हो
try:
    import fitz  # PyMuPDF for PDF & Images
except ImportError:
    fitz = None

try:
    from docx import Document  # python-docx for Word files
except ImportError:
    Document = None

try:
    import pytesseract  # Offline OCR for images
except ImportError:
    pytesseract = None

# पेज सेटिंग्स और लेआउट
st.set_page_config(page_title="Master Offline Model Paper PPT Maker", page_icon="📊", layout="centered")
st.markdown("""
    <style>
    [data-testid="stHeader"] {display: none;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge {display: none !important;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Master Offline Multi-Format & Direct Text to PPT Maker")
st.write("चाहे फाइल अपलोड करें या सीधे टेक्स्ट पेस्ट करें। यह ऑफलाइन AI और नए 3 प्रकार के विशेष प्रश्न लेआउट के साथ आपकी PPT तैयार करेगा!")

# PPT स्लाइड साइज चुनने का ऑप्शन
slide_format = st.selectbox(
    "PPT Slide Size Chunein",
    ["16:9 (Widescreen)", "20:9 (Cinematic)", "4:3 (Standard)"]
)

# इनपुट का तरीका चुनने के लिए विकल्प
input_choice = st.radio(
    "Data Input Ka Tarika Chunein:",
    ["📁 File Upload Karein (.txt, .pdf, .docx, Image)", "✍️ Direct Text Paste Karein"]
)

raw_text = ""

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
                if fitz is None:
                    st.error("⚠️ PyMuPDF (fitz) लाइब्रेरी इंस्टॉल नहीं है।")
                else:
                    doc = fitz.open(stream=file_bytes, filetype="pdf")
                    for page in doc:
                        raw_text += page.get_text() + "\n"
            elif file_extension == "docx":
                if Document is None:
                    st.error("⚠️ python-docx लाइब्रेरी इंस्टॉल नहीं है।")
                else:
                    doc = Document(io.BytesIO(file_bytes))
                    for para in doc.paragraphs:
                        raw_text += para.text + "\n"
            elif file_extension in ["png", "jpg", "jpeg"]:
                if pytesseract is None:
                    st.error("⚠️ pytesseract लाइब्रेरी इंस्टॉल नहीं है।")
                else:
                    image = Image.open(io.BytesIO(file_bytes))
                    raw_text = pytesseract.image_to_string(image, lang='hin+eng')
        except Exception as e:
            st.error(f"File reading error: {e}")
else:
    raw_text = st.text_area(
        "Yahan Apne Prashn, Options, Answer aur Explanation Paste Karein:",
        height=250,
        placeholder="प्रश्न 1: भारत की राजधानी क्या है?\n(A) मुंबई\n(B) दिल्ली\n(C) कोलकाता\n(D) चेन्नई\nउत्तर: (B) दिल्ली\nव्याख्या: दिल्ली भारत की राष्ट्रीय राजधानी है।"
    )

# --- स्मार्ट टेक्स्ट पार्सर जो 3 नए प्रश्न प्रकारों को भी पहचानता है ---
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
                while len(current_q['options']) < 4 and current_q['q_type'] == 'normal':
                    current_q['options'].append(f"({len(current_q['options']) + 1}) विकल्प उपलब्ध नहीं")
                questions.append(current_q)
            
            # प्रश्न के प्रकार की पहचान (3 नए प्रकार + सामान्य प्रकार)
            q_type = 'normal'
            lower_line = line.lower()
            if 'सूची' in line or 'मिलान' in line or 'match' in lower_line:
                q_type = 'match_type'
            elif 'कथन' in line or 'statement' in lower_line:
                q_type = 'statement_type'
            elif 'अभिकथन' in line or 'कारण' in line or 'assertion' in lower_line or 'reason' in lower_line:
                q_type = 'assertion_type'

            current_q = {'question': line, 'options': [], 'answer': '', 'explanation': [], 'q_type': q_type}
            continue
            
        if current_q is None:
            current_q = {'question': line, 'options': [], 'answer': '', 'explanation': [], 'q_type': 'normal'}
            continue

        # यदि लाइन के अंदर विशेष कीवर्ड मिलें
        if 'सूची' in line or 'मिलान करें' in line:
            current_q['q_type'] = 'match_type'
        elif 'कथन I' in line or 'कथन-1' in line or 'कथन ' in line:
            current_q['q_type'] = 'statement_type'
        elif 'अभिकथन' in line or 'कारण' in line:
            current_q['q_type'] = 'assertion_type'

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
        while len(current_q['options']) < 4 and current_q['q_type'] == 'normal':
            current_q['options'].append(f"({len(current_q['options']) + 1}) विकल्प उपलब्ध नहीं")
        questions.append(current_q)
        
    return questions

if st.button("🚀 Master PPT Generate Karein"):
    if not raw_text.strip():
        st.warning("⚠️ कृपया पहले कोई फाइल अपलोड करें या टेक्स्ट बॉक्स में प्रश्न/उत्तर पेस्ट करें!")
    else:
        with st.spinner("टेक्स्ट को पार्स किया जा रहा है और विशेष लेआउट तैयार किए जा रहे हैं..."):
            parsed_questions = double_verify_and_parse(raw_text)
            
            prs = Presentation()
            
            # साइज़ और डिज़ाइन सेटिंग्स
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
                q_font_size = Pt(36)
                opt_font_size = Pt(34)
                ans_font_size = Pt(28)
                exp_font_size = Pt(30)
                card_width = Inches(11.733)
                card_height = Inches(6.4)
                box_width = Inches(11.133)
                q_box_width = Inches(12.3)
                opt_box_width = Inches(12.0)
                exp_box_height = Inches(4.5)
                opt_left = Inches(0.8)
                opt_top = Inches(2.8)
                opt_space_before = Pt(6)

            elif slide_format == "4:3 (Standard)":
                prs.slide_width = Inches(10)
                prs.slide_height = Inches(7.5)
                q_font_size = Pt(26)
                opt_font_size = Pt(24)
                ans_font_size = Pt(24)
                exp_font_size = Pt(22)
                card_width = Inches(8.8)
                card_height = Inches(6.4)
                box_width = Inches(8.2)
                q_box_width = Inches(9.0)
                opt_box_width = Inches(8.8)
                exp_box_height = Inches(4.5)
                opt_left = Inches(0.5)
                opt_top = Inches(2.5)
                opt_space_before = Pt(4)

            blank_layout = prs.slide_layouts[6] 

            for idx, q in enumerate(parsed_questions):
                # ==========================================
                # SLIDE 1: Question Layout (Based on Question Type)
                # ==========================================
                slide1 = prs.slides.add_slide(blank_layout)

                q_type = q.get('q_type', 'normal')

                if q_type == 'statement_type':
                    # प्रकार 1: कथन आधारित प्रश्न (Statement-based - Blue Accent Theme)
                    stmt_banner = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(0.3), q_box_width, Inches(1.1))
                    stmt_banner.fill.solid()
                    stmt_banner.fill.fore_color.rgb = RGBColor(30, 58, 138) # Dark Blue
                    stmt_banner.line.color.rgb = RGBColor(30, 58, 138)
                    
                    tf_sb = stmt_banner.text_frame
                    tf_sb.word_wrap = True
                    p_sb = tf_sb.paragraphs[0]
                    p_sb.text = "📌 विशेष कथन आधारित प्रश्न:"
                    p_sb.font.name = 'Nirmala UI'
                    p_sb.font.size = Pt(20)
                    p_sb.font.bold = True
                    p_sb.font.color.rgb = RGBColor(255, 255, 255)

                    q_box = slide1.shapes.add_textbox(Inches(0.6), Inches(1.5), q_box_width, Inches(2.0))
                    tf1 = q_box.text_frame
                    tf1.word_wrap = True
                    p1 = tf1.paragraphs[0]
                    p1.text = q['question']
                    p1.font.name = 'Nirmala UI'
                    p1.font.size = q_font_size - Pt(4)
                    p1.font.bold = True
                    p1.font.color.rgb = RGBColor(15, 23, 42)
                    p1.line_spacing = 1.25

                elif q_type == 'match_type':
                    # प्रकार 2: सूची / मिलान वाले प्रश्न (Match the following - Purple Accent Theme)
                    match_banner = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(0.3), q_box_width, Inches(1.1))
                    match_banner.fill.solid()
                    match_banner.fill.fore_color.rgb = RGBColor(88, 28, 135) # Deep Purple
                    match_banner.line.color.rgb = RGBColor(88, 28, 135)
                    
                    tf_mb = match_banner.text_frame
                    tf_mb.word_wrap = True
                    p_mb = tf_mb.paragraphs[0]
                    p_mb.text = "🔄 सूची मिलान प्रश्न (Match the Following):"
                    p_mb.font.name = 'Nirmala UI'
                    p_mb.font.size = Pt(20)
                    p_mb.font.bold = True
                    p_mb.font.color.rgb = RGBColor(255, 255, 255)

                    q_box = slide1.shapes.add_textbox(Inches(0.6), Inches(1.5), q_box_width, Inches(2.0))
                    tf1 = q_box.text_frame
                    tf1.word_wrap = True
                    p1 = tf1.paragraphs[0]
                    p1.text = q['question']
                    p1.font.name = 'Nirmala UI'
                    p1.font.size = q_font_size - Pt(6)
                    p1.font.bold = True
                    p1.font.color.rgb = RGBColor(15, 23, 42)
                    p1.line_spacing = 1.2

                elif q_type == 'assertion_type':
                    # प्रकार 3: अभिकथन और कारण वाले प्रश्न (Assertion-Reason - Amber/Orange Accent Theme)
                    asrt_banner = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(0.3), q_box_width, Inches(1.1))
                    asrt_banner.fill.solid()
                    asrt_banner.fill.fore_color.rgb = RGBColor(180, 83, 9) # Amber/Orange
                    asrt_banner.line.color.rgb = RGBColor(180, 83, 9)
                    
                    tf_ab = asrt_banner.text_frame
                    tf_ab.word_wrap = True
                    p_ab = tf_ab.paragraphs[0]
                    p_ab.text = "⚡ अभिकथन और कारण (Assertion & Reason):"
                    p_ab.font.name = 'Nirmala UI'
                    p_ab.font.size = Pt(20)
                    p_ab.font.bold = True
                    p_ab.font.color.rgb = RGBColor(255, 255, 255)

                    q_box = slide1.shapes.add_textbox(Inches(0.6), Inches(1.5), q_box_width, Inches(2.0))
                    tf1 = q_box.text_frame
                    tf1.word_wrap = True
                    p1 = tf1.paragraphs[0]
                    p1.text = q['question']
                    p1.font.name = 'Nirmala UI'
                    p1.font.size = q_font_size - Pt(4)
                    p1.font.bold = True
                    p1.font.color.rgb = RGBColor(15, 23, 42)
                    p1.line_spacing = 1.25

                else:
                    # सामान्य प्रश्न (Original Style: Pure Red Question)
                    q_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.4), q_box_width, Inches(2.5))
                    tf1 = q_box.text_frame
                    tf1.word_wrap = True
                    p1 = tf1.paragraphs[0]
                    p1.text = q['question']
                    p1.font.name = 'Nirmala UI'
                    p1.font.size = q_font_size
                    p1.font.bold = True
                    p1.font.color.rgb = RGBColor(255, 0, 0) # Pure Red (#FF0000)
                    p1.line_spacing = 1.3

                # विकल्प बॉक्स (Options) सभी के लिए
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
                    p.font.name = 'Nirmala UI'
                    p.font.size = opt_font_size
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    p.line_spacing = 1.35

                # ==========================================
                # SLIDE 2: Answer + Explanation (Original Style Preserved)
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
                p_exp_title.font.color.rgb = RGBColor(30, 41, 59)

                expl_lines = q['explanation'][:3] if q['explanation'] else ["महत्वपूर्ण व्याख्या बिंदु यहाँ आएंगे।"]

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

            st.success("🎉 आपकी मास्टर PPT सफलतापूर्क तैयार हो गई है जिसमें नए प्रश्न प्रकारों के लिए अलग स्टाइल्स जोड़ी गई हैं!")
            st.download_button(
                label="📥 PPT Download Karein",
                data=ppt_stream,
                file_name="Master_Model_Paper.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
