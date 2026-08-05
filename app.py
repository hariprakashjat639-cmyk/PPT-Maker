import streamlit as st
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import pypdf

# पेज सेटिंग्स और सभी प्रोफाइल/GitHub आइकॉन/हेडर/फूटर छुपाने के लिए
st.set_page_config(page_title="Model Paper PPT Maker", page_icon="📊", layout="centered")
st.markdown("""
    <style>
    [data-testid="stHeader"] {display: none;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge {display: none !important;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Bilingual Text to PPT Converter App")
st.write("Apni Text (`.txt`) ya PDF file yahan upload karein aur bilkul tayar format wali PPT download karein.")

# PPT स्लाइड साइज चुनने का ऑप्शन (20:9 साइज शामिल)
slide_format = st.selectbox(
    "PPT Slide Size Chunein",
    ["16:9 (Widescreen)", "20:9 (Cinematic)", "4:3 (Standard)"]
)

# सुरक्षित फाइल अपलोडर (.txt और .pdf दोनों के लिए)
uploaded_file = st.file_uploader("Text ya PDF File Upload Karein", type=["txt", "pdf"])

content = ""
if uploaded_file is not None:
    if uploaded_file.name.lower().endswith(".pdf"):
        try:
            with st.spinner("बड़ी PDF फाइल पढ़ी जा रही है, कृपया थोड़ा इंतज़ार करें..."):
                pdf_reader = pypdf.PdfReader(uploaded_file)
                for page in pdf_reader.pages:
                    extracted_text = page.extract_text()
                    if extracted_text:
                        content += extracted_text + "\n"
        except Exception as e:
            st.error(f"PDF फाइल पढ़ने में समस्या आई: {e}")
    else:
        try:
            raw_bytes = uploaded_file.getvalue()
            try:
                content = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                content = raw_bytes.decode("latin-1")
        except Exception as e:
            st.error(f"Text फाइल पढ़ने में समस्या आई: {e}")

def parse_txt_content(text):
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
        questions.append(current_q)
        
    return questions

if content.strip():
    if st.button("🚀 PPT Generate Karein"):
        with st.spinner("PPT ban rahi hai, kripya intezar karein..."):
            parsed_questions = parse_txt_content(content)
            
            prs = Presentation()
            
            # यूजर के चयन के हिसाब से साइज सेट करना
            if slide_format == "20:9 (Cinematic)":
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(6.0)
            elif slide_format == "16:9 (Widescreen)":
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
            elif slide_format == "4:3 (Standard)":
                prs.slide_width = Inches(10)
                prs.slide_height = Inches(7.5)
                
            blank_layout = prs.slide_layouts[6] 

            for idx, q in enumerate(parsed_questions):
                # ==========================================
                # SLIDE 1: Question + Options
                # ==========================================
                slide1 = prs.slides.add_slide(blank_layout)

                q_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(2.5))
                tf1 = q_box.text_frame
                tf1.word_wrap = True
                p1 = tf1.paragraphs[0]
                p1.text = q['question']
                p1.font.name = 'Nirmala UI'
                p1.font.size = Pt(40)  # Size 40
                p1.font.bold = True
                p1.font.color.rgb = RGBColor(255, 0, 0)  # Pure Red (#FF0000)
                p1.line_spacing = 1.3  # Line Spacing 1.30

                opt_box = slide1.shapes.add_textbox(Inches(0.8), Inches(3.0), Inches(12.0), Inches(4.3))
                tf_opt = opt_box.text_frame
                tf_opt.word_wrap = True

                options = q['options'] if q['options'] else ["(A) विकल्प 1", "(B) विकल्प 2", "(C) विकल्प 3", "(D) विकल्प 4"]
                for opt_idx, opt in enumerate(options):
                    if opt_idx == 0:
                        p = tf_opt.paragraphs[0]
                    else:
                        p = tf_opt.add_paragraph()
                        p.space_before = Pt(8)
                    
                    p.text = opt
                    p.font.name = 'Nirmala UI'
                    p.font.size = Pt(40)  # Size 40
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    p.line_spacing = 1.4  # Line Spacing 1.40

                # ==========================================
                # SLIDE 2: Answer + Explanation
                # ==========================================
                slide2 = prs.slides.add_slide(blank_layout)

                card = slide2.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, 
                    Inches(0.8), Inches(0.5), Inches(11.733), Inches(6.4)
                )
                card.fill.solid()
                card.fill.fore_color.rgb = RGBColor(248, 250, 252)
                card.line.color.rgb = RGBColor(203, 213, 225)
                card.line.width = Pt(1.5)

                ans_banner = slide2.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Inches(1.1), Inches(0.8), Inches(11.133), Inches(1.1)
                )
                ans_banner.fill.solid()
                ans_banner.fill.fore_color.rgb = RGBColor(22, 163, 74) # Green Banner
                ans_banner.line.color.rgb = RGBColor(22, 163, 74)

                tf_ans = ans_banner.text_frame
                tf_ans.word_wrap = True
                p_ans = tf_ans.paragraphs[0]
                p_ans.text = q['answer'] if q['answer'] else "उत्तर: (सही विकल्प का नाम)"
                p_ans.font.name = 'Nirmala UI'
                p_ans.font.size = Pt(28)
                p_ans.font.bold = True
                p_ans.font.color.rgb = RGBColor(255, 255, 255)

                exp_box = slide2.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(11.133), Inches(4.5))
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
                    p_exp.line_spacing = 1.25  # Line Spacing 1.25
                    
                    p_exp.text = f"• {exp_line}" if not exp_line.startswith('•') else exp_line
                    p_exp.font.name = 'Nirmala UI'
                    p_exp.font.size = Pt(32)  # Size 32
                    p_exp.font.color.rgb = RGBColor(0, 0, 0)  # Full Deep Black

            ppt_stream = io.BytesIO()
            prs.save(ppt_stream)
            ppt_stream.seek(0)

            st.success("🎉 Aapki PPT safaltapoorvak tayar ho gayi hai!")
            st.download_button(
                label="📥 PPT Download Karein",
                data=ppt_stream,
                file_name="Model_Paper_Ready.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
