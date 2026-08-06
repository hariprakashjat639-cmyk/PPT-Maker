import streamlit as st
import re
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import fitz   # PyMuPDF
from PIL import Image
import google.generativeai as genai

# पेज सेटिंग्स और सभी लेआउट सेटिंग्स
st.set_page_config(page_title="Gemini Powered Model Paper PPT Maker", page_icon="📊", layout="centered")
st.markdown("""
    <style>
    [data-testid="stHeader"] {display: none;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge {display: none !important;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Gemini AI Powered Bilingual PDF/Photo to PPT Maker")
st.write("Gemini AI आपके पेपर को पढ़कर OCR की सभी गलतियाँ सुधारेगा, विकल्पों का क्रम सही करेगा और केवल डेटा होने पर ही 2nd स्लाइड बनाएगा!")

# API Key को secrets या environment variables से ऑटोमैटिक लोड करना
gemini_api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

# PPT स्लाइड साइज चुनने का ऑप्शन
slide_format = st.selectbox(
    "PPT Slide Size Chunein",
    ["16:9 (Widescreen)", "20:9 (Cinematic)", "4:3 (Standard)"]
)

# फाइल अपलोडर
uploaded_file = st.file_uploader("PDF, Photo या Text फाइल अपलोड करें", type=["txt", "pdf", "png", "jpg", "jpeg"])

# --- Gemini AI के जरिए स्मार्ट टेक्स्ट क्लीनिंग और स्ट्रक्चरिंग ---
def process_with_gemini(file_bytes, file_type, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = """
    आप एक विशेषज्ञ परीक्षा प्रश्न पत्र विश्लेषक हैं। इस फाइल/इमेज से सभी बहुविकल्पीय प्रश्न (MCQs), उनके विकल्प (Options), उत्तर (Answer) और व्याख्या (Explanation) को सही क्रम में निकालें।
    कठोर नियम:
    1. OCR की सभी गंदी गलतियों (जैसे कटे-फटे शब्द, 'FAT' जैसी अजीब अंग्रेजी गलतियाँ) को देवनागरी/हिंदी में बिल्कुल शुद्ध रूप में सुधारें।
    2. विकल्पों (Options) का क्रम हमेशा सही (1, 2, 3, 4 या A, B, C, D) क्रम से सेट करें, भले ही इमेज में वे आगे-पीछे हों।
    3. यदि किसी प्रश्न का उत्तर या व्याख्या पीडीएफ में उपलब्ध नहीं है, तो उसे खाली छोड़ दें (फालतू चीजें न जोड़ें)।
    4. आउटपुट केवल इसी फॉर्मेट में होना चाहिए ताकि इसे आसानी से पढ़ा जा सके:
    
    Q: [प्रश्न यहाँ लिखें]
    O1: [विकल्प 1]
    O2: [विकल्प 2]
    O3: [विकल्प 3]
    O4: [विकल्प 4]
    Ans: [उत्तर यदि है तो लिखें, अन्यथा खाली छोड़ें]
    Exp: [व्याख्या यदि है तो लिखें, अन्यथा खाली छोड़ें]
    ---
    """
    
    try:
        contents = []
        if file_type == "pdf":
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page in doc:
                pix = page.get_pixmap(dpi=200)
                img_bytes = pix.tobytes("png")
                contents.append(Image.open(io.BytesIO(img_bytes)))
        else:
            contents.append(Image.open(io.BytesIO(file_bytes)))
            
        contents.append(prompt)
        response = model.generate_content(contents)
        return response.text
    except Exception as e:
        st.error(f"Gemini API Error: {e}")
        return ""

# पार्सर जो Gemini के क्लीन आउटपुट को पढ़ेगा
def parse_gemini_output(text):
    questions = []
    blocks = text.split("---")
    
    for block in blocks:
        if not block.strip():
            continue
            
        q_data = {'question': '', 'options': [], 'answer': '', 'explanation': []}
        lines = [l.strip() for l in block.split('\n') if l.strip()]
        
        current_field = None
        for line in lines:
            if line.startswith("Q:"):
                q_data['question'] = line.replace("Q:", "").strip()
                current_field = 'q'
            elif line.startswith("O1:") or line.startswith("O2:") or line.startswith("O3:") or line.startswith("O4:") or line.startswith("("):
                opt_clean = re.sub(r'^O\d+:\s*', '', line)
                q_data['options'].append(opt_clean)
                current_field = 'opt'
            elif line.startswith("Ans:"):
                ans_text = line.replace("Ans:", "").strip()
                if ans_text:
                    q_data['answer'] = f"उत्तर: {ans_text}" if not ans_text.startswith("उत्तर") else ans_text
                current_field = 'ans'
            elif line.startswith("Exp:"):
                exp_text = line.replace("Exp:", "").strip()
                if exp_text:
                    q_data['explanation'].append(exp_text)
                current_field = 'exp'
            else:
                if current_field == 'q':
                    q_data['question'] += " " + line
                elif current_field == 'opt' and q_data['options']:
                    q_data['options'][-1] += " " + line
                elif current_field == 'ans':
                    q_data['answer'] += " " + line
                elif current_field == 'exp':
                    if q_data['explanation']:
                        q_data['explanation'][-1] += " " + line
                    else:
                        q_data['explanation'].append(line)
                        
        if q_data['question']:
            if not q_data['options']:
                q_data['options'] = ["(1) विकल्प 1", "(2) विकल्प 2", "(3) विकल्प 3", "(4) विकल्प 4"]
            questions.append(q_data)
            
    return questions

if uploaded_file is not None:
    if not gemini_api_key:
        st.error("⚠️ GEMINI_API_KEY सेट नहीं है। कृपया Streamlit secrets या environment variables में इसे कॉन्फ़िगर करें।")
    else:
        if st.button("🚀 Gemini AI से PPT Generate Karein"):
            with st.spinner("Gemini AI फाइल को समझ रहा है, शुद्ध कर रहा है और PPT बना रहा है..."):
                file_bytes = uploaded_file.getvalue()
                file_name = uploaded_file.name.lower()
                file_type = "pdf" if file_name.endswith(".pdf") else "image"
                
                gemini_raw_text = process_with_gemini(file_bytes, file_type, gemini_api_key)
                
                if gemini_raw_text:
                    parsed_questions = parse_gemini_output(gemini_raw_text)
                    
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

                    st.success("🎉 आपकी PPT बिलकुल सही शुद्धता और स्मार्ट स्लाइड रूल्स के साथ तैयार हो गई है!")
                    st.download_button(
                        label="📥 PPT Download Karein",
                        data=ppt_stream,
                        file_name="Gemini_Model_Paper.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
