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
st.write("Apni Text (.txt) ya PDF file yahan upload karein aur bilkul tayar format wali PPT download karein.")

# PPT स्लाइड साइज चुनने का ऑप्शन (20:9 साइज शामिल)
slide_format = st.selectbox(
    "PPT Slide Size Chunein",
    ["20:9 (Cinematic)", "16:9 (Widescreen)", "4:3 (Standard)"]
)

# सुरक्षित फाइल अपलोडर (.txt और .pdf दोनों के लिए)
uploaded_file = st.file_uploader("Text ya PDF File Upload Karein", type=["txt", "pdf"])

text = ""
if uploaded_file is not None:
    if uploaded_file.name.lower().endswith(".pdf"):
        try:
            pdf_reader = pypdf.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                extracted_text = page.extract_text()
                if extracted_text:
                    text += extracted_text + "\n"
        except Exception as e:
            st.error(f"PDF फाइल पढ़ने में समस्या आई: {e}")
    else:
        try:
            raw_bytes = uploaded_file.getvalue()
            try:
                text = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text = raw_bytes.decode("latin-1")
        except Exception as e:
            st.error(f"Text फाइल पढ़ने में समस्या आई: {e}")

# अगर टेक्स्ट मौजूद है तो PPT जनरेट करने का बटन दिखाएं
if text:
    st.success("फाइल सफलतापूर्वक पढ़ ली गई है!")
    
    if st.button("PPT Generate Karein"):
        prs = Presentation()
        
        # स्लाइड का साइज सेट करना (20:9 के लिए खास डाइमेंशन)
        if slide_format == "20:9 (Cinematic)":
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(6.0) # 20:9 आस्पेक्ट रेश्यो
        elif slide_format == "16:9 (Widescreen)":
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
        elif slide_format == "4:3 (Standard)":
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)
            
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # टेक्स्ट बॉक्स जोड़ना
        left = Inches(0.8)
        top = Inches(0.8)
        width = prs.slide_width - Inches(1.6)
        height = prs.slide_height - Inches(1.6)
        
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        
        # टेक्स्ट को लाइनों में तोड़कर अलग पैराग्राफ बनाना ताकि टेक्स्ट आपस में न मिले
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        for i, line_text in enumerate(lines[:40]):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = line_text
            p.font.size = Pt(14)
            p.space_after = Pt(10) # टेक्स्ट को आपस में मिलने से रोकने के लिए पर्याप्त गैप
        
        # फाइल सेव करने के लिए बफर
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        st.download_button(
            label="📥 Download PowerPoint Presentation",
            data=ppt_buffer,
            file_name="Model_Paper.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
