import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
import io
import re

# Functions from previous code (with minor adjustments for clarity)

def split_text(text):
    """텍스트를 문장 단위로 분리합니다."""
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences]

def group_sentences_to_slides(sentences, max_lines_per_slide=5, max_chars_per_line=35):
    """문장들을 슬라이드에 맞게 그룹화합니다."""

    slides = []
    current_slide_text = ""
    current_line_count = 0

    for sentence in sentences:
        # Calculate how many lines this sentence would take
        line_count = len(textwrap.wrap(sentence, width=max_chars_per_line))
        
        # Check if adding this sentence exceeds the limit
        if current_line_count + line_count > max_lines_per_slide and current_slide_text:
            slides.append(current_slide_text.strip())
            current_slide_text = sentence
            current_line_count = line_count
        else:
            current_slide_text += " " + sentence
            current_line_count += line_count
    
    if current_slide_text:
        slides.append(current_slide_text.strip())
    
    return slides

def create_ppt(slide_texts, max_chars_per_line=35):
    """슬라이드 텍스트를 사용하여 PPT를 생성합니다."""

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for text in slide_texts:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.33), Inches(6.5))
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        
        p = text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(40)
        p.font.name = '맑은 고딕'
        p.alignment = PP_ALIGN.CENTER
        
    return prs

# Streamlit UI
st.title("PPT 생성기")
text_input = st.text_area("텍스트 입력:", height=200)
max_lines_per_slide = st.slider("최대 줄 수 (슬라이드 당)", 3, 10, 5)
max_chars_per_line = st.slider("최대 글자 수 (줄 당)", 20, 100, 40)

if st.button("PPT 생성"):
    if text_input:
        sentences = split_text(text_input)
        slide_texts = group_sentences_to_slides(sentences, max_lines_per_slide, max_chars_per_line)
        prs = create_ppt(slide_texts, max_chars_per_line)
        
        ppt_bytes = io.BytesIO()
        prs.save(ppt_bytes)
        ppt_bytes.seek(0)
        
        st.download_button(label="PPT 다운로드", data=ppt_bytes, file_name="output.pptx")
    else:
        st.warning("텍스트를 입력해주세요.")