import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
import io
import re

def split_text_to_slides(text, max_lines=4):
    paragraphs = text.strip().split("\n")
    slides = []
    current_slide = []
    for para in paragraphs:
        if not para.strip():
            continue
        sentences = re.split(r'(?<=[.!?]) +', para.strip())
        for sentence in sentences:
            if sentence:
                current_slide.append(sentence.strip())
                if len(current_slide) >= max_lines:
                    slides.append(current_slide)
                    current_slide = []
    if current_slide:
        slides.append(current_slide)
    return slides

def create_ppt(slides):
    prs = Presentation()
    prs.slide_width = Inches(13.33)  # 16:9
    prs.slide_height = Inches(7.5)

    for idx, lines in enumerate(slides, 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 본문 텍스트 박스
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        tf.word_wrap = True
        tf.auto_size = False  # Changed from None to False
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = line
            p.font.size = Pt(54)
            p.font.name = '맑은 고딕'
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.CENTER

        # 우측 하단 페이지 번호
        footer_box = slide.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(1), Inches(0.4))
        footer_frame = footer_box.text_frame
        footer_frame.text = str(idx)
        footer_p = footer_frame.paragraphs[0]
        footer_p.font.size = Pt(18)
        footer_p.font.name = '맑은 고딕'
        footer_p.font.color.rgb = RGBColor(128, 128, 128)
        footer_p.alignment = PP_ALIGN.RIGHT

    return prs

# Streamlit UI
st.set_page_config(page_title="Paydo Kitty", layout="centered")
st.title("📄 Paydo Kitty - 텍스트를 PPT로 변환")

text_input = st.text_area("대본을 입력하세요:", height=300)

if st.button("PPT 만들기") and text_input.strip():
    slides = split_text_to_slides(text_input)
    ppt = create_ppt(slides)

    ppt_io = io.BytesIO()
    ppt.save(ppt_io)
    ppt_io.seek(0)

    st.download_button(
        label="📥 PPT 다운로드",
        data=ppt_io,
        file_name="paydo_kitty_output.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )