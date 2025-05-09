import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor  # RGBColor import 추가
import io
import re
import textwrap

def create_ppt(slide_texts, max_chars_per_line=18, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)  # 전체 슬라이드 수 계산

    for i, text in enumerate(slide_texts):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 텍스트 상단 정렬
        tf.word_wrap = True
        tf.clear()

        lines = textwrap.wrap(text, width=max_chars_per_line, break_long_words=False)  # 띄어쓰기 단위 줄바꿈
        p = tf.paragraphs[0]
        p.text = "\n".join(lines)
        p.font.size = Pt(font_size)
        p.font.name = 'Noto Color Emoji'
        p.alignment = PP_ALIGN.CENTER

        # 페이지 번호 추가
        add_page_number(slide, i + 1, total_slides)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

def add_page_number(slide, current_page, total_pages):
    """슬라이드에 페이지 번호 추가"""
    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_frame = footer_box.text_frame
    footer_frame.text = f"{current_page} / {total_pages}"
    footer_p = footer_frame.paragraphs[0]
    footer_p.font.size = Pt(18)
    footer_p.font.name = '맑은 고딕'
    footer_p.font.color.rgb = RGBColor(128, 128, 128)
    footer_p.alignment = PP_ALIGN.RIGHT

def split_and_group_text(text, separate_pattern=None):
    """텍스트를 분리하고 슬라이드 단위로 그룹화"""

    slides = []
    current_slide_text = ""
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())  # 문장 단위로 분리

    for sentence in sentences:
        if separate_pattern and re.search(separate_pattern, sentence):
            if current_slide_text:
                slides.append(current_slide_text.strip())
            slides.append(sentence.strip())  # 패턴 일치 문장은 새 슬라이드
            current_slide_text = ""
        else:
            current_slide_text += sentence + " "

    if current_slide_text:
        slides.append(current_slide_text.strip())

    return slides

st.title("🎬 Paydo 촬영 대본 PPT 생성기")
text_input = st.text_area("📝 촬영 대본을 입력하세요:", height=300)
separate_pattern_input = st.text_input("🔍 분리할 텍스트 패턴 (정규 표현식):")
max_chars_per_line_input = st.slider("📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=50, value=18)
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=120, value=54)

if st.button("🚀 PPT 만들기") and text_input.strip():
    slide_texts = split_and_group_text(text_input, separate_pattern_input)
    ppt_file = create_ppt(slide_texts, max_chars_per_line_input, font_size_input)
    st.download_button(
        label="📥 PPT 다운로드",
        data=ppt_file,
        file_name="paydo_script.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )