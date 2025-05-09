import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap

# 문장이 차지할 줄 수 계산 (단어 잘림 방지)
def sentence_line_count(sentence, max_chars_per_line=35):
    words = sentence.split()
    lines = 1
    current_line_length = 0
    for word in words:
        if current_line_length + len(word) + 1 <= max_chars_per_line:
            current_line_length += len(word) + 1
        else:
            lines += 1
            current_line_length = len(word)
            
    return lines

# 전체 입력을 문장 단위로 분해하고, 슬라이드 단위로 묶음
def split_and_group_text(text, max_lines_per_slide=5, min_chars_per_line=4, max_chars_per_line_in_ppt=18):
    slides = []
    current_slide_text = ""
    current_slide_lines = 0

    sentences = re.split(r'(?<=[.!?])\s+', text.strip())

    for sentence in sentences:
        lines_needed = sentence_line_count(sentence, max_chars_per_line_in_ppt)
        if current_slide_lines + lines_needed <= max_lines_per_slide:
            current_slide_text += sentence + " "
            current_slide_lines += lines_needed
        else:
            slides.append(current_slide_text.strip())
            current_slide_text = sentence + " "
            current_slide_lines = lines_needed

    if current_slide_text:
        slides.append(current_slide_text.strip())

    return slides

# PPT 생성 함수
def create_ppt(slide_texts, max_chars_per_line_in_ppt=18, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)
    
    for i, text in enumerate(slide_texts):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        tf.word_wrap = True
        tf.clear()

        lines = textwrap.wrap(text, width=max_chars_per_line_in_ppt, break_long_words=False)
        p = tf.paragraphs[0]
        p.text = "\n".join(lines)
        p.font.size = Pt(font_size)
        p.font.name = 'Noto Color Emoji'
        p.alignment = PP_ALIGN.CENTER

        # 페이지 번호 (현재 페이지/전체 페이지)
        footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
        footer_frame = footer_box.text_frame
        footer_frame.text = f"{i + 1} / {total_slides}"
        footer_p = footer_frame.paragraphs[0]
        footer_p.font.size = Pt(18)
        footer_p.font.name = '맑은 고딕'
        footer_p.font.color.rgb = RGBColor(128, 128, 128)
        footer_p.alignment = PP_ALIGN.RIGHT

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)

    return ppt_io

# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기")

text_input = st.text_area("📝 촬영 대본을 입력하세요:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
# PPT 텍스트 박스 내에서의 줄바꿈 글자 수 (실제 PPT에 표시될 때 적용)
max_chars_per_line_ppt_input = st.slider("📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=30, value=18, key="max_chars_slider_ppt")
min_chars_per_line_input = st.slider("🔤 한 줄당 최소 글자 수:", min_value=1, max_value=10, value=4, key="min_chars_slider")
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")


if st.button("🚀 PPT 만들기", key="create_ppt_button") and text_input.strip():
    # 수정된 함수 호출
    slide_texts = split_and_group_text(text_input,
                                        max_lines_per_slide=max_lines_per_slide_input,
                                        min_chars_per_line=min_chars_per_line_input,
                                        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input)
    ppt_file = create_ppt(slide_texts,
                        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
                        font_size=font_size_input)

    st.download_button(
        label="📥 PPT 다운로드",
        data=ppt_file,
        file_name="paydo_script.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="download_button"
    )