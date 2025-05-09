import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap

# 문장이 차지할 줄 수 계산 (띄어쓰기 기준 줄바꿈)
def sentence_line_count(sentence, max_chars_per_line):
    words = sentence.split()
    lines = 1
    current_line_length = 0
    for word in words:
        if current_line_length > 0:
            if current_line_length + len(word) + 1 <= max_chars_per_line:
                current_line_length += len(word) + 1
            else:
                lines += 1
                current_line_length = len(word) + 1
        else:
            if len(word) <= max_chars_per_line:
                current_line_length = len(word)
            else:
                lines += 1
                current_line_length = len(word)
    return lines

# 전체 입력을 문장 단위로 분해
def split_text(text):
    sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

# 슬라이드 텍스트 생성 (제목 처리, 최대 줄 수 제한)
def create_slide_texts(sentences, max_chars_per_line, max_lines_per_slide):
    slide_texts = []
    current_slide_text = []
    current_slide_lines = 0
    title = ""

    for sentence in sentences:
        lines_needed = sentence_line_count(sentence, max_chars_per_line)
        is_title = not re.search(r'[.!?]$', sentence.strip())

        if is_title:
            if current_slide_text:
                slide_texts.append("\n".join(current_slide_text))
            title = sentence.strip()
            current_slide_text = [title]
            current_slide_lines = lines_needed
            continue

        if current_slide_lines + lines_needed <= max_lines_per_slide - (1 if title else 0):
            current_slide_text.append(sentence)
            current_slide_lines += lines_needed
        else:
            slide_texts.append("\n".join(current_slide_text))
            current_slide_text = [title, sentence] if title else [sentence]
            current_slide_lines = lines_needed + (1 if title else 0)
            title = ""

    if current_slide_text:
        slide_texts.append("\n".join(current_slide_text))

    return slide_texts

# 실제 PPT 슬라이드 생성
def create_slide(prs, text, current_idx, total_slides, max_chars_per_line):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    tf = textbox.text_frame
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = textwrap.fill(text, width=max_chars_per_line, break_long_words=False)  # 수정: textwrap.fill 사용

    p.font.size = Pt(54)
    p.font.name = '맑은 고딕'
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    tf.paragraphs[0].vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_frame = footer_box.text_frame
    footer_frame.text = f"{current_idx} / {total_slides}"
    footer_p = footer_frame.paragraphs[0]
    footer_p.font.size = Pt(18)
    footer_p.font.name = '맑은 고딕'
    footer_p.font.color.rgb = RGBColor(128, 128, 128)
    footer_p.alignment = PP_ALIGN.RIGHT

    if current_idx == total_slides:
        add_end_mark(slide)

def add_end_mark(slide):
    end_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(10), Inches(6), Inches(2), Inches(1)
    )
    end_shape.fill.solid()
    end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)
    end_shape.line.color.rgb = RGBColor(0, 0, 0)

    end_text_frame = end_shape.text_frame
    end_text_frame.clear()
    end_paragraph = end_text_frame.paragraphs[0]
    end_paragraph.text = "끝"
    end_paragraph.font.size = Pt(36)
    end_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    end_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# PPT 생성 함수
def create_ppt(slide_texts, max_chars_per_line):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    total_slides = len(slide_texts)
    for idx, text in enumerate(slide_texts, 1):
        create_slide(prs, text, idx, total_slides, max_chars_per_line)

    return prs

# Streamlit UI
st.set_page_config(page_title="Paydo Kitty", layout="centered")
st.title("🎤 Paydo Kitty - 촬영용 대본 PPT 생성기")

text_input = st.text_area("촬영용 대본을 입력하세요:", height=300, key="text_input_area")

max_lines_per_slide_input = st.slider("슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=4, key="max_lines_slider")
max_chars_per_line_input = st.slider("한 줄당 최대 글자 수:", min_value=10, max_value=100, value=35, key="max_chars_slider")

if st.button("PPT 만들기", key="create_ppt_button") and text_input.strip():
    sentences = split_text(text_input)
    slide_texts = create_slide_texts(sentences, max_chars_per_line_input, max_lines_per_slide_input)  # 함수 이름 변경
    try:
        ppt = create_ppt(slide_texts, max_chars_per_line_input)

        if ppt:
            ppt_io = io.BytesIO()
            ppt.save(ppt_io)
            ppt_io.seek(0)

            st.download_button(
                label="📥 PPT 다운로드",
                data=ppt_io,
                file_name="paydo_kitty_script.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="download_button"
            )
        else:
            st.error("PPT 생성에 실패했습니다. PPT 객체가 None입니다.")

    except Exception as e:
        st.error(f"PPT 생성 중 오류가 발생했습니다: {e}")