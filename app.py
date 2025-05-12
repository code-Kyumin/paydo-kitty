import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap
import docx  # python-docx 라이브러리 추가

# Word 파일에서 텍스트 추출하는 함수
def extract_text_from_word(file_path):
    """Word 파일에서 모든 텍스트를 추출하여 하나의 문자열로 반환합니다."""
    doc = docx.Document(file_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return "\n".join(full_text)

# 문장이 차지할 줄 수 계산 (단어 잘림 방지)
def sentence_line_count(sentence, max_chars_per_line):
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

# 텍스트를 슬라이드로 분할 및 그룹화
def split_and_group_text(text, max_lines_per_slide, max_chars_per_line_ppt):
    slides = []
    current_slide_text = ""
    current_slide_lines = 0
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    split_flags = []

    for sentence in sentences:
        sentence = sentence.strip()
        sentence_lines = sentence_line_count(sentence, max_chars_per_line_ppt)

        if current_slide_lines + sentence_lines <= max_lines_per_slide:
            if current_slide_text:
                current_slide_text += " "
            current_slide_text += sentence
            current_slide_lines += sentence_lines
            split_flags.append(False)
        else:
            if not current_slide_text and sentence_lines > max_lines_per_slide:
                words = sentence.split()
                half_len = len(words) // 2
                first_half = " ".join(words[:half_len])
                second_half = " ".join(words[half_len:])
                slides.append(first_half)
                split_flags.append(True)
                current_slide_text = second_half
                current_slide_lines = sentence_line_count(second_half, max_chars_per_line_ppt)
                split_flags.append(True)
            else:
                slides.append(current_slide_text)
                current_slide_text = sentence
                current_slide_lines = sentence_lines
                split_flags.append(False)

    if current_slide_text:
        slides.append(current_slide_text)
        split_flags.append(False)

    return slides, split_flags

# PPT 생성 함수
def create_ppt(slide_texts, split_flags, max_chars_per_line_in_ppt=18, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)

    for i, text in enumerate(slide_texts):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER)
        add_slide_number(slide, i + 1, total_slides)
        if split_flags[i]:
            add_check_needed_shape(slide)
        if i == total_slides - 1:
            add_end_mark(slide)

    return prs

def add_text_to_slide(slide, text, font_size, alignment):
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    text_frame = textbox.text_frame
    text_frame.clear()
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    text_frame.word_wrap = True

    p = text_frame.add_paragraph()
    p.text = text
    p.font.size = Pt(font_size)
    p.font.name = 'Noto Color Emoji'
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = alignment

    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    wrapped_lines = textwrap.wrap(text, width=18, break_long_words=False)
    text_frame.clear()
    for line in wrapped_lines:
        p = text_frame.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.name = 'Noto Color Emoji'
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = alignment

    # 명시적으로 텍스트 프레임의 세로 정렬을 다시 설정
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP


def add_slide_number(slide, current, total):
    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_text_frame = footer_box.text_frame
    footer_text_frame.clear()
    p = footer_text_frame.paragraphs[0]
    p.text = f"{current} / {total}"
    p.font.size = Pt(18)
    p.font.name = '맑은 고딕'
    p.font.color.rgb = RGBColor(128, 128, 128)
    p.alignment = PP_ALIGN.RIGHT

def add_end_mark(slide):
    end_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(10),
        Inches(6),
        Inches(2),
        Inches(1)
    )
    end_shape.fill.solid()
    end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)
    end_shape.line.color.rgb = RGBColor(0, 0, 0)

    end_text_frame = end_shape.text_frame
    end_text_frame.clear()
    p = end_text_frame.paragraphs[0]
    p.text = "끝"
    p.font.size = Pt(36)
    p.font.color.rgb = RGBColor(255, 255, 255)
    end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

def add_check_needed_shape(slide):
    check_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.5),
        Inches(0.3),
        Inches(2),
        Inches(0.5)
    )
    check_shape.fill.solid()
    check_shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
    check_shape.line.color.rgb = RGBColor(0, 0, 0)

    check_text_frame = check_shape.text_frame
    check_text_frame.clear()
    p = check_text_frame.paragraphs[0]
    p.text = "확인 필요!"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    text_frame = check_shape.text_frame
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기")

# Word 파일 업로드 기능 추가
uploaded_file = st.file_uploader("📝 Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
max_chars_per_line_ppt_input = st.slider("📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=30, value=18, key="max_chars_slider_ppt")
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")

if st.button("🚀 PPT 만들기", key="create_ppt_button"):
    text = ""
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    slide_texts, split_flags = split_and_group_text(
        text,
        max_lines_per_slide=max_lines_per_slide_input,
        max_chars_per_line_ppt=max_chars_per_line_ppt_input
    )
    ppt = create_ppt(
        slide_texts,
        split_flags,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
        font_size=font_size_input
    )

    if ppt:
        ppt_io = io.BytesIO()
        ppt.save(ppt_io)
        ppt_io.seek(0)

        st.download_button(
            label="📥 PPT 다운로드",
            data=ppt_io,
            file_name="paydo_script.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="download_button"
        )
        if any(split_flags):
            split_slide_numbers = [i + 1 for i, flag in enumerate(split_flags) if flag]
            st.warning(f"❗️ 일부 슬라이드({split_slide_numbers})는 한 문장이 너무 길어 분할되었습니다. PPT를 확인하여 가독성을 검토해주세요.")
    else:
        st.error("❌ PPT 생성에 실패했습니다.")