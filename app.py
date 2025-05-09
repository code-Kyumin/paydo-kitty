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
def extract_text_from_word(file):
    doc = docx.Document(file)
    full_text = []
    for paragraph in doc.paragraphs:
        for p in doc.paragraphs:
            full_text.append(p.text)
    return "\n".join(full_text)

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

# 전체 입력을 문장 단위로 분해하고, 필요한 경우 슬라이드 분할 여부 표시
def split_and_group_text(text, max_lines_per_slide=5, min_chars_per_line=4, max_chars_per_line_in_ppt=18):
    slides = []
    original_sentence_flags = []  # 각 슬라이드가 원본 문장인지 여부 저장
    current_slide_text = ""
    current_slide_lines = 0
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    seen_sentences = set()  # [추가] 이미 처리한 문장을 저장할 set

    for sentence in sentences:
        sentence = sentence.strip()
        if sentence in seen_sentences:  # [추가] 중복 문장 확인
            continue
        seen_sentences.add(sentence)  # [추가] 처리한 문장 저장
        
        lines_needed = sentence_line_count(sentence, max_chars_per_line_in_ppt)

        if current_slide_lines + lines_needed <= max_lines_per_slide:
            if current_slide_text:
                current_slide_text += " "
            current_slide_text += sentence
            current_slide_lines += lines_needed
            original_sentence_flags.append(True)  # 원래 문장
        else:
            slides.append(current_slide_text)
            original_sentence_flags.append(False)  # 분할된 문장
            current_slide_text = sentence
            current_slide_lines = lines_needed

    if current_slide_text:
        slides.append(current_slide_text)
        original_sentence_flags.append(True)  # 마지막 슬라이드

    return slides, original_sentence_flags

# PPT 생성 함수
def create_ppt(slide_texts, original_sentence_flags, max_chars_per_line_in_ppt=18, max_lines_per_slide=5, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = 0
    current_slide_idx = 1
    slides_data = []
    check_needed_slides = []  # 확인이 필요한 슬라이드 번호 저장

    try:
        for i, original_text in enumerate(slide_texts):
            lines = textwrap.wrap(original_text, width=max_chars_per_line_in_ppt, break_long_words=False,
                                 fix_sentence_endings=True)
            slides_data.append({
                "text": original_text,
                "lines": lines,
                "original_sentence": original_sentence_flags[i]
            })
            total_slides += 1

        for i, data in enumerate(slides_data):
            create_slide(prs, data, current_slide_idx, total_slides, font_size)
            if not data["original_sentence"]:
                check_needed_slides.append(current_slide_idx)
            current_slide_idx += 1

        return prs, check_needed_slides

    except Exception as e:
        print(f"PPT 생성 중 오류 발생: {e}")
        return None, []

def create_slide(prs, data, current_idx, total_slides, font_size):
    """실제로 슬라이드를 생성하는 함수"""

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    tf = textbox.text_frame
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = data["text"]

    p.font.size = Pt(font_size)
    p.font.name = 'Noto Color Emoji'
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    # 페이지 번호 (현재 페이지/전체 페이지)
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
    if not data["original_sentence"]:
        add_check_needed_shape(slide)

def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가하는 함수"""

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
    end_paragraph = end_text_frame.paragraphs[0]
    end_paragraph.text = "끝"
    end_paragraph.font.size = Pt(36)
    end_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    end_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def add_check_needed_shape(slide):
    """슬라이드에 '확인 필요' 표시를 추가하는 함수"""

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
    check_paragraph = check_text_frame.paragraphs[0]
    check_paragraph.text = "확인 필요!"
    check_paragraph.font.size = Pt(18)
    check_paragraph.font.bold = True
    check_paragraph.font.color.rgb = RGBColor(0, 0, 0)
    check_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    check_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기")

# Word 파일 업로드 기능 추가
uploaded_file = st.file_uploader("📝 Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
max_chars_per_line_ppt_input = st.slider("📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=30, value=18, key="max_chars_slider_ppt")
min_chars_per_line_input = st.slider("🔤 한 줄당 최소 글자 수:", min_value=1, max_value=10, value=4, key="min_chars_slider")
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")

if st.button("🚀 PPT 만들기", key="create_ppt_button"):
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    slide_texts, original_sentence_flags = split_and_group_text(
        text,
        max_lines_per_slide=max_lines_per_slide_input,
        min_chars_per_line=min_chars_per_line_input,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input
    )
    ppt, check_needed_slides = create_ppt(
        slide_texts,
        original_sentence_flags,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
        max_lines_per_slide=max_lines_per_slide_input,
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
        if check_needed_slides:
            st.warning(f"❗️ 일부 슬라이드({check_needed_slides})는 한 문장이 너무 길어 분할되었습니다. PPT를 확인하여 가독성을 검토해주세요.")
    else:
        st.error("❌ PPT 생성에 실패했습니다.")