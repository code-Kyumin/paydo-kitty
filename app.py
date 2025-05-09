import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io
import re
import textwrap
import docx  # python-docx 라이브러리 추가


# Word 파일에서 텍스트 추출하는 함수
def extract_text_from_word(file):
    doc = docx.Document(file)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
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


# 전체 입력을 문장 단위로 분해하고, 슬라이드 단위로 묶음
def split_and_group_text(
    text, max_lines_per_slide=5, min_chars_per_line=4, max_chars_per_line_in_ppt=18
):
    slides = []
    current_slide_text = ""
    current_slide_lines = 0
    split_occurred = False  # 문장 분할 발생 여부 추적
    original_texts = []  # [추가] 원본 텍스트 저장을 위한 리스트

    sentences = re.split(r"(?<=[.!?])\s+", text.strip())

    for i, sentence in enumerate(sentences):
        lines_needed = sentence_line_count(sentence, max_chars_per_line_in_ppt)

        # 현재 슬라이드에 추가 가능한지 확인
        if current_slide_lines + lines_needed <= max_lines_per_slide:
            current_slide_text += sentence + " "
            current_slide_lines += lines_needed
        else:
            # 현재 슬라이드에 추가할 수 없는 경우
            if current_slide_text:
                slides.append(current_slide_text.strip())
                original_texts.append(
                    current_slide_text.strip()
                )  # [추가] 원본 텍스트 저장
            current_slide_text = sentence + " "
            current_slide_lines = lines_needed
            split_occurred = True  # 분할이 일어났음을 기록

        # 다음 문장과 합쳐도 최대 줄 수를 넘지 않는지 확인
        if i + 1 < len(sentences):
            next_sentence = sentences[i + 1]
            next_lines_needed = sentence_line_count(
                current_slide_text + next_sentence, max_chars_per_line_in_ppt
            )
            if next_lines_needed <= max_lines_per_slide:
                current_slide_text += next_sentence + " "
                current_slide_lines = next_lines_needed
                i += 1  # 다음 문장 처리 건너뜀

    if current_slide_text:
        slides.append(current_slide_text.strip())
        original_texts.append(current_slide_text.strip())  # [추가] 마지막 텍스트 저장

    return slides, split_occurred, original_texts  # [수정] 원본 텍스트 반환


# PPT 생성 함수
def create_ppt(
    slide_texts,
    original_texts,  # [추가] 원본 텍스트 받음
    max_chars_per_line_in_ppt=18,
    font_size=54,
    max_lines_per_slide=5,  # [추가] 최대 줄 수 인자 받음
):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)
    check_needed_slides = []  # [추가] 확인 필요 슬라이드 번호 저장

    for i, text in enumerate(slide_texts):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        tf.word_wrap = True
        tf.clear()

        lines = textwrap.wrap(text, width=max_chars_per_line_in_ppt, break_long_words=False)
        if len(lines) > max_lines_per_slide:  # [수정] 최대 줄 수 초과 확인
            check_needed_slides.append(i + 1)  # 슬라이드 번호 저장
            p = tf.paragraphs[0]
            p.text = text  # 원본 텍스트 그대로 표시
            p.font.size = Pt(font_size)
            p.font.name = "Noto Color Emoji"
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER
            add_check_needed_shape(slide)  # "확인 필요" 도형 추가
        else:
            p = tf.paragraphs[0]
            p.text = "\n".join(lines)
            p.font.size = Pt(font_size)
            p.font.name = "Noto Color Emoji"
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER

        # 페이지 번호 (현재 페이지/전체 페이지)
        footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
        footer_frame = footer_box.text_frame
        footer_frame.text = f"{i + 1} / {total_slides}"
        footer_p = footer_frame.paragraphs[0]
        footer_p.font.size = Pt(18)
        footer_p.font.name = "맑은 고딕"
        footer_p.font.color.rgb = RGBColor(128, 128, 128)
        footer_p.alignment = PP_ALIGN.RIGHT

        if i == total_slides - 1:  # 마지막 슬라이드에 '끝' 표시 추가
            add_end_mark(slide)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)

    return ppt_io, check_needed_slides  # [수정] 확인 필요 슬라이드 번호 반환


def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가하는 함수 (우측 하단, 도형 및 색상 추가)"""

    end_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(10), Inches(6), Inches(2), Inches(1)
    )
    end_shape.fill.solid()
    end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)  # 빨간색 배경
    end_shape.line.color.rgb = RGBColor(0, 0, 0)  # 검은색 테두리

    end_text_frame = end_shape.text_frame
    end_text_frame.clear()
    end_paragraph = end_text_frame.paragraphs[0]
    end_paragraph.text = "끝"
    end_paragraph.font.size = Pt(36)
    end_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # 흰색 글자
    end_paragraph.font.bold = True
    end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    end_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


def add_check_needed_shape(slide):
    """슬라이드에 '확인 필요' 도형을 추가하는 함수"""

    check_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.3), Inches(2), Inches(0.5)
    )
    check_shape.fill.solid()
    check_shape.fill.fore_color.rgb = RGBColor(255, 255, 0)  # 노란색 배경
    check_shape.line.color.rgb = RGBColor(0, 0, 0)  # 검은색 테두리

    check_text_frame = check_shape.text_frame
    check_text_frame.clear()
    check_paragraph = check_text_frame.paragraphs[0]
    check_paragraph.text = "확인 필요"
    check_paragraph.font.size = Pt(18)
    check_paragraph.font.bold = True
    check_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    check_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기")

# Word 파일 업로드 기능 추가
uploaded_file = st.file_uploader("📝 Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider(
    "📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider"
)
# PPT 텍스트 박스 내에서의 줄바꿈 글자 수 (실제 PPT에 표시될 때 적용)
max_chars_per_line_ppt_input = st.slider(
    "📏 한 줄당 최대 글자 수 (PPT 표시):",
    min_value=3,
    max_value=30,
    value=18,
    key="max_chars_slider_ppt",
)
min_chars_per_line_input = st.slider(
    "🔤 한 줄당 최소 글자 수:", min_value=1, max_value=10, value=4, key="min_chars_slider"
)
font_size_input = st.slider(
    "🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider"
)


if st.button("🚀 PPT 만들기", key="create_ppt_button"):
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    # 수정된 함수 호출
    slide_texts, split_occurred, original_texts = split_and_group_text(
        text,
        max_lines_per_slide=max_lines_per_slide_input,
        min_chars_per_line=min_chars_per_line_input,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
    )
    ppt_file, check_needed_slides = create_ppt(
        slide_texts,
        original_texts,  # [추가] 원본 텍스트 전달
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
        font_size=font_size_input,
        max_lines_per_slide=max_lines_per_slide_input,
    )

    st.download_button(
        label="📥 PPT 다운로드",
        data=ppt_file,
        file_name="paydo_script.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="download_button",
    )

    if split_occurred:
        st.info(
            "⚠️ 긴 문장으로 인해 일부 슬라이드가 자동으로 분할되었습니다. PPT를 확인하여 어색한 부분이 있는지 검토해주세요."
        )

    if check_needed_slides:  # [추가] 확인 필요 슬라이드 있는 경우 알림
        st.warning(
            f"❗️ 일부 슬라이드({check_needed_slides})는 최대 줄 수를 초과하여 텍스트가 나뉘었습니다. PPT를 확인하여 가독성을 검토해주세요."
        )