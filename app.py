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

# 전체 입력을 문장 단위로 분해하고, 특정 패턴을 별도 처리
def split_and_group_text(text, separate_pattern=None, max_lines_per_slide=5, min_chars_per_line=4):
    slides = []
    current_slide_sentences = []
    current_slide_lines = 0

    sentences = re.split(r'(?<=[.!?])\s+', text.strip())

    for sentence in sentences:
        sentence = sentence.strip()
        # 특정 패턴을 만족하는지 확인
        if separate_pattern and re.match(sentence):
            # 현재 슬라이드에 내용이 있으면 추가하고 새 슬라이드 시작
            if current_slide_sentences:
                slides.append("\n".join(current_slide_sentences))
            slides.append(sentence)  # 패턴에 맞는 텍스트는 단독 슬라이드로
            current_slide_sentences = []
            current_slide_lines = 0
        else:
            # 일반 문장의 경우, 줄 수를 계산하여 슬라이드에 추가
            words = sentence.split()
            lines_needed = len(words)  # 단어 수를 줄 수로 계산 (띄어쓰기 기준)
            
            # 최소 글자 수를 고려하여 줄 수를 보정
            lines_needed = max(1, (len(sentence) + min_chars_per_line - 1) // min_chars_per_line)
            
            if current_slide_lines + lines_needed <= max_lines_per_slide:
                current_slide_sentences.append(sentence)
                current_slide_lines += lines_needed
            else:
                slides.append("\n".join(current_slide_sentences))
                current_slide_sentences = [sentence]
                current_slide_lines = lines_needed

    # 마지막 슬라이드 내용 추가
    if current_slide_sentences:
        slides.append("\n".join(current_slide_sentences))

    return slides

# 전체 입력을 문장 단위로 분해
def split_text(text):
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

# PPT 생성 함수
def create_ppt(slide_texts, max_chars_per_line_in_ppt=18, max_lines_per_slide=5, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = 0  # 초기값 0으로 설정
    current_slide_idx = 1
    slides_data = []  # 슬라이드 데이터 저장

    try:
        for original_text in slide_texts:
            lines = textwrap.wrap(original_text, width=max_chars_per_line_in_ppt, break_long_words=False,
                                 fix_sentence_endings=True)
            slides_data.append({
                "text": original_text,
                "lines": lines
            })
            total_slides += 1

        # 실제 슬라이드 생성
        for data in slides_data:
            create_slide(prs, data["text"], current_slide_idx, total_slides, font_size)
            current_slide_idx += 1

        return prs

    except Exception as e:
        print(f"PPT 생성 중 오류 발생: {e}")
        return None

def create_slide(prs, text, current_idx, total_slides, font_size):
    """실제로 슬라이드를 생성하는 함수"""

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    tf = textbox.text_frame
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 상단 정렬
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = text

    p.font.size = Pt(font_size)  # 폰트 크기 동적으로 설정
    p.font.name = 'Noto Color Emoji'  # 이모지 지원 글꼴 설정
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

    if current_idx == total_slides:  # 마지막 슬라이드에 '끝' 도형 추가
        add_end_mark(slide)  # 끝 표시 추가 함수 호출

def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가하는 함수"""

    end_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(10),  # left
        Inches(6),   # top
        Inches(2),   # width
        Inches(1)    # height
    )
    end_shape.fill.solid()
    end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)  # 빨간색
    end_shape.line.color.rgb = RGBColor(0, 0, 0)  # 검은색 테두리

    end_text_frame = end_shape.text_frame
    end_text_frame.clear()
    end_paragraph = end_text_frame.paragraphs[0]
    end_paragraph.text = "끝"
    end_paragraph.font.size = Pt(36)
    end_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # 흰색 글자
    end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    end_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기")

text_input = st.text_area("📝 촬영 대본을 입력하세요:", height=300, key="text_input_area")

# "분리할 텍스트 패턴" 입력란에서 기본값 제거
separate_pattern_input = st.text_input("🔍 분리할 텍스트 패턴 (정규 표현식):", key="separate_pattern_input")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
# PPT 텍스트 박스 내에서의 줄바꿈 글자 수 (실제 PPT에 표시될 때 적용)
max_chars_per_line_ppt_input = st.slider("🔤 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=30, value=18, key="max_chars_slider_ppt")
min_chars_per_line_input = st.slider("🔤 한 줄당 최소 글자 수:", min_value=1, max_value=10, value=4, key="min_chars_slider")
font_size_input = st.slider("Aa 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")


if st.button("🚀 PPT 만들기", key="create_ppt_button") and text_input.strip():
    # 수정된 함수 호출
    slide_texts = split_and_group_text(text_input, separate_pattern=separate_pattern_input,
                                        max_lines_per_slide=max_lines_per_slide_input,
                                        min_chars_per_line=min_chars_per_line_input)
    ppt = create_ppt(slide_texts, max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
                    max_lines_per_slide=max_lines_per_slide_input,
                    font_size=font_size_input)

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
    else:
        st.error("❌ PPT 생성에 실패했습니다.")