# 1. 라이브러리 임포트
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap
import docx
# from sentence_transformers import SentenceTransformer, util  # 임베딩 관련 라이브러리 제거
import logging

# 로깅 설정 (디버깅용)
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 2. 함수 정의 (Word 파일 처리)
def extract_text_from_word(file_path):
    """Word 파일에서 모든 텍스트를 추출하여, 단락 단위로 분리하여 리스트로 반환합니다."""
    try:
        doc = docx.Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs]
        logging.debug(f"Word paragraphs extracted: {len(paragraphs)} paragraphs")  # Log 추출된 단락 수
        return paragraphs  # 단락 리스트 반환
    except FileNotFoundError:
        st.error(f"오류: Word 파일을 찾을 수 없습니다.")
        return None
    except docx.exceptions.PackageNotFoundError:
        st.error(f"오류: Word 파일이 유효하지 않습니다.")
        return None
    except Exception as e:
        st.error(f"오류: Word 파일 처리 중 오류 발생: {e}")
        return None

# 3. 함수 정의 (텍스트 처리)
def calculate_text_lines(text, max_chars_per_line):
    """텍스트의 줄 수를 계산합니다."""
    lines = 0
    paragraphs = text.split('\n')
    for paragraph in paragraphs:
        if not paragraph:
            lines += 1
        else:
            lines += len(textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=True))
    return lines

# def get_sentence_embeddings(text, model_name='paraphrase-multilingual-mpnet-base-v2'):
#     """텍스트에서 문장 임베딩을 추출합니다."""
#     model = SentenceTransformer(model_name)
#     sentences = smart_sentence_split(text)
#     embeddings = model.encode(sentences)
#     return sentences, embeddings

# def smart_sentence_split(text):
#     """문맥을 고려하여 자연스럽게 문장을 분할합니다."""
#     # 더 이상 사용하지 않음
#     paragraphs = text.split('\n')
#     sentences = []
#     for paragraph in paragraphs:
#         temp_sentences = re.split(r'(?<!\b\w)([.?!])(?=\s|$)', paragraph)
#         temp = []
#         for i in range(0, len(temp_sentences), 2):
#             if i + 1 < len(temp_sentences):
#                 temp.append(temp_sentences[i] + temp_sentences[i + 1])
#             else:
#                 temp.append(temp_sentences[i])
#         sentences.extend(temp)
#     sentences = [s.strip() for s in sentences if s.strip()]
#     logging.debug(f"Sentences split: {len(sentences)} sentences")
#     return sentences

# def smart_sub_split(sentence):
#     """더 복잡한 문장 구조를 고려하여 하위 문장으로 분리합니다."""
#     # 더 이상 사용하지 않음
#     sub_sentences = re.split(r',\s*(그리고|그러나|왜냐하면|예를 들어|즉|또는)\s+', sentence)
#     return sub_sentences

# 4. 함수 정의 (슬라이드 분할)
def split_text_into_slides(
    text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt
):
    """단락 리스트를 기반으로 슬라이드를 분할합니다."""

    slides = []
    current_slide_text = ""
    current_slide_lines = 0
    slide_number = 1

    for paragraph in text_paragraphs:
        paragraph_lines = calculate_text_lines(paragraph, max_chars_per_line_ppt)

        if current_slide_lines + paragraph_lines <= max_lines_per_slide:
            current_slide_text += paragraph + "\n\n"  # 단락 간 여백 추가
            current_slide_lines += paragraph_lines + 2  # 여백에 대한 줄 수 추가
        else:
            slides.append(current_slide_text.strip())
            logging.debug(f"Slide {slide_number}: {current_slide_text[:100]}...")  # Log 슬라이드 내용 (일부)
            slide_number += 1
            current_slide_text = paragraph + "\n\n"
            current_slide_lines = paragraph_lines + 2

    slides.append(current_slide_text.strip())  # 마지막 슬라이드 추가
    logging.debug(f"Slide {slide_number}: {current_slide_text[:100]}...")  # Log 마지막 슬라이드 내용 (일부)

    return slides, [False] * len(slides), list(range(1, len(slides) + 1))  # split_flags, slide_numbers 추가

# 5. 함수 정의 (PPT 생성 및 슬라이드 조작)
def create_ppt(slide_texts, split_flags, slide_numbers, max_chars_per_line_in_ppt=18, font_size=54):
    """슬라이드 텍스트를 기반으로 PPT를 생성하고, 슬라이드 번호, '끝' 표시 등을 추가합니다."""

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)

    for i, text in enumerate(slide_texts):
        try:
            logging.debug(f"슬라이드 {i+1}에 텍스트 추가")
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER, max_chars_per_line_in_ppt)
            add_slide_number(slide, slide_numbers[i], total_slides)
            if split_flags[i] and calculate_text_lines(text, max_chars_per_line_in_ppt) == 1:
                add_check_needed_shape(slide, slide_numbers[i], slide_numbers[i])
            if i == total_slides - 1:
                add_end_mark(slide)
        except Exception as e:
            st.error(f"오류: 슬라이드 생성 실패 (슬라이드 {i+1}): {e}")
            return None

    return prs

# 6. 함수 정의 (슬라이드 요소 추가)
def add_text_to_slide(slide, text, font_size, alignment, max_chars_per_line):
    """슬라이드에 텍스트를 추가하고, 폰트, 크기, 정렬 등을 설정합니다."""

    try:
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        text_frame.word_wrap = True

        # TODO: 텍스트 래핑 및 레이아웃 방식 추가 개선 (줄 간격, 자동 맞춤 등)
        wrapped_lines = textwrap.wrap(text, width=max_chars_per_line, break_long_words=True)
        for line in wrapped_lines:
            p = text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(font_size)
            p.font.name = 'Noto Color Emoji'
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = alignment
            p.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

        text_frame.auto_size = None  # 텍스트 프레임 자동 크기 조절 비활성화
        logging.debug(f"텍스트 추가됨")
    except Exception as e:
        st.error(f"오류: 슬라이드에 텍스트 추가 중 오류 발생: {e}")
        raise

def add_slide_number(slide, current, total):
    """슬라이드에 슬라이드 번호를 추가합니다."""
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
    """마지막 슬라이드에 '끝' 표시를 추가합니다."""

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

def add_check_needed_shape(slide, slide_number, ui_slide_number):
    """확인 필요한 슬라이드에 '확인 필요!' 상자를 추가합니다."""

    check_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.5),
        Inches(0.3),
        Inches(2.5),
        Inches(0.5)
    )
    check_shape.fill.solid()
    check_shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
    check_shape.line.color.rgb = RGBColor(0, 0, 0)

    check_text_frame = check_shape.text_frame
    check_text_frame.clear()
    p = check_text_frame.paragraphs[0]
    p.text = f"확인 필요 (슬라이드 {ui_slide_number})"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    text_frame = check_shape.text_frame
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

# 7. Streamlit UI
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
st.title("🎬 AI PPT 생성기")

# Word 파일 업로드
uploaded_file = st.file_uploader("Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력", height=300, key="text_input_area")

# UI 입력 슬라이더
max_lines_per_slide_input = st.slider(
    "슬라이드당 최대 줄 수", min_value=1, max_value=10, value=5, key="max_lines_slider"
)
max_chars_per_line_ppt_input = st.slider(
    "PPT 한 줄당 최대 글자 수", min_value=10, max_value=100, value=18, key="max_chars_slider_ppt"
)
font_size_input = st.slider("폰트 크기", min_value=10, max_value=60, value=54, key="font_size_slider")

similarity_threshold_input = st.slider(
    "문맥 유사도 기준",
    min_value=0.0, max_value=1.0, value=0.85, step=0.05,
    help="""
    문맥 유사도가 낮을 경우 슬라이드를 분리합니다.
    값이 낮을수록 슬라이드가 짧아지고 가독성이 높아집니다 (발표용).
    값이 높을수록 문맥이 유지되며 정보 밀도가 높아집니다 (강의용).
    """,
    key="similarity_threshold_input" # 이 부분은 수정되지 않도록 해줘.
)

# 8. PPT 생성 및 다운로드
if st.button("PPT 생성"):
    text = ""
    if uploaded_file is not None:
        text_paragraphs = extract_text_from_word(uploaded_file)  # 단락 리스트로 받음
        text = "\n\n".join(text_paragraphs) # text를 다시 만듦
    elif text_input.strip():
        text = text_input
        text_paragraphs = text.split("\n\n") # text_paragraphs 도 만듦
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    with st.spinner("PPT 생성 중..."):
        try:
            slide_texts, split_flags, slide_numbers = split_text_into_slides(  # 함수 이름 변경
                text_paragraphs,  # text_paragraphs 사용
                max_lines_per_slide=st.session_state.max_lines_slider,
                max_chars_per_line_ppt=st.session_state.max_chars_slider_ppt,
            )
            ppt = create_ppt(
                slide_texts, split_flags, slide_numbers,
                max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
                font_size=st.session_state.font_size_slider
            )
        except Exception as e:
            st.error(f"오류: PPT 생성 실패: {e}")
            st.error(f"오류 상세 내용: {str(e)}")
            st.stop()

    if ppt:
        ppt_io = io.BytesIO()
        try:
            ppt.save(ppt_io)
            ppt_io.seek(0)
        except Exception as e:
            st.error(f"오류: PPT 저장 실패: {e}")
            st.error(f"오류 상세 내용: {str(e)}")
        else:
            st.download_button(
                label="PPT 다운로드",
                data=ppt_io,
                file_name="paydo_script_ai.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )