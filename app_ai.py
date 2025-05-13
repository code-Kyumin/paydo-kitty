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
from sentence_transformers import SentenceTransformer, util
import logging

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 사용할 한국어 특화 모델 (무료, Streamlit 사용 가능)
model_name = 'jhgan/ko-sroberta-multitask'  # 또는 다른 적절한 모델

# 1. 함수 정의: Word 파일 처리
def extract_text_from_word(file_path):
    """Word 파일에서 텍스트 추출."""
    try:
        doc = docx.Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs]
        logging.debug(f"Word paragraphs extracted: {len(paragraphs)} paragraphs")
        return paragraphs
    except FileNotFoundError:
        st.error(f"오류: Word 파일을 찾을 수 없습니다.")
        return None
    except docx.exceptions.PackageNotFoundError:
        st.error(f"오류: Word 파일이 유효하지 않습니다.")
        return None
    except Exception as e:
        st.error(f"오류: Word 파일 처리 중 오류 발생: {e}")
        return None

# 2. 함수 정의: 텍스트 줄 수 계산
def calculate_text_lines(text, max_chars_per_line):
    """텍스트의 줄 수를 계산."""
    lines = 0
    paragraphs = text.split('\n')
    for paragraph in paragraphs:
        if not paragraph:
            lines += 1
        else:
            lines += len(textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=True))
    return lines

# 3. 함수 정의: 문장 분리 및 임베딩 생성
@st.cache_resource  # 모델 로딩을 캐싱하여 성능 향상
def load_embedding_model(model_name):
    """임베딩 모델 로드 (캐싱됨)."""
    return SentenceTransformer(model_name)

def get_sentence_embeddings(text, model):
    """텍스트에서 문장 임베딩 추출."""

    sentences = smart_sentence_split(text)
    embeddings = model.encode(sentences, convert_to_tensor=True)
    return sentences, embeddings

def smart_sentence_split(text):
    """문맥을 고려하여 자연스럽게 문장 분리."""

    paragraphs = text.split('\n')
    sentences = []
    for paragraph in paragraphs:
        temp_sentences = re.split(r'(?<!\b\w)([.?!])(?=\s|$)', paragraph)
        temp = []
        for i in range(0, len(temp_sentences), 2):
            if i + 1 < len(temp_sentences):
                temp.append(temp_sentences[i] + temp_sentences[i + 1])
            else:
                temp.append(temp_sentences[i])
        sentences.extend(temp)
    sentences = [s.strip() for s in sentences if s.strip()]
    logging.debug(f"Sentences split: {len(sentences)} sentences")
    return sentences

# 4. 함수 정의: 슬라이드 분할 (문맥 유사도 기반)
def split_text_into_slides_with_similarity(
    text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt,
    similarity_threshold=0.85, model=None  # 모델 인자로 받음
):
    """문맥 유사도를 기반으로 슬라이드 분할."""

    slides = []
    split_flags = []
    slide_numbers = []
    slide_number = 1
    current_slide_text = ""
    current_slide_lines = 0
    needs_check = False

    all_sentences = []
    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)
        all_sentences.extend(sentences)

    all_embeddings = model.encode(all_sentences, convert_to_tensor=True)  # 임베딩 생성

    embedding_index = 0
    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)

        for i, sentence in enumerate(sentences):
            sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

            if sentence_lines > max_lines_per_slide:
                # 긴 문장 분할 처리
                wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line_ppt, break_long_words=True)
                # ... (이전 코드와 동일)
            elif current_slide_lines + sentence_lines + 1 <= max_lines_per_slide:
                # 슬라이드에 추가 가능한 경우 유사도 검사
                if current_slide_text and i > 0:
                    similarity = util.cos_sim(all_embeddings[embedding_index + i - 1].unsqueeze(0), all_embeddings[embedding_index + i].unsqueeze(0))[0][0].item()
                    if similarity < similarity_threshold:
                        # ... (이전 코드와 동일)
                    else:
                        # ... (이전 코드와 동일)
                else:
                    # ... (이전 코드와 동일)
            else:
                # 슬라이드에 추가 불가능한 경우
                # ... (이전 코드와 동일)
        embedding_index += len(sentences)

    if current_slide_text:
        # 마지막 슬라이드 처리
        # ... (이전 코드와 동일)

    return slides, split_flags, slide_numbers

# 5. 함수 정의: PPT 생성
def create_ppt(slide_texts, split_flags, max_chars_per_line_in_ppt=18, font_size=54):
    """슬라이드 텍스트를 기반으로 PPT 생성."""
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)

    for i, text in enumerate(slide_texts):
        try:
            logging.debug(f"슬라이드 {i+1}에 텍스트 추가")
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER, max_chars_per_line_in_ppt)
            if split_flags[i]:
                add_check_needed_shape(slide)
            if i == total_slides - 1:
                add_end_mark(slide)
        except Exception as e:
            st.error(f"오류: 슬라이드 생성 실패 (슬라이드 {i+1}): {e}")
            return None

    return prs

# 6. 함수 정의: 슬라이드 요소 추가
def add_text_to_slide(slide, text, font_size, alignment, max_chars_per_line):
    """슬라이드에 텍스트 추가 및 스타일 설정."""
    try:
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        text_frame.word_wrap = True

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

        text_frame.auto_size = None
        logging.debug(f"텍스트 추가됨")
    except Exception as e:
        st.error(f"오류: 슬라이드에 텍스트 추가 중 오류 발생: {e}")
        raise

def add_end_mark(slide):
    """마지막 슬라이드에 '끝' 표시 추가."""
    # ... (이전 코드와 동일)

def add_check_needed_shape(slide):
    """확인 필요 슬라이드에 '확인 필요!' 상자 추가."""
    # ... (이전 코드와 동일)

# 7. Streamlit UI
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
st.title("🎬 AI PPT 생성기")

# Word 파일 업로드
uploaded_file = st.file_uploader("Word 파일 업로드", type=["docx"])

# 텍스트 직접 입력
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
    값이 낮을수록 슬라이드가 짧아지고 가독성이 높아집니다(발표용).
    값이 높을수록 문맥이 유지되며 정보 밀도가 높아집니다 (강의용).
    """,
    key="similarity_threshold_input"
)

# 8. PPT 생성 및 다운로드
if st.button("PPT 생성"):
    text = ""
    if uploaded_file is not None:
        text_paragraphs = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text_paragraphs = text_input.split("\n\n")
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    with st.spinner("PPT 생성 중..."):
        try:
            model = load_embedding_model(model_name)  # 모델 로드
            slide_texts, split_flags, slide_numbers = split_text_into_slides_with_similarity(
                text_paragraphs,
                max_lines_per_slide=st.session_state.max_lines_slider,
                max_chars_per_line_ppt=st.session_state.max_chars_slider_ppt,
                similarity_threshold=st.session_state.similarity_threshold_input,
                model=model  # 모델 전달
            )
            ppt = create_ppt(
                slide_texts, split_flags,
                max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
                font_size=st.session_state.font_size_slider
            )
            divided_slide_count = sum(split_flags)
        except Exception as e:
            st.error(f"오류: PPT 생성 실패: {e}")
            st.error(f"오류 상세 내용: {str(e)}")
            st.stop()

    if ppt:
        ppt_io = io.BytesIO()
        try:
            ppt.save(ppt_io)
            ppt_io.seek(0)
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

        st.subheader("생성 결과")
        st.write(f"총 {len(slide_texts)}개의 슬라이드가 생성되었습니다.")
        if divided_slide_count > 0:
            divided_slide_numbers = [i + 1 for i, flag in enumerate(split_flags) if flag]
            st.warning(f"이 중 {divided_slide_count}개의 슬라이드(번호: {divided_slide_numbers})는 나뉘어 졌으므로 확인이 필요합니다.")
        else:
            st.success("나뉘어진 슬라이드 없이 PPT가 성공적으로 생성되었습니다.")