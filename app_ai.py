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

# 2. 함수 정의 (Word 파일 처리)
def extract_text_from_word(file_path):
    """Word 파일에서 모든 텍스트를 추출하여, 단락 단위로 분리하여 리스트로 반환합니다."""
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

def get_sentence_embeddings(text, model_name='paraphrase-multilingual-mpnet-base-v2'):
    """텍스트에서 문장 임베딩을 추출합니다."""
    model = SentenceTransformer(model_name)
    sentences = smart_sentence_split(text)
    embeddings = model.encode(sentences)
    return sentences, embeddings

def smart_sentence_split(text):
    """문맥을 고려하여 자연스럽게 문장을 분할합니다."""
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

# 4. 함수 정의 (슬라이드 분할)
def split_text_into_slides_with_similarity(
    text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt, similarity_threshold=0.85
):
    """
    단락 및 문장 유사도를 기반으로 슬라이드를 분할합니다.
    유사도가 높은 문장들을 최대한 한 슬라이드에 넣되, 최대 줄 수를 넘지 않도록 합니다.
    """

    slides = []
    split_flags = []
    slide_numbers = []
    slide_number = 1
    current_slide_text = ""
    current_slide_lines = 0
    needs_check = False  # '확인 필요!' 표시 여부

    model = SentenceTransformer('paraphrase-multilingual-mpnet-base-v2')

    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)
        embeddings = model.encode(sentences) if sentences else []

        i = 0
        while i < len(sentences):
            sentence = sentences[i]
            sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

            # 한 문장이 최대 줄 수를 초과하는 경우 슬라이드 분리
            if sentence_lines > max_lines_per_slide:
                wrapped_lines = textwrap.wrap(
                    sentence, width=max_chars_per_line_ppt, break_long_words=True
                )
                temp_slide_text = ""
                temp_slide_lines = 0
                for line in wrapped_lines:
                    line_lines = calculate_text_lines(line, max_chars_per_line_ppt)
                    if temp_slide_lines + line_lines + 1 <= max_lines_per_slide:
                        temp_slide_text += line + "\n"
                        temp_slide_lines += line_lines + 1
                    else:
                        slides.append(temp_slide_text.strip())
                        split_flags.append(True)  # '확인 필요!' 표시
                        slide_numbers.append(slide_number)
                        logging.debug(
                            f"Slide {slide_number}: {temp_slide_text[:100]}..."
                        )
                        slide_number += 1
                        temp_slide_text = line + "\n"
                        temp_slide_lines = line_lines + 1
                slides.append(temp_slide_text.strip())
                split_flags.append(True)  # '확인 필요!' 표시
                slide_numbers.append(slide_number)
                logging.debug(f"Slide {slide_number}: {temp_slide_text[:100]}...")
                slide_number += 1
                current_slide_text = ""
                current_slide_lines = 0
                needs_check = True
                i += 1
            elif current_slide_lines + sentence_lines + 1 <= max_lines_per_slide:
                # 현재 슬라이드에 추가 가능한 경우, 유사도 높은 문장들을 최대한 추가
                current_slide_text += sentence + "\n"
                current_slide_lines += sentence_lines + 1
                j = i + 1
                while j < len(sentences):
                    next_sentence = sentences[j]
                    next_sentence_lines = calculate_text_lines(
                        next_sentence, max_chars_per_line_ppt
                    )
                    if (
                        current_slide_lines + next_sentence_lines + 1
                        <= max_lines_per_slide
                    ):
                        if i > 0:
                            similarity = util.cos_sim(
                                embeddings[i - 1], embeddings[j]
                            )[0][0].item()
                            if (
                                similarity >= similarity_threshold
                                or paragraph.endswith(next_sentence)
                            ):  # 유사도 검사 또는 단락 끝
                                current_slide_text += next_sentence + "\n"
                                current_slide_lines += next_sentence_lines + 1
                                i = j
                                j += 1
                            else:
                                break  # 유사도 부족하면 추가 중단
                        else:  # 첫 번째 문장은 유사도 검사 생략
                            current_slide_text += next_sentence + "\n"
                            current_slide_lines += next_sentence_lines + 1
                            i = j
                            j += 1
                    else:
                        break  # 최대 줄 수 초과하면 추가 중단
                i += 1
            else:
                # 현재 슬라이드에 추가 불가능한 경우
                slides.append(current_slide_text.strip())
                split_flags.append(needs_check)
                slide_numbers.append(slide_number)
                logging.debug(f"Slide {slide_number}: {current_slide_text[:100]}...")
                slide_number += 1
                current_slide_text = sentence + "\n"
                current_slide_lines = sentence_lines + 1
                needs_check = False
                i += 1

        if current_slide_text:  # 마지막 슬라이드 추가
            slides.append(current_slide_text.strip())
            split_flags.append(needs_check)
            slide_numbers.append(slide_number)
            logging.debug(f"Slide {slide_number}: {current_slide_text[:100]}...")

    return slides, split_flags, slide_numbers

# 5. 함수 정의 (PPT 생성 및 슬라이드 조작)
def create_ppt(slide_texts, split_flags, max_chars_per_line_in_ppt=18, font_size=54):
    """슬라이드 텍스트를 기반으로 PPT를 생성하고, '확인 필요!' 표시 등을 추가합니다."""

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
                add_check_needed_shape(slide)  # 슬라이드 번호 인자 제거
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

def add_check_needed_shape(slide):
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
    p.text = "확인 필요!"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    text_frame = check_shape.text_frame
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

# 7. Streamlit UI
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
st.title("🎬 AI PPT 생성기")

# CSS를 사용하여 요소를 가운데 정렬
st.markdown(
    """
    <style>
    .centered-container {
        display: flex;
        flex-direction: column;
        align-items: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# UI 레이아웃 분할
with st.container():
    st.markdown("<div class='centered-container'>", unsafe_allow_html=True)
    input_col, result_col = st.columns([1, 2])

    # 입력 옵션
    with input_col:
        with st.expander("입력 옵션 (클릭하여 펼치기)", expanded=True):
            uploaded_file = st.file_uploader("Word 파일 업로드", type=["docx"], key="file_uploader")
            text_input = st.text_area("또는 텍스트 직접 입력", height=200, key="text_input_area")
            max_lines_per_slide_input = st.slider(
                "슬라이드당 최대 줄 수", min_value=1, max_value=10, value=5, key="max_lines_slider"
            )
            max_chars_per_line_ppt_input = st.slider(
                "PPT 한 줄당 최대 글자 수", min_value=10, min_value=100, value=18, key="max_chars_slider_ppt"
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
                key="similarity_threshold_input"
            )

        if st.button("PPT 생성", key="generate_ppt_button"):
            text = ""
            if uploaded_file is not None:
                text_paragraphs = extract_text_from_word(uploaded_file)
            elif text_input.strip():
                text_paragraphs = text_input.split("\n\n")
            else:
                st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
                st.stopwith result_col:
                with st.spinner("PPT 생성 중..."):
                    try:
                        slide_texts, split_flags, slide_numbers = split_text_into_slides_with_similarity(
                            text_paragraphs,
                            max_lines_per_slide=st.session_state.max_lines_slider,
                            max_chars_per_line_ppt=st.session_state.max_chars_slider_ppt,
                            similarity_threshold=st.session_state.similarity_threshold_input
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
                    except Exception as e:
                        st.error(f"오류: PPT 저장 실패: {e}")
                        st.error(f"오류 상세 내용: {str(e)}")
                    else:
                        st.download_button(
                            label="PPT 다운로드",
                            data=ppt_io,
                            file_name="paydo_script_ai.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key="download_ppt_button"
                        )

                    if divided_slide_count > 0:
                        st.warning(
                            f"총 {len(slide_texts)}개의 슬라이드 중 <b>{divided_slide_count}개</b>의 슬라이드가 나뉘어 졌습니다. 확인이 필요합니다.",
                            unsafe_allow_html=True
                        )
                    else:
                        st.success("PPT가 성공적으로 생성되었습니다.")

                    st.subheader("생성된 슬라이드 목록")
                    for i in range(len(slide_texts)):
                        if split_flags[i]:
                            st.error(
                                f"슬라이드 {i + 1}/{len(slide_texts)} (나뉨): {slide_texts[i][:100]}...",
                                icon="⚠️",
                                key=f"split_slide_{i}"
                            )
                        else:
                            st.info(
                                f"슬라이드 {i + 1}/{len(slide_texts)}: {slide_texts[i][:100]}...",
                                icon="ℹ️",
                                key=f"normal_slide_{i}"
                            )
    st.markdown("</div>", unsafe_allow_html=True)  # centered-container 닫기