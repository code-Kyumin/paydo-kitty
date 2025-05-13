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
import logging
from typing import List
import os

# 공통 라이브러리 (두 기능 모두 사용)
common_libs = True  # 이 플래그는 실제로는 사용되지 않지만, 설명을 위해 존재

# 기존 기능 관련 라이브러리
try:
    from sentence_transformers import SentenceTransformer, util
    legacy_libs_available = True
except ImportError:
    legacy_libs_available = False

# gensim 관련 라이브러리 (별도 파일에서 처리)
gensim_libs_available = False  # 기본값으로 False 설정

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 사용할 한국어 특화 모델 (기존 기능)
model_name = 'jhgan/ko-sroberta-multitask'

# 2. 함수 정의 (Word 파일 처리)
def extract_text_from_word(file_path: str) -> List[str] or None:
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

# 3. 함수 정의 (텍스트 처리) (기존 기능)
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

# KoSBERT 임베딩 생성 함수 (기존 기능)
def get_kosbert_embeddings(sentences, model_name):
    """KoSBERT 임베딩을 생성합니다."""
    model = SentenceTransformer(model_name)
    embeddings = model.encode(sentences, convert_to_tensor=True)
    return embeddings

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

# 4. 함수 정의 (슬라이드 분할) (기존 기능)
def split_text_into_slides_with_similarity(
    text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt, similarity_threshold=0.85, model_name='jhgan/ko-sroberta-multitask'
):
    """
    단락 및 문장 유사도를 기반으로 슬라이드를 분할합니다.
    한 문장이 최대 줄 수를 초과하는 경우 슬라이드를 분리하고, 해당 슬라이드에 '확인 필요!' 표시를 합니다.
    """

    slides = []
    split_flags = []
    slide_numbers = []
    slide_number = 1
    current_slide_text = ""
    current_slide_lines = 0
    needs_check = False  # '확인 필요!' 표시 여부

    all_sentences = []
    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)
        all_sentences.extend(sentences)

    # 모든 문장의 KoSBERT 임베딩 생성
    all_embeddings = get_kosbert_embeddings(all_sentences, model_name)

    embedding_index = 0
    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)

        for i, sentence in enumerate(sentences):
            sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

            # 한 문장이 최대 줄 수를 초과하는 경우 슬라이드 분리
            if sentence_lines > max_lines_per_slide:
                wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line_ppt, break_long_words=True)
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
                        logging.debug(f"Slide {slide_number}: {temp_slide_text[:100]}...")
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
            elif current_slide_lines + sentence_lines + 1 <= max_lines_per_slide:
                # 현재 슬라이드에 추가 가능한 경우
                # 첫 번째 문장이 아니면 이전 문장과의 유사도 검사
                if current_slide_text and i > 0:
                    similarity = util.cos_sim(embeddings[embedding_index + i - 1].unsqueeze(0), embeddings[embedding_index + i].unsqueeze(0))[0][0].item()
                    if similarity < similarity_threshold:
                        slides.append(current_slide_text.strip())
                        split_flags.append(needs_check)
                        slide_numbers.append(slide_number)
                        logging.debug(f"Slide {slide_number}: {current_slide_text[:100]}...")
                        slide_number += 1
                        current_slide_text = sentence + "\n"
                        current_slide_lines = sentence_lines + 1
                        needs_check = False
                    else:
                        current_slide_text += sentence + "\n"
                        current_slide_lines += sentence_lines + 1
                else:
                    current_slide_text += sentence + "\n"
                    current_slide_lines += sentence_lines + 1
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
        embedding_index += len(sentences)

    if current_slide_text:  # 마지막 슬라이드 추가
        slides.append(current_slide_text.strip())
        split_flags.append(needs_check)
        slide_numbers.append(slide_number)
        logging.debug(f"Slide {slide_number}: {current_slide_text[:100]}...")

    return slides, split_flags, slide_numbers

# 5. 함수 정의 (PPT 생성 및 슬라이드 조작)
def create_ppt(slide_texts, split_flags, max_chars_per_line_in_ppt=18, font_size=54, font_name="Noto Color Emoji"):
    """슬라이드 텍스트를 기반으로 PPT를 생성하고, '확인 필요!' 표시 등을 추가합니다."""

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)
    prs.core_properties.title = "AI Script Reader"  # PPT 제목 설정 (전체 PPT 제목)

    # 슬라이드 레이아웃 설정 (제목 및 내용)
    title_slide_layout = prs.slide_layouts[0]  # 제목 슬라이드 레이아웃
    content_slide_layout = prs.slide_layouts[5]  # 제목+내용 슬라이드 레이아웃

    # 첫 번째 슬라이드 (제목 슬라이드)
    title_slide = prs.slides.add_slide(title_slide_layout)
    title = title_slide.shapes.title
    subtitle = title_slide.shapes.placeholders[1]  # subtitle placeholder
    title.text = "AI Script Reader"
    subtitle.text = "AI가 분석한 대본입니다."

    for i, text in enumerate(slide_texts):
        try:
            logging.debug(f"슬라이드 {i+1}에 텍스트 추가")
            # slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide = prs.slides.add_slide(content_slide_layout)

            # 제목 추가
            title_text_frame = slide.shapes.title.text_frame
            title_text_frame.clear()  # 기존 텍스트 제거
            title_para = title_text_frame.paragraphs[0]
            title_para.text = f"#{i + 1}"  # 슬라이드 번호
            title_para.font.size = Pt(font_size)
            title_para.font.name = font_name

            # 텍스트 박스 추가
            body_shape = slide.shapes.placeholders[1]
            text_frame = body_shape.text_frame
            text_frame.text = text
            set_text_box_style(text_frame, font_size, font_name)

            # add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER, max_chars_per_line_in_ppt, font_name)
            if split_flags[i]:
                add_check_needed_shape(slide)  # 슬라이드 번호 인자 제거
            if i == total_slides - 1:
                add_end_mark(slide)
        except Exception as e:
            st.error(f"오류: 슬라이드 생성 실패 (슬라이드 {i+1}): {e}")
            return None

    return prs

# 6. 함수 정의 (슬라이드 요소 추가)
def set_text_box_style(text_frame, font_size, font_name):
    """텍스트 박스 스타일을 설정합니다."""
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 상단 정렬
    for paragraph in text_frame.paragraphs:
        paragraph.font.name = font_name
        paragraph.font.size = Pt(font_size)
        paragraph.alignment = PP_ALIGN.LEFT

def add_text_to_slide(slide, text, font_size, alignment, max_chars_per_line, font_name):
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
            p.font.name = font_name
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = alignment
            p.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

        text_frame.auto_size = None
        logging.debug(f"텍스트 추가됨")
    except Exception as e:
        st.error(f"오류: 슬라이드에 텍스트 추가 중 오류 발생: {e}")
        raise

def add_end_mark(slide, font_name="Noto Color Emoji"):
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
    p.font.name = font_name
    p.font.color.rgb = RGBColor(255, 255, 255)
    end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

def add_check_needed_shape(slide, font_name="Noto Color Emoji"):
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
    p.font.name = font_name
    p.font.color.rgb = RGBColor(0, 0, 0)
    check_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

# 7. Streamlit UI
def main():
    st.title("AI Script Reader")

    # 기능 선택
    app_mode = st.sidebar.selectbox(
        "기능 선택",
        ["기존 PPT 생성", "새로운 PPT 생성 (AI 유사도 분석)"]
    )

    if app_mode == "기존 PPT 생성":
        st.header("기존 PPT 생성 기능")
        if not legacy_libs_available:
            st.error("기존 PPT 생성 기능에 필요한 라이브러리가 설치되지 않았습니다. requirements.txt 를 설치하세요.")
            return

        uploaded_file = st.file_uploader("Word 파일을 업로드하세요", type=["docx"])

        if uploaded_file is not None:
            text_paragraphs = extract_text_from_word(uploaded_file.name)
            if text_paragraphs is None:
                return  # 오류 발생 시 여기서 종료

            full_text = " ".join(text_paragraphs)  # 모든 단락을 합쳐서 유사도 분석에 사용

            st.subheader("PPT 슬라이드 생성 설정")
            max_lines_per_slide = st.slider("슬라이드 당 최대 줄 수", 1, 20, 10)
            st.session_state.max_chars_slider_ppt = st.slider("PPT 한 줄 최대 문자 수", 20, 100, 40)
            st.session_state.font_size_slider = st.slider("PPT 폰트 크기", 10, 80, 54)
            font_choice = st.selectbox("PPT 폰트 선택", ["맑은 고딕", "나눔고딕", "Arial"])  # 폰트 선택 추가

            if st.button("PPT 생성"):
                slide_texts, split_flags, slide_numbers = split_text_into_slides_with_similarity(
                    text_paragraphs,
                    max_lines_per_slide,
                    max_chars_per_line_ppt=st.session_state.max_chars_slider_ppt
                )
                ppt = None  # ppt 변수를 미리 선언
                divided_slide_count = 0  # 분할된 슬라이드 수 초기화
                try:
                    ppt = create_ppt(
                        slide_texts, split_flags,
                        max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
                        font_size=st.session_state.font_size_slider,
                        font_name=font_choice
                    )
                    divided_slide_count = sum(split_flags)  # 분할된 슬라이드 수 계산
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

                    # 분할된 슬라이드 정보 표시 (슬라이드 수, 번호)
                    st.subheader("생성 결과")
                    st.write(f"총 {len(slide_texts)}개의 슬라이드가 생성되었습니다.")
                    if divided_slide_count > 0:
                        divided_slide_numbers = [i + 1 for i, flag in enumerate(split_flags) if flag == 1]
                        st.write(f"이 중 {divided_slide_count}개의 슬라이드가 분할되었습니다. (분할된 슬라이드 번호: {divided_slide_numbers})")

    elif app_mode == "새로운 PPT 생성 (AI 유사도 분석)":
        st.header("새로운 PPT 생성 기능 (AI 유사도 분석)")
        # gensim 관련 코드는 별도의 함수 또는 파일로 분리
        show_gensim_functionality()

# 8. gensim 관련 기능 (별도 함수로 분리)
def show_gensim_functionality():
    """gensim 관련 기능을 처리하는 함수"""

    try:
        from gensim.models import Word2Vec
        from gensim.utils import simple_preprocess
        from gensim.similarities import Similarity
        import numpy as np
        import requests  # download_file 함수에서 사용
        gensim_libs_available = True
    except ImportError:
        st.error("새로운 PPT 생성 기능에 필요한 라이브러리가 설치되지 않았습니다. requirements_gensim.txt 를 설치하세요.")
        return

    # ko.bin 파일 다운로드 URL (실제 URL로 변경)
    ko_bin_url = "https://drive.google.com/uc?id=1SYB0v_qbww78TTv8WnW5FvggzU7XigdE"  # 여기에 실제 URL을 넣으세요!
    local_ko_bin_path = "ko.bin"

    # ko.bin 파일 다운로드 및 Word2Vec 모델 로드
    if download_file(ko_bin_url, local_ko_bin_path):
        model = load_word2vec_model(local_ko_bin_path)
    else:
        st.error("ko.bin 파일 로드 실패. 프로그램을 종료합니다.")
        return

    if model is None:
        st.error("Word2Vec 모델 로드 실패. 프로그램을 종료합니다.")
        return

    # 파일 업로드
    uploaded_file = st.file_uploader("Word 파일을 업로드하세요", type=["docx"], key="gensim_uploader")  # key 추가

    if uploaded_file is not None:
        # 파일 처리 및 텍스트 추출
        text_paragraphs = extract_text_from_word(uploaded_file.name)
        if text_paragraphs is None:
            return  # 오류 발생 시 여기서 종료

        full_text = " ".join(text_paragraphs)  # 모든 단락을 합쳐서 유사도 분석에 사용
        sentences = smart_sentence_split(full_text)

        # 텍스트 유사도 분석
        st.subheader("텍스트 유사도 분석")
        compare_text = st.text_area("비교할 텍스트를 입력하세요 (선택 사항)", "", key="gensim_compare")  # key 추가

        if compare_text:
            similarity = calculate_similarity(full_text, compare_text, model)
            st.write(f"입력된 텍스트와의 유사도: {similarity:.2f}")

        # 슬라이드 생성 설정
        st.subheader("PPT 슬라이드 생성 설정")
        max_chars_per_slide = st.slider("슬라이드 당 최대 문자 수", 100, 1000, 400, key="gensim_max_chars")  # key 추가
        st.session_state.max_chars_slider_ppt = st.slider("PPT 한 줄 최대 문자 수", 20, 100, 60, key="gensim_max_chars_ppt")  # key 추가
        st.session_state.font_size_slider = st.slider("PPT 폰트 크기", 10, 30, 18, key="gensim_font_size")  # key 추가
        font_choice = st.selectbox("PPT 폰트 선택", ["맑은 고딕", "나눔고딕", "Arial"], key="gensim_font_choice")  # 폰트 선택 추가

        # 슬라이드 분할 제안 및 생성
        split_flags = [0] * len(sentences)  # 기본적으로 분할하지 않음
        suggested_breaks = suggest_slide_breaks(sentences, max_chars_per_slide)

        st.subheader("슬라이드 분할 지점 (선택)")
        for i, sentence in enumerate(sentences):
            if i in suggested_breaks:
                if st.checkbox(f"#{i + 1} 슬라이드 분할", key=f"gensim_split_{i}"):  # key 추가
                    split_flags[i] = 1

        if st.button("PPT 생성", key="gensim_generate_button"):  # key 추가
            ppt = None  # ppt 변수를 미리 선언
            divided_slide_count = 0  # 분할된 슬라이드 수 초기화
            try:
                ppt = create_ppt(
                    sentences, split_flags,
                    max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
                    font_size=st.session_state.font_size_slider,
                    font_name=font_choice  # 선택된 폰트 적용
                )
                divided_slide_count = sum(split_flags)  # 분할된 슬라이드 수 계산
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

                # 분할된 슬라이드 정보 표시 (슬라이드 수, 번호)
                st.subheader("생성 결과")
                st.write(f"총 {len(sentences)}개의 슬라이드가 생성되었습니다.")
                if divided_slide_count > 0:
                    divided_slide_numbers = [i + 1 for i, flag in enumerate(split_flags) if flag == 1]
                    st.write(f"이 중 {divided_slide_count}개의 슬라이드가 분할되었습니다. (분할된 슬라이드 번호: {divided_slide_numbers})")

if __name__ == "__main__":
    main()