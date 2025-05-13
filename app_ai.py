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
from sentence_transformers import SentenceTransformer, util
import logging

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 사용할 한국어 특화 모델
model_name = 'jhgan/ko-sroberta-multitask'

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
def calculate_similarity(text1, text2):
    """두 텍스트 간의 유사도를 계산합니다."""

    model = SentenceTransformer(model_name)
    embedding1 = model.encode(text1, convert_to_tensor=True)
    embedding2 = model.encode(text2, convert_to_tensor=True)

    cosine_similarity = util.pytorch_cos_sim(embedding1, embedding2)
    return cosine_similarity.item()

def split_text_into_sentences(text):
    """텍스트를 문장 단위로 분리합니다."""
    # 마침표, 물음표, 느낌표 뒤에 오는 공백을 기준으로 분리
    sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!)\s', text)
    sentences = [s.strip() for s in sentences if s.strip()]  # 빈 문자열 제거
    logging.debug(f"Sentences split: {len(sentences)} sentences")
    return sentences

def suggest_slide_breaks(sentences, max_chars_per_slide):
    """문장 리스트를 입력받아 슬라이드 분할 지점을 제안합니다."""

    slide_breaks = []
    current_slide_length = 0
    for i, sentence in enumerate(sentences):
        # 문장 길이 업데이트 (공백 포함)
        sentence_length = len(sentence) + 1
        # 현재 슬라이드에 문장 추가 가능 여부 확인
        if current_slide_length + sentence_length > max_chars_per_slide and current_slide_length > 0:
            slide_breaks.append(i)
            current_slide_length = sentence_length
        else:
            current_slide_length += sentence_length
    return slide_breaks

def create_ppt(sentences, split_flags, max_chars_per_line_in_ppt, font_size):
    """문장 리스트와 분할 플래그를 입력받아 PPTX 파일을 생성합니다."""

    prs = Presentation()
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

    # 텍스트 박스 스타일 정의
    def set_text_box_style(text_frame, font_size):
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 상단 정렬
        for paragraph in text_frame.paragraphs:
            paragraph.font.name = "맑은 고딕"
            paragraph.font.size = Pt(font_size)
            paragraph.alignment = PP_ALIGN.LEFT

    # 나머지 슬라이드 (내용 슬라이드)
    current_slide_text = ""
    slide_count = 1  # 슬라이드 번호
    original_slide_count = 1 # 원본 슬라이드 번호
    for i, sentence in enumerate(sentences):
        # 현재 슬라이드 텍스트에 문장 추가
        wrapped_sentence = textwrap.fill(sentence, width=max_chars_per_line_in_ppt)
        current_slide_text += wrapped_sentence + "\n"

        # 슬라이드 분할 조건 확인
        if i < len(split_flags) and split_flags[i] == 1:
            # 새 슬라이드 추가
            content_slide = prs.slides.add_slide(content_slide_layout)
            title_text_frame = content_slide.shapes.title.text_frame
            title_text_frame.clear()  # 기존 텍스트 제거
            title_para = title_text_frame.paragraphs[0]
            title_para.text = f"#{original_slide_count} → #{slide_count}"  # 슬라이드 번호
            title_para.font.size = Pt(font_size)
            title_para.font.name = "맑은 고딕"
            
            body_shape = content_slide.shapes.placeholders[1]
            text_frame = body_shape.text_frame
            text_frame.text = current_slide_text
            set_text_box_style(text_frame, font_size)
            current_slide_text = ""
            slide_count += 1
        original_slide_count += 1

    # 마지막 슬라이드 처리
    if current_slide_text:
        content_slide = prs.slides.add_slide(content_slide_layout)
        title_text_frame = content_slide.shapes.title.text_frame
        title_text_frame.clear()
        title_para = title_text_frame.paragraphs[0]
        title_para.text = f"#{original_slide_count} → #{slide_count}"
        title_para.font.size = Pt(font_size)
        title_para.font.name = "맑은 고딕"
        
        body_shape = content_slide.shapes.placeholders[1]
        text_frame = body_shape.text_frame
        text_frame.text = current_slide_text
        set_text_box_style(text_frame, font_size)

    return prs

# 4. Streamlit UI
def main():
    st.title("AI Script Reader")

    # 파일 업로드
    uploaded_file = st.file_uploader("Word 파일을 업로드하세요", type=["docx"])

    if uploaded_file is not None:
        # 파일 처리 및 텍스트 추출
        text_paragraphs = extract_text_from_word(uploaded_file)
        if text_paragraphs is None:
            return  # 오류 발생 시 여기서 종료

        full_text = " ".join(text_paragraphs)  # 모든 단락을 합쳐서 유사도 분석에 사용
        sentences = split_text_into_sentences(full_text)

        # 텍스트 유사도 분석
        st.subheader("텍스트 유사도 분석")
        compare_text = st.text_area("비교할 텍스트를 입력하세요 (선택 사항)", "")

        if compare_text:
            similarity = calculate_similarity(full_text, compare_text)
            st.write(f"입력된 텍스트와의 유사도: {similarity:.2f}")

        # 슬라이드 생성 설정
        st.subheader("PPT 슬라이드 생성 설정")
        max_chars_per_slide = st.slider("슬라이드 당 최대 문자 수", 100, 1000, 400)
        st.session_state.max_chars_slider_ppt = st.slider("PPT 한 줄 최대 문자 수", 20, 100, 60)
        st.session_state.font_size_slider = st.slider("PPT 폰트 크기", 10, 30, 18)

        # 슬라이드 분할 제안 및 생성
        split_flags = [0] * len(sentences)  # 기본적으로 분할하지 않음
        suggested_breaks = suggest_slide_breaks(sentences, max_chars_per_slide)

        st.subheader("슬라이드 분할 지점 (선택)")
        for i, sentence in enumerate(sentences):
            if i in suggested_breaks:
                if st.checkbox(f"#{i + 1} 슬라이드 분할", key=f"split_{i}"):
                    split_flags[i] = 1

        if st.button("PPT 생성"):
            ppt = None  # ppt 변수를 미리 선언
            divided_slide_count = 0  # 분할된 슬라이드 수 초기화
            try:
                ppt = create_ppt(
                    sentences, split_flags,
                    max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
                    font_size=st.session_state.font_size_slider
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