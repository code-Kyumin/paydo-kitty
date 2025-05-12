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
from sentence_transformers import SentenceTransformer, util  # Sentence Transformers 임포트

# Word 파일에서 텍스트 추출하는 함수
def extract_text_from_word(file_path):
    """Word 파일에서 모든 텍스트를 추출하여 하나의 문자열로 반환합니다."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        return "\n".join(full_text)
    except FileNotFoundError:
        st.error(f"Error: Word file not found at {file_path}")
        return None
    except docx.exceptions.PackageNotFoundError:
        st.error(f"Error: Invalid Word file at {file_path}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred while processing the Word file: {e}")
        return None

# 문장이 차지할 줄 수 계산
def calculate_text_lines(text, max_chars_per_line):
    lines = 0
    paragraphs = text.split('\n')
    for paragraph in paragraphs:
        if not paragraph:
            lines += 1
        else:
            lines += len(textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=True))
    return lines

# 문장 임베딩 생성 함수
def get_sentence_embeddings(text, model_name='paraphrase-multilingual-mpnet-base-v2'):
    """
    텍스트에서 문장 임베딩을 추출합니다.

    Args:
        text (str): 입력 텍스트.
        model_name (str, optional): 사용할 Sentence Transformers 모델 이름.
                                     기본값은 'paraphrase-multilingual-mpnet-base-v2'입니다.

    Returns:
        tuple: (문장 리스트, 임베딩 벡터 리스트)
    """
    model = SentenceTransformer(model_name)
    sentences = text.split('\n')  # 문장 단위로 분리 (일단은 개행 기준으로 분리)
    embeddings = model.encode(sentences)
    return sentences, embeddings

# 텍스트를 슬라이드로 분할 및 그룹화 (AI 기반)
def split_and_group_text_with_embeddings(text, max_lines_per_slide, max_chars_per_line_ppt, similarity_threshold=0.75):
    """
    문장 임베딩을 사용하여 문맥을 고려하며 텍스트를 슬라이드로 분할 및 그룹화합니다.

    Args:
        text (str): 입력 텍스트.
        max_lines_per_slide (int): 슬라이드당 최대 줄 수.
        max_chars_per_line_ppt (int): PPT 한 줄당 최대 문자 수.
        similarity_threshold (float, optional): 문장 간 유사도 임계값.
                                             이 값보다 낮으면 슬라이드를 나누는 것을 고려합니다.

    Returns:
        tuple: (분할된 텍스트 슬라이드 리스트, 각 슬라이드가 강제로 분할되었는지 여부를 나타내는 불리언 리스트)
    """

    slides = []
    split_flags = []
    sentences, embeddings = get_sentence_embeddings(text)
    current_slide_text = ""
    current_slide_lines = 0
    is_forced_split = False

    for i, sentence in enumerate(sentences):
        sentence = sentence.strip()
        sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

        if not slides:
            slides.append(sentence)
            split_flags.append(is_forced_split)
            current_slide_text = sentence
            current_slide_lines = sentence_lines
        elif current_slide_lines + sentence_lines <= max_lines_per_slide:
            if i > 0:
                similarity = util.cos_sim(embeddings[i-1], embeddings[i])[0][0].item()
                if similarity < similarity_threshold:
                    slides.append(sentence)
                    split_flags.append(True)
                    current_slide_text = sentence
                    current_slide_lines = sentence_lines
                else:
                    slides[-1] += "\n" + sentence
                    split_flags[-1] = is_forced_split
                    current_slide_text += "\n" + sentence
                    current_slide_lines += sentence_lines
            else:
                slides[-1] += "\n" + sentence
                split_flags[-1] = is_forced_split
                current_slide_text += "\n" + sentence
                current_slide_lines += sentence_lines
        else:
            slides.append(sentence)
            split_flags.append(True)
            current_slide_text = sentence
            current_slide_lines = sentence_lines

    final_slides_result = []
    final_split_flags_result = []

    for i, slide_text in enumerate(slides):
        if calculate_text_lines(slide_text, max_chars_per_line_ppt) > max_lines_per_slide:
            original_sentence = slide_text.replace('\n', ' ')
            sub_sentences = re.split(r'(?<=[.?!;])\s+', original_sentence.strip())
            temp_slide_text = ""
            temp_slide_lines = 0
            is_forced_split = False
            for sub_sentence in sub_sentences:
                sub_sentence = sub_sentence.strip()
                sub_sentence_lines = calculate_text_lines(sub_sentence, max_chars_per_line_ppt)
                if temp_slide_lines + sub_sentence_lines <= max_lines_per_slide:
                    if temp_slide_text:
                        temp_slide_text += " "
                    temp_slide_text += sub_sentence
                    temp_slide_lines += sub_sentence_lines
                else:
                    final_slides_result.append(temp_slide_text)
                    final_split_flags_result.append(is_forced_split)
                    temp_slide_text = sub_sentence
                    temp_slide_lines = sub_sentence_lines
                    is_forced_split = False

            if temp_slide_text:
                if calculate_text_lines(temp_slide_text, max_chars_per_line_ppt) > max_lines_per_slide:
                    words = temp_slide_text.split()
                    segment = ""
                    for word in words:
                        if len(segment.replace(" ", "")) + len(word) + (1 if segment else 0) <= 60:
                            if segment:
                                segment += " "
                            segment += word
                        else:
                            final_slides_result.append(segment)
                            final_split_flags_result.append(True)
                            segment = word
                            is_forced_split = True
                    if segment:
                        final_slides_result.append(segment)
                        final_split_flags_result.append(True)
                else:
                    final_slides_result.append(temp_slide_text)
                    final_split_flags_result.append(False)
        else:
            final_slides_result.append(slide_text)
            final_split_flags_result.append(False)

    return final_slides_result, final_split_flags_result

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
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 상단 정렬 명시적으로 설정
    text_frame.word_wrap = True

    wrapped_lines = textwrap.wrap(text, width=18, break_long_words=True)  # 긴 단어 분리 활성화
    text_frame.clear()
    for line in wrapped_lines:
        p = text_frame.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.name = 'Noto Color Emoji'
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = alignment
        p.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

    # 텍스트 박스의 자동 맞춤 기능 제거 (상단 정렬에 영향 줄 수 있음)
    text_frame.auto_size = None
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
st.set_page_config(page_title="Paydo AI PPT", layout="centered")  # 페이지 제목 변경
st.title("🎬 AI 기반 촬영 대본 PPT 자동 생성기")  # 타이틀 변경

# Word 파일 업로드 기능 추가
uploaded_file = st.file_uploader("📝 Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
max_chars_per_line_ppt_input = st.slider("📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=30, value=18, key="max_chars_slider_ppt")
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")

similarity_threshold_input = st.slider(
    "📚 문장 나누기 민감도:",
    min_value=0.0,
    max_value=1.0,
    value=0.75,
    step=0.05,
    help="""
    이 값을 조절하여 문장을 얼마나 세밀하게 나눌지 결정합니다.
    - 값을 낮추면 (0.0에 가깝게) 문장을 더 잘게 쪼개서 슬라이드를 많이 만듭니다.
    - 값을 높이면 (1.0에 가깝게) 문장을 덜 쪼개서 슬라이드를 적게 만듭니다.
    """
)

if st.button("🚀 AI 기반 PPT 만들기", key="create_ppt_button"):  # 버튼 텍스트 변경
    text = ""
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    slide_texts, split_flags = split_and_group_text_with_embeddings(  # AI 기반 함수 호출
        text,
        max_lines_per_slide=max_lines_per_slide_input,
        max_chars_per_line_ppt=max_chars_per_line_ppt_input,
        similarity_threshold=similarity_threshold_input  # UI에서 입력받은 값 사용
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
            file_name="paydo_script_ai.pptx",  # 파일 이름에 "ai" 추가
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="download_button"
        )
        if any(split_flags):
            split_slide_numbers = [i + 1 for i, flag in enumerate(split_flags) if flag]
            st.warning(f"❗️ 일부 슬라이드({split_slide_numbers})는 문장 유사도에 따라 분할되었습니다. PPT를 확인하여 가독성을 검토해주세요.")  # 경고 메시지 수정
    else:
        st.error("❌ PPT 생성에 실패했습니다.")