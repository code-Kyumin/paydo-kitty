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

# 2. 함수 정의 (Word 파일 처리)
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

# 3. 함수 정의 (텍스트 처리)
def calculate_text_lines(text, max_chars_per_line):
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
    """문맥을 고려하여 더 자연스럽게 문장을 분할합니다."""
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
    return [s.strip() for s in sentences if s.strip()]

def smart_sub_split(sentence):
    """더 복잡한 문장 구조를 고려하여 하위 문장으로 분리합니다."""
    sub_sentences = re.split(r',\s*(그리고|그러나|왜냐하면|예를 들어|즉|또는)\s+', sentence)
    return sub_sentences

# 4. 함수 정의 (AI 기반 슬라이드 분할)
def split_and_group_text_with_embeddings(
    text, max_lines_per_slide, max_chars_per_line_ppt,
    similarity_threshold=0.85, max_slide_length=100
):
    """문장 임베딩을 사용하여 문맥을 고려하며 텍스트를 슬라이드로 분할/그룹화합니다."""

    slides = []
    split_flags = []
    slide_numbers = []
    sentences, embeddings = get_sentence_embeddings(text)
    current_slide_text = ""
    current_slide_lines = 0
    current_slide_length = 0
    is_forced_split = False
    slide_number = 1

    for i, sentence in enumerate(sentences):
        sentence = sentence.strip()
        sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)
        sentence_length = len(sentence)
        current_slide_length_with_spaces = len(current_slide_text) if current_slide_text else 0
        sentence_length_with_spaces = len(sentence)

        if not slides:
            slides.append(sentence)
            split_flags.append(is_forced_split)
            slide_numbers.append(slide_number)
            current_slide_text = sentence
            current_slide_lines = sentence_lines
            current_slide_length = sentence_length
        elif (
            current_slide_lines + sentence_lines <= max_lines_per_slide
            and current_slide_length_with_spaces + sentence_length_with_spaces <= max_slide_length
        ):
            if i > 0:
                similarity = util.cos_sim(embeddings[i - 1], embeddings[i])[0][0].item()
                if similarity < similarity_threshold:
                    slides.append(sentence)
                    split_flags.append(True)
                    slide_numbers.append(++slide_number)
                    current_slide_text = sentence
                    current_slide_lines = sentence_lines
                    current_slide_length = sentence_length
                    is_forced_split = True
                else:
                    slides[-1] += " " + sentence
                    split_flags[-1] = is_forced_split
                    slide_numbers[-1] = slide_number
                    current_slide_text += " " + sentence
                    current_slide_lines += sentence_lines
                    current_slide_length += sentence_length
            else:
                split_point = -1
                if ", " in current_slide_text:
                    split_point = current_slide_text.rfind(", ")
                elif ". " in current_slide_text:
                    split_point = current_slide_text.rfind(". ")

                if split_point != -1:
                    slides.append(current_slide_text[:split_point])
                    slides.append(current_slide_text[split_point + 2:] + " " + sentence)
                    split_flags.extend([True, True])
                    slide_numbers.extend([slide_number, ++slide_number])
                    current_slide_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)
                    current_slide_length = sentence_length
                    is_forced_split = True
                else:
                    slides.append(sentence)
                    split_flags.append(True)
                    slide_numbers.append(++slide_number)
                    current_slide_text = sentence
                    current_slide_lines = sentence_lines
                    current_slide_length = sentence_length
                    is_forced_split = True
        # st.write(f"Slide {slide_number}: {slides[-1]}")  # 디버깅: 슬라이드 내용 확인

    final_slides_result = [slide for slide in slides if slide.strip()]
    final_split_flags_result = split_flags[:len(final_slides_result)]
    final_slide_numbers_result = slide_numbers[:len(final_slides_result)]

    return final_slides_result, final_split_flags_result, final_slide_numbers_result

# 5. 함수 정의 (PPT 생성 및 슬라이드 조작)
def create_ppt(slide_texts, split_flags, slide_numbers, max_chars_per_line_in_ppt=18, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)

    for i, text in enumerate(slide_texts):
        try:
            print(f"Adding text to slide {i+1}: {text[:50]}...")  # 디버깅용 출력
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER)
            add_slide_number(slide, slide_numbers[i], total_slides)
            if split_flags[i] and calculate_text_lines(text, max_chars_per_line_in_ppt) == 1:
                add_check_needed_shape(slide, slide_numbers[i], slide_numbers[i])
            if i == total_slides - 1:
                add_end_mark(slide)
        except Exception as e:
            st.error(f"Error creating slide {i+1}: {e}")
            return None

    return prs

def add_text_to_slide(slide, text, font_size, alignment):
    try:
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        text_frame.word_wrap = True

        wrapped_lines = textwrap.wrap(text, width=18, break_long_words=True)
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

        text_frame.auto_size = None
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    except Exception as e:
        st.error(f"Error adding text to slide: {e}")
        raise

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

def add_check_needed_shape(slide, slide_number, ui_slide_number):
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
    p.text = f"확인 필요! (슬라이드 {ui_slide_number})"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    text_frame = check_shape.text_frame
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

# 6. Streamlit UI
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
st.title("🎬 AI 기반 촬영 대본 PPT 자동 생성기")

# Word 파일 업로드 기능 추가
uploaded_file = st.file_uploader("📝 Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력:", height=300, key="text_input_area")

# UI 입력 슬라이더
max_lines_per_slide_input = st.slider(
    "📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider"
)
max_chars_per_line_ppt_input = st.slider(
    "📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=10, max_value=100, value=18, key="max_chars_slider_ppt"
)
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")

similarity_threshold_input = st.slider(
    "📚 문맥 유지 민감도:",
    min_value=0.0, max_value=1.0, value=0.85, step=0.05,
    help="""
    이 값보다 낮은 문맥 유사도를 가지는 문장 사이에서 슬라이드를 나누는 것을 고려합니다.
    1.0에 가까울수록 문맥을 최대한 유지하며 슬라이드를 분할합니다 (강의용에 적합).
    0.0에 가까울수록 슬라이드를 더 짧게 나누어 가독성을 높입니다 (발표용에 적합).
    """,
    key="similarity_threshold_input"
)

# 7. PPT 생성 및 다운로드
if st.button("🚀 AI 기반 PPT 만들기", key="create_ppt_button"):
    text = ""
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    with st.spinner("PPT 생성 중..."):
        slide_texts, split_flags, slide_numbers = split_and_group_text_with_embeddings(
            text, max_lines_per_slide=st.session_state.max_lines_slider,
            max_chars_per_line_ppt=st.session_state.max_chars_slider_ppt,
            similarity_threshold=st.session_state.similarity_threshold_input,
            max_slide_length=st.session_state.max_chars_slider_ppt
        )
        ppt = create_ppt(
            slide_texts, split_flags, slide_numbers,
            max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
            font_size=st.session_state.font_size_slider
        )

    if ppt:
        ppt_io = io.BytesIO()
        ppt.save(ppt_io)
        ppt_io.seek(0)

        st.download_button(
            label="📥 PPT 다운로드",
            data=ppt_io,
            file_name="paydo_script_ai.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )