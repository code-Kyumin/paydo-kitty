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

# 로깅 설정 (디버깅용)
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 2. 함수 정의 (Word 파일 처리)
def extract_text_from_word(file_path):
    """Word 파일에서 모든 텍스트를 추출하여 하나의 문자열로 반환합니다."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        text = "\n".join(full_text)
        logging.debug(f"Word text extracted: {text[:100]}...")  # Log 추출된 텍스트
        return text
    except FileNotFoundError:
        st.error(f"Error: Word file not found.")
        return None
    except docx.exceptions.PackageNotFoundError:
        st.error(f"Error: Invalid Word file.")
        return None
    except Exception as e:
        st.error(f"Error: Word processing error: {e}")
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
    # TODO: 더 강력한 문장 분할 로직 구현 또는 라이브러리 사용 검토
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
    logging.debug(f"Sentences split: {len(sentences)} sentences")  # Log 분할된 문장 수
    return sentences

def smart_sub_split(sentence):
    """더 복잡한 문장 구조를 고려하여 하위 문장으로 분리합니다."""
    # TODO: 하위 문장 분할 로직 개선 (필요한 경우)
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

    logging.debug(f"Total slides: {len(final_slides_result)}")  # Log 최종 슬라이드 개수
    return final_slides_result, final_split_flags_result, final_slide_numbers_result

# 5. 함수 정의 (PPT 생성 및 슬라이드 조작)
def create_ppt(slide_texts, split_flags, slide_numbers, max_chars_per_line_in_ppt=18, font_size=54):
    """슬라이드 텍스트를 기반으로 PPT를 생성하고, 슬라이드 번호, '끝' 표시 등을 추가합니다."""

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)

    for i, text in enumerate(slide_texts):
        try:
            logging.debug(f"Adding text to slide {i+1}")
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER, max_chars_per_line_in_ppt)
            add_slide_number(slide, slide_numbers[i], total_slides)
            if split_flags[i] and calculate_text_lines(text, max_chars_per_line_in_ppt) == 1:
                add_check_needed_shape(slide, slide_numbers[i], slide_numbers[i])
            if i == total_slides - 1:
                add_end_mark(slide)
        except Exception as e:
            st.error(f"Error: Slide creation failed (slide {i+1}): {e}")
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

        # TODO: textwrap.wrap 대신 다른 텍스트 래핑 방식 고려
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
        logging.debug(f"Text added")
    except Exception as e:
        st.error(f"Error: Adding text to slide: {e}")
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
    p.text = f"Check needed (slide {ui_slide_number})"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    text_frame = check_shape.text_frame
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

# 7. Streamlit UI
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
st.title("🎬 AI PPT Generator")

# Word file upload
uploaded_file = st.file_uploader("Upload Word file", type=["docx"])

text_input = st.text_area("Or enter text directly", height=300, key="text_input_area")

# UI input sliders
max_lines_per_slide_input = st.slider(
    "Max lines per slide", min_value=1, max_value=10, value=5, key="max_lines_slider"
)
max_chars_per_line_ppt_input = st.slider(
    "Max chars per line (PPT)", min_value=10, max_value=100, value=18, key="max_chars_slider_ppt"
)
font_size_input = st.slider("Font size", min_value=10, max_value=60, value=54, key="font_size_slider")

similarity_threshold_input = st.slider(
    "Context similarity threshold",
    min_value=0.0, max_value=1.0, value=0.85, step=0.05, # 쉼표 추가
    help="""
    Consider splitting slides between sentences with lower context similarity.
    Lower values create shorter, more readable slides (e.g., for presentations).
    Higher values maintain context for longer, information-dense slides (e.g., for lectures).
    """,
    key="similarity_threshold_input" # 이 부분은 수정되지 않도록 해줘.
)

# 8. PPT generation and download
if st.button("Generate PPT"):
    text = ""
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Please upload a Word file or enter text.")
        st.stop()

    with st.spinner("Generating PPT..."):
        try:
            slide_texts, split_flags, slide_numbers = split_and_group_text_with_embeddings(
                text,
                max_lines_per_slide=st.session_state.max_lines_slider,
                max_chars_per_line_ppt=st.session_state.max_chars_slider_ppt,
                similarity_threshold=st.session_state.similarity_threshold_input,
                max_slide_length=100
            )
            ppt = create_ppt(
                slide_texts, split_flags, slide_numbers,
                max_chars_per_line_in_ppt=st.session_state.max_chars_slider_ppt,
                font_size=st.session_state.font_size_slider
            )
        except Exception as e:
            st.error(f"Error: PPT generation failed: {e}")
            st.error(f"Error details: {str(e)}")
            st.stop()

    if ppt:
        ppt_io = io.BytesIO()
        try:
            ppt.save(ppt_io)
            ppt_io.seek(0)
        except Exception as e:
            st.error(f"Error: PPT save failed: {e}")
            st.error(f"Error details: {str(e)}")
        else:
            st.download_button(
                label="Download PPT",
                data=ppt_io,
                file_name="paydo_script_ai.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )