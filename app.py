import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap
import docx  # python-docx 라이브러리 추가
from konlpy.tag import Kkma  # KoNLPy에서 Kkma 형태소 분석기 임포트

# Word 파일에서 텍스트 추출하는 함수
def extract_text_from_word(file_path):
    """Word 파일에서 모든 텍스트를 추출하여 하나의 문자열로 반환합니다."""

    doc = docx.Document(file_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return "\n".join(full_text)

# 문장이 차지할 줄 수 계산 (단어 잘림 방지)
def calculate_text_lines(text, max_chars_per_line):
    """주어진 텍스트가 지정된 최대 문자 길이를 기준으로 몇 줄을 차지하는지 계산합니다."""

    lines = 0
    if not text:
        return lines

    words = text.split()
    current_line_length = 0
    lines += 1  # 최소 1줄
    for word in words:
        word_length = len(word)
        if current_line_length + word_length + 1 <= max_chars_per_line:
            current_line_length += word_length + 1
        else:
            lines += 1
            current_line_length = word_length
            
    return lines

def split_text_into_slides_konlpy(text, max_lines_per_slide, max_chars_per_line_in_ppt):
    """KoNLPy를 사용하여 입력 텍스트를 슬라이드에 맞게 분할하고, 각 슬라이드가 원본 문장인지 여부를 반환합니다."""

    kkma = Kkma()
    slides = []
    original_sentence_flags = []
    current_slide_text = ""
    current_slide_lines = 0
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    seen_sentences = set() # [추가] 중복 문장 제거를 위한 set

    for sentence in sentences:
        sentence = sentence.strip()
        if sentence in seen_sentences: # [추가] 중복 문장 확인
            continue
        seen_sentences.add(sentence) # [추가] 처리한 문장 저장

        lines_needed = calculate_text_lines(sentence, max_chars_per_line_in_ppt)

        if current_slide_lines + lines_needed <= max_lines_per_slide:
            if current_slide_text:
                current_slide_text += " "
            current_slide_text += sentence
            current_slide_lines += lines_needed
            original_sentence_flags.append(True)  # 원래 문장
        else:
            # 슬라이드 분할 로직 (KoNLPy 활용)
            split_points = []
            pos_result = kkma.pos(current_slide_text + " " + sentence)  # 형태소 분석
            for i, (word, pos) in enumerate(pos_result):
                # 조사나 어미 앞, 접속사 뒤에서 분할 시도
                if pos.startswith("J") or pos.startswith("E") or (
                    pos == "MA" and word in ["그리고", "그러나", "그래서"]
                ):
                    split_points.append(i)

            if split_points:
                # 분할 가능한 지점 중, 현재 슬라이드에 가장 적합한 지점 선택
                best_split_idx = max(
                    (
                        idx
                        for idx in split_points
                        if calculate_text_lines(
                            "".join(p[0] for p in pos_result[:idx]),
                            max_chars_per_line_in_ppt,
                        )
                        <= max_lines_per_slide
                    ),
                    default=0,
                )
                if best_split_idx > 0:
                    split_text = "".join(p[0] for p in pos_result[:best_split_idx]).strip()
                    if split_text:
                        slides.append(split_text)
                    current_slide_text = (
                        "".join(p[0] for p in pos_result[best_split_idx:]).strip()
                    )
                    current_slide_lines = calculate_text_lines(
                        current_slide_text, max_chars_per_line_in_ppt
                    )
                    original_sentence_flags.append(
                        False
                    )  # [수정] 분할된 문장으로 표시
                else:
                    # 분할 가능한 지점이 없으면, 단어 단위로 분할
                    slides.append(current_slide_text.strip())
                    current_slide_text = sentence
                    current_slide_lines = lines_needed
                    original_sentence_flags.append(
                        False
                    )  # [수정] 분할된 문장으로 표시
            else:
                # 분할 가능한 지점이 없으면, 단어 단위로 분할
                slides.append(current_slide_text.strip())
                current_slide_text = sentence
                current_slide_lines = lines_needed
                original_sentence_flags.append(
                    False
                )  # [수정] 분할된 문장으로 표시
        # [수정] 다음 슬라이드를 위해 초기화
        current_slide_text = current_slide_text.strip()
        if current_slide_text:
            current_slide_text += " "

    if current_slide_text:
        slides.append(current_slide_text.strip())
        original_sentence_flags.append(True)  # 원래 문장

    return slides, original_sentence_flags

def create_powerpoint(slides, original_sentence_flags, max_chars_per_line_in_ppt, max_lines_per_slide, font_size):
    """분할된 텍스트 슬라이드와 문장 분할 정보를 바탕으로 PowerPoint 프레젠테이션을 생성합니다."""

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    check_needed_slides = []  # 확인이 필요한 슬라이드 번호 저장

    for i, text in enumerate(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6번 레이아웃 (빈 슬라이드) 사용
        add_text_to_slide(slide, text, font_size)
        add_slide_number(slide, i + 1, len(slides))
        
        # 분할된 문장인 경우 '확인 필요' 도형 추가
        if not original_sentence_flags[i]:
            add_check_needed_shape(slide)
            check_needed_slides.append(i + 1)
            
        # 마지막 슬라이드인 경우 '끝' 표시 추가
        if i == len(slides) - 1:
            add_end_mark(slide)

    return prs, check_needed_slides

def add_text_to_slide(slide, text, font_size):
    """슬라이드에 텍스트를 추가하고 서식을 설정합니다."""

    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    text_frame = textbox.text_frame
    text_frame.clear()  # 기존 텍스트 제거
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 텍스트 상단 정렬
    text_frame.word_wrap = True

    paragraph = text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.size = Pt(font_size)
    paragraph.font.name = 'Noto Color Emoji'
    paragraph.font.bold = True
    paragraph.font.color.rgb = RGBColor(0, 0, 0)
    paragraph.alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE # 중앙 정렬

def add_slide_number(slide, current, total):
    """슬라이드에 페이지 번호를 추가합니다."""

    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_text_frame = footer_box.text_frame
    footer_text_frame.clear()
    paragraph = footer_text_frame.paragraphs[0]
    paragraph.text = f"{current} / {total}"
    paragraph.font.size = Pt(18)
    paragraph.font.name = '맑은 고딕'
    paragraph.font.color.rgb = RGBColor(128, 128, 128)
    paragraph.alignment = PP_ALIGN.RIGHT

def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가합니다."""

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
    paragraph = end_text_frame.paragraphs[0]
    paragraph.text = "끝"
    paragraph.font.size = Pt(36)
    paragraph.font.color.rgb = RGBColor(255, 255, 255)
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    paragraph.alignment = PP_ALIGN.CENTER

def add_check_needed_shape(slide):
    """슬라이드에 '확인 필요' 표시를 추가합니다."""

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
    paragraph = check_text_frame.paragraphs[0]
    paragraph.text = "확인 필요!"
    paragraph.font.size = Pt(18)
    paragraph.font.bold = True
    paragraph.font.color.rgb = RGBColor(0, 0, 0)
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    paragraph.alignment = PP_ALIGN.CENTER

# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기 (KoNLPy)")

# Word 파일 업로드 기능 추가
uploaded_file = st.file_uploader("📝 Word 파일 업로드", type=["docx"])

text_input = st.text_area("또는 텍스트 직접 입력:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
max_chars_per_line_ppt_input = st.slider("📏 한 줄당 최대 글자 수 (PPT 표시):", min_value=3, max_value=30, value=18, key="max_chars_slider_ppt")
min_chars_per_line_input = st.slider("🔤 한 줄당 최소 글자 수:", min_value=1, max_value=10, value=4, key="min_chars_slider")
font_size_input = st.slider("🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider")

if st.button("🚀 PPT 만들기", key="create_ppt_button"):
    if uploaded_file is not None:
        text = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        text = text_input
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    slide_texts, original_sentence_flags = split_text_into_slides_konlpy(
        text,
        max_lines_per_slide=max_lines_per_slide_input,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input
    )
    ppt, check_needed_slides = create_powerpoint(
        slide_texts,
        original_sentence_flags,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
        max_lines_per_slide=max_lines_per_slide_input,
        font_size=font_size_input
    )

    if ppt:
        ppt_io = io.BytesIO()
        ppt.save(ppt_io)
        ppt_io.seek(0)

        st.download_button(
            label="📥 PPT 다운로드",
            data=ppt_io,
            file_name="paydo_script_konlpy.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="download_button"
        )
        if check_needed_slides:
            st.warning(f"❗️ 일부 슬라이드({check_needed_slides})는 한 문장이 너무 길어 분할되었습니다. PPT를 확인하여 가독성을 검토해주세요.")
    else:
        st.error("❌ PPT 생성에 실패했습니다.")