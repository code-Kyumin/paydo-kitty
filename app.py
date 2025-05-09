import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap

def split_text(text):
    """텍스트를 문장 단위로 분리합니다. ✂️"""
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences]

def is_sentence(text):
    """텍스트가 문장인지 확인합니다."""
    return bool(re.search(r'[.!?]', text))

def split_long_sentence(sentence, min_chars, max_chars_per_line):
    """긴 문장을 최소/최대 글자 수 기준에 맞게 분할합니다. 📏"""

    wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line, break_long_words=False)
    
    final_lines = []
    for line in wrapped_lines:
        if len(line) < min_chars:
            words = line.split()
            if len(words) > 1:
                mid_point = len(words) // 2
                final_lines.extend([" ".join(words[:mid_point]), " ".join(words[mid_point:])])
            else:
                final_lines.append(line)  # 분할할 수 없는 경우 그대로 추가
        else:
            final_lines.append(line)
    return final_lines

def group_sentences_to_slides(sentences, max_lines_per_slide, max_chars_per_line, min_chars):
    """문장들을 슬라이드에 맞게 그룹화합니다. 📦"""

    slides = []
    current_slide_text = ""
    current_line_count = 0

    for sentence in sentences:
        if not is_sentence(sentence):
            # 문장이 아니면 강제로 새 슬라이드
            if current_slide_text:
                slides.append(current_slide_text.strip())
            slides.append(sentence.strip())
            current_slide_text = ""
            current_line_count = 0
        else:
            lines = split_long_sentence(sentence, min_chars, max_chars_per_line)
            line_count = len(lines)

            if current_line_count + line_count > max_lines_per_slide and current_slide_text:
                slides.append(current_slide_text.strip())
                current_slide_text = "\n".join(lines)
                current_line_count = line_count
            else:
                if current_slide_text:
                    current_slide_text += "\n"  # Add newline between sentences
                current_slide_text += "\n".join(lines)
                current_line_count += line_count

    if current_slide_text:
        slides.append(current_slide_text.strip())
    
    return slides

def create_ppt(slide_texts, max_chars_per_line):
    """슬라이드 텍스트를 사용하여 PPT를 생성합니다. 📝"""

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)  # 총 슬라이드 수 계산

    for i, text in enumerate(slide_texts):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.33), Inches(6.5))
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        
        p = text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(54)
        p.font.name = '맑은 고딕'
        p.font.bold = True  # 텍스트 볼드 처리
        p.alignment = PP_ALIGN.CENTER
        
        # 페이지 번호 추가
        add_page_number(slide, i + 1, total_slides)

        if i == total_slides - 1:  # 마지막 슬라이드에 '끝' 표시
            add_end_mark(slide)
        
    return prs

def add_page_number(slide, current_idx, total_slides):
    """슬라이드에 페이지 번호를 추가합니다."""
    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_frame = footer_box.text_frame
    footer_frame.text = f"{current_idx} / {total_slides}"
    footer_p = footer_frame.paragraphs[0]
    footer_p.font.size = Pt(18)
    footer_p.font.name = '맑은 고딕'
    footer_p.font.color.rgb = RGBColor(128, 128, 128)
    footer_p.alignment = PP_ALIGN.RIGHT

def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가합니다."""

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
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기")  # UI 제목 변경
text_input = st.text_area("텍스트 입력: ✍️", height=200)

# Updated slider ranges and defaults
max_lines_per_slide = st.slider("최대 줄 수 (슬라이드 당) 📄", 1, 10, 4)
max_chars_per_line = st.slider("최대 글자 수 (줄 당) 🔡", 3, 20, 18)
min_chars = st.slider("최소 글자 수 📏", 1, 10, 3)  # Slider for minimum characters, range 1-10

if st.button("PPT 생성 🚀"):
    if text_input:
        sentences = split_text(text_input)
        slide_texts = group_sentences_to_slides(sentences, max_lines_per_slide, max_chars_per_line, min_chars)
        prs = create_ppt(slide_texts, max_chars_per_line)
        
        ppt_bytes = io.BytesIO()
        prs.save(ppt_bytes)
        ppt_bytes.seek(0)
        
        st.download_button(label="PPT 다운로드 📥", data=ppt_bytes, file_name="paydo_kitty_script.pptx")  # 파일명 변경
    else:
        st.warning("텍스트를 입력해주세요. ⚠️")