import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap

# 문장이 차지할 줄 수 계산 (단어 잘림 방지)
def sentence_line_count(sentence, max_chars_per_line=20):
    wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line, break_long_words=False,
                                 fix_sentence_endings=True)
    return max(1, len(wrapped_lines))

# 전체 입력을 문장 단위로 분해
def split_text(text):
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

# PPT 생성 함수
def create_ppt(slide_texts, max_chars_per_line_in_ppt=20, max_lines_per_slide=5):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = 0  # 초기값 0으로 설정
    current_slide_idx = 1
    slides_data = []  # 슬라이드 데이터 저장

    try:
        for original_text in slide_texts:
            lines = textwrap.wrap(original_text, width=max_chars_per_line_in_ppt,
                                  break_long_words=False, replace_whitespace=False)
            num_lines = len(lines)

            if num_lines <= max_lines_per_slide:
                slides_data.append({
                    "text": original_text,
                    "lines": lines
                })
                total_slides += 1
            else:
                # 현재 문장이 최대 줄 수를 초과하는 경우, 강제로 새 슬라이드 생성
                temp_text = ""
                temp_lines = []
                for line in lines:
                    temp_lines.append(line)
                    if len(temp_lines) == max_lines_per_slide:
                        slides_data.append({
                            "text": "\n".join(temp_lines),
                            "lines": temp_lines
                        })
                        total_slides += 1
                        temp_lines = []
                if temp_lines:  # 남은 줄이 있으면 추가
                    slides_data.append({
                        "text": "\n".join(temp_lines),
                        "lines": temp_lines
                    })
                    total_slides += 1

        # 실제 슬라이드 생성
        for data in slides_data:
            create_slide(prs, data["text"], current_slide_idx, total_slides)
            current_slide_idx += 1

        return prs

    except Exception as e:
        print(f"PPT 생성 중 오류 발생: {e}")
        return None

def create_slide(prs, text, current_idx, total_slides):
    """실제로 슬라이드를 생성하는 함수"""

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    tf = textbox.text_frame
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = text

    p.font.size = Pt(54)
    p.font.name = '맑은 고딕'
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    # 페이지 번호 (현재 페이지/전체 페이지)
    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_frame = footer_box.text_frame
    footer_frame.text = f"{current_idx} / {total_slides}"
    footer_p = footer_frame.paragraphs[0]
    footer_p.font.size = Pt(18)
    footer_p.font.name = '맑은 고딕'
    footer_p.font.color.rgb = RGBColor(128, 128, 128)
    footer_p.alignment = PP_ALIGN.RIGHT

    if current_idx == total_slides:  # 마지막 슬라이드에 '끝' 도형 추가
        add_end_mark(slide)  # 끝 표시 추가 함수 호출

def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가하는 함수"""

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
st.set_page_config(page_title="Paydo Kitty", layout="centered")