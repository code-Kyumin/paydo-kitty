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
def sentence_line_count(sentence, max_chars_per_line=15):  # 한 줄 최대 15자
    wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line, break_long_words=False,
                                 fix_sentence_endings=True)
    return max(1, len(wrapped_lines))

# 문장 단위로 나누고 슬라이드당 최대 줄 수 제한
def group_sentences_to_slides(sentences, max_lines_per_slide=5, max_chars_per_line=15):  # 한 줄 최대 15자
    slides = []
    current_slide_sentences = []
    current_slide_lines = 0

    for sentence in sentences:
        lines_for_sentence = sentence_line_count(sentence, max_chars_per_line)

        if current_slide_lines + lines_for_sentence > max_lines_per_slide and current_slide_sentences:
            slides.append("\n".join(current_slide_sentences))
            current_slide_sentences = [sentence]
            current_slide_lines = lines_for_sentence
        else:
            current_slide_sentences.append(sentence)
            current_slide_lines += lines_for_sentence

    if current_slide_sentences:
        slides.append("\n".join(current_slide_sentences))

    return slides

# 전체 입력을 문장 단위로 분해
def split_text(text):
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

# PPT 생성 함수
def create_ppt(slide_texts, max_chars_per_line_in_ppt=15):  # 한 줄 최대 15자
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)  # 전체 슬라이드 수

    try:
        for idx, text_for_slide in enumerate(slide_texts, 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
            tf = textbox.text_frame
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
            tf.word_wrap = True
            tf.clear()

            p = tf.paragraphs[0]
            wrapped_text = textwrap.fill(text_for_slide, width=max_chars_per_line_in_ppt, break_long_words=False,
                                         fix_sentence_endings=True, replace_whitespace=False)
            p.text = wrapped_text

            p.font.size = Pt(54)
            p.font.name = '맑은 고딕'
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.CENTER  # 가운데 정렬

            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

            # 페이지 번호 (현재 페이지/전체 페이지)
            footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
            footer_frame = footer_box.text_frame
            footer_frame.text = f"{idx} / {total_slides}"
            footer_p = footer_frame.paragraphs[0]
            footer_p.font.size = Pt(18)
            footer_p.font.name = '맑은 고딕'
            footer_p.font.color.rgb = RGBColor(128, 128, 128)
            footer_p.alignment = PP_ALIGN.RIGHT

            if idx == total_slides:  # 마지막 슬라이드에 '끝' 도형 추가
                end_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6), Inches(6.5), Inches(1.5),
                                                   Inches(0.5))
                end_shape.fill.solid()
                end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)  # 빨간색
                end_shape.line.color.rgb = RGBColor(0, 0, 0)  # 검은색 테두리

                end_text_frame = end_shape.text_frame
                end_text_frame.clear()
                end_paragraph = end_text_frame.paragraphs[0]
                end_paragraph.text = "끝"
                end_paragraph.font.size = Pt(24)
                end_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # 흰색 글자
                end_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                end_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        return prs

    except Exception as e:
        print(f"PPT 생성 중 오류 발생: {e}")
        return None

# Streamlit UI
st.set_page_config(page_title="Paydo Kitty", layout="centered")
st.title("🎤 Paydo Kitty - 촬영용 대본 PPT 생성기")

text_input = st.text_area("촬영용 대본을 입력하세요:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider")
max_chars_per_line_input = st.slider("한 줄당 최대 글자 수 (줄 수 계산 시):", min_value=10, max_value=100, value=35,
                                     key="max_chars_slider_logic")
# PPT 텍스트 박스 내에서의 줄바꿈 글자 수 (실제 PPT에 표시될 때 적용)
max_chars_per_line_ppt_input = st.slider("한 줄당 최대 글자 수 (PPT 표시용):", min_value=10, max_value=100, value=15,
                                         key="max_chars_slider_ppt")  # 기본값 15로 변경

if st.button("PPT 만들기", key="create_ppt_button") and text_input.strip():
    sentences = split_text(text_input)
    slide_texts = group_sentences_to_slides(sentences, max_lines_per_slide=max_lines_per_slide_input,
                                             max_chars_per_line=max_chars_per_line_input)
    print("slide_texts 내용:", slide_texts)
    ppt = create_ppt(slide_texts, max_chars_per_line_in_ppt=max_chars_per_line_ppt_input)

    if ppt:
        ppt_io = io.BytesIO()
        ppt.save(ppt_io)
        ppt_io.seek(0)

        st.download_button(
            label="📥 PPT 다운로드",
            data=ppt_io,
            file_name="paydo_kitty_script.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="download_button"
        )
    else:
        st.error("PPT 생성에 실패했습니다. 입력 데이터를 확인하거나 잠시 후 다시 시도해주세요.")