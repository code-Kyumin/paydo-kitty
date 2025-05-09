import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
import io
import re
import textwrap

# 긴 단어를 강제로 잘라서 줄바꿈 힌트 추가
def force_wrap_long_words(sentence, max_word_length=30):
    words = sentence.split(" ")
    wrapped = []
    for word in words:
        if len(word) > max_word_length:
            chunks = [word[i:i+max_word_length] for i in range(0, len(word), max_word_length)]
            wrapped.append("\n".join(chunks))
        else:
            wrapped.append(word)
    return " ".join(wrapped)

# 문장이 차지할 줄 수 계산
def sentence_line_count(sentence, max_chars_per_line=35):
    return max(1, len(textwrap.wrap(sentence, width=max_chars_per_line, break_long_words=False)))

# 문장 단위로 나누고 슬라이드당 최대 줄 수 제한
def group_sentences_to_slides(sentences, max_lines_per_slide=5):
    slides = []
    current_slide = []
    current_lines = 0

    for sentence in sentences:
        sentence = force_wrap_long_words(sentence)
        lines = sentence_line_count(sentence)
        if current_lines + lines > max_lines_per_slide:
            slides.append(current_slide)
            current_slide = [sentence]
            current_lines = lines
        else:
            current_slide.append(sentence)
            current_lines += lines

    if current_slide:
        slides.append(current_slide)

    return slides

# 전체 입력을 문장 단위로 분해
def split_text(text):
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

# PPT 생성 함수
def create_ppt(slides):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for idx, lines in enumerate(slides, 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = line
            p.font.size = Pt(54)
            p.font.name = '맑은 고딕'
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.CENTER

        # 페이지 번호
        footer_box = slide.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(1), Inches(0.4))
        footer_frame = footer_box.text_frame
        footer_frame.text = str(idx)
        footer_p = footer_frame.paragraphs[0]
        footer_p.font.size = Pt(18)
        footer_p.font.name = '맑은 고딕'
        footer_p.font.color.rgb = RGBColor(128, 128, 128)
        footer_p.alignment = PP_ALIGN.RIGHT

    return prs

# Streamlit UI
st.set_page_config(page_title="Paydo Kitty", layout="centered")
st.title("🎤 Paydo Kitty - 촬영용 대본 PPT 생성기")

text_input = st.text_area("촬영용 대본을 입력하세요:", height=300)

if st.button("PPT 만들기") and text_input.strip():
    sentences = split_text(text_input)
    slides = group_sentences_to_slides(sentences)
    ppt = create_ppt(slides)

    ppt_io = io.BytesIO()
    ppt.save(ppt_io)
    ppt_io.seek(0)

    st.download_button(
        label="📥 PPT 다운로드",
        data=ppt_io,
        file_name="paydo_kitty_script.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
