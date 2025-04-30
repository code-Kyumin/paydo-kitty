
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import re
import textwrap
import io

MAX_LINES_PER_SLIDE = 4
CHARS_PER_LINE = 35

def split_into_sentences(text):
    return [s.strip() for s in re.split(r'(?<=[.!?])\s+', text.strip()) if s.strip()]

def sentence_line_count(sentence, chars_per_line=CHARS_PER_LINE):
    return max(1, len(textwrap.wrap(sentence, width=chars_per_line)))

def chunk_sentences_by_line_limit(sentences, max_lines=MAX_LINES_PER_SLIDE):
    chunks = []
    current_chunk = []
    current_line_count = 0

    for sentence in sentences:
        line_count = sentence_line_count(sentence)
        if current_line_count + line_count <= max_lines:
            current_chunk.append(sentence)
            current_line_count += line_count
        else:
            if current_chunk:
                chunks.append(current_chunk)
            current_chunk = [sentence]
            current_line_count = line_count
    if current_chunk:
        chunks.append(current_chunk)
    return chunks

def create_ppt(slide_chunks):
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]

    for chunk in slide_chunks:
        slide = prs.slides.add_slide(blank_layout)
        left = Inches(0.5)
        top = Inches(1)
        width = Inches(9)
        height = Inches(5.5)

        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.text = ""
        tf.vertical_anchor = PP_ALIGN.MIDDLE

        for sentence in chunk:
            p = tf.add_paragraph()
            p.text = sentence
            p.font.size = Pt(48)
            p.alignment = PP_ALIGN.CENTER
            p.font.name = 'Arial'
            p.font.color.rgb = RGBColor(0, 0, 0)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

st.title("🐾 paydo kitty - 텍스트 → 프롬프트용 PPT 변환기")
st.markdown("입력한 스크립트를 **문장 단위로 나눠**, 슬라이드당 최대 4줄로 구성된 PPT를 생성합니다.")

text_input = st.text_area("📝 스크립트를 입력하세요", height=300)

if st.button("🎬 PPT 만들기"):
    if not text_input.strip():
        st.warning("텍스트를 입력해주세요.")
    else:
        sentences = split_into_sentences(text_input)
        chunks = chunk_sentences_by_line_limit(sentences)
        ppt_file = create_ppt(chunks)

        st.success("✅ PPT 생성 완료!")
        st.download_button(
            label="📥 PPT 다운로드",
            data=ppt_file,
            file_name="paydo_kitty_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
