import streamlit as st
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import textwrap
import io
import time


def extract_text_from_word(file):
    try:
        doc = docx.Document(file)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip() != '']
        return '\n'.join(paragraphs)
    except Exception as e:
        st.error(f"Word 파일을 읽는 중 오류가 발생했습니다: {e}")
        return ""


def split_and_group_text(text, max_lines=5, max_chars=100):
    paragraphs = text.split('\n')
    grouped_texts = []

    for para in paragraphs:
        if not para.strip():
            continue
        wrapped = textwrap.wrap(para, width=max_chars, replace_whitespace=False)
        for i in range(0, len(wrapped), max_lines):
            chunk = wrapped[i:i + max_lines]
            grouped_texts.append('\n'.join(chunk))

    return grouped_texts


def add_text_to_slide(slide, text):
    left = Inches(1)
    top = Inches(1.5)
    width = Inches(8)
    height = Inches(5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.clear()

    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = text
    font = run.font
    font.name = '맑은 고딕'
    font.size = Pt(28)
    font.color.rgb = RGBColor(0, 0, 0)


def create_ppt(slide_texts):
    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)

    for idx, text in enumerate(slide_texts):
        slide_layout = prs.slide_layouts[6]  # 빈 슬라이드
        slide = prs.slides.add_slide(slide_layout)
        add_text_to_slide(slide, text)

    return prs


def main():
    st.set_page_config(layout="wide")
    st.title("촬영 대본용 PPT 자동 생성기")

    col1, col2 = st.columns([2, 1])

    with col1:
        uploaded_file = st.file_uploader("Word 파일 업로드 (.docx)", type="docx")
        text_input = st.text_area("또는 직접 텍스트 입력", height=300, help="문단 단위로 작성해주세요.")

    with col2:
        max_lines = st.slider("슬라이드 당 최대 줄 수", 1, 10, 5)
        max_chars = st.slider("한 줄당 최대 글자 수", 30, 120, 80)

    if uploaded_file is None and not text_input.strip():
        st.warning("Word 파일을 업로드하거나 직접 텍스트를 입력해주세요.")
        return

    if uploaded_file:
        text = extract_text_from_word(uploaded_file)
    else:
        text = text_input.strip()

    if not text:
        st.warning("입력된 텍스트가 비어 있습니다.")
        return

    with st.spinner("텍스트를 슬라이드로 분할 중..."):
        slide_texts = split_and_group_text(text, max_lines, max_chars)
        st.success(f"총 {len(slide_texts)}개의 슬라이드가 생성됩니다.")

    if st.button("PPT 생성하기"):
        with st.spinner("PPT 생성 중입니다..."):
            prs = create_ppt(slide_texts)
            st.progress(100, text=f"PPT 생성 완료 - 총 {len(slide_texts)}개 슬라이드")

            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)

            st.success("PPT 파일 생성 완료!")
            st.download_button("📥 다운로드", data=ppt_io, file_name="촬영대본.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if __name__ == "__main__":
    main()
