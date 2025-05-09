import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
import io
import re
import textwrap

# 문장이 차지할 줄 수 계산 (단어 잘림 방지)
def sentence_line_count(sentence, max_chars_per_line=35):  # 이 값을 조정하여 한 줄의 글자 수 변경
    # textwrap.wrap은 단어를 자르지 않고 줄바꿈을 시도합니다.
    # break_long_words=False가 기본값이지만 명시적으로 표현했습니다.
    wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line, break_long_words=False, fix_sentence_endings=True)
    return max(1, len(wrapped_lines))

# 문장 단위로 나누고 슬라이드당 최대 줄 수 제한
def group_sentences_to_slides(sentences, max_lines_per_slide=4, max_chars_per_line=35):  # 슬라이드당 최대 줄 수, 줄당 최대 글자 수
    slides = []
    current_slide_sentences = []
    current_slide_lines = 0

    for sentence in sentences:
        # 문장이 길 경우, 문장 자체를 여러 줄로 나누어 계산합니다.
        # 이 때, 단어가 잘리지 않도록 합니다.
        lines_for_sentence = sentence_line_count(sentence, max_chars_per_line)

        if current_slide_lines + lines_for_sentence > max_lines_per_slide and current_slide_sentences:
            slides.append("\n".join(current_slide_sentences))  # 각 문장을 개행으로 합쳐 한 슬라이드의 텍스트로 만듦
            current_slide_sentences = [sentence]
            current_slide_lines = lines_for_sentence
        else:
            current_slide_sentences.append(sentence)
            current_slide_lines += lines_for_sentence

    if current_slide_sentences:  # 남은 문장들이 있다면 마지막 슬라이드에 추가
        slides.append("\n".join(current_slide_sentences))

    return slides

# 전체 입력을 문장 단위로 분해
def split_text(text):
    # 문장 분리 시 마침표, 물음표, 느낌표 뒤에 공백이 오는 경우를 기준으로 합니다.
    # 다양한 문장 부호와 상황에 맞춰 정규식을 개선할 수 있습니다.
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

# PPT 생성 함수
def create_ppt(slide_texts, max_chars_per_line_in_ppt=35):  # PPT 내부 텍스트 박스용 줄당 글자 수
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    try:  # PPT 생성 과정에서 발생할 수 있는 오류를 처리하기 위해 try-except 블록을 사용
        for idx, text_for_slide in enumerate(slide_texts, 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # 빈 슬라이드 레이아웃 사용
            textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
            tf = textbox.text_frame
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
            tf.word_wrap = True  # 자동 줄 바꿈 활성화
            # tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE # 텍스트에 맞춰 도형 크기 조정 (필요시 주석 해제)
            tf.clear()  # 기존 텍스트 프레임 내용 삭제

            p = tf.paragraphs[0]  # 첫 번째 단락 사용
            # textwrap.fill을 사용하여 단어 단위로 줄바꿈 된 텍스트를 만듭니다.
            # 이 때, break_long_words=False로 설정하여 단어가 중간에 잘리는 것을 방지합니다.
            wrapped_text = textwrap.fill(text_for_slide, width=max_chars_per_line_in_ppt, break_long_words=False,
                                         fix_sentence_endings=True, replace_whitespace=False)
            p.text = wrapped_text

            p.font.size = Pt(54)
            p.font.name = '맑은 고딕'
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.LEFT  # 왼쪽 정렬

            # 텍스트 프레임 내에서 상하 정렬 (상단 정렬)
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

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

    except Exception as e:
        print(f"PPT 생성 중 오류 발생: {e}")  # 오류 메시지 출력 (디버깅용)
        return None  # 오류 발생 시 None 반환 또는 다른 적절한 처리

# Streamlit UI
st.set_page_config(page_title="Paydo Kitty", layout="centered")
st.title("🎤 Paydo Kitty - 촬영용 대본 PPT 생성기")

text_input = st.text_area("촬영용 대본을 입력하세요:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider("슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=4, key="max_lines_slider")
max_chars_per_line_input = st.slider("한 줄당 최대 글자 수 (줄 수 계산 시):", min_value=10, max_value=100, value=35, key="max_chars_slider_logic")
# PPT 텍스트 박스 내에서의 줄바꿈 글자 수 (실제 PPT에 표시될 때 적용)
max_chars_per_line_ppt_input = st.slider("한 줄당 최대 글자 수 (PPT 표시용):", min_value=10, max_value=100, value=35, key="max_chars_slider_ppt")

if st.button("PPT 만들기", key="create_ppt_button") and text_input.strip():
    sentences = split_text(text_input)
    # 사용자가 UI에서 설정한 값을 group_sentences_to_slides 함수에 전달
    slide_texts = group_sentences_to_slides(sentences, max_lines_per_slide=max_lines_per_slide_input,
                                             max_chars_per_line=max_chars_per_line_input)
    print("slide_texts 내용:", slide_texts)  # 추가: slide_texts 내용 확인
    ppt = create_ppt(slide_texts, max_chars_per_line_ppt=max_chars_per_line_ppt_input)

    if ppt:  # ppt가 None이 아닌 경우에만 다운로드 버튼 생성
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