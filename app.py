import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap
from konlpy.tag import Kkma  # KoNLPy에서 Kkma 형태소 분석기 임포트


# 문장이 차지할 줄 수 계산 (단어 잘림 방지)
def sentence_line_count(sentence, max_chars_per_line=35):
    words = sentence.split()
    lines = 1
    current_line_length = 0
    for word in words:
        if current_line_length + len(word) + 1 <= max_chars_per_line:
            current_line_length += len(word) + 1
        else:
            lines += 1
            current_line_length = len(word)
    return lines


# KoNLPy를 사용하여 문장 분할 및 슬라이드 그룹화
def split_and_group_text_ko(
    text, max_lines_per_slide=5, min_chars_per_line=4, max_chars_per_line_in_ppt=18
):
    kkma = Kkma()
    slides = []
    current_slide_text = ""
    current_slide_lines = 0
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    original_sentence_flags = (
        []
    )  # [추가] 원래 문장 여부 추적 (True: 원래 문장, False: 분할된 문장)

    for sentence in sentences:
        words = sentence.split()
        lines_needed = sentence_line_count(sentence, max_chars_per_line_in_ppt)

        if current_slide_lines + lines_needed <= max_lines_per_slide:
            current_slide_text += sentence + " "
            current_slide_lines += lines_needed
            original_sentence_flags.append(True)  # 원래 문장
        else:
            # 슬라이드 분할 로직 (KoNLPy 활용)
            split_points = []
            pos_result = kkma.pos(current_slide_text + sentence)  # 형태소 분석
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
                        if sentence_line_count(
                            "".join(p[0] for p in pos_result[:idx]),
                            max_chars_per_line_in_ppt,
                        )
                        <= max_lines_per_slide
                    ),
                    default=0,
                )
                if best_split_idx > 0:
                    slides.append(
                        "".join(p[0] for p in pos_result[:best_split_idx]).strip()
                    )
                    current_slide_text = (
                        "".join(p[0] for p in pos_result[best_split_idx:]).strip() + " "
                    )
                    current_slide_lines = sentence_line_count(
                        current_slide_text, max_chars_per_line_in_ppt
                    )
                    original_sentence_flags.append(
                        False
                    )  # [수정] 분할된 문장으로 표시
                else:
                    # 분할 가능한 지점이 없으면, 단어 단위로 분할
                    slides.append(current_slide_text.strip())
                    current_slide_text = sentence + " "
                    current_slide_lines = lines_needed
                    original_sentence_flags.append(
                        False
                    )  # [수정] 분할된 문장으로 표시
            else:
                # 분할 가능한 지점이 없으면, 단어 단위로 분할
                slides.append(current_slide_text.strip())
                current_slide_text = sentence + " "
                current_slide_lines = lines_needed
                original_sentence_flags.append(
                    False
                )  # [수정] 분할된 문장으로 표시

    if current_slide_text:
        slides.append(current_slide_text.strip())
        original_sentence_flags.append(True)  # 원래 문장

    return slides, original_sentence_flags


# PPT 생성 함수
def create_ppt(
    slide_texts,
    original_sentence_flags,  # [추가] 원래 문장 여부 정보 받음
    max_chars_per_line_in_ppt=18,
    max_lines_per_slide=5,
    font_size=54,
):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = 0  # 초기값 0으로 설정
    current_slide_idx = 1
    slides_data = []  # 슬라이드 데이터 저장
    check_needed_slides = []  # [추가] 확인 필요 슬라이드 번호 저장

    try:
        for i, original_text in enumerate(slide_texts):
            lines = textwrap.wrap(
                original_text,
                width=max_chars_per_line_in_ppt,
                break_long_words=False,
                fix_sentence_endings=True,
            )
            slides_data.append(
                {
                    "text": original_text,
                    "lines": lines,
                    "original_sentence": original_sentence_flags[i],
                }
            )  # [수정] 원래 문장 여부 정보 저장
            total_slides += 1

        # 실제 슬라이드 생성
        for i, data in enumerate(slides_data):
            create_slide(prs, data, current_slide_idx, total_slides, font_size)
            if not data["original_sentence"]:  # [수정] 분할된 문장인 경우
                check_needed_slides.append(current_slide_idx)  # 슬라이드 번호 저장
            current_slide_idx += 1

        return prs, check_needed_slides  # [수정] 확인 필요 슬라이드 번호 반환

    except Exception as e:
        print(f"PPT 생성 중 오류 발생: {e}")
        return None, []


def create_slide(prs, data, current_idx, total_slides, font_size):
    """실제로 슬라이드를 생성하는 함수"""

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    tf = textbox.text_frame
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 상단 정렬
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = data["text"]

    p.font.size = Pt(font_size)  # 폰트 크기 동적으로 설정
    p.font.name = "Noto Color Emoji"  # 이모지 지원 글꼴 설정
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
    footer_p.font.name = "맑은 고딕"
    footer_p.font.color.rgb = RGBColor(128, 128, 128)
    footer_p.alignment = PP_ALIGN.RIGHT

    if current_idx == total_slides:  # 마지막 슬라이드에 '끝' 도형 추가
        add_end_mark(slide)  # 끝 표시 추가 함수 호출
    if not data["original_sentence"]:  # [수정] 분할된 문장인 경우
        add_check_needed_shape(slide)  # "확인 필요" 도형 추가


def add_end_mark(slide):
    """슬라이드에 '끝' 표시를 추가하는 함수"""

    end_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(10), Inches(6), Inches(2), Inches(1)
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


def add_check_needed_shape(slide):
    """슬라이드에 '확인 필요' 도형을 추가하는 함수"""

    check_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.3), Inches(2), Inches(0.5)
    )
    check_shape.fill.solid()
    check_shape.fill.fore_color.rgb = RGBColor(255, 255, 0)  # 노란색 배경
    check_shape.line.color.rgb = RGBColor(0, 0, 0)  # 검은색 테두리

    check_text_frame = check_shape.text_frame
    check_text_frame.clear()
    check_paragraph = check_text_frame.paragraphs[0]
    check_paragraph.text = "확인 필요"
    check_paragraph.font.size = Pt(18)
    check_paragraph.font.bold = True
    check_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    check_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


# Streamlit UI
st.set_page_config(page_title="Paydo", layout="centered")
st.title("🎬 Paydo 촬영 대본 PPT 자동 생성기 (KoNLPy)")

text_input = st.text_area("📝 촬영 대본을 입력하세요:", height=300, key="text_input_area")

# UI에서 사용자로부터 직접 값을 입력받도록 슬라이더 추가
max_lines_per_slide_input = st.slider(
    "📄 슬라이드당 최대 줄 수:", min_value=1, max_value=10, value=5, key="max_lines_slider"
)
# PPT 텍스트 박스 내에서의 줄바꿈 글자 수 (실제 PPT에 표시될 때 적용)
max_chars_per_line_ppt_input = st.slider(
    "📏 한 줄당 최대 글자 수 (PPT 표시):",
    min_value=3,
    max_value=30,
    value=18,
    key="max_chars_slider_ppt",
)
min_chars_per_line_input = st.slider(
    "🔤 한 줄당 최소 글자 수:", min_value=1, max_value=10, value=4, key="min_chars_slider"
)
font_size_input = st.slider(
    "🅰️ 폰트 크기:", min_value=10, max_value=60, value=54, key="font_size_slider"
)

if st.button("🚀 PPT 만들기", key="create_ppt_button") and text_input.strip():
    # KoNLPy를 사용한 문장 분할 함수 호출
    slide_texts, original_sentence_flags = split_and_group_text_ko(
        text_input,
        max_lines_per_slide=max_lines_per_slide_input,
        min_chars_per_line=min_chars_per_line_input,
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
    )
    ppt, check_needed_slides = create_ppt(
        slide_texts,
        original_sentence_flags,  # 원래 문장 여부 정보 전달
        max_chars_per_line_in_ppt=max_chars_per_line_ppt_input,
        max_lines_per_slide=max_lines_per_slide_input,
        font_size=font_size_input,
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
            key="download_button",
        )
        if check_needed_slides:  # 확인 필요 슬라이드 있는 경우 알림
            st.warning(
                f"❗️ 일부 슬라이드({check_needed_slides})는 최대 줄 수를 초과하여 텍스트가 나뉘었습니다. PPT를 확인하여 가독성을 검토해주세요."
            )
    else:
        st.error("❌ PPT 생성에 실패했습니다.")