def create_slide(prs, text, current_idx, total_slides, max_chars_per_line):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    tf = textbox.text_frame
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 텍스트 상자 상단 정렬
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = textwrap.fill(text, width=max_chars_per_line, break_long_words=False)

    p.font.size = Pt(54)
    p.font.name = '맑은 고딕'
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

    # tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE  # 이 줄을 제거하거나 주석 처리
    tf.paragraphs[0].vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE # 문단 내에서 수직 가운데 정렬

    footer_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4))
    footer_frame = footer_box.text_frame
    footer_frame.text = f"{current_idx} / {total_slides}"
    footer_p = footer_frame.paragraphs[0]
    footer_p.font.size = Pt(18)
    footer_p.font.name = '맑은 고딕'
    footer_p.font.color.rgb = RGBColor(128, 128, 128)
    footer_p.alignment = PP_ALIGN.RIGHT

    if current_idx == total_slides:
        add_end_mark(slide)