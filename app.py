def create_ppt(slides):
    prs = Presentation()
    prs.slide_width = Inches(13.33)  # 16:9 비율
    prs.slide_height = Inches(7.5)

    for idx, lines in enumerate(slides, 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        tf.word_wrap = True  # ← 자동 줄바꿈
        tf.auto_size = None
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = line
            p.font.size = Pt(54)
            p.font.name = '맑은 고딕'
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.CENTER

        # 우측 하단 페이지 번호
        footer_box = slide.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(1), Inches(0.4))
        footer_frame = footer_box.text_frame
        footer_frame.text = str(idx)
        footer_p = footer_frame.paragraphs[0]
        footer_p.font.size = Pt(18)
        footer_p.font.name = '맑은 고딕'
        footer_p.font.color.rgb = RGBColor(128, 128, 128)
        footer_p.alignment = PP_ALIGN.RIGHT

    return prs
