import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Генератор курсовой ЭБ", layout="centered")

st.title("📘 Полноценный генератор курсовой (по МР СПбУ МВД 2026)")

with st.form("course_form"):
    title = st.text_area("Полное название курсовой работы", 
                        "Экономическая безопасность предприятия в условиях санкционного давления (на примере ...)",
                        height=120)
    student = st.text_input("ФИО курсанта полностью")
    group = st.text_input("Группа", "ЭБ-XXX")
    teacher = st.text_input("ФИО руководителя")
    teacher_title = st.text_input("Звание и должность руководителя", "полковник полиции, доцент")

    submitted = st.form_submit_button("Создать ПОЛНУЮ курсовую работу", type="primary")

if submitted and title and student:
    with st.spinner("Генерирую полную курсовую по методичке..."):
        doc = Document()

        # Настройки по МР
        section = doc.sections[0]
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.paragraph_format.first_line_indent = Cm(1.25)

        def centered(text, size=14, bold=False):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.font.size = Pt(size)
            if bold: run.bold = True

        # Титульный лист
        centered("МИНИСТЕРСТВО ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ", 14)
        centered("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ КАЗЕННОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ", 14)
        centered("ВЫСШЕГО ОБРАЗОВАНИЯ", 14)
        centered("«САНКТ-ПЕТЕРБУРГСКИЙ УНИВЕРСИТЕТ МИНИСТЕРСТВА ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ»", 14)
        doc.add_paragraph()
        centered("КАФЕДРА ЭКОНОМИЧЕСКОЙ БЕЗОПАСНОСТИ", 14)
        for _ in range(7): doc.add_paragraph()
        centered(title.upper(), 16, True)
        for _ in range(8): doc.add_paragraph()
        p = doc.add_paragraph(f"Курсант {student}\n{group}")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p = doc.add_paragraph(f"Научный руководитель:\n{teacher_title}\n{teacher}")
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        doc.add_page_break()

        # Содержание