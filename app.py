import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Генератор курсовой ЭБ", page_icon="📘", layout="centered")

st.title("📘 Генератор курсовой работы (по МР СПбУ МВД 2026)")
st.markdown("**Специальность 38.05.01 Экономическая безопасность**")

with st.form("course_form"):
    col1, col2 = st.columns(2)
    with col1:
        title = st.text_area("Название курсовой работы", height=100)
        student = st.text_input("ФИО курсанта (полностью)")
        group = st.text_input("Группа / Курс")
        rank = st.text_input("Звание (если есть)", value="курсант")
    with col2:
        teacher = st.text_input("ФИО руководителя")
        teacher_rank = st.text_input("Звание и должность руководителя", value="полковник полиции, доцент")
        university = st.text_input("Университет", value="Санкт-Петербургский университет МВД России")
        year = st.number_input("Год", value=2026)

    submitted = st.form_submit_button("Создать курсовую по МР", type="primary")

if submitted:
    if not title or not student:
        st.error("Укажите название и ФИО")
    else:
        with st.spinner("Создаём документ строго по методичке..."):
            doc = Document()

            # === Настройки по МР ===
            section = doc.sections[0]
            section.top_margin = Mm(20)
            section.bottom_margin = Mm(20)
            section.left_margin = Mm(30)
            section.right_margin = Mm(15)

            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'  # PT Astra Serif часто заменяется на TNR
            style.font.size = Pt(14)
            style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            style.paragraph_format.space_after = Pt(0)
            style.paragraph_format.first_line_indent = Cm(1.25)

            # === Титульный лист (по Приложению 2) ===
            def centered(text, size=14, bold=False):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                run.font.size = Pt(size)
                if bold: run.bold = True

            centered("МИНИСТЕРСТВО ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ", 14)
            centered("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ КАЗЕННОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ", 14)
            centered("ВЫСШЕГО ОБРАЗОВАНИЯ", 14)
            centered("«САНКТ-ПЕТЕРБУРГСКИЙ УНИВЕРСИТЕТ МИНИСТЕРСТВА ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ»", 14)
            
            doc.add_paragraph()
            centered("КАФЕДРА ЭКОНОМИЧЕСКОЙ БЕЗОПАСНОСТИ", 14)
            
            for _ in range(8): doc.add_paragraph()
            
            centered(title.upper(), 16, True)
            
            for _ in range(8): doc.add_paragraph()
            
            p = doc.add_paragraph(f"{rank} {student}\n{group}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            p = doc.add_paragraph(f"Научный руководитель:\n{teacher_rank}\n{teacher}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            doc.add_page_break()

            # === Содержание ===
            centered("СОДЕРЖАНИЕ", 14, True)
            doc.add_paragraph("Введение ........................................ 3")
            doc.add_paragraph("1. Теоретические основы .......................... 4")
            doc.add_paragraph("2. Анализ состояния .............................. 15")
            doc.add_paragraph("Заключение ...................................... 30")
            doc.add_paragraph("Список использованной литературы ........... 32")
            doc.add_paragraph("Приложения ...................................... 35")
            doc.add_page_break()

            # === Наполнение (примерное, под экономическую безопасность) ===
            doc.add_heading("ВВЕДЕНИЕ", level=1)
            doc.add_paragraph("Актуальность темы исследования обусловлена...")  # можно доработать

            doc.add_heading("1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ЭКОНОМИЧЕСКОЙ БЕЗОПАСНОСТИ", level=1)
            doc.add_heading("1.1. Сущность и содержание экономической безопасности", level=2)
            doc.add_paragraph("...")

            doc.add_heading("2. АНАЛИЗ ЭКОНОМИЧЕСКОЙ БЕЗОПАСНОСТИ [ОБЪЕКТ]", level=1)
            doc.add_paragraph("...")

            doc.add_heading("ЗАКЛЮЧЕНИЕ", level=1)
            doc.add_paragraph("В результате проведенного исследования можно сделать следующие выводы...")

            doc.add_heading("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ", level=1)
            doc.add_paragraph("1. Конституция Российской Федерации...")
            # Добавь минимум 40 источников позже вручную

            doc.add_heading("ПРИЛОЖЕНИЯ", level=1)

            # Сохранение
            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)

            filename = f"Курсовая_{student.split()[-1]}_{datetime.now().strftime('%d%m%Y')}.docx"

            st.success("✅ Документ создан строго по МР!")
            st.download_button("📥 Скачать .docx", bio, filename, type="primary")

st.info(""" 
**После скачивания:**
1. Откройте в Word
2. Обновите оглавление (правой кнопкой → Обновить поле)
3. Замените примерный текст на свой
4. Добавьте реальные источники (не менее 40)
""")