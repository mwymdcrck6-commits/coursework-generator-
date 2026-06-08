import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Генератор курсовой",
    page_icon="📘",
    layout="centered",
    initial_sidebar_state="expanded"
)

st.title("📘 Генератор курсовой работы")
st.markdown("**По методическим рекомендациям российских вузов (ГОСТ)**")

st.sidebar.header("Настройки документа")
left_margin = st.sidebar.slider("Левое поле (см)", 2.0, 4.0, 3.0)
right_margin = st.sidebar.slider("Правое поле (см)", 1.0, 3.0, 1.5)
top_bottom_margin = st.sidebar.slider("Верхнее/нижнее поле (см)", 1.5, 3.0, 2.0)

with st.form("course_form"):
    col1, col2 = st.columns(2)
    with col1:
        title = st.text_area("Название курсовой работы", placeholder="Анализ эффективности внедрения...", height=100)
        student = st.text_input("ФИО студента (полностью)")
        group = st.text_input("Группа / Специальность")
    with col2:
        teacher = st.text_input("ФИО научного руководителя")
        university = st.text_input("Полное название университета", value="ФГБОУ ВО «Название университета»")
        city = st.text_input("Город", value="Москва")
        year = st.number_input("Год защиты", value=datetime.now().year)

    submitted = st.form_submit_button("🚀 Создать курсовую работу", type="primary")

if submitted:
    if not title.strip() or not student.strip():
        st.error("⚠️ Пожалуйста, заполните название работы и ФИО студента")
    else:
        with st.spinner("Создаём документ по ГОСТ..."):
            doc = Document()

            # === Форматирование ===
            section = doc.sections[0]
            section.top_margin = Cm(top_bottom_margin)
            section.bottom_margin = Cm(top_bottom_margin)
            section.left_margin = Cm(left_margin)
            section.right_margin = Cm(right_margin)

            normal_style = doc.styles['Normal']
            normal_style.font.name = 'Times New Roman'
            normal_style.font.size = Pt(14)
            normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            normal_style.paragraph_format.space_after = Pt(0)

            # === Титульный лист ===
            def centered(text, size=14, bold=False):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                run.font.size = Pt(size)
                if bold:
                    run.bold = True

            centered("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ")
            centered(university)
            doc.add_paragraph()
            centered("КАФЕДРА ___________________________")
            
            for _ in range(6): doc.add_paragraph()
            
            centered(title.upper(), size=16, bold=True)
            
            for _ in range(8): doc.add_paragraph()
            
            p = doc.add_paragraph(f"Студент: {student}\nГруппа: {group}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            p = doc.add_paragraph(f"Научный руководитель: {teacher}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            doc.add_page_break()

            # === Оглавление ===
            centered("ОГЛАВЛЕНИЕ", size=14, bold=True)
            doc.add_paragraph("Введение ........................................ 3")
            for i in range(1, 4):
                doc.add_paragraph(f"{i}. {['Теоретические основы', 'Анализ предметной области', 'Практическая часть'][i-1]} ................")
            doc.add_paragraph("Заключение ......................................")
            doc.add_paragraph("Список использованной литературы ......")
            doc.add_paragraph("Приложения ......................................")
            doc.add_page_break()

            # === Основное содержание ===
            doc.add_heading("ВВЕДЕНИЕ", level=1)
            doc.add_paragraph("Актуальность темы исследования обусловлена...")

            doc.add_heading("1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ИССЛЕДОВАНИЯ", level=1)
            doc.add_heading("1.1. Подраздел (пример)", level=2)
            doc.add_paragraph("Текст главы...")

            doc.add_heading("2. ПРАКТИЧЕСКАЯ ЧАСТЬ / АНАЛИЗ", level=1)
            doc.add_paragraph("В данном разделе проводится...")

            doc.add_heading("ЗАКЛЮЧЕНИЕ", level=1)
            doc.add_paragraph("В ходе выполнения курсовой работы были решены поставленные задачи...")

            doc.add_heading("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ", level=1)
            doc.add_paragraph("1. Иванов И. И. Название источника. – М.: Издательство, 2025.")

            doc.add_heading("ПРИЛОЖЕНИЯ", level=1)
            doc.add_paragraph("Приложение А. Таблицы, графики, код и т.д.")

            # === Скачивание ===
            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)

            filename = f"Курсовая_{student.split()[-1]}_{datetime.now().strftime('%d%m%Y')}.docx"

            st.success("✅ Курсовая успешно создана!")
            st.download_button(
                label="📥 Скачать .docx файл",
                data=bio,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )

st.info("**После скачивания:** Откройте файл в Microsoft Word → правой кнопкой по оглавлению → «Обновить поле».")
st.caption("Приложение полностью работает в браузере, включая iOS Safari.")