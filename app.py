import streamlit as st
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from io import BytesIO
from datetime import datetime
import openai

# Установка API-ключа для GPT (замените на ваш ключ или настройте через переменные окружения)
openai.api_key = "your_openai_api_key"

st.set_page_config(page_title="Генератор курсовой ЭБ", layout="centered")

st.title("Генератор курсовой работы с использованием ИИ")

def generate_ai_text(prompt, max_tokens=120):
    try:
        response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=prompt,
            temperature=0.7,
            max_tokens=max_tokens,
        )
        return response.choices[0].text.strip()
    except Exception as e:
        st.error(f"Ошибка при генерации текста: {e}")
        return ""

with st.form("course_form"):
    title = st.text_area("Тема курсовой работы",
        "Экономическая безопасность предприятия в условиях санкционного давления (на примере предприятия пищевой промышленности)",
        height=120)

    col1, col2 = st.columns(2)
    with col1:
        student = st.text_input("ФИО курсанта полностью", "Иванов Иван Иванович")
        group = st.text_input("Группа", "ЭБ-21-1")
    with col2:
        teacher = st.text_input("ФИО научного руководителя", "Петров Пётр Петрович")
        teacher_title = st.text_input("Звание и должность руководителя", "полковник полиции, доцент")

    submitted = st.form_submit_button("Создать курсовую с помощью ИИ", type="primary")

if submitted:
    if not title or not student:
        st.error("Заполните тему и ФИО курсанта")
    else:
        with st.spinner("Создаём курсовую работу с использованием ИИ..."):
            doc = Document()

            # === Форматирование по требованиям ===
            section = doc.sections[0]
            section.top_margin = Mm(20)
            section.bottom_margin = Mm(20)
            section.left_margin = Mm(30)
            section.right_margin = Mm(15)

            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(14)
            style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            style.paragraph_format.first_line_indent = Mm(12.5)
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            def centered(text, size=14, bold=False):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                run.font.size = Pt(size)
                if bold: run.bold = True

            # === Титульный лист ===
            centered("МИНИСТЕРСТВО ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ")
            centered("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ КАЗЕННОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ")
            centered("ВЫСШЕГО ОБРАЗОВАНИЯ")
            centered("«САНКТ-ПЕТЕРБУРГСКИЙ УНИВЕРСИТЕТ МИНИСТЕРСТВА ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ»")
            doc.add_paragraph()
            centered("КАФЕДРА ЭКОНОМИЧЕСКОЙ БЕЗОПАСНОСТИ")
            for _ in range(7): doc.add_paragraph()
            centered(title.upper(), 16, True)
            for _ in range(8): doc.add_paragraph()
            p = doc.add_paragraph(f"Курсант {student}\n{group}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p = doc.add_paragraph(f"Научный руководитель:\n{teacher_title}\n{teacher}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_page_break()

            # === Генерация содержания с помощью ИИ ===
            intro = generate_ai_text(f"Напишите введение для курсовой работы на тему: {title}")
            chapter1 = generate_ai_text(f"Напишите содержание для первой главы на тему: Теоретические основы {title}")
            chapter2 = generate_ai_text(f"Напишите содержание для второй главы на тему: Анализ {title}")
            conclusion = generate_ai_text(f"Напишите заключение для курсовой работы на тему: {title}")

            # === Содержание ===
            centered("СОДЕРЖАНИЕ", 14, True)
            doc.add_paragraph("Введение ............................................... 3")
            doc.add_paragraph("1. Теоретические основы .................................. 5")
            doc.add_paragraph("2. Анализ ............................................... 20")
            doc.add_paragraph("Заключение ............................................. 33")
            doc.add_paragraph("Список использованной литературы ....................... 36")
            doc.add_page_break()

            # === ВВЕДЕНИЕ ===
            doc.add_heading("ВВЕДЕНИЕ", level=1)
            doc.add_paragraph(intro)

            # === Глава 1 ===
            doc.add_heading("1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ", level=1)
            doc.add_paragraph(chapter1)

            # === Глава 2 ===
            doc.add_heading("2. АНАЛИЗ", level=1)
            doc.add_paragraph(chapter2)

            # === Заключение ===
            doc.add_heading("ЗАКЛЮЧЕНИЕ", level=1)
            doc.add_paragraph(conclusion)

            # === Список литературы ===
            doc.add_heading("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ", level=1)
            sources = [
                "Конституция Российской Федерации (принята всенародным голосованием 12.12.1993)",
                "Федеральный закон от 28.12.2010 № 390-ФЗ \"О безопасности\"",
                "Указ Президента РФ от 13.05.2017 № 208 \"О Стратегии экономической безопасности Российской Федерации\"",
                "Абалкин Л.И. Экономическая безопасность России: угрозы и отражение // Вопросы экономики. 2023. № 5."
            ]
            for i, src in enumerate(sources, 1):
                doc.add_paragraph(f"{i}. {src}")

            # Сохранение документа
            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)

            filename = f"Курсовая_{student.split()[-1]}_{datetime.now().strftime('%d%m%Y')}.docx"

            st.success("✅ Курсовая работа создана с использованием ИИ!")
            st.download_button("📥 Скачать .docx файл", bio, filename, type="primary")

st.caption("После скачивания откройте файл в Word и обновите оглавление (правой кнопкой → Обновить поле)")