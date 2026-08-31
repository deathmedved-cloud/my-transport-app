import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Генерация Актов и Счетов", layout="wide")
st.title("🚚 Генератор Счёта и Акта (ООО «АВРОРА-ТРАНЗИТ»)")

# --- ИСХОДНЫЕ ДАННЫЕ ВВОДА ---
st.subheader("1. Данные по Договору-Заявке")

col1, col2 = st.columns(2)

with col1:
    doc_num = st.text_input("Номер заявки", value="20")
    doc_date = st.text_input("Дата заявки", value="23.07.2026")
    customer_name = st.text_input("Заказчик", value='ООО «ФорГлэйд»')
    customer_details = st.text_area("Реквизиты Заказчика", value="УНП: 590940553, РБ 231940, Гродненская обл, Зельвенский р-н, д. Бережки 1В")

with col2:
    route = st.text_input("Маршрут перевозки", value="г. Щелково (РФ) — д. Бережки (РБ)")
    amount = st.text_input("Сумма фрахта", value="1080,00")
    currency = st.selectbox("Валюта", ["белорусских рублей (BYN)", "росс. руб. (RUB)", "евро (EUR)"])
    nds = st.selectbox("НДС", ["Без НДС", "0%", "20%"])

# --- ФУНКЦИИ ГЕНЕРАЦИИ DOCX ---

def generate_invoice():
    doc = Document()
    
    # Заголовок реквизитов
    p = doc.add_paragraph()
    p.add_run("Исполнитель: ").bold = True
    p.add_run("ООО \"АВРОРА-ТРАНЗИТ\", ИНН 6700042504, КПП 670001001\nРФ, Смоленская обл., г. Смоленск, ул. Карбышева, д. 15а, корп. 2, оф. К 52\n")
    p.add_run("Р/с: 40702933501130000115 в ЗАО «АЛЬФА-БАНК», БИК ALFABY2X\n").italic = True
    
    doc.add_paragraph("-" * 50)
    
    # Название документа
    h = doc.add_paragraph(f"СЧЁТ НА ОПЛАТУ № {doc_num} от {doc_date} г.")
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.runs[0].bold = True
    h.runs[0].font.size = Pt(14)

    # Заказчик и Основание
    doc.add_paragraph(f"Заказчик: {customer_name} ({customer_details})")
    doc.add_paragraph(f"Основание: Договор-заявка № {doc_num} от {doc_date} г.")

    # Таблица услуг
    table = doc.add_table(rows=2, cols=5)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '№'
    hdr_cells[1].text = 'Наименование услуги'
    hdr_cells[2].text = 'Кол-во'
    hdr_cells[3].text = 'НДС'
    hdr_cells[4].text = f'Сумма ({currency})'

    row_cells = table.rows[1].cells
    row_cells[0].text = '1'
    row_cells[1].text = f'Транспортные услуги по маршруту: {route}'
    row_cells[2].text = '1'
    row_cells[3].text = nds
    row_cells[4].text = amount

    doc.add_paragraph()
    doc.add_paragraph(f"Всего к оплате: {amount} {currency}").bold = True
    
    # Исполнитель
    p_sign = doc.add_paragraph("\nДиректор ООО «АВРОРА-ТРАНЗИТ»: _________________ / Галенда С.В. /")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_act():
    doc = Document()
    
    h = doc.add_paragraph(f"АКТ ВЫПОЛНЕННЫХ РАБОТ № {doc_num} от {doc_date} г.")
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.runs[0].bold = True
    h.runs[0].font.size = Pt(14)

    doc.add_paragraph(f"Исполнитель: ООО «АВРОРА-ТРАНЗИТ»")
    doc.add_paragraph(f"Заказчик: {customer_name}")
    doc.add_paragraph(f"Основание: Договор-заявка № {doc_num} от {doc_date} г.")
    doc.add_paragraph("-" * 50)

    p_body = doc.add_paragraph(
        f"Мы, нижеподписавшиеся, подтверждаем, что Исполнитель выполнил автомобильную перевозку "
        f"груза по маршруту: {route}. Услуги оказаны в полном объёме и в срок. "
        f"Общая стоимость оказанных услуг составляет {amount} {currency} ({nds})."
    )

    doc.add_paragraph("\n" * 2)
    
    table = doc.add_table(rows=1, cols=2)
    cells = table.rows[0].cells
    cells[0].text = "Исполнитель:\nООО «АВРОРА-ТРАНЗИТ»\n\n_________________ /Галенда С.В./"
    cells[1].text = f"Заказчик:\n{customer_name}\n\n_________________ /___________/"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- КНОПКИ СКАЧИВАНИЯ ---
st.subheader("2. Сформировать и скачать документы")

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    invoice_file = generate_invoice()
    st.download_button(
        label="📄 Скачать Счёт (.docx)",
        data=invoice_file,
        file_name=f"Счет_{doc_num}_от_{doc_date}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with col_btn2:
    act_file = generate_act()
    st.download_button(
        label="📝 Скачать Акт выполненных работ (.docx)",
        data=act_file,
        file_name=f"Акт_{doc_num}_от_{doc_date}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
