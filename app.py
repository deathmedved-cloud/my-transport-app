import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Генератор Счетов и Актов", layout="wide")
st.title("🚚 Документы ООО «АВРОРА-ТРАНЗИТ»")

# --- ПОЛЯ ВВОДА ---
st.subheader("1. Основные реквизиты договора и перевозки")

col1, col2 = st.columns(2)

with col1:
    doc_num = st.text_input("Номер заявки / счета / акта", value="14")
    doc_date = st.text_input("Дата документа", value="15.07.2026")
    order_num = st.text_input("Номер транспортного заказа (для Акта)", value="01-07")
    order_date = st.text_input("Дата заказа", value="10.07.2026")
    city = st.text_input("Город составления Акта", value="г. Смоленск")
    
    st.markdown("---")
    st.markdown("**Данные Заказчика:**")
    customer_name = st.text_input("Наименование Заказчика", value="ООО «Егорчик»")
    customer_director = st.text_input("ФИО директора (в родительном падеже)", value="Насирова Олега Холмуродовича")
    customer_details = st.text_area("Полные реквизиты/адрес Заказчика", value="230026, РБ, г. Гродно, ул. Победы 17, УНП 591503882")
    customer_signer = st.text_input("Подпись Заказчика (ФИО)", value="Насиров О.Х.")

with col2:
    route = st.text_input("Маршрут перевозки", value="РБ, ТЛЦ Брузги (место перецепки) - РБ")
    transport_details = st.text_input("Детали ТС / перецепка (для Акта)", value="в том числе на этапе перевозки после перецепки на ТЛЦ Брузги, ТС М226ЕО67 / 3TA7860")
    cmr_numbers = st.text_input("Номера CMR (через запятую)", value="R456159, MA456028, MA455671, MA453832, R455456, MA455721, VB455624, R455078, IB454302")
    
    st.markdown("---")
    st.markdown("**Финансы:**")
    amount_num = st.text_input("Сумма цифрами", value="1 450,00")
    amount_words = st.text_input("Сумма прописью", value="Одна тысяча четыреста пятьдесят")
    currency = st.selectbox("Валюта", ["евро", "белорусских рублей", "российских рублей"])
    nds = st.text_input("НДС", value="Без НДС")
    payment_terms = st.text_input("Условия конвертации", value="в российских рублях по курсу ЦБ РФ на дату оплаты")

# --- ФУНКЦИЯ: СЧЕТ НА ОПЛАТУ (ПО ОБРАЗЦУ INVOICE-25) ---
def generate_invoice():
    doc = Document()
    
    # Шапка банка (таблица реквизитов)
    table_bank = doc.add_table(rows=2, cols=2)
    table_bank.style = 'Table Grid'
    
    c0 = table_bank.rows[0].cells[0]
    c0.text = "АО \"АЛЬФА-БАНК\"\nБанк получателя"
    c1 = table_bank.rows[0].cells[1]
    c1.text = "БИК: 044525593\nСч.№: 30101810200000000593"
    
    c2 = table_bank.rows[1].cells[0]
    c2.text = "ИНН: 6700042504 | КПП: 670001001\nООО \"АВРОРА-ТРАНЗИТ\"\nПолучатель"
    c3 = table_bank.rows[1].cells[1]
    c3.text = "Сч.№: 40702810901130005079"

    doc.add_paragraph()
    
    # Исполнитель и Заказчик
    p_party = doc.add_paragraph()
    p_party.add_run("Исполнитель: ").bold = True
    p_party.add_run("ООО \"АВРОРА-ТРАНЗИТ\", 214022, РОССИЯ, Смоленская область, Смоленск, ул Карбышева, 15а, 2, К 52\n")
    p_party.add_run("Заказчик: ").bold = True
    p_party.add_run(f"{customer_name}, {customer_details}\n")
    p_party.add_run("Комментарий: ").bold = True
    p_party.add_run(f"Договор-заявка № {doc_num} от {doc_date}г. {payment_terms}.")

    # Заголовок Счета
    h = doc.add_paragraph(f"\nСчет на оплату №{doc_num} от {doc_date} г.")
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h.runs[0].bold = True
    h.runs[0].font.size = Pt(14)

    # Таблица услуг
    t_services = doc.add_table(rows=2, cols=6)
    t_services.style = 'Table Grid'
    
    headers = ["№", "Название услуги", "Кол-во", "Ед.изм", "НДС", "Сумма"]
    for i, title in enumerate(headers):
        t_services.rows[0].cells[i].text = title
        
    row = t_services.rows[1].cells
    row[0].text = "1"
    row[1].text = f"Автомобильная перевозка груза по маршруту: {route}"
    row[2].text = "1"
    row[3].text = "шт."
    row[4].text = nds
    row[5].text = amount_num

    # Итоги
    doc.add_paragraph()
    p_total = doc.add_paragraph()
    p_total.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_total.add_run(f"Итого: {amount_num}\nСумма НДС: 0,00\nВсего к оплате: {amount_num} {currency}").bold = True
    
    doc.add_paragraph(f"Всего наименований 1 на сумму {amount_num} {currency}\n{amount_words} {currency}").italic = True
    doc.add_paragraph("\n\n(должность) _________________ (подпись) / Галенда С.В. /")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- ФУНКЦИЯ: АКТ ВЫПОЛНЕННЫХ РАБОТ (ПО ОБРАЗЦУ АКТ 14) ---
def generate_act():
    doc = Document()
    
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = h.add_run(f"Акт выполненных работ № {doc_num}\nоб оказании услуг\nпо транспортному заказу № {order_num} от {order_date} г.")
    r1.bold = True
    r1.font.size = Pt(12)
    
    p_meta = doc.add_paragraph()
    p_meta.add_run(f"{city}\t\t\t\t\t\t\t\t\t{doc_date} г.").bold = True
    
    p_preamble = doc.add_paragraph()
    p_preamble.paragraph_format.first_line_indent = Inches(0.4)
    p_preamble.add_run(
        f"{customer_name}, в лице директора {customer_director}, действующей на основании Устава, "
        f"именуемое в дальнейшем «Заказчик» с одной стороны, и Общество с ограниченной ответственностью «АВРОРА-ТРАНЗИТ», "
        f"в лице директора Галенда Сергея Владимировича, действующего на основании Устава, именуемое в дальнейшем «Исполнитель» "
        f"с другой стороны, составили настоящий акт о нижеследующем:"
    )
    
    p_b1 = doc.add_paragraph()
    p_b1.paragraph_format.first_line_indent = Inches(0.4)
    p_b1.add_run(
        f"В установленные транспортным заказом сроки ООО «АВРОРА-ТРАНЗИТ» оказало транспортные услуги по выполнению "
        f"международной перевозки груза, {transport_details} по маршруту: {route}, по CMR № {cmr_numbers}."
    )
    
    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.first_line_indent = Inches(0.4)
    p_b2.add_run(
        "На основании изложенного «Заказчик» и «Исполнитель» заявляют, что оказанные транспортные услуги выполнены в полном объеме, "
        "и в срок. Заказчик претензий по объёму, качеству и срокам оказания услуг не имеет."
    )
    
    p_pay = doc.add_paragraph()
    p_pay.paragraph_format.first_line_indent = Inches(0.4)
    p_pay.add_run(
        f"Указанную в заявке сумму в размере {amount_num} ({amount_words}) {currency}, в том числе НДС {nds} следует перечислить "
        f"{payment_terms} ООО«АВРОРА-ТРАНЗИТ» по следующим реквизитам:\n"
        "Номер счёта:\n"
        "р/с RUR: 40702810901130005079\n"
        "Банк: АО \"АЛЬФА-БАНК\"\n"
        "ИНН банка: 7728168971\n"
        "БИК: 044525593\n"
        "К/с: 30101810200000000593\n"
        "Адрес банка: 214004, Смоленская обл., г. Смоленск, ул. Николаева, д.8."
    )
    
    doc.add_paragraph("\n")
    
    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = f"{customer_name}\n{customer_details}"
    table.rows[0].cells[1].text = "ООО «АВРОРА-ТРАНЗИТ»\n14022, РФ г. Смоленск\nул. Карбышева д.15А, стр.2, помещ. К 52\nОГРН 1266700001501"
    
    table.rows[1].cells[0].text = f"\n\n___________________ {customer_signer}"
    table.rows[1].cells[1].text = "\n\n___________________ Галенда С.В."
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- КНОПКИ СКАЧИВАНИЯ ---
st.subheader("2. Сформировать и скачать файлы")

c_btn1, c_btn2 = st.columns(2)

with c_btn1:
    st.download_button(
        label="📄 Скачать Счёт (.docx)",
        data=generate_invoice(),
        file_name=f"Счет_{doc_num}_от_{doc_date}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with c_btn2:
    st.download_button(
        label="📝 Скачать Акт (.docx)",
        data=generate_act(),
        file_name=f"Акт_{doc_num}_от_{doc_date}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
