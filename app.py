import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Генератор Актов", layout="wide")
st.title("🚚 Генератор Акта выполненных работ (по образцу)")

st.subheader("1. Параметры Акта")

col1, col2 = st.columns(2)

with col1:
    act_num = st.text_input("Номер Акта", value="14")
    act_date = st.text_input("Дата Акта", value="15.07.2026")
    order_num = st.text_input("Номер транспортного заказа", value="01-07")
    order_date = st.text_input("Дата заказа", value="10.07.2026")
    city = st.text_input("Город составления", value="г. Смоленск")
    
    st.markdown("---")
    customer_name = st.text_input("Заказчик (полное наименование)", value="Общество с ограниченной ответственностью «Егорчик»")
    customer_director = st.text_input("ФИО директора Заказчика (в родительном падеже)", value="Насирова Олега Холмуродовича")
    customer_basis = st.text_input("Основание Заказчика", value="Устава")
    customer_address = st.text_area("Реквизиты/Адрес Заказчика", value="230026, РБ, г. Гродно\nул. Победы 17\nУНП 591503882")
    customer_signer = st.text_input("Фамилия И.О. Заказчика под подпись", value="Насиров О.Х.")

with col2:
    route = st.text_input("Маршрут", value="РБ, ТЛЦ Брузги (место перецепки) - РБ")
    transport_details = st.text_input("Детали ТС / перецепки", value="в том числе на этапе перевозки после перецепки на ТЛЦ Брузги, ТС М226ЕО67 / 3TA7860")
    cmr_numbers = st.text_area("Номера CMR", value="R456159, MA456028, MA455671, MA453832, R455456, MA455721, VB455624, R455078, IB454302")
    
    st.markdown("---")
    amount_num = st.text_input("Сумма цифрами", value="450")
    amount_words = st.text_input("Сумма прописью", value="четыреста пятьдесят")
    currency = st.text_input("Валюта", value="евро")
    nds = st.text_input("Ставка НДС", value="0%")
    payment_terms = st.text_input("Условия конвертации/оплаты", value="в российских рублях по курсу ЦБ РФ на дату оплаты")

def generate_act():
    doc = Document()
    
    # Шапка
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = h.add_run(f"Акт выполненных работ № {act_num}\n")
    r1.bold = True
    r1.font.size = Pt(13)
    r2 = h.add_run("об оказании услуг\n")
    r2.bold = True
    r3 = h.add_run(f"по транспортному заказу № {order_num} от {order_date} г.")
    r3.bold = True
    
    # Город и дата
    p_meta = doc.add_paragraph()
    p_meta.add_run(f"{city}\t\t\t\t\t\t\t\t\t{act_date} г.").bold = True
    
    # Преамбула
    p_preamble = doc.add_paragraph()
    p_preamble.paragraph_format.first_line_indent = Inches(0.4)
    p_preamble.add_run(
        f"{customer_name}, в лице директора {customer_director}, действующей на основании {customer_basis}, "
        f"именуемое в дальнейшем «Заказчик» с одной стороны, и Общество с ограниченной ответственностью «АВРОРА-ТРАНЗИТ», "
        f"в лице директора Галенда Сергея Владимировича, действующего на основании Устава, именуемое в дальнейшем «Исполнитель» "
        f"с другой стороны, составили настоящий акт о нижеследующем:"
    )
    
    # Текст перевозки
    p_b1 = doc.add_paragraph()
    p_b1.paragraph_format.first_line_indent = Inches(0.4)
    p_b1.add_run(
        f"В установленные транспортным заказом сроки ООО «АВРОРА-ТРАНЗИТ» оказало транспортные услуги по выполнению "
        f"международной перевозки груза, {transport_details} по маршруту: {route}, по CMR № {cmr_numbers}."
    )
    
    # Отсутствие претензий
    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.first_line_indent = Inches(0.4)
    p_b2.add_run(
        "На основании изложенного «Заказчик» и «Исполнитель» заявляют, что оказанные транспортные услуги выполнены в полном объеме, "
        "и в срок. Заказчик претензий по объёму, качеству и срокам оказания услуг не имеет."
    )
    
    # Банковские реквизиты и сумма
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
    
    # Блок подписей
    table = doc.add_table(rows=2, cols=2)
    
    cell_c = table.rows[0].cells[0]
    cell_c.text = f"{customer_name}\n{customer_address}"
    
    cell_e = table.rows[0].cells[1]
    cell_e.text = "ООО «АВРОРА-ТРАНЗИТ»\n214022, РФ г. Смоленск\nул. Карбышева д.15А, стр.2, помещ. К 52\nОГРН 1266700001501"
    
    cell_sig_c = table.rows[1].cells[0]
    cell_sig_c.text = f"\n\n___________________ {customer_signer}"
    
    cell_sig_e = table.rows[1].cells[1]
    cell_sig_e.text = "\n\n___________________ Галенда С.В."
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.subheader("2. Сформировать документ")
st.download_button(
    label="📝 Скачать Акт выполненных работ (.docx)",
    data=generate_act(),
    file_name=f"Акт_{act_num}_от_{act_date}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

