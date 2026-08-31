import io
import json
import os
import tempfile
import subprocess
import mimetypes
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Автоматизация Документов", layout="wide")
st.title("🚚 Автоматическое формирование Актов и Счетов")

# --- НАСТРОЙКА КЛЮЧА GEMINI ---
with st.sidebar:
    st.header("⚙️ Настройки AI")
    default_key = st.secrets.get("GEMINI_API_KEY", "")
    api_key = st.text_input("Gemini API Key", value=default_key, type="password")

# --- РАСПОЗНАВАНИЕ ЧЕРЕЗ GEMINI ---
st.subheader("1. Загрузка документа (любой формат)")
uploaded_file = st.file_uploader("Загрузите файл документа (PDF, фото, DOCX, TXT и др.)")

if "doc_data" not in st.session_state:
    st.session_state.doc_data = {
        "doc_num": "20",
        "doc_date": "23.07.2026",
        "order_num": "20",
        "order_date": "23.07.2026",
        "city": "г. Смоленск",
        "customer_name": "ООО «ФорГлэйд»",
        "customer_director": "Осипчика М.П.",
        "customer_details": "УНП: 590940553, РБ 231940, Гродненская обл, Зельвенский р-н, д. Бережки 1В",
        "customer_signer": "Осипчик М.П.",
        "route": "Московская обл., г. Щелково — д. Бережки, Зельвенский р-н",
        "transport_details": "Авто: Renault Premium 450.19T/SCHMITZ SPR 24/L-13.62EB, Гос. номер: Р712HX67/AM385467, Водитель: Галенда Сергей Владимирович",
        "cmr_numbers": "б/н",
        "amount_num": "1 080,00",
        "amount_words": "Одна тысяча восемьдесят",
        "currency": "белорусских рублей",
        "nds": "0%",
        "payment_terms": "в день предоставления акта выполненных работ"
    }

if st.button("🤖 Распознать документ через AI"):
    final_key = api_key or default_key
    if not final_key:
        st.error("Укажите Gemini API Key!")
    elif not uploaded_file:
        st.warning("Сначала выберите файл для загрузки.")
    else:
        try:
            genai.configure(api_key=final_key)
            model = genai.GenerativeModel("gemini-3.6-flash")

            with st.spinner("Нейросеть считывает данные из файла..."):
                file_bytes = uploaded_file.getvalue()
                mime_type = uploaded_file.type
                if not mime_type:
                    mime_type, _ = mimetypes.guess_type(uploaded_file.name)
                if not mime_type:
                    mime_type = "application/octet-stream"

                prompt = """
                Проанализируй документ (заявку/договор на грузоперевозку) и верни данные В СТРОГОМ JSON-формате без лишних слов и без разметки markdown.
                Структура JSON:
                {
                    "doc_num": "номер заявки/договора",
                    "doc_date": "дата документа в формате ДД.ММ.ГГГГ",
                    "order_num": "номер заказа",
                    "order_date": "дата заказа",
                    "city": "г. Смоленск",
                    "customer_name": "название Заказчика",
                    "customer_director": "ФИО директора Заказчика в родительном падеже",
                    "customer_details": "УНП, юридический адрес и банковские реквизиты Заказчика",
                    "customer_signer": "Фамилия И.О. подписывающего со стороны Заказчика",
                    "route": "Маршрут перевозки (откуда - куда)",
                    "transport_details": "Марка, гос. номер авто, прицепа и ФИО водителя",
                    "cmr_numbers": "номера CMR (если указаны, иначе б/н)",
                    "amount_num": "сумма цифрами",
                    "amount_words": "сумма прописью",
                    "currency": "валюта платежа",
                    "nds": "ставка НДС",
                    "payment_terms": "условия и сроки оплаты"
                }
                """

                response = model.generate_content([
                    {"mime_type": mime_type, "data": file_bytes},
                    prompt
                ])

                raw_text = response.text.strip().replace("```json", "").replace("```", "")
                parsed_json = json.loads(raw_text)
                st.session_state.doc_data.update(parsed_json)
                st.success("Данные успешно извлечены!")

        except Exception as e:
            st.error(f"Ошибка распознавания: {e}")

# --- ФОРМА ПРОВЕРКИ И РЕДАКТИРОВАНИЯ ---
st.subheader("2. Проверка распознанных данных")

data = st.session_state.doc_data
col1, col2 = st.columns(2)

with col1:
    doc_num = st.text_input("Номер заявки / акта / счета", value=data.get("doc_num", ""))
    doc_date = st.text_input("Дата документа", value=data.get("doc_date", ""))
    order_num = st.text_input("Номер заказа", value=data.get("order_num", ""))
    order_date = st.text_input("Дата заказа", value=data.get("order_date", ""))
    city = st.text_input("Город составления", value=data.get("city", "г. Смоленск"))
    
    st.markdown("---")
    customer_name = st.text_input("Заказчик", value=data.get("customer_name", ""))
    customer_director = st.text_input("Директор Заказчика (в родительном падеже)", value=data.get("customer_director", ""))
    customer_details = st.text_area("Реквизиты Заказчика", value=data.get("customer_details", ""))
    customer_signer = st.text_input("Подпись Заказчика", value=data.get("customer_signer", ""))

with col2:
    route = st.text_input("Маршрут", value=data.get("route", ""))
    transport_details = st.text_input("Детали ТС и водитель", value=data.get("transport_details", ""))
    cmr_numbers = st.text_input("Номера CMR", value=data.get("cmr_numbers", "б/н"))
    
    st.markdown("---")
    amount_num = st.text_input("Сумма цифрами", value=data.get("amount_num", ""))
    amount_words = st.text_input("Сумма прописью", value=data.get("amount_words", ""))
    currency = st.text_input("Валюта", value=data.get("currency", "белорусских рублей"))
    nds = st.text_input("НДС", value=data.get("nds", "0%"))
    payment_terms = st.text_input("Условия оплаты", value=data.get("payment_terms", ""))

# --- ФУНКЦИИ ГЕНЕРАЦИИ СЧЕТА И АКТА ---
def build_invoice_doc():
    doc = Document()
    table_bank = doc.add_table(rows=2, cols=2)
    table_bank.style = 'Table Grid'
    
    table_bank.rows[0].cells[0].text = "АО \"АЛЬФА-БАНК\"\nБанк получателя"
    table_bank.rows[0].cells[1].text = "БИК: 044525593\nСч.№: 30101810200000000593"
    table_bank.rows[1].cells[0].text = "ИНН: 6700042504 | КПП: 670001001\nООО \"АВРОРА-ТРАНЗИТ\"\nПолучатель"
    table_bank.rows[1].cells[1].text = "Сч.№: 40702810901130005079"

    doc.add_paragraph()
    p_party = doc.add_paragraph()
    p_party.add_run("Исполнитель: ").bold = True
    p_party.add_run("ООО \"АВРОРА-ТРАНЗИТ\", 214022, РОССИЯ, Смоленская область, Смоленск, ул Карбышева, 15а, 2, К 52\n")
    p_party.add_run("Заказчик: ").bold = True
    p_party.add_run(f"{customer_name}, {customer_details}\n")
    p_party.add_run("Комментарий: ").bold = True
    p_party.add_run(f"Договор-заявка № {doc_num} от {doc_date}г. {payment_terms}.")

    h = doc.add_paragraph(f"\nСчет на оплату №{doc_num} от {doc_date} г.")
    h.runs[0].bold = True
    h.runs[0].font.size = Pt(14)

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

def build_act_doc():
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
    table.rows[0].cells[1].text = "ООО «АВРОРА-ТРАНЗИТ»\n214022, РФ г. Смоленск\nул. Карбышева д.15А, стр.2, помещ. К 52\nОГРН 1266700001501"
    
    table.rows[1].cells[0].text = f"\n\n___________________ {customer_signer}"
    table.rows[1].cells[1].text = "\n\n___________________ Галенда С.В."
    
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- КОНВЕРТЕР ФОРМАТОВ ---
def convert_doc(doc_buf, target_fmt):
    docx_bytes = doc_buf.getvalue()
    
    if target_fmt == "PDF":
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in:
                tmp_in.write(docx_bytes)
                tmp_in_path = tmp_in.name
            
            out_dir = tempfile.gettempdir()
            subprocess.run(["soffice", "--headless", "--convert-to", "pdf", tmp_in_path, "--outdir", out_dir], check=True)
            pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(tmp_in_path))[0] + ".pdf")
            
            with open(pdf_path, "rb") as f:
                res_bytes = f.read()
            return res_bytes, "application/pdf", "pdf"
        except Exception:
            st.error("Ошибка генерации PDF. Убедитесь, что добавлен файл packages.txt с 'libreoffice'.")
            return docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx"
            
    elif target_fmt == "TXT":
        doc = Document(io.BytesIO(docx_bytes))
        lines = [p.text for p in doc.paragraphs if p.text.strip()]
        for t in doc.tables:
            for row in t.rows:
                lines.append(" | ".join([cell.text.replace('\n', ' ') for cell in row.cells]))
        return "\n".join(lines).encode('utf-8'), "text/plain", "txt"
        
    else: # DOCX
        return docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx"

# --- КНОПКИ СКАЧИВАНИЯ ---
st.subheader("3. Сформировать и скачать файлы")

output_format = st.selectbox("Выберите формат для скачивания:", ["DOCX (.docx)", "PDF (.pdf)", "TXT (.txt)"])
fmt_code = output_format.split()[0]

c_btn1, c_btn2 = st.columns(2)

invoice_data, inv_mime, inv_ext = convert_doc(build_invoice_doc(), fmt_code)
act_data, act_mime, act_ext = convert_doc(build_act_doc(), fmt_code)

with c_btn1:
    st.download_button(
        label=f"📄 Скачать Счёт ({fmt_code})",
        data=invoice_data,
        file_name=f"Счет_{doc_num}_от_{doc_date}.{inv_ext}",
        mime=inv_mime
    )

with c_btn2:
    st.download_button(
        label=f"📝 Скачать Акт ({fmt_code})",
        data=act_data,
        file_name=f"Акт_{doc_num}_от_{doc_date}.{act_ext}",
        mime=act_mime
    )
