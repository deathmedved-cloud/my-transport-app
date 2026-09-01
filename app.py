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
        "app_num": "б/н",
        "app_date": "06.08.2026",
        "doc_num": "25",
        "doc_date": "10.08.2026",
        "city": "г. Смоленск",
        "customer_name": "ИП Саук Д.М.",
        "customer_director": "директора Саук Д.М.",
        "customer_basis": "свидетельства о государственной регистрации 590625501",
        "customer_details": "230026, РБ г. Гродно ул. Пестрака, 22-1, УНП 590625501",
        "customer_signer": "Саук Д.М.",
        "route": "РБ, ТЛЦ №2 г. Брест - РФ, Московская обл. г. Реутов",
        "transport_details": "ТС Р712НХ67 / АМ385467",
        "amount_num": "1 658,00",
        "amount_words": "Одна тысяча шестьсот пятьдесят восемь",
        "currency": "евро",
        "nds": "0%",
        "payment_note": "Оплата в рублях РФ по курсу НБ РФ на дату оплаты."
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
            
            model = genai.GenerativeModel(
                model_name="gemini-3.6-flash",
                generation_config={"temperature": 0.0}
            )

            with st.spinner("Нейросеть детально считывает документ..."):
                file_bytes = uploaded_file.getvalue()
                mime_type = uploaded_file.type
                if not mime_type:
                    mime_type, _ = mimetypes.guess_type(uploaded_file.name)
                if not mime_type:
                    mime_type = "application/octet-stream"

                prompt = """
                Шаг 1. Внимательно изучи и полностью распознай весь текст на изображении/скане/документе.
                
                Шаг 2. Извлеки данные В СТРОГОМ JSON-формате без лишних слов и без разметки markdown (без ```json).

                СТРОГИЕ ПРАВИЛА ОПРЕДЕЛЕНИЯ СТОРОН И ДАННЫХ:
                1. ИСПОЛНИТЕЛЬ (Перевозчик) — это ВСЕГДА ООО «АВРОРА-ТРАНЗИТ».
                2. ЗАКАЗЧИК — это ВТОРАЯ сторона по договору/заявке (клиент, ИП, экспедитор, плательщик).
                3. "app_num" — номер договора/заявки (например "б/н" или "25").
                4. "app_date" — дата договора/заявки.
                5. "doc_num" — номер формируемого акта/счета.
                6. "doc_date" — дата акта/счета.
                7. "customer_director" — в лице кого действуют со стороны Заказчика (например "директора Саук Д.М." или "Иванова И.И.").
                8. "customer_basis" — основание полномочий (например "свидетельства о государственной регистрации 590625501" или "Устава").
                9. "payment_note" — фраза об условиях расчетов (например "Оплата в рублях РФ по курсу НБ РФ на дату оплаты.").

                Структура JSON:
                {
                    "app_num": "номер договора-заявки",
                    "app_date": "дата договора-заявки",
                    "doc_num": "номер акта и счета",
                    "doc_date": "дата акта и счета",
                    "city": "г. Смоленск",
                    "customer_name": "название Заказчика",
                    "customer_director": "в лице кого (напр.: директора Саук Д.М.)",
                    "customer_basis": "действующего на основании (напр.: Устава / свидетельства...)",
                    "customer_details": "адрес, УНП/ИНН Заказчика",
                    "customer_signer": "Фамилия И.О. подписывающего",
                    "route": "Маршрут перевозки",
                    "transport_details": "Информация по ТС и авто (напр.: ТС Р712НХ67 / АМ385467)",
                    "amount_num": "сумма цифрами (напр.: 1 658,00)",
                    "amount_words": "сумма прописью",
                    "currency": "валюта (евро / белорусских рублей / RUR)",
                    "nds": "0%",
                    "payment_note": "условия курса/оплаты (если есть)"
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
st.subheader("2. Проверка и редактирование данных")

data = st.session_state.doc_data
col1, col2 = st.columns(2)

with col1:
    app_num = st.text_input("Номер договора-заявки", value=data.get("app_num", "б/н"))
    app_date = st.text_input("Дата договора-заявки", value=data.get("app_date", ""))
    doc_num = st.text_input("Номер Акта и Счета", value=data.get("doc_num", "25"))
    doc_date = st.text_input("Дата Акта и Счета", value=data.get("doc_date", ""))
    city = st.text_input("Город составления", value=data.get("city", "г. Смоленск"))
    
    st.markdown("---")
    customer_name = st.text_input("Заказчик (Название)", value=data.get("customer_name", ""))
    customer_director = st.text_input("В лице кого (Руководитель)", value=data.get("customer_director", ""))
    customer_basis = st.text_input("Действующего на основании", value=data.get("customer_basis", "Устава"))
    customer_details = st.text_area("Реквизиты и адрес Заказчика", value=data.get("customer_details", ""))
    customer_signer = st.text_input("Подпись Заказчика (ФИО)", value=data.get("customer_signer", ""))

with col2:
    route = st.text_input("Маршрут", value=data.get("route", ""))
    transport_details = st.text_input("Детали ТС", value=data.get("transport_details", ""))
    
    st.markdown("---")
    amount_num = st.text_input("Сумма цифрами", value=data.get("amount_num", ""))
    amount_words = st.text_input("Сумма прописью", value=data.get("amount_words", ""))
    currency = st.text_input("Валюта", value=data.get("currency", "евро"))
    nds = st.text_input("НДС", value=data.get("nds", "0%"))
    payment_note = st.text_input("Условия оплаты/курса", value=data.get("payment_note", ""))

# --- ФУНКЦИЯ ГЕНЕРАЦИИ СЧЕТА ПО ОБРАЗЦУ ---
def build_invoice_doc():
    doc = Document()
    
    # Шапка банка (Таблица 2x2)
    table_bank = doc.add_table(rows=2, cols=2)
    table_bank.style = 'Table Grid'
    
    table_bank.rows[0].cells[0].text = "АО \"АЛЬФА-БАНК\"\n\nБанк получателя"
    table_bank.rows[0].cells[1].text = "БИК: 044525593\nСч. №: 30101810200000000593"
    table_bank.rows[1].cells[0].text = "ИНН 6700042504    КПП 670001001\n\nООО \"АВРОРА-ТРАНЗИТ\"\n\nПолучатель"
    table_bank.rows[1].cells[1].text = "Сч. №: 40702810901130005079"

    doc.add_paragraph()
    
    # Заголовок
    h = doc.add_paragraph(f"Счет на оплату № {doc_num} от {doc_date} г.")
    h.runs[0].bold = True
    h.runs[0].font.size = Pt(14)

    # Исполнитель / Заказчик / Комментарий
    p_party = doc.add_paragraph()
    p_party.add_run("Исполнитель: ").bold = True
    p_party.add_run("ООО \"АВРОРА-ТРАНЗИТ\", 214022, Смоленская область, г. о. город Смоленск, г. Смоленск, ул. Карбышева, д. 15А, стр. 2, помещ. К 52\n")
    p_party.add_run("Заказчик: ").bold = True
    p_party.add_run(f"{customer_name}, {customer_details}\n")
    p_party.add_run("Комментарий: ").bold = True
    p_party.add_run(f"Договор-заявка {app_num} от {app_date}г.\n{payment_note}")

    # Таблица услуг (7 колонок по образцу)
    t_services = doc.add_table(rows=2, cols=7)
    t_services.style = 'Table Grid'
    
    headers = ["№", "Название услуги", "Кол-во", "Ед.изм.", "Цена", "НДС", "Сумма"]
    for i, title in enumerate(headers):
        cell = t_services.rows[0].cells[i]
        cell.text = title
        cell.paragraphs[0].runs[0].bold = True
        
    row = t_services.rows[1].cells
    row[0].text = "1"
    row[1].text = f"Автомобильная перевозка груза по маршруту:\n{route}"
    row[2].text = "1"
    row[3].text = "шт."
    row[4].text = amount_num
    row[5].text = nds
    row[6].text = amount_num

    doc.add_paragraph()
    
    # Итоги
    p_total_text = doc.add_paragraph()
    p_total_text.add_run(f"Всего наименований 1 на сумму {amount_num} {currency}\n").bold = True
    p_total_text.add_run(f"{amount_words} {currency}").italic = True

    p_summary = doc.add_paragraph()
    p_summary.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_summary.add_run(f"Итого:  {amount_num}\nСумма НДС:  {nds}\nВсего к оплате:  {amount_num} {currency}").bold = True

    doc.add_paragraph("\n\n(должность) _________________ (подпись) _________________ (расшифровка подписи) / Галенда С.В. /")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- ФУНКЦИЯ ГЕНЕРАЦИИ АКТА ПО ОБРАЗЦУ ---
def build_act_doc():
    doc = Document()
    
    # Заголовок Акта
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = h.add_run(f"Акт выполненных работ № {doc_num}\nоб оказании транспортных услуг\nпо договору {app_num} от {app_date}г.")
    r1.bold = True
    r1.font.size = Pt(13)
    
    # Город и дата
    p_meta = doc.add_paragraph()
    p_meta.add_run(f"{city}\t\t\t\t\t\t\t\t\t\t{doc_date}г.").bold = True
    
    # Преамбула
    p_preamble = doc.add_paragraph()
    p_preamble.paragraph_format.first_line_indent = Inches(0.3)
    p_preamble.add_run(
        f"{customer_name}, в лице {customer_director} действующего на основании {customer_basis}, "
        f"именуемое в дальнейшем Заказчик с одной стороны, и ООО«АВРОРА-ТРАНЗИТ», в лице директора Галенда С.В., "
        f"действующего на основании Устава, именуемое в дальнейшем Исполнитель с другой стороны, "
        f"составили настоящий акт о нижеследующем:"
    )
    
    # Пункт 1
    p_b1 = doc.add_paragraph()
    p_b1.paragraph_format.first_line_indent = Inches(0.3)
    p_b1.add_run(
        f"В установленные Договором сроки ООО«АВРОРА-ТРАНЗИТ» оказало транспортные услуги по выполнению "
        f"международной перевозки груза, {transport_details}, по маршруту: {route}."
    )
    
    # Пункт 2
    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.first_line_indent = Inches(0.3)
    p_b2.add_run(
        "На основании изложенного Заказчик и Исполнитель заявляют, что оказанные транспортные услуги выполнены в полном объеме, "
        "и в срок. Заказчик претензий по объёму, качеству и срокам оказания услуг не имеет."
    )
    
    # Реквизиты для оплаты
    p_pay = doc.add_paragraph()
    p_pay.paragraph_format.first_line_indent = Inches(0.3)
    p_pay.add_run(
        f"Указанную в договоре-заявке сумму в размере {amount_num} ({amount_words}) {currency}, "
        f"в том числе НДС {nds} следует перечислить ООО«АВРОРА-ТРАНЗИТ» по следующим реквизитам:\n"
        "Номер счёта: 40702810901130005079\n"
        "Валюта: RUR\n"
        "Банк: АО \"АЛЬФА-БАНК\"\n"
        "ИНН банка: 7728168971\n"
        "БИК: 044525593\n"
        "Адрес банка: 214004, Смоленская обл., г. Смоленск, ул. Николаева, д.8.\n"
        f"{payment_note}"
    )
    
    doc.add_paragraph("\n")
    
    # Подписи сторон
    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = f"{customer_name}\n{customer_details}"
    table.rows[0].cells[1].text = "ООО «АВРОРА-ТРАНЗИТ»\n214022, РФ г. Смоленск\nул. Карбышева д. 15А, стр.2, помещ. К 52\nОГРН 1266700001501"
    
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
