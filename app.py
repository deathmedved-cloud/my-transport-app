import streamlit as st

st.set_page_config(page_title="Автоматизация Актов", layout="wide")
st.title("🚚 Автоматизация Актов и Счетов")

st.subheader("1. Ручной ввод данных договора-заявки")

col1, col2 = st.columns(2)

with col1:
    number = st.text_input("Номер договора/заявки")
    date = st.date_input("Дата")
    executor = st.text_input("Исполнитель")

with col2:
    customer = st.text_input("Заказчик")
    amount = st.text_input("Сумма (руб.)")
    route = st.text_input("Маршрут")

if st.button("💾 Сохранить и сформировать документ"):
    st.success("Данные успешно внесены!")
    st.subheader("2. Результат")
    st.markdown(f"""
    * **Номер:** {number}
    * **Дата:** {date}
    * **Исполнитель:** {executor}
    * **Заказчик:** {customer}
    * **Сумма:** {amount} руб.
    * **Маршрут:** {route}
    """)
