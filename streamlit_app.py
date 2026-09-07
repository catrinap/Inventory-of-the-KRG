import streamlit as st
from report_builder import build_report, ReportBuildError, SHEET_DATA, SHEET_REPORT, SHEET_RESULT

st.set_page_config(page_title="Автоматизация отчёта по ТГ (КРГ)", layout="centered")

st.title("Автоматизация «Отчёт по ТГ» с учётом КРГ")

st.markdown(
    f"""
Загрузите Excel-файл, в котором есть листы:

- **«{SHEET_DATA}»** — построчные данные пересчёта (нужны столбцы
  `ТоварнаяГруппа`, `Дельта`, `СуммаРозница`, `СуммаПриход`);
- **«{SHEET_REPORT}»** — готовый «Отчёт по ТГ» в виде Excel-таблицы, где есть
  столбцы для дельты КРГ (`Дельта КРГ, шт`, `Дельта КРГ в средневзвеш.розн.
  ценах, руб.`, `Дельта КРГ в приходных ценах, руб.`).

Приложение само:
1. Построит сводную по «{SHEET_DATA}» в разрезе Товарной группы (сумма
   Дельта / СуммаРозница / СуммаПриход).
2. Создаст лист **«{SHEET_RESULT}»** — точную копию «{SHEET_REPORT}» со всеми
   формулами и форматированием.
3. Впишет сводные цифры в столбцы дельты КРГ по соответствующим Товарным
   группам.
4. К фактическим остаткам (шт., средневзвеш. розница, приход) прибавит
   дельту КРГ формулой — существующие формулы никуда не денутся.
"""
)

uploaded = st.file_uploader("Excel-файл (.xlsx)", type=["xlsx"])

if uploaded is not None:
    st.write(f"Файл: **{uploaded.name}**")
    if st.button("Собрать «Готовый отчёт по ТГ»", type="primary"):
        with st.spinner("Считаю сводную и собираю отчёт..."):
            try:
                result_bytes, info = build_report(uploaded.getvalue())
            except ReportBuildError as e:
                st.error(f"Не получилось: {e}")
            except Exception as e:  # noqa: BLE001
                st.error(f"Непредвиденная ошибка: {e}")
            else:
                st.success("Готово! Лист «Готовый отчёт по ТГ» сформирован.")

                if info["matched"]:
                    st.write(f"✅ Данные КРГ учтены по {len(info['matched'])} товарным группам:")
                    st.write(", ".join(info["matched"]))

                if info["not_in_report"]:
                    st.warning(
                        "⚠️ В «Данные КРГ» встретились товарные группы, которых нет "
                        "в таблице «Отчёт по ТГ» — их дельта нигде не учтена:\n\n"
                        + ", ".join(info["not_in_report"])
                    )

                out_name = uploaded.name.rsplit(".", 1)[0] + "_готовый_отчет.xlsx"
                st.download_button(
                    "⬇️ Скачать файл с готовым отчётом",
                    data=result_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.divider()
st.caption(
    "Запуск локально: `pip install streamlit openpyxl` затем "
    "`streamlit run streamlit_app.py`. "

)
