from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


WORKBOOK_PATH = Path(__file__).with_name("Simulador_HC_Escala_5x2.xlsx")
STORE_SHEET = "Simulador por loja"
PREMISES_SHEET = "Premissas"
VALIDATION_SHEET = "Validacao 5x2"
STORE_COLUMNS = (
    "Canal",
    "Loja",
    "HC atual",
    "Líderes",
    "Unidades de cobertura simultâneas (input)",
    "Pessoas/turno seg-sex",
    "Pessoas/turno sáb-dom",
    "HC operacional c/reserva",
    "Liderança mínima sugerida",
    "HC recomendado estimado",
    "Diferença vs HC atual",
)


def load_public_data(
    path: Path = WORKBOOK_PATH,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any]]:
    if not path.is_file():
        raise FileNotFoundError("A base agregada do simulador não foi encontrada.")

    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        missing_sheets = {
            sheet
            for sheet in (STORE_SHEET, PREMISES_SHEET, VALIDATION_SHEET)
            if sheet not in workbook.sheetnames
        }
        if missing_sheets:
            raise ValueError(
                "Abas necessárias não encontradas: "
                + ", ".join(sorted(missing_sheets))
            )

        store_sheet = workbook[STORE_SHEET]
        rows = store_sheet.iter_rows(values_only=True)
        headers = next(rows, None)
        if headers is None:
            raise ValueError("A aba agregada de lojas está vazia.")
        header_indexes = {
            str(header).strip(): index
            for index, header in enumerate(headers)
            if header is not None
        }
        missing_columns = set(STORE_COLUMNS).difference(header_indexes)
        if missing_columns:
            raise ValueError(
                "Colunas agregadas necessárias não encontradas: "
                + ", ".join(sorted(missing_columns))
            )

        store_rows = []
        for row in rows:
            values = {
                column: row[header_indexes[column]]
                if header_indexes[column] < len(row)
                else None
                for column in STORE_COLUMNS
            }
            if not values["Canal"] or not values["Loja"]:
                continue
            if str(values["Loja"]).strip().casefold() == "total / consolidado":
                continue
            store_rows.append(values)
        if not store_rows:
            raise ValueError("Nenhuma loja foi encontrada na base agregada.")
        stores = pd.DataFrame(store_rows)

        for column in STORE_COLUMNS[2:]:
            stores[column] = pd.to_numeric(stores[column], errors="coerce")

        premise_sheet = workbook[PREMISES_SHEET]
        premise_rows = list(premise_sheet.iter_rows(values_only=True))
        premise_header_index = next(
            (
                index
                for index, row in enumerate(premise_rows)
                if len(row) > 2
                and row[0] == "Premissa"
                and row[1] == "Valor"
            ),
            None,
        )
        if premise_header_index is None:
            raise ValueError("O cabeçalho das premissas não foi encontrado.")
        premises = {}
        for row in premise_rows[premise_header_index + 1 :]:
            if not row[0] or row[0] == "Nota metodológica":
                continue
            if len(row) > 1 and row[1] is not None:
                premises[str(row[0]).strip()] = row[1]

        validation_sheet = workbook[VALIDATION_SHEET]
        validation = {
            "days_worked": validation_sheet["B4"].value,
            "days_off": validation_sheet["B5"].value,
            "premise_result": validation_sheet["B9"].value,
            "hc_current": validation_sheet["B13"].value,
            "active_count": validation_sheet["B14"].value,
            "headcount_difference": validation_sheet["B15"].value,
            "stores_valid": validation_sheet["G50"].value,
            "stores_review": validation_sheet["G51"].value,
        }
        return stores, premises, validation
    finally:
        workbook.close()


def main() -> None:
    st.set_page_config(
        page_title="Simulador de HC · Escala 5x2",
        page_icon="📅",
        layout="wide",
    )
    st.markdown(
        """
        <style>
        .stApp {
            color: #fff;
            background-image:
                linear-gradient(rgba(0, 30, 19, 0.45), rgba(0, 22, 14, 0.62)),
                url("/app/static/lacoste-background.png");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
            background-color: transparent;
        }
        .brand-logo {
            margin: 0 auto 2rem;
            text-align: center;
        }
        .brand-logo img {
            display: inline-block;
            width: min(260px, 55vw);
            height: auto;
        }
        [data-testid="stAppViewContainer"],
        [data-testid="stMarkdownContainer"],
        [data-testid="stWidgetLabel"],
        [data-testid="stCaptionContainer"],
        [data-testid="stExpander"] summary,
        [data-testid="stMetricLabel"],
        [data-testid="stMetricValue"] {
            color: #fff;
        }
        [data-testid="stCaptionContainer"] p,
        [data-testid="stMarkdownContainer"] p {
            color: #fff !important;
        }
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3 {
            color: #fff;
            text-shadow: 0 2px 8px rgba(0, 0, 0, 0.55);
        }
        [data-testid="stMetric"] {
            padding: 1rem 1.1rem;
            border: 1px solid rgba(255, 255, 255, 0.24);
            border-radius: 0.8rem;
            background: rgba(0, 39, 25, 0.82);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
        }
        div[data-baseweb="select"] > div {
            color: #fff;
            border-color: rgba(255, 255, 255, 0.45);
            background-color: rgba(0, 38, 25, 0.92);
        }
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] svg {
            color: #fff;
            fill: #fff;
        }
        [data-testid="stAlert"] {
            border: 1px solid rgba(255, 255, 255, 0.2);
            background-color: rgba(0, 38, 25, 0.94);
        }
        [data-testid="stExpander"] {
            border-color: rgba(255, 255, 255, 0.32);
            border-radius: 0.7rem;
            background: rgba(0, 35, 23, 0.84);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="brand-logo"><img src="/app/static/lacoste-logo.png" '
        'alt="Lacoste"></div>',
        unsafe_allow_html=True,
    )
    st.title("Simulador de HC · Escala 5x2")
    st.caption(
        "Visão pública agregada por loja. Não exibe nomes nem escalas individuais."
    )

    try:
        stores, premises, validation = load_public_data()
    except (FileNotFoundError, InvalidFileException, ValueError, KeyError) as error:
        st.error(f"Não foi possível carregar a base agregada do simulador: {error}")
        st.stop()

    channels = sorted(stores["Canal"].dropna().astype(str).unique())
    filters = st.columns(2)
    with filters[0]:
        selected_channel = st.selectbox("Canal", ["Todos", *channels])
    channel_stores = (
        stores
        if selected_channel == "Todos"
        else stores[stores["Canal"].eq(selected_channel)]
    )
    store_options = sorted(channel_stores["Loja"].astype(str).unique())
    with filters[1]:
        selected_store = st.selectbox("Loja", ["Todas", *store_options])
    visible = (
        channel_stores
        if selected_store == "Todas"
        else channel_stores[channel_stores["Loja"].eq(selected_store)]
    )

    summary = st.columns(5)
    summary[0].metric("Lojas no filtro", len(visible))
    summary[1].metric("HC atual", f"{visible['HC atual'].sum():,.0f}".replace(",", "."))
    summary[2].metric(
        "HC recomendado",
        f"{visible['HC recomendado estimado'].sum():,.0f}".replace(",", "."),
    )
    summary[3].metric(
        "Diferença estimada",
        f"{visible['Diferença vs HC atual'].sum():+,.0f}".replace(",", "."),
    )
    summary[4].metric(
        "Ativos na base",
        f"{int(validation['active_count']):,}".replace(",", "."),
    )

    days_worked = int(validation["days_worked"])
    days_off = int(validation["days_off"])
    weekly_coverage_days = int(premises["Dias de baixo fluxo por semana"]) + int(
        premises["Dias de fim de semana por semana"]
    )
    if (
        validation["premise_result"] == "5x2 OK"
        and days_worked == 5
        and days_off == 2
        and weekly_coverage_days == 7
        and validation["stores_review"] == 0
    ):
        st.success(
            "Validação aprovada: 5 dias trabalhados, 2 folgas e cobertura "
            f"considerada nos 7 dias da semana; {validation['stores_valid']} lojas "
            "com cálculos consistentes."
        )
    else:
        st.warning(
            "As premissas 5x2 ou os cálculos por loja precisam de revisão."
        )

    if validation["headcount_difference"]:
        st.warning(
            f"A base lista {int(validation['active_count'])} ativos, enquanto "
            f"o HC atual somado nas lojas é {int(validation['hc_current'])} "
            f"(diferença de {int(validation['headcount_difference'])}). "
            "Reconcilie as bases antes de tomar decisões de quadro."
        )

    st.subheader("Dimensionamento agregado por loja")
    display_columns = {
        "Canal": "Canal",
        "Loja": "Loja",
        "HC atual": "HC atual",
        "Líderes": "Líderes",
        "Unidades de cobertura simultâneas (input)": "Unidades simultâneas",
        "Pessoas/turno seg-sex": "Pessoas/turno seg-sex",
        "Pessoas/turno sáb-dom": "Pessoas/turno sáb-dom",
        "HC operacional c/reserva": "HC com reserva",
        "Liderança mínima sugerida": "Liderança mínima",
        "HC recomendado estimado": "HC recomendado",
        "Diferença vs HC atual": "Diferença",
    }
    table = visible.loc[:, list(display_columns)].rename(columns=display_columns)
    st.dataframe(table, hide_index=True, width="stretch")
    st.download_button(
        "Exportar resumo agregado · CSV",
        data=table.to_csv(index=False).encode("utf-8-sig"),
        file_name="simulador_hc_escala_5x2_agregado.csv",
        mime="text/csv",
        width="stretch",
    )

    with st.expander("Premissas consideradas"):
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Premissa": "Dias trabalhados por pessoa/semana",
                        "Valor": days_worked,
                    },
                    {
                        "Premissa": "Folgas por pessoa/semana",
                        "Valor": days_off,
                    },
                    {
                        "Premissa": "Jornada semanal (h)",
                        "Valor": premises.get("Jornada semanal (h)", "Não informada"),
                    },
                    {
                        "Premissa": "Dias de baixa cobertura por semana",
                        "Valor": premises["Dias de baixo fluxo por semana"],
                    },
                    {
                        "Premissa": "Dias de fim de semana por semana",
                        "Valor": premises["Dias de fim de semana por semana"],
                    },
                    {
                        "Premissa": "Reserva para férias/ausências",
                        "Valor": premises.get("Reserva para férias/ausências"),
                    },
                ]
            ),
            hide_index=True,
            width="stretch",
        )
    st.caption(
        "O app apresenta totais por loja, não uma escala nominal por colaborador."
    )


if __name__ == "__main__":
    main()
