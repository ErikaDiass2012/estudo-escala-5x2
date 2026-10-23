import calendar
from datetime import date
from pathlib import Path
import re
from typing import Any
import unicodedata

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


WORKBOOK_PATH = Path(__file__).with_name("Simulador_HC_Escala_5x2.xlsx")
ACTIVE_SHEET = "Ativos"
ASSUMPTIONS_SHEET = "Premissas"
WEEKDAYS = ("seg", "ter", "qua", "qui", "sex", "sáb", "dom")
OFF_DAY_PAIRS = ((6, 0), (0, 1), (1, 2), (2, 3), (3, 4))
MONTHS = (
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
)


def load_workbook_data() -> tuple[list[dict[str, str]], list[dict[str, Any]], str | None]:
    if not WORKBOOK_PATH.is_file():
        raise FileNotFoundError(
            f"Não foi encontrado o arquivo {WORKBOOK_PATH.name} na pasta do aplicativo."
        )

    workbook = load_workbook(WORKBOOK_PATH, data_only=True, read_only=True)
    try:
        missing_sheets = {
            sheet
            for sheet in (ACTIVE_SHEET, ASSUMPTIONS_SHEET)
            if sheet not in workbook.sheetnames
        }
        if missing_sheets:
            raise ValueError(
                "Abas não encontradas no arquivo: "
                + ", ".join(sorted(missing_sheets))
                + "."
            )

        active_sheet = workbook[ACTIVE_SHEET]
        active_rows = active_sheet.iter_rows(values_only=True)
        headers = next(active_rows, None)
        if headers is None or len(headers) < 5:
            raise ValueError(
                f"A aba “{ACTIVE_SHEET}” não contém as colunas esperadas "
                "nas posições A, B, D e E."
            )
        if (
            headers[0] != "Canal"
            or headers[1] not in ("Empresa", "Loja")
            or headers[3] not in ("Funcionário", "Colaborador", "Colaboradores")
            or headers[4] != "Cargo"
        ):
            raise ValueError(
                f"Os cabeçalhos da aba “{ACTIVE_SHEET}” devem identificar "
                "Canal na coluna A, Loja na B, Colaboradores na D e Cargo na E."
            )

        active_data: list[dict[str, str]] = []
        for row_number, row in enumerate(active_rows, start=2):
            if not any(value is not None for value in row):
                continue
            channel, store, employee, role = row[0], row[1], row[3], row[4]
            if not all(
                value is not None and str(value).strip()
                for value in (channel, store)
            ):
                raise ValueError(
                    f"A linha {row_number} da aba “{ACTIVE_SHEET}” está "
                    "sem canal ou loja nas colunas A ou B."
                )
            active_data.append(
                {
                    "Canal": str(channel).strip(),
                    "Loja": str(store).strip(),
                    "Colaborador": str(employee).strip() if employee else "",
                    "Cargo": str(role).strip() if role else "",
                }
            )
        if not active_data:
            raise ValueError(f"Nenhum colaborador foi encontrado na aba “{ACTIVE_SHEET}”.")

        assumptions_sheet = workbook[ASSUMPTIONS_SHEET]
        assumption_rows = list(assumptions_sheet.iter_rows(values_only=True))
        header_index = next(
            (
                index
                for index, row in enumerate(assumption_rows)
                if tuple(row[:2]) == ("Premissa", "Valor")
                and row[2] in ("Escala", "Orientação")
            ),
            None,
        )
        if header_index is None:
            raise ValueError(
                f"O cabeçalho de premissas não foi encontrado na aba "
                f"“{ASSUMPTIONS_SHEET}”."
            )

        assumption_columns = [
            str(column)
            for column in assumption_rows[header_index][:4]
            if column is not None
        ]
        assumptions: list[dict[str, Any]] = []
        methodology_note = None
        for row in assumption_rows[header_index + 1 :]:
            values = dict(zip(assumption_columns, row[: len(assumption_columns)]))
            premise = values.get("Premissa")
            if premise == "Nota metodológica":
                methodology_note = next(
                    (
                        str(next_row[0]).strip()
                        for next_row in assumption_rows[header_index + 2 :]
                        if next_row[0]
                        and str(next_row[0]).strip().startswith("O simulador calcula")
                    ),
                    None,
                )
                break
            if premise and any(
                value is not None
                for key, value in values.items()
                if key != "Premissa"
            ):
                assumptions.append(values)
        return active_data, assumptions, methodology_note
    finally:
        workbook.close()


def format_premise_value(value: Any, scale: Any) -> str:
    if value is None:
        return "—"
    if scale == "Horas" and isinstance(value, (int, float)):
        hours = int(value)
        minutes = round((value - hours) * 100)
        if minutes == 60:
            hours += 1
            minutes = 0
        return f"{hours}h{minutes:02d}"
    if isinstance(value, float):
        return f"{value:g}".replace(".", ",")
    return str(value)


def clean_store_name(store: str) -> str:
    name = re.sub(r"^\s*\d+\s*[-–]\s*", "", store).strip()
    return re.sub(r"^(?:BOUTIQUE|OUTLET)\s+", "", name, flags=re.IGNORECASE)


def is_leadership_role(role: str) -> bool:
    normalized_role = unicodedata.normalize("NFKD", role.casefold())
    normalized_role = "".join(
        character
        for character in normalized_role
        if not unicodedata.combining(character)
    )
    return any(
        title in normalized_role
        for title in ("gerente", "sub gerente", "respons", "lider operacional")
    )


def allocate_weekly_days_off(
    employees: list[dict[str, str]],
    leadership_ids: set[int],
    sunday_off_leader_id: int | None,
    week_index: int,
) -> dict[int, tuple[int, int]]:
    """Allocate consecutive days off for one week, rotating leaders' Sundays."""
    days_off_count = [0] * 7
    assignments: dict[int, tuple[int, int]] = {}
    shift_headcount = len(employees)
    ordered_employees = sorted(
        employees,
        key=lambda employee: (
            0 if id(employee) == sunday_off_leader_id else 1,
            0 if id(employee) in leadership_ids else 1,
            employee["Cargo"].casefold(),
            employee["Colaborador"].casefold(),
        ),
    )

    for employee_index, employee in enumerate(ordered_employees):
        employee_id = id(employee)
        if employee_id == sunday_off_leader_id:
            candidate_pairs = ((6, 0),)
        elif employee_id in leadership_ids:
            candidate_pairs = OFF_DAY_PAIRS[1:]
        else:
            candidate_pairs = OFF_DAY_PAIRS

        scored_pairs = []
        for pair_index, pair in enumerate(candidate_pairs):
            updated_days_off = days_off_count.copy()
            for weekday in pair:
                updated_days_off[weekday] += 1
            coverage_shortfall = sum(
                max(0, 2 - (shift_headcount - updated_days_off[weekday])) ** 2
                for weekday in range(4)
            )
            day_imbalance = sum(
                (
                    6 * updated_days_off[weekday]
                    - 2 * shift_headcount
                )
                ** 2
                for weekday in (0, 1, 2, 3, 4, 6)
            )
            tie_breaker = (
                pair_index - week_index - employee_index
            ) % len(candidate_pairs)
            scored_pairs.append(
                (
                    coverage_shortfall,
                    day_imbalance,
                    tie_breaker,
                    pair,
                    updated_days_off,
                )
            )

        _, _, _, selected_pair, days_off_count = min(
            scored_pairs,
            key=lambda scored_pair: scored_pair[:3],
        )
        assignments[employee_id] = selected_pair

    return assignments


def build_month_schedule(
    employees: list[dict[str, str]], year: int, month: int
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, list[str]]:
    days_in_month = calendar.monthrange(year, month)[1]
    day_columns = [
        f"{day:02d} {WEEKDAYS[calendar.weekday(year, month, day)]}"
        for day in range(1, days_in_month + 1)
    ]
    ordered_employees = [dict(employee) for employee in employees]
    for employee in ordered_employees:
        employee["Cargo ativo"] = employee["Cargo"]
    exceptions: list[str] = []
    if ordered_employees and ordered_employees[0]["Canal"].casefold() == "boutique":
        manager_count = 0
        responsible_seller_count = 0
        for employee in ordered_employees:
            role = employee["Cargo"]
            normalized_role = role.strip().casefold()
            if "gerente" in normalized_role:
                manager_count += 1
                if manager_count > 1:
                    employee["Cargo"] = "AJUSTE DE QUADRO — GERENTE DE LOJA"
            elif "vendedor" in normalized_role and "respons" in normalized_role:
                responsible_seller_count += 1
                if responsible_seller_count > 2:
                    employee["Cargo"] = "AJUSTE DE QUADRO — VENDEDOR RESPONSÁVEL"

        if manager_count > 1:
            exceptions.append(
                f"{manager_count} gerentes constam nos ativos; a posição de "
                "gerência está limitada a 1. Os excedentes foram mantidos "
                "na escala como ajuste de quadro."
            )
        if responsible_seller_count > 2:
            exceptions.append(
                f"{responsible_seller_count} vendedores responsáveis constam "
                "nos ativos; o limite é 2. Os excedentes foram mantidos na "
                "escala como ajuste de quadro."
            )

    ordered_employees.sort(
        key=lambda employee: (
            employee["Cargo"].casefold(),
            employee["Colaborador"].casefold(),
        )
    )
    employees_by_role: dict[str, list[dict[str, str]]] = {}
    for employee in ordered_employees:
        employees_by_role.setdefault(employee["Cargo"], []).append(employee)

    role_order = sorted(employees_by_role, key=str.casefold)
    interleaved_employees = [
        employees_by_role[role][index]
        for index in range(max(map(len, employees_by_role.values()), default=0))
        for role in role_order
        if index < len(employees_by_role[role])
    ]
    headcount = len(interleaved_employees)
    if headcount >= 8:
        t1_count = min(headcount - 4, max(3, (2 * headcount + 2) // 5))
    elif headcount >= 2:
        t1_count = headcount // 2
    else:
        t1_count = 0

    employees_by_shift = {
        "T1": interleaved_employees[:t1_count],
        "T2": interleaved_employees[t1_count:],
    }
    shift_assignments: dict[int, str] = {}
    for shift, shift_employees in employees_by_shift.items():
        for employee in shift_employees:
            shift_assignments[id(employee)] = shift

    leadership_ids = {
        id(employee)
        for employee in interleaved_employees
        if is_leadership_role(employee["Cargo ativo"])
    }
    leadership_employees = [
        employee
        for employee in interleaved_employees
        if id(employee) in leadership_ids
    ]
    if not leadership_employees:
        exceptions.append(
            "Não há colaborador ativo em cargo de liderança "
            "(gerente, subgerente, vendedor responsável ou líder operacional); "
            "não é possível garantir cobertura aos domingos."
        )

    calendar_weeks = calendar.monthcalendar(year, month)
    day_week_indexes = {
        day: week_index
        for week_index, calendar_week in enumerate(calendar_weeks)
        for day in calendar_week
        if day
    }
    weekly_off_pairs: dict[tuple[int, int], tuple[int, int]] = {}
    for week_index in range(len(calendar_weeks)):
        sunday_off_leader_id = None
        if len(leadership_employees) > 1:
            sunday_off_leader_id = id(
                leadership_employees[week_index % len(leadership_employees)]
            )
        for shift_employees in employees_by_shift.values():
            week_assignments = allocate_weekly_days_off(
                shift_employees,
                leadership_ids,
                sunday_off_leader_id,
                week_index,
            )
            for employee_id, off_pair in week_assignments.items():
                weekly_off_pairs[(employee_id, week_index)] = off_pair

    role_shift_counts: dict[str, dict[str, int]] = {}
    employee_rows: list[dict[str, Any]] = []
    coverage = {
        "T1 · manhã": [0] * days_in_month,
        "T2 · tarde": [0] * days_in_month,
    }
    leadership_sunday_coverage = [0] * days_in_month
    for employee in interleaved_employees:
        role = employee["Cargo"]
        source_role = employee["Cargo ativo"]
        shift = shift_assignments[id(employee)]
        role_shift_counts.setdefault(role, {"T1 · manhã": 0, "T2 · tarde": 0})
        role_shift_counts[role][f"{shift} · {'manhã' if shift == 'T1' else 'tarde'}"] += 1

        row: dict[str, Any] = {
            "Colaborador": employee["Colaborador"],
            "Cargo ativo": source_role,
            "Turno": shift,
        }
        for day in range(1, days_in_month + 1):
            weekday = calendar.weekday(year, month, day)
            week_index = day_week_indexes[day]
            off_pair = weekly_off_pairs[(id(employee), week_index)]
            on_leave = weekday in off_pair
            row[day_columns[day - 1]] = "F" if on_leave else shift
            if not on_leave:
                coverage[f"{shift} · {'manhã' if shift == 'T1' else 'tarde'}"][day - 1] += 1
                if weekday == 6 and id(employee) in leadership_ids:
                    leadership_sunday_coverage[day - 1] += 1
        employee_rows.append(row)

    role_rows = [
        {
            "Cargo": role,
            "T1 · manhã": counts["T1 · manhã"],
            "T2 · tarde": counts["T2 · tarde"],
            "Total": counts["T1 · manhã"] + counts["T2 · tarde"],
        }
        for role, counts in sorted(role_shift_counts.items(), key=lambda item: item[0].casefold())
    ]
    role_rows.append(
        {
            "Cargo": "TOTAL",
            "T1 · manhã": sum(counts["T1 · manhã"] for counts in role_shift_counts.values()),
            "T2 · tarde": sum(counts["T2 · tarde"] for counts in role_shift_counts.values()),
            "Total": len(ordered_employees),
        }
    )

    coverage_rows = [
        {"Turno": shift, **dict(zip(day_columns, counts))}
        for shift, counts in coverage.items()
    ]
    coverage_rows.append(
        {
            "Turno": "TOTAL",
            **{
                day_columns[index]: coverage["T1 · manhã"][index]
                + coverage["T2 · tarde"][index]
                for index in range(days_in_month)
            },
        }
    )
    for shift, daily_counts in coverage.items():
        understaffed_days = [
            day_columns[day - 1]
            for day in range(1, days_in_month + 1)
            if calendar.weekday(year, month, day) < 4
            and daily_counts[day - 1] < 2
        ]
        if understaffed_days:
            exceptions.append(
                f"Não é possível manter 2 pessoas no {shift} de segunda a "
                f"quinta nos dias: {', '.join(understaffed_days)}. "
                "Ajuste a quantidade de ativos ou a distribuição de folgas."
            )
    uncovered_sundays = [
        day_columns[day - 1]
        for day in range(1, days_in_month + 1)
        if calendar.weekday(year, month, day) == 6
        and leadership_employees
        and leadership_sunday_coverage[day - 1] == 0
    ]
    if uncovered_sundays:
        exceptions.append(
            "Não foi possível manter liderança trabalhando nos domingos: "
            f"{', '.join(uncovered_sundays)}."
        )

    return (
        pd.DataFrame(
            employee_rows,
            columns=("Colaborador", "Cargo ativo", "Turno", *day_columns),
        ),
        pd.DataFrame(role_rows),
        pd.DataFrame(coverage_rows),
        exceptions,
    )


def main() -> None:
    st.set_page_config(page_title="Escala 5x2 por loja", page_icon="📅", layout="wide")
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
        [data-testid="stAppViewContainer"] {
            color: #fff;
        }
        [data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"],
        [data-testid="stCaptionContainer"], [data-testid="stExpander"] summary,
        [data-testid="stMetricLabel"], [data-testid="stMetricValue"] {
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
    st.title("Escala 5x2 por loja")
    st.caption(
        "Escala mensal por cargo e turno, baseada nos colaboradores ativos "
        "da planilha."
    )

    try:
        active_data, assumptions_data, methodology_note = load_workbook_data()
    except (FileNotFoundError, InvalidFileException, ValueError) as error:
        st.error(str(error))
        st.stop()

    st.subheader("Escala mensal por loja")
    channels = sorted({employee["Canal"] for employee in active_data}, key=str.casefold)
    today = date.today()
    selector_columns = st.columns((1, 2, 1, 1))
    with selector_columns[0]:
        selected_channel = st.selectbox("Canal", channels)
    channel_employees = [
        employee for employee in active_data if employee["Canal"] == selected_channel
    ]
    store_labels = {
        employee["Loja"]: clean_store_name(employee["Loja"])
        for employee in channel_employees
    }
    label_to_store = {label: store for store, label in store_labels.items()}
    if len(label_to_store) != len(store_labels):
        label_to_store = {store: store for store in store_labels}
    with selector_columns[1]:
        selected_store_label = st.selectbox(
            "Loja",
            sorted(label_to_store, key=str.casefold),
        )
    selected_store = label_to_store[selected_store_label]
    with selector_columns[2]:
        selected_month = st.selectbox(
            "Mês",
            MONTHS,
            index=today.month - 1,
        )
    with selector_columns[3]:
        selected_year = st.selectbox(
            "Ano",
            list(range(today.year - 2, today.year + 3)),
            index=2,
        )

    employees = [employee for employee in channel_employees if employee["Loja"] == selected_store]
    month_number = MONTHS.index(selected_month) + 1
    schedule, role_summary, daily_coverage, role_exceptions = build_month_schedule(
        employees, selected_year, month_number
    )
    metric_columns = st.columns(3)
    metric_columns[0].metric("Colaboradores ativos", len(employees))
    metric_columns[1].metric("Cargos", role_summary.shape[0] - 1)
    metric_columns[2].metric("Mês", f"{selected_month.capitalize()} {selected_year}")

    for exception in role_exceptions:
        st.warning(exception)

    st.markdown("**Ativos por cargo e turno**")
    role_styler = role_summary.style.set_properties(
        **{
            "background-color": "#073d2b",
            "color": "#fff",
            "border": "1px solid #789487",
        }
    ).apply(
        lambda row: [
            "background-color: #ffeb00; color: #111; font-weight: bold"
            if row["Cargo"] == "TOTAL"
            else ""
            for _ in row
        ],
        axis=1,
    )
    st.dataframe(role_styler, hide_index=True, width="stretch")
    st.caption(
        "Cada linha do quadro representa um colaborador ativo; nomes e cargos "
        "são preservados conforme a aba “Ativos”. A escala distribui mais "
        "colaboradores no T2 (tarde), reforçando o período de maior fluxo, e "
        "busca manter ao menos 2 pessoas por turno de segunda a quinta. "
        "Quando o HC não permite atingir esse mínimo, os dias são sinalizados. "
        "As folgas são em pares consecutivos; ninguém folga aos sábados. "
        "Quando há mais de uma liderança ativa, as folgas de domingo se "
        "alternam entre elas para manter cobertura."
    )

    day_columns = list(schedule.columns[3:])
    schedule_styler = (
        schedule.style.map(
            lambda value: (
                "background-color: #00b050; color: #000"
                if value == "F"
                else "background-color: #fff; color: #000"
            ),
            subset=day_columns,
        )
        .set_properties(
            subset=["Colaborador", "Cargo ativo"],
            **{"background-color": "#f1f0e8", "color": "#12271d"},
        )
        .set_properties(
            subset=["Turno"],
            **{
                "background-color": "#dce9df",
                "color": "#063b29",
                "text-align": "center",
                "font-weight": "bold",
            },
        )
        .set_properties(subset=["Turno", *day_columns], **{"text-align": "center"})
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#1f3b67"),
                        ("color", "#fff"),
                        ("text-align", "center"),
                    ],
                },
                {"selector": "td", "props": [("border", "1px solid #777")]},
            ]
        )
    )
    day_config = {
        column: st.column_config.TextColumn(column, width="small")
        for column in day_columns
    }
    st.dataframe(
        schedule_styler,
        hide_index=True,
        width="stretch",
        height=680,
        column_config={
            "Cargo": st.column_config.TextColumn("Cargo", width="medium"),
            "Colaborador": st.column_config.TextColumn("Colaborador", width="medium"),
            "Cargo ativo": st.column_config.TextColumn("Cargo ativo", width="medium"),
            "Turno": st.column_config.TextColumn("Turno", width="small"),
            **day_config,
        },
    )

    st.markdown("**Colaboradores escalados por dia**")
    coverage_styler = daily_coverage.style.set_properties(
        **{
            "background-color": "#073d2b",
            "color": "#fff",
            "border": "1px solid #789487",
        }
    ).apply(
        lambda row: [
            "background-color: #ffeb00; color: #111; font-weight: bold"
            if row["Turno"] == "TOTAL"
            else ""
            for _ in row
        ],
        axis=1,
    )
    st.dataframe(
        coverage_styler,
        hide_index=True,
        width="stretch",
        column_config={
            column: st.column_config.NumberColumn(column, width="small")
            for column in daily_coverage.columns[1:]
        },
    )

    with st.expander("Premissas consideradas"):
        assumptions_data = [
            {
                **assumption,
                "Valor": format_premise_value(
                    assumption.get("Valor"), assumption.get("Escala")
                ),
            }
            for assumption in assumptions_data
        ]
        assumptions_styler = pd.DataFrame(assumptions_data).style.set_properties(
            **{
                "background-color": "#073d2b",
                "color": "#fff",
                "border": "1px solid #789487",
            }
        )
        st.dataframe(
            assumptions_styler,
            column_order=tuple(assumptions_data[0].keys())
            if assumptions_data
            else None,
            hide_index=True,
            width="stretch",
        )
        if methodology_note:
            st.info(methodology_note)


if __name__ == "__main__":
    main()