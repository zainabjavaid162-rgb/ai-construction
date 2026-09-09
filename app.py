
import os
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from groq import Groq
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# AI CONSTRUCTION ASSISTANT
# Streamlit Cloud ready - NO Plotly dependency
#
# Source workbook:
#   BHATTI 7 MARLA CPLAN.xlsx
# Source worksheet:
#   5.32C-full costr cost-HB
# ============================================================

APP_TITLE = "AI Construction Assistant"
SOURCE_FILE = "BHATTI 7 MARLA CPLAN.xlsx"
SOURCE_SHEET = "5.32C-full costr cost-HB"

MARLA_SQFT = 272.25

# The supplied worksheet is based on 30 x 40 = 1,200 sq.ft
# and contains Basement + GF + 1st + 2nd + Top = 5 reference levels.
REFERENCE_FLOOR_SQFT = 1200.0
REFERENCE_LEVELS = 5
REFERENCE_TOTAL_SQFT = REFERENCE_FLOOR_SQFT * REFERENCE_LEVELS


# ============================================================
# Utility functions
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def to_number(value):
    try:
        if value is None or str(value).strip() == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def money(value):
    try:
        return f"PKR {float(value):,.0f}"
    except (TypeError, ValueError):
        return "PKR 0"


def infer_category(row_number, description):
    """Map source worksheet rows into BOQ categories."""
    d = description.lower()

    if 11 <= row_number <= 22:
        if row_number <= 16:
            return "Waterproofing & Site Works"
        return "Excavation & Grey Labour"

    if 26 <= row_number <= 31:
        return "Grey Structure"

    if 39 <= row_number <= 66:
        return "Plumbing"

    if 71 <= row_number <= 104:
        return "Electrical"

    if 109 <= row_number <= 137:
        if any(x in d for x in [
            "door", "window", "wood", "wardrobe", "cabinet"
        ]):
            return "Doors, Windows & Woodwork"
        if any(x in d for x in [
            "metal", "gate", "railing", "staircase"
        ]):
            return "Metal Works"
        return "Finishing"

    if 142 <= row_number <= 162:
        return "Fittings & Fixtures"

    return "Other"


def is_heading_or_total(description):
    """Skip section headings and total rows from the source sheet."""
    d = description.upper().strip()

    if not d:
        return True

    if d.startswith("TOTAL "):
        return True

    headings = [
        "ROOF INSULATION",
        "EXCAVATION & BACKFILL",
        "BRICKS/CEMENT/SAND/KRUSH/IRON",
        "PLUMBING WORKS",
        "ELECTRICAL WORKS",
        "WOOD, METAL",
        "FITTING AND FIXTURE",
        "CONSTRUCTION COST",
        "COLD AND HOT WATER",
        "INSULATION ARMAFLEX",
        "DRAINAGE SYSTEM PIPING",
        "GAS PIPING",
        "PUMPS",
        "AC DRAINAGE SYSTEM",
        "UPVC PIPELINE",
        "SWITCHES AND BACK BOXES",
        "CONDUITS & ACCESSORIES",
        "WIRES & CABLES",
        "LIGHTS, FANS",
        "PAINT AND CEILING",
        "TILE WORK",
        "WOODEN DOORS",
        "WINDOWS",
        "METAL WORKS",
        "WARDROBES & CABINETS",
        "KITCHEN",
        "BATHROOM",
        "GEYSER & MANHOLE",
    ]

    return any(d.startswith(h) for h in headings)


# ============================================================
# Excel source handling
# ============================================================

def load_source_workbook(uploaded_file=None):
    """
    Load the exact supplied workbook and exact named worksheet.
    """
    if uploaded_file is not None:
        raw = uploaded_file.getvalue()
        values_wb = load_workbook(BytesIO(raw), data_only=True)
        formulas_wb = load_workbook(BytesIO(raw), data_only=False)
    else:
        if not os.path.exists(SOURCE_FILE):
            return None, None, (
                f"'{SOURCE_FILE}' was not found. Upload the workbook "
                "using the sidebar."
            )

        values_wb = load_workbook(SOURCE_FILE, data_only=True)
        formulas_wb = load_workbook(SOURCE_FILE, data_only=False)

    if SOURCE_SHEET not in values_wb.sheetnames:
        return None, None, (
            f"The workbook does not contain '{SOURCE_SHEET}'. "
            f"Available sheets: {', '.join(values_wb.sheetnames)}"
        )

    return (
        values_wb[SOURCE_SHEET],
        formulas_wb[SOURCE_SHEET],
        None,
    )


def extract_reference_items(ws_values, ws_formulas):
    """
    Read the source BOQ from columns C:G:
        C = Description
        D = Unit
        E = Rate
        F = Quantity
        G = Amount

    Broken #REF! formulas are reported rather than guessed.
    """
    items = []
    warnings = []

    for row in range(1, ws_values.max_row + 1):
        description = clean_text(ws_values.cell(row, 3).value)
        unit = clean_text(ws_values.cell(row, 4).value)

        rate = to_number(ws_values.cell(row, 5).value)
        quantity = to_number(ws_values.cell(row, 6).value)
        amount = to_number(ws_values.cell(row, 7).value)

        formula_quantity = clean_text(ws_formulas.cell(row, 6).value)
        formula_amount = clean_text(ws_formulas.cell(row, 7).value)

        if is_heading_or_total(description):
            continue

        if not unit:
            continue

        if "#REF!" in formula_quantity or "#REF!" in formula_amount:
            warnings.append({
                "Source Row": row,
                "Description": description,
                "Quantity Formula": formula_quantity,
                "Amount Formula": formula_amount,
            })

        # If rate is missing but amount exists, preserve the row as a
        # reference lump sum rather than inventing a unit rate.
        lump_sum = False

        if rate is None and amount is not None and amount > 0:
            lump_sum = True
            reference_quantity = 1.0
            reference_rate = amount
        else:
            reference_quantity = (
                quantity if quantity is not None else 0.0
            )
            reference_rate = rate if rate is not None else 0.0

        if reference_quantity <= 0 and not lump_sum:
            continue

        items.append({
            "source_row": row,
            "category": infer_category(row, description),
            "description": description,
            "unit": unit,
            "rate": reference_rate,
            "reference_quantity": reference_quantity,
            "reference_amount": (
                amount
                if amount is not None
                else reference_quantity * reference_rate
            ),
            "lump_sum": lump_sum,
        })

    return items, warnings


# ============================================================
# BOQ calculation
# ============================================================

def room_adjustment(
    item,
    base_quantity,
    bedrooms,
    washrooms,
    drawing_rooms,
    lobby,
    laundry,
    basement,
):
    """
    Residential adjustments for naturally room-dependent items.

    Area-based construction quantities continue to scale from the
    supplied reference worksheet.
    """
    d = item["description"].lower()
    category = item["category"]

    # Basement-related items are removed when basement is not selected.
    if "basement" in d and not basement:
        return 0.0

    # Laundry items only when laundry is selected.
    if "laundry" in d and not laundry:
        return 0.0

    # Bathroom fixtures scale with washroom count.
    if category == "Fittings & Fixtures":
        bathroom_items = [
            "water closet",
            "vanity basin",
            "basin mixer",
            "muslim shower",
            "shower set",
            "toilet paper holder",
            "beveled edge mirror",
            "soap dish",
        ]

        if any(term in d for term in bathroom_items):
            return float(max(washrooms, 0))

        if "geyser" in d:
            return float(max(washrooms, 1))

        if "kitchen sink" in d or "sink mixer" in d:
            return 1.0

    # Doors and woodwork.
    if category == "Doors, Windows & Woodwork":
        if "room doors" in d:
            return float(max(bedrooms, 0))

        if "washroom" in d and "door" in d:
            return float(max(washrooms, 0))

        if "main entrance" in d:
            return 1.0

        if "wardrobe" in d:
            return float(max(bedrooms, 0))

        if "kitchen cabinet" in d:
            return 1.0

    # Basic residential window assumptions.
    if "window" in d:
        if "bathroom" in d:
            return float(max(washrooms, 0))

        if "room windows" in d:
            return float(max(bedrooms * 2, 0))

    # Electrical fixtures.
    if category == "Electrical":
        if "ceiling fans" in d:
            return float(
                max(bedrooms + drawing_rooms + (1 if lobby else 0), 1)
            )

        if "bathroom lights" in d:
            return float(max(washrooms, 0))

        if "drawing room lights" in d:
            return float(max(drawing_rooms, 0))

        if "kitchen lights" in d:
            return 1.0

    return base_quantity


def calculate_boq(
    reference_items,
    sqft_per_floor,
    floors,
    bedrooms,
    washrooms,
    drawing_rooms,
    lobby,
    laundry,
    basement,
):
    total_area = float(sqft_per_floor) * int(floors)

    # Reference worksheet represents five 1,200 sq.ft levels.
    scale = total_area / REFERENCE_TOTAL_SQFT

    rows = []

    for item in reference_items:
        base_quantity = item["reference_quantity"] * scale

        quantity = room_adjustment(
            item=item,
            base_quantity=base_quantity,
            bedrooms=bedrooms,
            washrooms=washrooms,
            drawing_rooms=drawing_rooms,
            lobby=lobby,
            laundry=laundry,
            basement=basement,
        )

        if quantity <= 0:
            continue

        amount = quantity * item["rate"]

        rows.append({
            "Category": item["category"],
            "Description": item["description"],
            "Unit": (
                "Reference lump sum"
                if item["lump_sum"]
                else item["unit"]
            ),
            "Quantity": round(quantity, 2),
            "Rate (PKR)": round(item["rate"], 2),
            "Amount (PKR)": round(amount, 2),
            "Source Row": item["source_row"],
        })

    return pd.DataFrame(rows), scale, total_area


def category_summary(boq_df):
    if boq_df.empty:
        return pd.DataFrame(
            columns=["Category", "Subtotal (PKR)"]
        )

    result = (
        boq_df.groupby("Category", as_index=False)["Amount (PKR)"]
        .sum()
        .rename(columns={"Amount (PKR)": "Subtotal (PKR)"})
        .sort_values("Subtotal (PKR)", ascending=False)
    )

    result["Subtotal (PKR)"] = result["Subtotal (PKR)"].round(0)
    return result


# ============================================================
# Excel / CSV export
# ============================================================

def create_excel(boq_df, summary_df, project_inputs):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        boq_df.to_excel(
            writer,
            index=False,
            sheet_name="BOQ",
        )

        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Cost Summary",
        )

        inputs_df = pd.DataFrame(
            [
                {"Input": key, "Value": value}
                for key, value in project_inputs.items()
            ]
        )

        inputs_df.to_excel(
            writer,
            index=False,
            sheet_name="Project Inputs",
        )

        workbook = writer.book

        money_format = workbook.add_format({
            "num_format": "#,##0"
        })

        quantity_format = workbook.add_format({
            "num_format": "#,##0.00"
        })

        boq_sheet = writer.sheets["BOQ"]

        boq_sheet.set_column("A:A", 30)
        boq_sheet.set_column("B:B", 65)
        boq_sheet.set_column("C:C", 20)
        boq_sheet.set_column("D:D", 14, quantity_format)
        boq_sheet.set_column("E:F", 18, money_format)
        boq_sheet.set_column("G:G", 12)

    output.seek(0)
    return output.getvalue()


# ============================================================
# Groq
# ============================================================

def get_groq_client(api_key):
    if not api_key:
        return None
    return Groq(api_key=api_key)


def ask_groq(client, prompt, model):
    response = client.chat.completions.create(
        model=model,
        temperature=0.2,
        max_tokens=1800,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a Pakistan-specific residential construction "
                    "assistant. Python calculations are authoritative. "
                    "Never invent spreadsheet rates or quantities. "
                    "Use the provided BOQ to answer. Clearly identify "
                    "assumptions and recommend contractor/engineer verification "
                    "for decisions involving structural safety or current "
                    "market prices."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
    )

    return response.choices[0].message.content


# ============================================================
# Streamlit page
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ AI Construction Assistant")

st.caption(
    "Pakistan-specific residential construction planner using "
    f"'{SOURCE_SHEET}' from the supplied workbook."
)


# ============================================================
# Sidebar
# ============================================================

st.sidebar.header("🏠 Project Inputs")

uploaded_file = st.sidebar.file_uploader(
    "Upload source Excel",
    type=["xlsx", "xlsm"],
    help=(
        f"Upload '{SOURCE_FILE}'. The app reads only the "
        f"'{SOURCE_SHEET}' worksheet."
    ),
)

if uploaded_file is not None:
    ws_values, ws_formulas, error = load_source_workbook(
        uploaded_file
    )
else:
    ws_values, ws_formulas, error = load_source_workbook()


if error:
    st.error(error)
    st.info(
        f"Upload '{SOURCE_FILE}' using the sidebar to continue."
    )
    st.stop()


reference_items, source_warnings = extract_reference_items(
    ws_values,
    ws_formulas,
)


marla = st.sidebar.number_input(
    "Plot size (Marla)",
    min_value=1.0,
    max_value=100.0,
    value=5.0,
    step=0.5,
)

automatic_area = marla * MARLA_SQFT

manual_area = st.sidebar.checkbox(
    "Manually override covered area",
    value=False,
)

if manual_area:
    sqft_per_floor = st.sidebar.number_input(
        "Covered area per floor (sq.ft)",
        min_value=100.0,
        max_value=10000.0,
        value=float(min(automatic_area, 1200.0)),
        step=25.0,
    )
else:
    sqft_per_floor = st.sidebar.number_input(
        "Covered area per floor (sq.ft)",
        min_value=100.0,
        max_value=10000.0,
        value=float(round(automatic_area, 2)),
        step=25.0,
        disabled=True,
    )


floors = st.sidebar.selectbox(
    "No. of levels / floors",
    options=[1, 2, 3, 4, 5],
    index=1,
    format_func=lambda x: {
        1: "Ground",
        2: "G+1",
        3: "G+2",
        4: "G+3",
        5: "G+4",
    }[x],
)

bedrooms = st.sidebar.number_input(
    "Bedrooms",
    min_value=0,
    max_value=20,
    value=3,
    step=1,
)

washrooms = st.sidebar.number_input(
    "Washrooms",
    min_value=0,
    max_value=20,
    value=3,
    step=1,
)

drawing_rooms = st.sidebar.number_input(
    "Drawing rooms",
    min_value=0,
    max_value=10,
    value=1,
    step=1,
)

lobby = st.sidebar.checkbox(
    "Lobby",
    value=True,
)

laundry = st.sidebar.checkbox(
    "Laundry room",
    value=False,
)

basement = st.sidebar.checkbox(
    "Basement",
    value=False,
)


st.sidebar.divider()
st.sidebar.header("🤖 Groq AI")

api_key = st.sidebar.text_input(
    "Groq API key",
    value=os.getenv("GROQ_API_KEY", ""),
    type="password",
)

model = st.sidebar.text_input(
    "Groq model",
    value="llama-3.3-70b-versatile",
)


# ============================================================
# Calculate BOQ
# ============================================================

boq_df, scale_factor, total_area = calculate_boq(
    reference_items=reference_items,
    sqft_per_floor=sqft_per_floor,
    floors=floors,
    bedrooms=bedrooms,
    washrooms=washrooms,
    drawing_rooms=drawing_rooms,
    lobby=lobby,
    laundry=laundry,
    basement=basement,
)

summary_df = category_summary(boq_df)

grand_total = (
    float(boq_df["Amount (PKR)"].sum())
    if not boq_df.empty
    else 0.0
)


# ============================================================
# Metrics
# ============================================================

m1, m2, m3, m4 = st.columns(4)

m1.metric(
    "Plot",
    f"{marla:g} Marla",
)

m2.metric(
    "Area / Floor",
    f"{sqft_per_floor:,.0f} sq.ft",
)

m3.metric(
    "Total Covered Area",
    f"{total_area:,.0f} sq.ft",
)

m4.metric(
    "Estimated Cost",
    money(grand_total),
)


st.info(
    f"Source reference = 30 × 40 = 1,200 sq.ft per level. "
    f"Reference levels = {REFERENCE_LEVELS}. "
    f"Reference total = {REFERENCE_TOTAL_SQFT:,.0f} sq.ft. "
    f"Current project scale factor = {scale_factor:.4f}."
)


if source_warnings:
    with st.expander(
        f"⚠️ {len(source_warnings)} source spreadsheet warning(s)"
    ):
        st.write(
            "The supplied workbook contains broken Excel formulas "
            "(#REF!). The app does not invent values for those formulas."
        )

        st.dataframe(
            pd.DataFrame(source_warnings),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# Tabs
# ============================================================

tab_boq, tab_ai, tab_summary = st.tabs(
    [
        "📋 BOQ",
        "🤖 AI Assistant",
        "📊 Cost Summary",
    ]
)


# ============================================================
# BOQ TAB
# ============================================================

with tab_boq:
    st.subheader("Bill of Quantities")

    if boq_df.empty:
        st.warning("No BOQ items were generated.")
    else:
        st.dataframe(
            boq_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Quantity": st.column_config.NumberColumn(
                    "Quantity",
                    format="%.2f",
                ),
                "Rate (PKR)": st.column_config.NumberColumn(
                    "Rate (PKR)",
                    format="%.0f",
                ),
                "Amount (PKR)": st.column_config.NumberColumn(
                    "Amount (PKR)",
                    format="%.0f",
                ),
            },
        )

        st.subheader("Category Subtotals")

        st.dataframe(
            summary_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Subtotal (PKR)": st.column_config.NumberColumn(
                    "Subtotal (PKR)",
                    format="%.0f",
                )
            },
        )

        st.success(
            f"Grand Total: {money(grand_total)}"
        )

        project_inputs = {
            "Plot Size (Marla)": marla,
            "Marla Conversion (sq.ft)": MARLA_SQFT,
            "Covered Area/Floor (sq.ft)": sqft_per_floor,
            "Number of Floors": floors,
            "Total Covered Area (sq.ft)": total_area,
            "Bedrooms": bedrooms,
            "Washrooms": washrooms,
            "Drawing Rooms": drawing_rooms,
            "Lobby": lobby,
            "Laundry": laundry,
            "Basement": basement,
            "Reference Sheet": SOURCE_SHEET,
            "Scale Factor": scale_factor,
            "Grand Total (PKR)": grand_total,
        }

        excel_data = create_excel(
            boq_df,
            summary_df,
            project_inputs,
        )

        csv_data = boq_df.to_csv(
            index=False
        ).encode("utf-8")

        c1, c2 = st.columns(2)

        with c1:
            st.download_button(
                "⬇️ Download BOQ Excel",
                data=excel_data,
                file_name="ai_construction_boq.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

        with c2:
            st.download_button(
                "⬇️ Download BOQ CSV",
                data=csv_data,
                file_name="ai_construction_boq.csv",
                mime="text/csv",
            )


# ============================================================
# AI TAB
# ============================================================

with tab_ai:
    st.subheader("🤖 AI Construction Assistant")

    if not api_key:
        st.warning(
            "Enter your Groq API key in the sidebar to use AI features."
        )
    else:
        client = get_groq_client(api_key)

        if not boq_df.empty:
            boq_context = boq_df[
                [
                    "Category",
                    "Description",
                    "Unit",
                    "Quantity",
                    "Rate (PKR)",
                    "Amount (PKR)",
                ]
            ].to_string(index=False)

            analysis_prompt = f"""
Analyze this residential construction estimate.

Project:
- Plot: {marla} marla
- Covered area/floor: {sqft_per_floor:,.0f} sq.ft
- Floors: {floors}
- Total covered area: {total_area:,.0f} sq.ft
- Bedrooms: {bedrooms}
- Washrooms: {washrooms}
- Drawing rooms: {drawing_rooms}
- Lobby: {lobby}
- Laundry: {laundry}
- Basement: {basement}

Grand estimate: {money(grand_total)}

BOQ:
{boq_context}

Provide:
1. Short cost summary.
2. Three largest cost categories.
3. Important assumptions.
4. Items that should be verified.
5. Practical cost-control suggestions.
"""

            if st.button(
                "✨ Generate AI Cost Analysis"
            ):
                with st.spinner(
                    "Analyzing construction estimate..."
                ):
                    try:
                        answer = ask_groq(
                            client,
                            analysis_prompt,
                            model,
                        )
                        st.markdown(answer)
                    except Exception as exc:
                        st.error(
                            f"Groq error: {exc}"
                        )

        st.divider()

        st.subheader("Ask a Construction Question")

        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []

        for message in st.session_state.chat_history:
            with st.chat_message(
                message["role"]
            ):
                st.markdown(
                    message["content"]
                )

        question = st.chat_input(
            "Example: Which category is costing the most?"
        )

        if question:
            st.session_state.chat_history.append(
                {
                    "role": "user",
                    "content": question,
                }
            )

            if boq_df.empty:
                context = "No BOQ was calculated."
            else:
                context = boq_df[
                    [
                        "Category",
                        "Description",
                        "Unit",
                        "Quantity",
                        "Rate (PKR)",
                        "Amount (PKR)",
                    ]
                ].to_string(index=False)

            question_prompt = f"""
Answer the user's construction question using the calculated BOQ.

Project total: {money(grand_total)}
Total covered area: {total_area:,.0f} sq.ft
Source worksheet: {SOURCE_SHEET}

BOQ:
{context}

User question:
{question}

Rules:
- Python-calculated quantities and costs are authoritative.
- Do not invent rates that are absent from the workbook.
- If a requested price is not available, say so.
- Give practical Pakistan-specific guidance where appropriate.
"""

            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    try:
                        answer = ask_groq(
                            client,
                            question_prompt,
                            model,
                        )

                        st.markdown(answer)

                        st.session_state.chat_history.append(
                            {
                                "role": "assistant",
                                "content": answer,
                            }
                        )

                    except Exception as exc:
                        st.error(
                            f"Groq error: {exc}"
                        )


# ============================================================
# SUMMARY TAB
# ============================================================

with tab_summary:
    st.subheader("📊 Cost Summary")

    if summary_df.empty:
        st.warning("No cost data available.")
    else:
        # Streamlit-native bar chart: no Plotly required.
        chart_df = summary_df.set_index(
            "Category"
        )[["Subtotal (PKR)"]]

        st.bar_chart(
            chart_df,
            use_container_width=True,
        )

        st.subheader("Category Distribution")

        total = summary_df["Subtotal (PKR)"].sum()

        distribution = summary_df.copy()

        if total > 0:
            distribution["Percentage"] = (
                distribution["Subtotal (PKR)"] / total * 100
            )
        else:
            distribution["Percentage"] = 0

        distribution["Percentage"] = distribution[
            "Percentage"
        ].round(2)

        st.dataframe(
            distribution,
            use_container_width=True,
            hide_index=True,
        )


st.divider()

st.caption(
    f"Source: {SOURCE_FILE} → {SOURCE_SHEET}. "
    "This tool is an estimate/planning assistant. "
    "Verify quantities, structural requirements, labour rates and "
    "current market prices with qualified local professionals."
)
