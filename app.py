import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Universal Dashboard",
    layout="wide",
    page_icon="📊"
)


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div style='text-align:center;
background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%);
padding:15px;
border-radius:10px;
color:white'>

<h2 style='margin:0; color:white'>
📊 Universal Dashboard - AUTO HEADER
</h2>

</div>
""", unsafe_allow_html=True)


# =========================================================
# AUTO HEADER DETECTION
# =========================================================

def find_best_header(file, sheet):
    best_row, max_cols = 0, 0

    for r in range(8):
        try:
            file.seek(0)

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=r
            )

            df = df.dropna(how='all').dropna(axis=1, how='all')

            cols = len([
                c for c in df.columns
                if 'Unnamed' not in str(c)
            ])

            col_str = ' '.join(
                [str(c).lower() for c in df.columns]
            )

            score = cols

            if any(
                x in col_str
                for x in [
                    'status',
                    'row label',
                    'wir',
                    'system',
                    'originator',
                    'approved'
                ]
            ):
                score += 10

            if score > max_cols:
                max_cols = score
                best_row = r

        except:
            pass

    return best_row


# =========================================================
# EXCEL UPLOAD
# =========================================================

f = st.file_uploader(
    "📂 Excel Upload Karo",
    type=["xlsx", "xls"]
)

if not f:
    st.stop()


xls = pd.ExcelFile(f)

sheet = st.selectbox(
    f"Sheet Select Karo ({len(xls.sheet_names)} sheets)",
    xls.sheet_names
)


# =========================================================
# HEADER DETECT
# =========================================================

auto_header = find_best_header(f, sheet)

st.info(
    f"🤖 Auto Detected Header Row: {auto_header} "
    f"(Aap change kar sakte ho)"
)

header_row = st.number_input(
    "Header Row Number",
    0,
    10,
    value=auto_header
)


f.seek(0)

df = pd.read_excel(
    f,
    sheet_name=sheet,
    header=header_row
)

df = df.dropna(how='all')


# Remove unnamed columns
df = df.loc[
    :,
    ~df.columns.astype(str).str.contains(
        '^Unnamed',
        na=False,
        regex=True
    )
]


# =========================================================
# FALLBACK HEADER
# =========================================================

if len(df.columns) == 0 or len(df.columns) < 2:

    for r in range(8):

        f.seek(0)

        temp = pd.read_excel(
            f,
            sheet_name=sheet,
            header=r
        )

        temp = temp.dropna(how='all')
        temp = temp.dropna(axis=1, how='all')

        if len(temp.columns) >= 3:

            df = temp
            header_row = r
            break


st.success(
    f"✅ Loaded: {len(df)} Records | "
    f"{len(df.columns)} Columns | "
    f"Header Row: {header_row}"
)

st.write(
    "Columns:",
    list(df.columns)[:10]
)


# =========================================================
# DASHBOARD
# =========================================================

is_pivot = 'Row Labels' in df.columns


# =========================================================
# PIVOT FILE
# =========================================================

if is_pivot:

    df_count = df[
        [c for c in df.columns if '.1' not in str(c)]
    ]

    grand_col = [
        c for c in df_count.columns
        if 'Grand Total' in str(c)
    ]

    grand_col = grand_col[0] if grand_col else None

    row_col = [
        c for c in df_count.columns
        if 'Row Labels' in str(c)
    ]

    row_col = (
        row_col[0]
        if row_col
        else df_count.columns[0]
    )


    if grand_col:

        grand_idx = df_count.columns.get_loc(
            grand_col
        )

        status_cols = list(
            df_count.columns[1:grand_idx]
        )

        df_data = df_count[
            (df_count[row_col] != 'Grand Total') &
            (df_count[row_col].notna())
        ].copy()


        for c in status_cols + [grand_col]:

            df_data[c] = pd.to_numeric(
                df_data[c],
                errors='coerce'
            ).fillna(0)


        c1, c2, c3 = st.columns(3)

        c1.metric(
            "SYSTEMS",
            len(df_data)
        )

        c2.metric(
            "TOTAL WIR",
            int(df_data[grand_col].sum())
        )

        c3.metric(
            "STATUS TYPES",
            len(status_cols)
        )


        a, b = st.columns(2)


        # TOP 10
        with a:

            top10 = (
                df_data
                .sort_values(
                    grand_col,
                    ascending=False
                )
                .head(10)
            )

            fig = px.bar(
                top10,
                x=grand_col,
                y=row_col,
                orientation='h',
                text=grand_col,
                title="Top 10 Systems - Status Wise"
            )

            fig.update_layout(
                yaxis={
                    'categoryorder':
                    'total ascending'
                }
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


        # STATUS DONUT
        with b:

            status_totals = [
                {
                    'Status': sc,
                    'Count': int(df_data[sc].sum())
                }
                for sc in status_cols
                if df_data[sc].sum() > 0
            ]

            status_df = pd.DataFrame(
                status_totals
            )

            fig2 = px.pie(
                status_df,
                values='Count',
                names='Status',
                hole=0.5,
                title="STATUS WISE - A vs B vs C"
            )

            fig2.update_traces(
                textinfo='percent+label'
            )

            st.plotly_chart(
                fig2,
                use_container_width=True
            )


        final_df = df_data

    else:

        final_df = df


# =========================================================
# NORMAL EXCEL
# =========================================================

else:

    text_cols = df.select_dtypes(
        include=['object']
    ).columns.tolist()


    if len(text_cols) == 0 and len(df.columns) > 0:
        text_cols = list(df.columns[:4])


    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "ROWS",
        len(df)
    )

    c2.metric(
        "COLS",
        len(df.columns)
    )

    c3.metric(
        "TEXT COLS",
        len(text_cols)
    )

    c4.metric(
        "NUMERIC",
        len(
            df.select_dtypes(
                include='number'
            ).columns
        )
    )


    cols = st.columns(2)


    for i, col in enumerate(text_cols[:4]):

        try:

            vc = (
                df[col]
                .astype(str)
                .str.strip()
            )

            vc = vc[
                ~vc.isin([
                    'nan',
                    'None',
                    '',
                    'NaT'
                ])
            ]

            vc = (
                vc.value_counts()
                .head(10)
                .reset_index()
            )

            if len(vc) == 0:
                continue


            vc.columns = [
                col,
                'Count'
            ]


            with cols[i % 2]:

                if i == 0:

                    fig = px.pie(
                        vc,
                        values='Count',
                        names=col,
                        hole=0.5,
                        title=f"{col} - Breakdown"
                    )

                    fig.update_traces(
                        textinfo='percent+label'
                    )

                else:

                    fig = px.bar(
                        vc,
                        x='Count',
                        y=col,
                        orientation='h',
                        text='Count',
                        title=f"{col} - Top 10"
                    )

                    fig.update_layout(
                        yaxis={
                            'categoryorder':
                            'total ascending'
                        }
                    )


                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

        except:
            pass


    final_df = df


# =========================================================
# DATA PREVIEW
# =========================================================

st.markdown("### 📋 Data Preview")

st.dataframe(
    final_df.head(50),
    use_container_width=True
)


# =========================================================
# CREATE EXCEL DASHBOARD
# =========================================================

def get_excel():

    out = io.BytesIO()


    # -----------------------------------------------------
    # WRITE DATA
    # -----------------------------------------------------

    with pd.ExcelWriter(
        out,
        engine='openpyxl'
    ) as writer:

        final_df.to_excel(
            writer,
            index=False,
            sheet_name='Data'
        )


    out.seek(0)


    # -----------------------------------------------------
    # OPEN WORKBOOK
    # -----------------------------------------------------

    wb = load_workbook(out)


    # Remove old dashboard if exists
    if 'Dashboard' in wb.sheetnames:
        del wb['Dashboard']

    if 'Summary' in wb.sheetnames:
        del wb['Summary']


    ws_data = wb['Data']


    # -----------------------------------------------------
    # CREATE SHEETS
    # -----------------------------------------------------

    ws_dash = wb.create_sheet(
        'Dashboard',
        0
    )

    ws_sum = wb.create_sheet(
        'Summary'
    )


    # -----------------------------------------------------
    # DASHBOARD TITLE
    # -----------------------------------------------------

    ws_dash['A1'] = "📊 Universal Dashboard"

    ws_dash['A1'].font = Font(
        size=20,
        bold=True,
        color="FFFFFF"
    )

    ws_dash['A1'].fill = PatternFill(
        "solid",
        fgColor="185ADB"
    )

    ws_dash['A1'].alignment = Alignment(
        horizontal="center"
    )

    ws_dash.merge_cells(
        'A1:L2'
    )


    ws_dash['A4'] = "Total Records"
    ws_dash['B4'] = len(final_df)

    ws_dash['D4'] = "Total Columns"
    ws_dash['E4'] = len(final_df.columns)


    # -----------------------------------------------------
    # PIVOT DASHBOARD
    # -----------------------------------------------------

    if is_pivot and 'Grand Total' in ' '.join(
        map(str, final_df.columns)
    ):

        grand_col = [
            c for c in final_df.columns
            if 'Grand Total' in str(c)
        ]

        row_col = [
            c for c in final_df.columns
            if 'Row Labels' in str(c)
        ]


        if grand_col:

            grand_col = grand_col[0]

            row_col = (
                row_col[0]
                if row_col
                else final_df.columns[0]
            )


            # =============================================
            # TOP 10 SUMMARY
            # =============================================

            top10 = (
                final_df
                .sort_values(
                    grand_col,
                    ascending=False
                )
                .head(10)
            )


            start_row = 1

            ws_sum.cell(
                start_row,
                1,
                row_col
            )

            ws_sum.cell(
                start_row,
                2,
                "Count"
            )


            for i, (_, row) in enumerate(
                top10.iterrows(),
                start=2
            ):

                ws_sum.cell(
                    i,
                    1,
                    str(row[row_col])
                )

                ws_sum.cell(
                    i,
                    2,
                    float(row[grand_col])
                )


            # Excel Bar Chart

            chart = BarChart()

            chart.type = "bar"

            chart.style = 10

            chart.title = (
                "Top 10 Systems - Status Wise"
            )

            chart.y_axis.title = row_col
            chart.x_axis.title = "Count"

            data = Reference(
                ws_sum,
                min_col=2,
                min_row=1,
                max_row=11
            )

            cats = Reference(
                ws_sum,
                min_col=1,
                min_row=2,
                max_row=11
            )

            chart.add_data(
                data,
                titles_from_data=True
            )

            chart.set_categories(cats)

            chart.height = 8
            chart.width = 14

            ws_dash.add_chart(
                chart,
                "A7"
            )


            # =============================================
            # STATUS SUMMARY
            # =============================================

            status_cols = [
                c for c in final_df.columns
                if c not in [row_col, grand_col]
            ]


            status_cols = [
                c for c in status_cols
                if not str(c).startswith('.')
            ]


            status_start = 15

            ws_sum.cell(
                status_start,
                1,
                "Status"
            )

            ws_sum.cell(
                status_start,
                2,
                "Count"
            )


            status_row = status_start + 1


            for c in status_cols:

                try:

                    total = pd.to_numeric(
                        final_df[c],
                        errors='coerce'
                    ).fillna(0).sum()

                    if total > 0:

                        ws_sum.cell(
                            status_row,
                            1,
                            str(c)
                        )

                        ws_sum.cell(
                            status_row,
                            2,
                            float(total)
                        )

                        status_row += 1

                except:
                    pass


            if status_row > status_start + 1:

                donut = DoughnutChart()

                donut.title = (
                    "STATUS WISE"
                )

                data = Reference(
                    ws_sum,
                    min_col=2,
                    min_row=status_start,
                    max_row=status_row - 1
                )

                labels = Reference(
                    ws_sum,
                    min_col=1,
                    min_row=status_start + 1,
                    max_row=status_row - 1
                )

                donut.add_data(
                    data,
                    titles_from_data=True
                )

                donut.set_categories(
                    labels
                )

                donut.holeSize = 50

                donut.height = 8
                donut.width = 12

                ws_dash.add_chart(
                    donut,
                    "J7"
                )


    # -----------------------------------------------------
    # NORMAL EXCEL DASHBOARD
    # -----------------------------------------------------

    else:

        text_cols = final_df.select_dtypes(
            include=['object']
        ).columns.tolist()


        if len(text_cols) == 0:
            text_cols = list(
                final_df.columns[:4]
            )


        chart_positions = [
            "A7",
            "J7",
            "A25",
            "J25"
        ]


        summary_col = 1


        for i, col in enumerate(
            text_cols[:4]
        ):

            try:

                vc = (
                    final_df[col]
                    .astype(str)
                    .str.strip()
                )

                vc = vc[
                    ~vc.isin([
                        'nan',
                        'None',
                        '',
                        'NaT'
                    ])
                ]

                vc = (
                    vc.value_counts()
                    .head(10)
                    .reset_index()
                )

                if len(vc) == 0:
                    continue


                vc.columns = [
                    str(col),
                    'Count'
                ]


                # =========================================
                # WRITE SUMMARY DATA
                # =========================================

                ws_sum.cell(
                    1,
                    summary_col,
                    str(col)
                )

                ws_sum.cell(
                    1,
                    summary_col + 1,
                    "Count"
                )


                for r, (_, row) in enumerate(
                    vc.iterrows(),
                    start=2
                ):

                    ws_sum.cell(
                        r,
                        summary_col,
                        str(row[col])
                    )

                    ws_sum.cell(
                        r,
                        summary_col + 1,
                        int(row['Count'])
                    )


                # =========================================
                # CREATE CHART
                # =========================================

                position = chart_positions[i]


                if i == 0:

                    chart = DoughnutChart()

                    chart.title = (
                        f"{col} - Breakdown"
                    )

                    chart.holeSize = 50


                else:

                    chart = BarChart()

                    chart.type = "bar"

                    chart.style = 10

                    chart.title = (
                        f"{col} - Top 10"
                    )

                    chart.x_axis.title = "Count"
                    chart.y_axis.title = str(col)


                data = Reference(
                    ws_sum,
                    min_col=summary_col + 1,
                    min_row=1,
                    max_row=len(vc) + 1
                )

                labels = Reference(
                    ws_sum,
                    min_col=summary_col,
                    min_row=2,
                    max_row=len(vc) + 1
                )


                chart.add_data(
                    data,
                    titles_from_data=True
                )

                chart.set_categories(
                    labels
                )


                if i == 0:

                    chart.dataLabels = DataLabelList()

                    chart.dataLabels.showPercent = True
                    chart.dataLabels.showLeaderLines = True


                chart.height = 8
                chart.width = 14


                ws_dash.add_chart(
                    chart,
                    position
                )


                summary_col += 3


            except:
                pass


    # -----------------------------------------------------
    # FORMAT DATA SHEET
    # -----------------------------------------------------

    for ws in [ws_data, ws_sum]:

        ws.freeze_panes = "A2"


        for cell in ws[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="185ADB"
            )

            cell.alignment = Alignment(
                horizontal="center"
            )


        # Auto width

        for column_cells in ws.columns:

            max_length = 0

            column_letter = get_column_letter(
                column_cells[0].column
            )

            for cell in column_cells:

                try:

                    value_length = len(
                        str(cell.value)
                    )

                    if value_length > max_length:
                        max_length = value_length

                except:
                    pass


            ws.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                40
            )


    # -----------------------------------------------------
    # DASHBOARD COLUMN WIDTH
    # -----------------------------------------------------

    for col in range(1, 14):

        ws_dash.column_dimensions[
            get_column_letter(col)
        ].width = 14


    # -----------------------------------------------------
    # SAVE
    # -----------------------------------------------------

    final_output = io.BytesIO()

    wb.save(
        final_output
    )

    final_output.seek(0)

    return final_output.getvalue()


# =========================================================
# DOWNLOAD BUTTON
# =========================================================

st.download_button(
    f"📥 DOWNLOAD EXCEL + CHARTS - {len(final_df)} Records",
    get_excel(),
    file_name=f"Report_{len(final_df)}_Dashboard.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True,
    type="primary"
)
