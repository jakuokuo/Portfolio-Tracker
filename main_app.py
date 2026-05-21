"""Futures Portfolio Tracker — Dash app.

Upload a "Futs Portfolio" deal report (.xlsm / .xlsx / .xls) and explore the
net position of the portfolio over time, by hub and by product.

Deploy to Posit Connect (push-deploy from CLI):
    rsconnect deploy dash --entrypoint main_app:app \\
        --server https://connect.teainc.org --api-key mxe1rbmYxDF30zmNEgI5MnDW4EyEdiho \\
        --title "Futures Portfolio Tracker" .

Git-backed deploy (preferred): commit manifest.json + push to GitHub, then
in Connect choose New Content → Import from Git.
"""

from __future__ import annotations

import base64
import io
import re
from typing import Optional, Tuple

import dash
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Input, Output, State, dash_table, dcc, html

# --------------------------------------------------------------------------- #
# Dash app setup
# --------------------------------------------------------------------------- #

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY, dbc.icons.BOOTSTRAP],
    suppress_callback_exceptions=True,
    title="Futures Portfolio Tracker",
)
server = app.server  # WSGI entrypoint Posit Connect looks for

PLOT_TEMPLATE = "plotly_white"
LONG_COLOR = "#2ca02c"
SHORT_COLOR = "#d62728"
NET_COLOR = "#1f77b4"

REQUIRED_COLUMNS = [
    "Trade Date",
    "B/S",
    "Product",
    "Hub",
    "Contract",
    "Begin Date",
    "End Date",
    "Total Quantity",
    "Price",
]

# Strip "(1 MW)" / "(25 MW)" volume tags so block sizes aggregate together.
_PRODUCT_VOLUME_SUFFIX = re.compile(r"\s*\(\d+\s*MW\)\s*", flags=re.IGNORECASE)

# Hub-name fragment that marks short-tenor daily hubs (filtered out).
_DAILY_HUB_TAG = re.compile(r"\(\s*daily\s*\)", flags=re.IGNORECASE)

# Minimum delivery-window length (days) for a deal to count as a "monthly" position.
MONTHLY_MIN_DAYS = 28


def _normalize_product(name: str) -> str:
    """Collapse 'Peak Futures (25 MW)' / 'Peak Futures (1 MW)' to 'Peak Futures'."""
    cleaned = _PRODUCT_VOLUME_SUFFIX.sub(" ", str(name))
    return re.sub(r"\s+", " ", cleaned).strip()


def _is_daily_hub(name: str) -> bool:
    return bool(_DAILY_HUB_TAG.search(str(name)))


def _weighted_avg(prices: pd.Series, weights: pd.Series) -> float:
    """Volume-weighted average. Returns NaN when weights sum to zero."""
    mask = prices.notna() & weights.notna()
    w = weights[mask]
    if w.sum() == 0:
        return float("nan")
    return float((prices[mask] * w).sum() / w.sum())


# --------------------------------------------------------------------------- #
# Parsing & shaping
# --------------------------------------------------------------------------- #

def _read_deal_report(buffer: bytes, filename: str) -> pd.DataFrame:
    """Load the DealReport sheet using the correct engine for the extension.

    The Futs Portfolio export uses three preamble rows: a title ("Cleared Deals"),
    a blank row, and then the actual column headers on row 3 (zero-based index 2).
    """
    ext = (filename or "").lower().rsplit(".", 1)[-1]
    engine = "xlrd" if ext == "xls" else "openpyxl"

    xl = pd.ExcelFile(io.BytesIO(buffer), engine=engine)
    sheet = next(
        (s for s in xl.sheet_names if "dealreport" in s.lower().replace(" ", "")),
        xl.sheet_names[0],
    )
    return xl.parse(sheet, header=2)


def parse_uploaded_file(
    contents: Optional[str], filename: Optional[str]
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """Decode a Dash upload payload into a cleaned deals DataFrame."""
    if contents is None:
        return None, "No file uploaded."

    try:
        _, b64 = contents.split(",", 1)
        raw = base64.b64decode(b64)
        df = _read_deal_report(raw, filename or "")
    except Exception as exc:  # noqa: BLE001 — surface any parse error to the UI
        return None, f"Could not read file: {exc}"

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return None, (
            "The uploaded file is missing required columns: "
            + ", ".join(missing)
        )

    df["Trade Date"] = pd.to_datetime(df["Trade Date"], errors="coerce")
    df["Begin Date"] = pd.to_datetime(df["Begin Date"], errors="coerce")
    df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")
    df["Total Quantity"] = pd.to_numeric(df["Total Quantity"], errors="coerce")
    if "Price" in df.columns:
        df["Price"] = pd.to_numeric(df["Price"], errors="coerce")
    if "Lots" in df.columns:
        df["Lots"] = pd.to_numeric(df["Lots"], errors="coerce")

    df = df.dropna(subset=["Begin Date", "End Date", "B/S", "Total Quantity"])

    for col in ("Product", "Hub", "Contract", "B/S"):
        df[col] = (
            df[col].astype(str).str.replace("\xa0", " ", regex=False).str.strip()
        )

    df["Product"] = df["Product"].map(_normalize_product)

    df = df.loc[~df["Hub"].map(_is_daily_hub)].copy()

    df["Direction"] = df["B/S"].str.lower().str.startswith("b").map(
        {True: 1, False: -1}
    )
    df["Signed Qty (MWh)"] = df["Total Quantity"] * df["Direction"]
    df["Lots"] = df["Lots"].fillna(0)
    df["Signed Lots"] = df["Lots"] * df["Direction"]

    delivery_days = (df["End Date"] - df["Begin Date"]).dt.days + 1
    df = df.loc[delivery_days >= MONTHLY_MIN_DAYS].copy()

    df["Delivery Month"] = df["Begin Date"].dt.to_period("M").dt.to_timestamp()

    return df.reset_index(drop=True), None


def aggregate_monthly(
    df: pd.DataFrame, by: Optional[list] = None
) -> pd.DataFrame:
    """Aggregate deals by Delivery Month (+ optional extra dimensions).

    Returns one row per group with Long/Short/Net Lots and a Trade VWAP
    (volume-weighted average $/MWh across all trades in the group).
    """
    if df is None or df.empty:
        return pd.DataFrame()

    by = by or []
    keys = ["Delivery Month"] + list(by)

    rows = []
    for key_vals, grp in df.groupby(keys, sort=True):
        if not isinstance(key_vals, tuple):
            key_vals = (key_vals,)
        long_g = grp[grp["Direction"] == 1]
        short_g = grp[grp["Direction"] == -1]
        long_lots = float(long_g["Lots"].sum())
        short_lots = float(short_g["Lots"].sum())
        row = dict(zip(keys, key_vals))
        row.update(
            {
                "Long Lots": long_lots,
                "Short Lots": short_lots,
                "Net Lots": long_lots - short_lots,
                "Trade VWAP": _weighted_avg(grp["Price"], grp["Total Quantity"]),
            }
        )
        rows.append(row)
    return pd.DataFrame(rows).sort_values(keys).reset_index(drop=True)


def build_position_summary(df: pd.DataFrame) -> pd.DataFrame:
    """One row per Hub × Product × Delivery Month with lots, MWh and prices."""
    if df is None or df.empty:
        return pd.DataFrame()

    rows = []
    for (hub, product, month), grp in df.groupby(
        ["Hub", "Product", "Delivery Month"], sort=True
    ):
        long_g = grp[grp["Direction"] == 1]
        short_g = grp[grp["Direction"] == -1]
        long_lots = float(long_g["Lots"].sum())
        short_lots = float(short_g["Lots"].sum())
        long_mwh = float(long_g["Total Quantity"].sum())
        short_mwh = float(short_g["Total Quantity"].sum())
        rows.append(
            {
                "Delivery Month": month,
                "Hub": hub,
                "Product": product,
                "Long Lots": long_lots,
                "Short Lots": short_lots,
                "Net Lots": long_lots - short_lots,
                "Long MWh": long_mwh,
                "Short MWh": short_mwh,
                "Net MWh": long_mwh - short_mwh,
                "Avg Buy ($/MWh)": _weighted_avg(
                    long_g["Price"], long_g["Total Quantity"]
                ),
                "Avg Sell ($/MWh)": _weighted_avg(
                    short_g["Price"], short_g["Total Quantity"]
                ),
                "Trade VWAP ($/MWh)": _weighted_avg(
                    grp["Price"], grp["Total Quantity"]
                ),
            }
        )

    summary = pd.DataFrame(rows)
    if summary.empty:
        return summary
    return summary.sort_values(["Delivery Month", "Hub", "Product"]).reset_index(
        drop=True
    )


# --------------------------------------------------------------------------- #
# Layout
# --------------------------------------------------------------------------- #

navbar = dbc.Navbar(
    dbc.Container(
        dbc.NavbarBrand(
            [
                html.I(className="bi bi-graph-up-arrow me-2"),
                "Futures Portfolio Tracker",
            ],
            className="fw-semibold text-white",
        ),
        fluid=True,
    ),
    color="primary",
    dark=True,
    sticky="top",
    className="mb-3",
)

upload_card = dbc.Card(
    dbc.CardBody(
        [
            html.H5("Upload Deal Report", className="card-title"),
            html.P(
                "Export the Futs Portfolio DealReport (Excel) and drop it here.",
                className="text-muted small mb-2",
            ),
            dcc.Upload(
                id="upload-data",
                children=html.Div(
                    [
                        html.I(className="bi bi-cloud-upload fs-3 d-block"),
                        html.Span("Drag & drop or "),
                        html.A("browse for a file", className="text-primary"),
                        html.Div(
                            ".xlsm / .xlsx / .xls",
                            className="text-muted small",
                        ),
                    ]
                ),
                style={
                    "width": "100%",
                    "padding": "24px",
                    "borderWidth": "2px",
                    "borderStyle": "dashed",
                    "borderRadius": "10px",
                    "borderColor": "#adb5bd",
                    "textAlign": "center",
                    "backgroundColor": "#f8f9fa",
                    "cursor": "pointer",
                },
                multiple=False,
            ),
            html.Div(id="upload-status", className="mt-2"),
        ]
    ),
    className="shadow-sm",
)

filter_card = dbc.Card(
    dbc.CardBody(
        [
            html.H6("Filters", className="card-title"),
            dbc.Label("Hubs", className="mt-1 fw-semibold"),
            dcc.Dropdown(
                id="hub-filter", multi=True, placeholder="All hubs", clearable=True
            ),
            dbc.Label("Products", className="mt-2 fw-semibold"),
            dcc.Dropdown(
                id="product-filter",
                multi=True,
                placeholder="All products",
                clearable=True,
            ),
            dbc.Label("Direction", className="mt-2 fw-semibold"),
            dbc.Checklist(
                id="direction-filter",
                options=[
                    {"label": "Long (Bought)", "value": 1},
                    {"label": "Short (Sold)", "value": -1},
                ],
                value=[1, -1],
                inline=False,
                switch=True,
            ),
            dbc.Label("Delivery date range", className="mt-2 fw-semibold"),
            dcc.DatePickerRange(
                id="date-filter",
                display_format="YYYY-MM-DD",
                style={"width": "100%"},
            ),
        ]
    ),
    className="shadow-sm",
)


def _placeholder_figure(message: str) -> go.Figure:
    fig = go.Figure()
    fig.update_layout(
        template=PLOT_TEMPLATE,
        xaxis={"visible": False},
        yaxis={"visible": False},
        annotations=[
            dict(
                text=message,
                showarrow=False,
                xref="paper",
                yref="paper",
                x=0.5,
                y=0.5,
                font=dict(size=16, color="#6c757d"),
            )
        ],
        height=320,
    )
    return fig


def _graph(graph_id: str, height: int = 360) -> dcc.Loading:
    return dcc.Loading(
        dcc.Graph(id=graph_id, style={"height": f"{height}px"}),
        type="circle",
    )


app.layout = html.Div(
    [
        navbar,
        dcc.Store(id="store-deals"),
        dbc.Container(
            [
                dbc.Row(dbc.Col(upload_card), className="mb-3"),
                dbc.Row(
                    [
                        dbc.Col(filter_card, lg=3, md=4),
                        dbc.Col(
                            [
                                dbc.Row(id="kpi-row", className="g-3 mb-3"),
                                dbc.Card(
                                    dbc.CardBody(_graph("hub-graph", 380)),
                                    className="shadow-sm mb-3",
                                ),
                                dbc.Card(
                                    dbc.CardBody(_graph("cumulative-graph", 320)),
                                    className="shadow-sm mb-3",
                                ),
                                dbc.Row(
                                    [
                                        dbc.Col(
                                            dbc.Card(
                                                dbc.CardBody(_graph("position-graph", 360)),
                                                className="shadow-sm",
                                            ),
                                            md=6,
                                        ),
                                        dbc.Col(
                                            dbc.Card(
                                                dbc.CardBody(_graph("product-graph", 360)),
                                                className="shadow-sm",
                                            ),
                                            md=6,
                                        ),
                                    ],
                                    className="g-3",
                                ),
                            ],
                            lg=9,
                            md=8,
                        ),
                    ],
                    className="g-3",
                ),
                dbc.Row(
                    dbc.Col(
                        dbc.Card(
                            dbc.CardBody(
                                [
                                    html.H5(
                                        "Position Summary",
                                        className="card-title",
                                    ),
                                    html.P(
                                        "Net Lots and volume-weighted average prices "
                                        "for each Hub × Product × Delivery Month. "
                                        "Trade VWAP is the current $/MWh that the "
                                        "desk has been doing business at for that "
                                        "position.",
                                        className="text-muted small mb-2",
                                    ),
                                    dcc.Loading(html.Div(id="position-summary-table")),
                                ]
                            ),
                            className="shadow-sm",
                        ),
                    ),
                    className="my-3",
                ),
                dbc.Row(
                    dbc.Col(
                        dbc.Card(
                            dbc.CardBody(
                                [
                                    html.H5("Deal Detail", className="card-title"),
                                    dcc.Loading(html.Div(id="trade-table")),
                                ]
                            ),
                            className="shadow-sm",
                        ),
                    ),
                    className="my-3",
                ),
                html.Footer(
                    html.Small(
                        "Built with Dash · Includes deals with delivery windows of "
                        "≥ 28 days, excluding any '(Daily)' hubs. Positions are "
                        "tallied by net Lots per delivery month (Bought = +, "
                        "Sold = −). All prices are volume-weighted by MWh.",
                        className="text-muted",
                    ),
                    className="text-center my-4",
                ),
            ],
            fluid=True,
        ),
    ]
)


# --------------------------------------------------------------------------- #
# Callbacks
# --------------------------------------------------------------------------- #

@app.callback(
    Output("store-deals", "data"),
    Output("upload-status", "children"),
    Output("hub-filter", "options"),
    Output("product-filter", "options"),
    Output("date-filter", "min_date_allowed"),
    Output("date-filter", "max_date_allowed"),
    Output("date-filter", "start_date"),
    Output("date-filter", "end_date"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True,
)
def on_upload(contents, filename):
    df, err = parse_uploaded_file(contents, filename)
    if err:
        return (
            dash.no_update,
            dbc.Alert(err, color="danger", className="mb-0"),
            [],
            [],
            None,
            None,
            None,
            None,
        )

    if df.empty:
        return (
            dash.no_update,
            dbc.Alert(
                "No monthly-or-longer deals found in this file.",
                color="warning",
                className="mb-0",
            ),
            [],
            [],
            None,
            None,
            None,
            None,
        )

    hubs = sorted(df["Hub"].dropna().unique().tolist())
    products = sorted(df["Product"].dropna().unique().tolist())
    min_d = df["Begin Date"].min().date()
    max_d = df["End Date"].max().date()

    status = dbc.Alert(
        [
            html.I(className="bi bi-check-circle me-2"),
            f"Loaded {filename} — {len(df):,} monthly+ deals from "
            f"{min_d} to {max_d}.",
        ],
        color="success",
        className="mb-0",
    )
    return (
        df.to_json(date_format="iso", orient="split"),
        status,
        [{"label": h, "value": h} for h in hubs],
        [{"label": p, "value": p} for p in products],
        min_d,
        max_d,
        min_d,
        max_d,
    )


def _kpi(title: str, value: str, color: str, icon: str) -> dbc.Col:
    return dbc.Col(
        dbc.Card(
            dbc.CardBody(
                [
                    html.Div(
                        [
                            html.I(className=f"bi {icon} me-2"),
                            html.Span(title, className="text-muted small"),
                        ]
                    ),
                    html.H4(value, className=f"fw-bold mb-0 {color}"),
                ]
            ),
            className="shadow-sm h-100",
        ),
        lg=2,
        md=4,
        sm=6,
    )


def _fmt_price(value: float) -> str:
    return "—" if value is None or pd.isna(value) else f"${value:,.2f}"


_DIM_HOVER = (
    "<b>%{x|%b %Y} · %{fullData.name}</b><br>"
    "Net Lots: %{y:+,.0f}<br>"
    "Long Lots: %{customdata[1]:,.0f}<br>"
    "Short Lots: %{customdata[2]:,.0f}<br>"
    "Trade VWAP: $%{customdata[0]:,.2f}/MWh"
    "<extra></extra>"
)

_TOTAL_HOVER = (
    "<b>%{x|%b %Y}</b><br>"
    "Net Lots: %{y:+,.0f}<br>"
    "Long Lots: %{customdata[1]:,.0f}<br>"
    "Short Lots: %{customdata[2]:,.0f}<br>"
    "Trade VWAP: $%{customdata[0]:,.2f}/MWh"
    "<extra></extra>"
)


@app.callback(
    Output("kpi-row", "children"),
    Output("position-graph", "figure"),
    Output("cumulative-graph", "figure"),
    Output("hub-graph", "figure"),
    Output("product-graph", "figure"),
    Output("position-summary-table", "children"),
    Output("trade-table", "children"),
    Input("store-deals", "data"),
    Input("hub-filter", "value"),
    Input("product-filter", "value"),
    Input("direction-filter", "value"),
    Input("date-filter", "start_date"),
    Input("date-filter", "end_date"),
)
def update_views(data, hubs, products, directions, start_d, end_d):
    if not data:
        placeholder = _placeholder_figure("Upload a deal report to get started.")
        empty_msg = html.Div(
            "Upload a file above to populate the dashboard.",
            className="text-muted",
        )
        return (
            [],
            placeholder,
            placeholder,
            placeholder,
            placeholder,
            empty_msg,
            empty_msg,
        )

    df = pd.read_json(io.StringIO(data), orient="split")
    df["Trade Date"] = pd.to_datetime(df["Trade Date"])
    df["Begin Date"] = pd.to_datetime(df["Begin Date"])
    df["End Date"] = pd.to_datetime(df["End Date"])
    if "Delivery Month" in df.columns:
        df["Delivery Month"] = pd.to_datetime(df["Delivery Month"])
    else:
        df["Delivery Month"] = df["Begin Date"].dt.to_period("M").dt.to_timestamp()

    if hubs:
        df = df[df["Hub"].isin(hubs)]
    if products:
        df = df[df["Product"].isin(products)]
    if directions:
        df = df[df["Direction"].isin(directions)]
    if start_d:
        df = df[
            df["Delivery Month"]
            >= pd.to_datetime(start_d).to_period("M").to_timestamp()
        ]
    if end_d:
        df = df[
            df["Delivery Month"]
            <= pd.to_datetime(end_d).to_period("M").to_timestamp()
        ]

    if df.empty:
        empty = _placeholder_figure("No deals match the current filters.")
        msg = html.Div("No deals match the current filters.", className="text-muted")
        return ([], empty, empty, empty, empty, msg, msg)

    long_df = df[df["Direction"] == 1]
    short_df = df[df["Direction"] == -1]
    long_lots = float(long_df["Lots"].sum())
    short_lots = float(short_df["Lots"].sum())
    net_lots = long_lots - short_lots
    avg_buy = _weighted_avg(long_df["Price"], long_df["Total Quantity"])
    avg_sell = _weighted_avg(short_df["Price"], short_df["Total Quantity"])

    kpis = [
        _kpi("Deals", f"{len(df):,}", "text-dark", "bi-clipboard-data"),
        _kpi("Long Lots", f"{long_lots:,.0f}", "text-success", "bi-arrow-up-right"),
        _kpi("Short Lots", f"{short_lots:,.0f}", "text-danger", "bi-arrow-down-right"),
        _kpi(
            "Net Lots",
            f"{net_lots:+,.0f}",
            "text-primary" if net_lots >= 0 else "text-danger",
            "bi-bar-chart-line",
        ),
        _kpi("Avg Buy $/MWh", _fmt_price(avg_buy), "text-success", "bi-cash-coin"),
        _kpi("Avg Sell $/MWh", _fmt_price(avg_sell), "text-danger", "bi-cash-stack"),
    ]

    month_total = aggregate_monthly(df)
    hub_month = aggregate_monthly(df, ["Hub"])
    product_month = aggregate_monthly(df, ["Product"])

    # ---- Hub graph: grouped mini-bars per month, VWAP in hover ---- #
    fig_hub = px.bar(
        hub_month,
        x="Delivery Month",
        y="Net Lots",
        color="Hub",
        barmode="group",
        custom_data=["Trade VWAP", "Long Lots", "Short Lots"],
        template=PLOT_TEMPLATE,
        title="Net Position by Hub (Lots)",
    )
    fig_hub.update_traces(hovertemplate=_DIM_HOVER)
    fig_hub.update_layout(
        yaxis_title="Net Lots",
        xaxis_title=None,
        legend_title=None,
        bargap=0.1,
        bargroupgap=0.04,
    )
    fig_hub.update_xaxes(dtick="M1", tickformat="%b %Y")
    fig_hub.add_hline(y=0, line_width=1, line_color="#adb5bd")

    # ---- Cumulative net lots across delivery months ---- #
    cum_df = month_total.sort_values("Delivery Month").copy()
    cum_df["Cumulative Net Lots"] = cum_df["Net Lots"].cumsum()
    fig_cum = go.Figure()
    fig_cum.add_trace(
        go.Scatter(
            x=cum_df["Delivery Month"],
            y=cum_df["Cumulative Net Lots"],
            mode="lines+markers",
            line=dict(color=NET_COLOR, width=2.5),
            fill="tozeroy",
            fillcolor="rgba(31,119,180,0.15)",
            customdata=cum_df[["Trade VWAP", "Net Lots"]].to_numpy(),
            hovertemplate=(
                "<b>%{x|%b %Y}</b><br>"
                "Cumulative Net Lots: %{y:+,.0f}<br>"
                "Month Net Lots: %{customdata[1]:+,.0f}<br>"
                "Month Trade VWAP: $%{customdata[0]:,.2f}/MWh"
                "<extra></extra>"
            ),
            name="Cumulative",
        )
    )
    fig_cum.update_layout(
        template=PLOT_TEMPLATE,
        title="Cumulative Net Lots over Delivery Months",
        yaxis_title="Cumulative Net Lots",
        xaxis_title=None,
    )
    fig_cum.update_xaxes(dtick="M1", tickformat="%b %Y")

    # ---- Net position by month (single bar per month, colored by side) ---- #
    bar_df = month_total.copy()
    bar_df["Position"] = bar_df["Net Lots"].apply(
        lambda v: "Long" if v >= 0 else "Short"
    )
    fig_pos = px.bar(
        bar_df,
        x="Delivery Month",
        y="Net Lots",
        color="Position",
        color_discrete_map={"Long": LONG_COLOR, "Short": SHORT_COLOR},
        custom_data=["Trade VWAP", "Long Lots", "Short Lots"],
        template=PLOT_TEMPLATE,
        title="Net Position by Month (Lots)",
    )
    fig_pos.update_traces(hovertemplate=_TOTAL_HOVER)
    fig_pos.update_layout(
        yaxis_title="Net Lots",
        xaxis_title=None,
        legend_title=None,
        bargap=0.1,
    )
    fig_pos.update_xaxes(dtick="M1", tickformat="%b %Y")
    fig_pos.add_hline(y=0, line_width=1, line_color="#adb5bd")

    # ---- Net position by product (mini-bars per month) ---- #
    fig_product = px.bar(
        product_month,
        x="Delivery Month",
        y="Net Lots",
        color="Product",
        barmode="group",
        custom_data=["Trade VWAP", "Long Lots", "Short Lots"],
        template=PLOT_TEMPLATE,
        title="Net Position by Product (Lots)",
    )
    fig_product.update_traces(hovertemplate=_DIM_HOVER)
    fig_product.update_layout(
        yaxis_title="Net Lots",
        xaxis_title=None,
        legend_title=None,
        bargap=0.1,
        bargroupgap=0.04,
    )
    fig_product.update_xaxes(dtick="M1", tickformat="%b %Y")
    fig_product.add_hline(y=0, line_width=1, line_color="#adb5bd")

    summary_df = build_position_summary(df)
    if summary_df.empty:
        summary_table = html.Div(
            "No positions to summarize.", className="text-muted"
        )
    else:
        display = summary_df.copy()
        display["Delivery Month"] = display["Delivery Month"].dt.strftime("%b %Y")
        for col in ("Long Lots", "Short Lots", "Net Lots"):
            display[col] = display[col].round(0)
        for col in ("Avg Buy ($/MWh)", "Avg Sell ($/MWh)", "Trade VWAP ($/MWh)"):
            display[col] = display[col].round(2)

        summary_table = dash_table.DataTable(
            data=display.to_dict("records"),
            columns=[
                {"name": "Delivery Month", "id": "Delivery Month"},
                {"name": "Hub", "id": "Hub"},
                {"name": "Product", "id": "Product"},
                {
                    "name": "Long Lots",
                    "id": "Long Lots",
                    "type": "numeric",
                    "format": {"specifier": ",.0f"},
                },
                {
                    "name": "Short Lots",
                    "id": "Short Lots",
                    "type": "numeric",
                    "format": {"specifier": ",.0f"},
                },
                {
                    "name": "Net Lots",
                    "id": "Net Lots",
                    "type": "numeric",
                    "format": {"specifier": "+,.0f"},
                },
                {
                    "name": "Avg Buy ($/MWh)",
                    "id": "Avg Buy ($/MWh)",
                    "type": "numeric",
                    "format": {"specifier": ",.2f"},
                },
                {
                    "name": "Avg Sell ($/MWh)",
                    "id": "Avg Sell ($/MWh)",
                    "type": "numeric",
                    "format": {"specifier": ",.2f"},
                },
                {
                    "name": "Trade VWAP ($/MWh)",
                    "id": "Trade VWAP ($/MWh)",
                    "type": "numeric",
                    "format": {"specifier": ",.2f"},
                },
            ],
            page_size=15,
            sort_action="native",
            filter_action="native",
            style_table={"overflowX": "auto"},
            style_cell={
                "padding": "6px 10px",
                "fontSize": 13,
                "fontFamily": "Inter, system-ui, sans-serif",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#f1f3f5",
                "border": "none",
            },
            style_cell_conditional=[
                {
                    "if": {"column_id": "Trade VWAP ($/MWh)"},
                    "backgroundColor": "#f8f9fa",
                    "fontWeight": "600",
                },
            ],
            style_data_conditional=[
                {
                    "if": {
                        "filter_query": "{Net Lots} > 0",
                        "column_id": "Net Lots",
                    },
                    "color": LONG_COLOR,
                    "fontWeight": "600",
                },
                {
                    "if": {
                        "filter_query": "{Net Lots} < 0",
                        "column_id": "Net Lots",
                    },
                    "color": SHORT_COLOR,
                    "fontWeight": "600",
                },
            ],
        )

    show_cols = [
        c
        for c in [
            "Trade Date",
            "B/S",
            "Product",
            "Hub",
            "Contract",
            "Begin Date",
            "End Date",
            "Price",
            "Lots",
            "Total Quantity",
        ]
        if c in df.columns
    ]
    tbl_df = df[show_cols].copy().sort_values("Trade Date", ascending=False)
    for col in ("Trade Date", "Begin Date", "End Date"):
        if col in tbl_df.columns:
            tbl_df[col] = pd.to_datetime(tbl_df[col]).dt.strftime("%Y-%m-%d")

    table = dash_table.DataTable(
        data=tbl_df.to_dict("records"),
        columns=[{"name": c, "id": c} for c in tbl_df.columns],
        page_size=15,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={
            "padding": "6px 10px",
            "fontSize": 13,
            "fontFamily": "Inter, system-ui, sans-serif",
        },
        style_header={
            "fontWeight": "bold",
            "backgroundColor": "#f1f3f5",
            "border": "none",
        },
        style_data_conditional=[
            {
                "if": {"filter_query": '{B/S} = "Bought"'},
                "color": LONG_COLOR,
            },
            {
                "if": {"filter_query": '{B/S} = "Sold"'},
                "color": SHORT_COLOR,
            },
        ],
    )

    return kpis, fig_pos, fig_cum, fig_hub, fig_product, summary_table, table


if __name__ == "__main__":
    app.run(debug=True)
