

import os
import dash
import dash_bootstrap_components as dbc
import dash_table as dt
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
from apscheduler.schedulers.background import BackgroundScheduler
import requests
import zipfile
import io
import xml.etree.ElementTree as ET
import pandas as pd
import datetime
import time
import threading
import random
import atexit
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version

# Run this in powershell
# rsconnect deploy dash --server https://connect.teainc.org/ --api-key NrOWfiwUBd3XBym4H871XkTzr96i1Pu1 ./

# Directory to save the data files in
# DATA_DIR = r"C:/Users/jkuo/OneDrive - TEA/Documents/CAISO-ATC-Tracker/data"
DATA_DIR = r"/mnt/SEAFS1/Jasmine - Copy/CAISO-ATC-Tracker_datapulls"
atc_data = {}  # Cache for DAM, HASP, RTPD DataFrames

# Uncomment when deploying
username = os.environ.get('WINDOWSUSER')
password = os.environ.get('WINDOWSPASS')

# For local testing, set your SharePoint credentials here. DO NOT commit to a repository with these visible. Get them from LastPass.


# Auth with SharePoint Online
authcookie = Office365('https://theenergyauthorityinc.sharepoint.com', username=username, password=password).GetCookies()
site = Site('https://theenergyauthorityinc.sharepoint.com/sites/WestTrading', version=Version.v365, authcookie=authcookie)

# Accessing the document library
target_folder = site.Folder('Shared Documents/Process Automation Documents/CAISO-ATC-Tracker Files')

rtpd_lock = threading.Lock()  # Lock for RTPD data refresh
hasp_lock = threading.Lock()  # Lock for HASP data refresh
dam_lock = threading.Lock()  # Lock for DAM data refresh

# Track background updates (markets + time)
data_update_flag = {
    "updated": False,
    "markets": [],
    "timestamp": ""
}

# Function to pull data from Oasis and converting the .xml file to a .csv file
def get_lmp_data_csv(start_dt, end_dt, node="TH_NP15_GEN-APND", market="DAM", save_dir="./", csv_filename="caiso_lmp.csv"):
    url = "http://oasis.caiso.com/oasisapi/SingleZip"
    params = {
        "queryname": "TRNS_ATC",
        "market_run_id": market,
        "ti_id": "ALL",
        "startdatetime": start_dt,
        "enddatetime": end_dt,
        "version": "1"
    }
    response = requests.get(url, params=params)
    if response.status_code != 200:
        raise Exception(f"Failed to download OASIS data: Status {response.status_code}")

    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        xml_name = [f for f in z.namelist() if f.endswith(".xml")][0]
        xml_file = z.open(xml_name)

        tree = ET.parse(xml_file)

    ns = {'ns': 'http://www.caiso.com/soa/OASISReport_v1.xsd'}
    root = tree.getroot()

    data = []
    for report_item in root.findall(".//ns:REPORT_ITEM", ns):
        for report_data in report_item.findall("ns:REPORT_DATA", ns):
            try:
                row = {
                    "DATA_ITEM": report_data.find("ns:DATA_ITEM", ns).text,
                    "RESOURCE_NAME": report_data.find("ns:RESOURCE_NAME", ns).text,
                    "DIRECTION": report_data.find("ns:DIRECTION", ns).text,
                    "OPR_DATE": report_data.find("ns:OPR_DATE", ns).text,
                    "INTERVAL_NUM": report_data.find("ns:INTERVAL_NUM", ns).text,
                    "INTERVAL_START_GMT": report_data.find("ns:INTERVAL_START_GMT", ns).text,
                    "INTERVAL_END_GMT": report_data.find("ns:INTERVAL_END_GMT", ns).text,
                    "VALUE": float(report_data.find("ns:VALUE", ns).text),
                }
                data.append(row)
            except Exception as e:
                print(f"Skipping row due to error: {e}")

    df = pd.DataFrame(data)

    df_pivot = df.pivot_table(index=["OPR_DATE", "INTERVAL_NUM", "INTERVAL_START_GMT", "INTERVAL_END_GMT", "RESOURCE_NAME", "DIRECTION"],
                               columns="DATA_ITEM",
                               values="VALUE").reset_index()
    
    df_pivot["INTERVAL_NUM"] = df_pivot["INTERVAL_NUM"].astype(int)
    df_pivot = df_pivot.sort_values(by="INTERVAL_NUM").reset_index(drop=True)

    # os.makedirs(save_dir, exist_ok=True)
    # full_path = os.path.join(save_dir, csv_filename)

    # # Delete files with the same name if they exist
    # if os.path.exists(full_path):
    #     os.remove(full_path)
    #     print(f"Removed existing file: {full_path}")

    # # Save the dataframe to the data directory
    # df_pivot.to_csv(full_path, index=False)
    # print(f"Saved CSV to {full_path}")       
    file_content = df_pivot.to_csv(index=False).encode('utf-8')
    
    # Upload the file content to the document library
    target_folder.upload_file(file_content, csv_filename)

    return df_pivot

from office365.sharepoint.files.file import File

def refresh_rtpd_data():
    global atc_data

    # Authenticate with SharePoint Online
    authcookie = Office365(
        'https://theenergyauthorityinc.sharepoint.com',
        username=username,
        password=password
    ).GetCookies()

    site = Site(
        'https://theenergyauthorityinc.sharepoint.com/sites/WestTrading',
        version=Version.v365,
        authcookie=authcookie
    )

    with rtpd_lock:
        now = datetime.datetime.now()
        print(f"Refreshing RTPD data at {now.strftime('%Y-%m-%d %H:%M:%S')}")

        rtpd_filename = f"TRNS_ATC_RTPD_{datetime.date.today()}.csv"
        target_folder = site.Folder('Shared Documents/Process Automation Documents/CAISO-ATC-Tracker Files')

        # Check if file exists in SharePoint
        try:
            existing_files = [f['Name'] for f in target_folder.files]
            if rtpd_filename in existing_files:
                print(f"File found on SharePoint: {rtpd_filename}")

                time.sleep(1) # Sleep for a second before deleting the old file to avoid deleting all of them

                # Delete existing file
                target_folder.delete_file(rtpd_filename)
                print(f"Deleted existing file: {rtpd_filename}")
            else:
                print(f"No existing RTPD file to delete.")
        except Exception as e:
            print(f"Error checking/deleting RTPD file: {e}")

        # Now download fresh data from CAISO
        try:
            atc_data["RTPD"] = get_lmp_data_csv(
                start_dt=now.strftime("%Y%m%dT07:00-0000"),
                end_dt=(now + datetime.timedelta(days=1)).strftime("%Y%m%dT07:00-0000"),
                market="RTPD",
                save_dir=DATA_DIR,
                csv_filename=rtpd_filename
            )
            print("RTPD data refreshed.")

            # Upload the new file to SharePoint
            file_content = atc_data["RTPD"].to_csv(index=False).encode('utf-8')
            target_folder.upload_file(file_content, rtpd_filename)
            print(f"Uploaded refreshed RTPD file to SharePoint: {rtpd_filename}")

            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data_update_flag["updated"] = True
            data_update_flag["markets"].append("RTPD")
            data_update_flag["timestamp"] = now_str

        except Exception as e:
            if "Status 429" in str(e):
                print("CAISO API throttled (429). Retrying in 10 seconds...")
                time.sleep(10)
            else:
                raise


def refresh_hasp_data():
    global atc_data

    # Authenticate with SharePoint Online
    authcookie = Office365(
        'https://theenergyauthorityinc.sharepoint.com',
        username=username,
        password=password
    ).GetCookies()

    site = Site(
        'https://theenergyauthorityinc.sharepoint.com/sites/WestTrading',
        version=Version.v365,
        authcookie=authcookie
    )

    with hasp_lock:
        now = datetime.datetime.now()
        print(f"Refreshing HASP data at {now.strftime('%Y-%m-%d %H:%M:%S')}")

        hasp_filename = f"TRNS_ATC_HASP_{datetime.date.today()}.csv"
        target_folder = site.Folder('Shared Documents/Process Automation Documents/CAISO-ATC-Tracker Files')

        # Check if file exists in SharePoint
        try:
            existing_files = [f['Name'] for f in target_folder.files]
            if hasp_filename in existing_files:
                print(f"File found on SharePoint: {hasp_filename}")

                # Delete existing file
                target_folder.delete_file(hasp_filename)
                print(f"Deleted existing file: {hasp_filename}")
            else:
                print(f"No existing HASP file to delete.")
        except Exception as e:
            print(f"Error checking/deleting HASP file: {e}")

        # Now download fresh data from CAISO
        try:
            atc_data["HASP"] = get_lmp_data_csv(
                start_dt=now.strftime("%Y%m%dT07:00-0000"),
                end_dt=(now + datetime.timedelta(days=1)).strftime("%Y%m%dT07:00-0000"),
                market="HASP",
                save_dir=DATA_DIR,
                csv_filename=hasp_filename
            )
            print("HASP data refreshed.")

            # Upload the new file to SharePoint
            file_content = atc_data["HASP"].to_csv(index=False).encode('utf-8')
            target_folder.upload_file(file_content, hasp_filename)
            print(f"Uploaded refreshed HASP file to SharePoint: {hasp_filename}")

            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data_update_flag["updated"] = True
            data_update_flag["markets"].append("HASP")
            data_update_flag["timestamp"] = now_str

        except Exception as e:
            if "Status 429" in str(e):
                print("CAISO API throttled (429). Retrying in 10 seconds...")
                time.sleep(10)
            else:
                raise


def refresh_dam_data():
    global atc_data

    # Authenticate with SharePoint Online
    authcookie = Office365(
        'https://theenergyauthorityinc.sharepoint.com',
        username=username,
        password=password
    ).GetCookies()

    site = Site(
        'https://theenergyauthorityinc.sharepoint.com/sites/WestTrading',
        version=Version.v365,
        authcookie=authcookie
    )

    with rtpd_lock:
        now = datetime.datetime.now()
        print(f"Refreshing DAM data at {now.strftime('%Y-%m-%d %H:%M:%S')}")

        dam_filename = f"TRNS_ATC_DAM_{datetime.date.today()}.csv"
        target_folder = site.Folder('Shared Documents/Process Automation Documents/CAISO-ATC-Tracker Files')

        # Check if file exists in SharePoint
        try:
            existing_files = [f['Name'] for f in target_folder.files]
            if dam_filename in existing_files:
                print(f"File found on SharePoint: {dam_filename}")

                # Delete existing file
                target_folder.delete_file(dam_filename)
                print(f"Deleted existing file: {dam_filename}")
            else:
                print(f"No existing DAM file to delete.")
        except Exception as e:
            print(f"Error checking/deleting DAM file: {e}")

        # Now download fresh data from CAISO
        try:
            atc_data["DAM"] = get_lmp_data_csv(
                start_dt=now.strftime("%Y%m%dT07:00-0000"),
                end_dt=(now + datetime.timedelta(days=1)).strftime("%Y%m%dT07:00-0000"),
                market="DAM",
                save_dir=DATA_DIR,
                csv_filename=dam_filename
            )
            print("DAM data refreshed.")

            # Upload the new file to SharePoint
            file_content = atc_data["DAM"].to_csv(index=False).encode('utf-8')
            target_folder.upload_file(file_content, dam_filename)
            print(f"Uploaded refreshed DAM file to SharePoint: {dam_filename}")

            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data_update_flag["updated"] = True
            data_update_flag["markets"].append("DAM")
            data_update_flag["timestamp"] = now_str

        except Exception as e:
            if "Status 429" in str(e):
                print("CAISO API throttled (429). Retrying in 10 seconds...")
                time.sleep(10)
            else:
                raise



# Ensure data is not stale
def is_data_stale(market, site, max_age_minutes=15):
    """
    Check if a file in SharePoint is stale (older than max_age_minutes) or missing.

    Args:
        market (str): "RTPD", "HASP", or "DAM"
        site (Site): Authenticated SharePoint site object
        max_age_minutes (int): Threshold for staleness
    Returns:
        bool: True if file is missing or stale, False otherwise
    """
    fname = f"TRNS_ATC_{market}_{datetime.date.today()}.csv"
    sharepoint_folder = site.Folder('Shared Documents/Process Automation Documents/CAISO-ATC-Tracker Files')

    try:
        # Get metadata for all files in folder
        files = sharepoint_folder.files
        file_meta = next((f for f in files if f['Name'] == fname), None)

        if file_meta is None:
            print(f"{fname} does not exist in SharePoint")
            return True  # Treat missing file as stale

        # Parse modified date
        file_mtime_str = file_meta['TimeLastModified']
        file_mtime_dt = datetime.datetime.strptime(file_mtime_str, "%Y-%m-%dT%H:%M:%SZ")
        file_mtime_dt = file_mtime_dt.replace(tzinfo=datetime.timezone.utc).astimezone(tz=None)

        # Calculate file age
        age_minutes = (datetime.datetime.now() - file_mtime_dt).total_seconds() / 60
        print(f"{fname} last modified {age_minutes:.2f} minutes ago")

        return age_minutes > max_age_minutes  # True = stale
    except Exception as e:
        print(f"Error checking staleness of {fname}: {e}")
        return True  # Conservative: treat as stale on error



# Set up a background scheduler to refresh data every 15 minutes
scheduler = BackgroundScheduler()

def start_scheduler():

    if scheduler.get_jobs():
        print("Stopping old scheduler jobs...")
        scheduler.remove_all_jobs()

    print("Starting APScheduler with fresh jobs...")
    scheduler.add_job(refresh_rtpd_data, 'cron', minute='0,15,30,45', second=45)  # Stagger API pulls to avoid throttling errors
    scheduler.add_job(refresh_hasp_data, 'cron', minute=0, second=50)
    scheduler.add_job(refresh_dam_data, 'cron', hour=0, minute=0, second=55)  # Daily at midnight, since traders usually pull data at the start of the day
    scheduler.start()

# Call it outside __main__ so it works in WSGI
start_scheduler()

atexit.register(lambda: scheduler.shutdown(wait=False))

def load_csv_from_sharepoint(market, site):
        
        today = datetime.date.today()
        fname = f"TRNS_ATC_{market}_{today}.csv"

        # Path to folder in SharePoint
        sharepoint_folder = site.Folder('Shared Documents/Process Automation Documents/CAISO-ATC-Tracker Files')

        try:
            # Get file content from SharePoint
            file_bytes = sharepoint_folder.get_file(fname)

            # Read content into pandas DataFrame
            print(f"Loading {fname} from SharePoint")
            return pd.read_csv(io.BytesIO(file_bytes))
        except Exception as e:
            if "cannot find file" in str(e).lower() or "not found" in str(e).lower():
                print(f"No file found for {market} in SharePoint")
                return pd.DataFrame()
            else:
                print(f"Error loading {fname} from SharePoint: {e}")
                return pd.DataFrame()
            

def preload_data():
    """
    Preload existing CSVs from SharePoint into atc_data.
    """
    print("Preloading existing CSVs from SharePoint...")

    # Load each market's data
    atc_data["RTPD"] = load_csv_from_sharepoint("RTPD", site)
    atc_data["HASP"] = load_csv_from_sharepoint("HASP", site)
    atc_data["DAM"] = load_csv_from_sharepoint("DAM", site)
    print("Preloading complete. Data loaded into atc_data cache.")



# -------------------------- Dash app setup-------------------------- #

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])


important_itc = ['IPPDCADLN_ITC', 'MALIN500_ISL', 'ADLANTOVICTVL-SP_ITC', 'ADLANTO-SP_ITC', 'GONIPPDC_ITC', 
                 'IPPUTAH_ITC', 'MONAIPPDC_ITC', 'NOB_ITC', 'SYLMAR-AC_ITC', 'PALOVRDE_ITC', 
                 'MEAD_ITC', 'BLYTHE_ITC', 'MEADMKTPC_ITC', 'MKTPCADLN_ITC', 'WSTWGMEAD_ITC']


# top nav bar
nav = dbc.Navbar(
    dbc.NavbarBrand("CAISO ATC Tracker", style={"padding-left": "20px"}),  # optional: ms-3 adds extra margin
    color="blue",
    sticky="top",
)

# left side grouping of selction options
form_card_group = dbc.Card(
    [
        dbc.Row(
            [
                dbc.Label("Choose an ITC", width=10, style={"fontWeight": "bold"}),
                dbc.Col(
                    dcc.Dropdown(
                        id="itc-select",
                        options=[
                            {
                                "label": itc,
                                "value": itc,
                            }
                            for itc in important_itc
                        ],
                        multi=False,
                        value=important_itc[0],
                    ),
                ),
            ]
        ),
        dbc.Row(
            [
                dbc.Label("Market", width="auto", style={"fontWeight": "bold"}),
                dbc.Col(
                    dbc.Checklist(
                        id="itc-market-select",
                        options=[
                            {
                                "label": "DAM",
                                "value": "DAM",
                            },
                            {
                                "label": "HASP",
                                "value": "HASP",
                            },
                            {
                                "label": "RTPD",
                                "value": "RTPD",
                            },
                        ],
                        value=["DAM"],
                        inline=True
                    ),
                    width=10,
                ),
            ]
        ),
        dbc.Row(
            [
                dbc.Label("Flow Direction", width="auto", style={"fontWeight": "bold"}),
                dbc.Col(
                    dbc.Checklist(
                        id="flow-direction-select",
                        options=[
                            {"label": "Imports", "value": "I"},
                            {"label": "Exports", "value": "E"},
                        ],
                        value=["E"],  # Default: Exports
                        inline=True
                    ),
                    width=10,
                ),
            ]
        ),
        html.Div(
            [
                html.Pre(id="selected-data"),
            ],
        ),
        dbc.Button("Refresh Data", id="manual-refresh-button", color="primary", className="mt-3"),
    ],
    body=True,
)

# sidebar
SIDEBAR_STYLE = {
    "float": "left",
    "top": "50px",
    "left": 0,
    "bottom": 0,
    "width": "28rem",
    "padding": "2rem 1rem",
}

sidebar = html.Div(
    form_card_group,
    style=SIDEBAR_STYLE,
)

# price and volume graphs
graphs = [
    dbc.Alert(
        "📊 Data is updated every 15 minutes for RTPD, hourly for HASP, and daily for DAM. Hit \"Refresh Data\" to get the latest updates.",
        color="info",
    ),
    dcc.Loading(
        id="loading-graph",
        type="circle",  # or "default", "dot"
        children=dcc.Graph(id="atc-graph", style={"height": "600px"})
    )
]

atc_table = dcc.Loading(
    id="loading-tables",
    type="circle",  # You can also use "default", "dot", or "cube"
    children=html.Div(id="atc-tables-container")
)

body_container = dbc.Container(
    [
        html.Div(
            children=[
                dbc.Row(
                    [
                        dbc.Col(
                            sidebar,
                            md=4,
                        ),
                        dbc.Col(
                            graphs,
                            md=8,
                        ),
                    ],
                ),
            ],
            className="m-4",
        ),
        html.Div(
            [
                dbc.Row([dbc.Col([atc_table])]),
            ],
        ),
        html.Div(id="dummy-output", style={"display": "none"})
    ],
    fluid=True,
)

preload_data()  # Load existing data on startup

# main app ui entry
app.layout = html.Div([nav,
                       dcc.Store(id="data-refresh-status", data={"alert": False}),
                       dcc.Store(id="manual-refresh-store"),
                       dcc.Store(id="alert-store", data={"show_alert": False}),
                        dbc.Alert(
                            id="new-data-alert",
                            is_open=False,
                            color="success",
                            duration=None,
                            children=dcc.Markdown(id="alert-text"),
                            style={"margin": "10px"}
                        ),
                        body_container,
                        dcc.Interval(id="interval-check", interval=30*1000, n_intervals=0)])


# Update the ATC graph

@app.callback(
    Output("atc-graph", "figure"),
    [Input("itc-select", "value"),
     Input("itc-market-select", "value"),
     Input("flow-direction-select", "value"),
     Input("manual-refresh-store", "data")],
)
def update_atc_graph(itc_name, selected_markets, selected_directions, refresh_data):
    fig = go.Figure()
    colors = {"DAM": "blue", "HASP": "green", "RTPD": "orange"}
    
    for market_name in selected_markets:
        df = atc_data.get(market_name)
        if df is None:
            continue  # Skip if no data yet
        df_filtered = df[df["RESOURCE_NAME"] == itc_name]

        for direction in selected_directions:
            df_dir = df_filtered[df_filtered["DIRECTION"] == direction].copy()
            if df_dir.empty:
                continue
            df_dir["INTERVAL_START_GMT"] = pd.to_datetime(df_dir["INTERVAL_START_GMT"]) - pd.Timedelta(hours=6)

            dir_label = "Imports" if direction == "I" else "Exports"
            line_style = "solid" if direction == "I" else "dot"

            fig.add_trace(go.Scatter(
                x=df_dir["INTERVAL_START_GMT"],
                y=df_dir["ATC_MW"],
                mode="lines+markers",
                name=f"{market_name} {dir_label}",
                line=dict(color=colors[market_name], dash=line_style)
            ))

    fig.update_layout(
        title=f"ATC for {itc_name}",
        xaxis_title="Interval Start Time (PST)",
        yaxis_title="Available Transfer Capability (MW)",
        hovermode="x unified",
        legend=dict(x=0, y=1)
    )

    return fig



# Update the ATC table

@app.callback(
    Output("atc-tables-container", "children"),
    [Input("itc-select", "value"),
     Input("itc-market-select", "value"),
     Input("manual-refresh-store", "data")]
)
def update_atc_tables(itc_name, selected_markets, refresh_data):
    tables = []

    for market_name in selected_markets:
        df = atc_data.get(market_name)
        if df is None or df.empty:
            continue

        df_filtered = df[df["RESOURCE_NAME"] == itc_name].copy()

        if market_name in ["RTPD", "HASP"]:
            df_filtered["DIRECTION"] = df_filtered["DIRECTION"].replace({"I": "Import", "E": "Export"})
            df_filtered["HE"] = (df_filtered["INTERVAL_NUM"] - 1) // 4 + 1
            df_filtered["QTR"] = ((df_filtered["INTERVAL_NUM"] - 1) % 4 + 1)  # 1=00, 2=15, 3=30, 4=45


            # Create a complete multi-index for all possible hours and quarters
            full_index = pd.MultiIndex.from_product(
                [df_filtered["DIRECTION"].unique(), range(1, 25), range(1, 5)],
                names=["DIRECTION", "HE", "QTR"]
            )

            # Pivot and reindex to fill missing intervals
            df_pivot = df_filtered.pivot_table(
                index=["DIRECTION", "HE", "QTR"],
                values="ATC_MW",
                aggfunc="mean"  # or "last" if you prefer
            ).reindex(full_index).reset_index()

            # Format column names nicely
            df_pivot["Interval"] = df_pivot.apply(
                lambda row: f"HE{str(row['HE']).zfill(2)}:{['00', '15', '30', '45'][row['QTR'] - 1]}",
                axis=1
            )

            df_pivot = df_pivot.pivot_table(index="DIRECTION", columns="Interval", values="ATC_MW")

        else:  # DAM
            df_filtered["HE"] = df_filtered["INTERVAL_NUM"]
            df_pivot = df_filtered.pivot_table(
                index="DIRECTION",
                columns="HE",
                values="ATC_MW"
            )
            df_pivot.columns = [f"HE{str(he).zfill(2)}" for he in df_pivot.columns]

        df_pivot = df_pivot.rename(index={"I": "Import", "E": "Export", "DIRECTION": "Direction"}).reset_index()

        table_data = df_pivot.to_dict("records")
        table_columns = [{"name": col, "id": col} for col in df_pivot.columns]

        # Add title and table
        tables.append(
            dbc.Card(
                dbc.CardBody([
                    html.H5(f"{market_name} Market Table", className="card-title"),
                    dt.DataTable(
                        data=table_data,
                        columns=table_columns,
                        style_table={"overflowX": "auto", "marginBottom": "30px"},
                        style_cell={"textAlign": "center", "padding": "5px"},
                        style_header={"backgroundColor": "lightgrey", "fontWeight": "bold"},
                    )
                ]),
                className="mb-4"
            )
        )

    if not tables:
        return html.Div("No data available for the selected markets.", style={"marginTop": "20px"})

    return tables


@app.callback(
    [Output("new-data-alert", "is_open"),
     Output("alert-text", "children"),
     Output("alert-store", "data")],
    [Input("interval-check", "n_intervals"),
     Input("manual-refresh-button", "n_clicks")],
    [dash.dependencies.State("alert-store", "data")]
)
def manage_new_data_alert(n_intervals, n_clicks, alert_data):
    ctx = dash.callback_context

    if not ctx.triggered:
        return dash.no_update, dash.no_update, alert_data

    triggered_id = ctx.triggered[0]["prop_id"].split(".")[0]

    if triggered_id == "interval-check" and data_update_flag["updated"]:
        markets = ", ".join(data_update_flag["markets"])
        timestamp = data_update_flag["timestamp"]
        message = f"**New data available** for: **{markets}**\n\n Updated at: {timestamp}\n\nHit 'Refresh Data' to refresh."
        # Set alert to show
        return True, message, {"show_alert": True}

    if triggered_id == "manual-refresh-button":
        # User refreshed; close alert
        return False, "", {"show_alert": False}

    return dash.no_update, dash.no_update, alert_data



@app.callback(
    [Output("dummy-output", "children"),
     Output("manual-refresh-store", "data")],
    Input("manual-refresh-button", "n_clicks"),
    prevent_initial_call=True
)
def manual_refresh(n_clicks=None):
    global atc_data

    authcookie = Office365(
        'https://theenergyauthorityinc.sharepoint.com',
        username=username,
        password=password
    ).GetCookies()
    site = Site(
        'https://theenergyauthorityinc.sharepoint.com/sites/WestTrading',
        version=Version.v365,
        authcookie=authcookie
    )

    today = datetime.date.today()

    def load_or_pull(market):
        if is_data_stale(market, site):
            print(f"{market} data stale or missing. Pulling new data...")
            refresh_func = {
                "RTPD": refresh_rtpd_data,
                "HASP": refresh_hasp_data,
                "DAM": refresh_dam_data
            }[market]
            refresh_func()
        else:
            print(f"{market} data is fresh. Loading from SharePoint.")

        # Load latest file from SharePoint
        try:
            atc_data[market] = load_csv_from_sharepoint(market, site)
        except Exception as e:
            print(f"Failed to load {market} from SharePoint: {e}")
            atc_data[market] = pd.DataFrame()

    for market in ["RTPD", "HASP", "DAM"]:
        load_or_pull(market)

    print("Manual or auto data refresh complete.")
    data_update_flag["updated"] = False
    data_update_flag["markets"] = []
    data_update_flag["timestamp"] = ""

    return "", {"refreshed": True}



if __name__ == "__main__":
    app.run(debug=True)
