import streamlit as st
import pandas as pd
from matplotlib import pyplot as plt
import seaborn as sns
from plotly import express as px, graph_objects as go, figure_factory as ff
from plotly.subplots import make_subplots
from math import ceil, floor
import os
import openpyxl


def fetchExcelFile(file_name_temp: str):
    try:
        wb_temp = openpyxl.load_workbook(file_name_temp)
    except FileNotFoundError as err:
        print(f"File not found: {err}")
        wb_temp = None
    except InvalidFileException as err:
        print(f"File must be .xlsx: {err}")
    
    return wb_temp

def loadDataFrame(wb_temp, event_name: str) -> pd.DataFrame:
    try:
        df_temp: pd.DataFrame = pd.DataFrame(wb_temp[event_name].values)
        df_temp.columns = df_temp.iloc[0]
        df_temp = df_temp[1:].reset_index(drop=True)
    except KeyError as err:
        print(f"Invalid worksheet name: {err}")
        df_temp = None
    return df_temp

def cleanDataFrame(df_temp: pd.DataFrame) -> pd.DataFrame:
    df_temp.columns = [
        col.lower().replace(
            " ", "_"
        ).replace(
            "(", ""
        ).replace(
            ")", ""
        ).replace(
            ".", ""
        ).replace(
            "/", "_per_"
        )
        for col in df_temp.columns
    ]

    df_temp["elapsed_time_sec"] = df_temp.elapsed_time.apply(
        lambda x: x.hour*3600 + x.minute*60 + x.second + x.microsecond*1e-6
    )
    df_temp["split_sec"] = df_temp.split_gps.apply(
        lambda x: x.hour*3600 + x.minute*60 + x.second + x.microsecond*1e-6
    )

    return df_temp

def createLinePlotSpeedStrokeRate(
        df_temp: pd.DataFrame,
        strokes_to_ignore_temp: int = 5,
        breakdown_temp: bool = False):# -> go._figure.Figure:
    """
    Adapted secondary y-axis from:
    https://stackoverflow.com/questions/62853539/how-to-plot-on-secondary-y-axis-with-plotly-express
    User: derflo
    On: 6/14/25
    """

    if strokes_to_ignore_temp > 0:
        df_temp = df_temp.loc[df_temp.total_strokes > strokes_to_ignore_temp, :]

    all_figs = make_subplots(specs=[[{"secondary_y": True}]])

    fig1_temp = px.line(
        df_temp,
        x="distance_gps",
        y="speed_gps",
        hover_data=["split_gps", "elapsed_time", "stroke_rate", "distance_per_stroke_gps", "total_strokes"],
        labels={
            "speed_gps": "Speed (m/s)",
            "distance_gps": "Distance (m)",
            "stroke_rate": "Stroke Rate",
            "elapsed_time": "Time",
            "distance_per_stroke_gps": "Meters per Stroke",
            "total_strokes": "Stroke Count"
        },
        # title=f"{file_name.split('.')[0]}: WSRC {sheet_name}",
    )
    fig1_temp.update_traces({'name': "Speed"})

    fig2_temp = px.line(
        df_temp,
        x="distance_gps",
        y="stroke_rate",
        hover_data=["elapsed_time", "stroke_rate", "distance_per_stroke_gps", "total_strokes"],
        labels={
            "speed_gps": "Speed (m/s)",
            "distance_gps": "Distance (m)",
            "stroke_rate": "Stroke Rate",
            "elapsed_time": "Time",
            "distance_per_stroke_gps": "Meters per Stroke",
            "total_strokes": "Stroke Count"
        },
        # title=f"{file_name.split('.')[0]}: WSRC {sheet_name}",
    )
    fig2_temp.update_traces({'name': "Stroke Rate"}, yaxis="y2")

    all_figs.add_traces(fig1_temp.data + fig2_temp.data)
    all_figs.layout.xaxis.title="Distance (m)"
    all_figs.layout.yaxis.title="Speed (m/s)"
    all_figs.layout.yaxis2.title="Stroke Rate"
    # all_figs.layout.title=f"{file_name.split('.')[0]}: WSRC {sheet_name}"

    all_figs.for_each_trace(lambda t: t.update(
        line=dict(color=t.marker.color),
        showlegend=True
    ))

    if breakdown_temp:
        speed_lim_max = 5 * floor((500 / df_temp.loc[df_temp.total_strokes > 5, :].speed_gps.max())/5)
        speed_lim_min = 5 * ceil((500 / df_temp.loc[df_temp.total_strokes > 5, :].speed_gps.min())/5)

        start_speed = speed_lim_max
        while start_speed <= speed_lim_min:
            speed_str = f"{start_speed//60}:{start_speed - 60*(start_speed//60):02d}"

            txt_loc = "bottom left" if start_speed == speed_lim_max else "top left"

            all_figs.add_hline(
                y=500 / start_speed,
                line_dash="dash",
                line_width=1,
                annotation_text=speed_str,
                annotation_position=txt_loc
            )
            start_speed += 5

        high_strokes_first, *_, high_strokes_last = df_temp.loc[
            (df_temp.total_strokes >= 6) & (df_temp.total_strokes <= 10), :
        ].distance_gps.values

        all_figs.add_vrect(
            x0=high_strokes_first,
            x1=high_strokes_last,
            # annotation_text="high strokes",
            # annotation_position="top left",
            fillcolor="blue",
            opacity=0.1
        )

        all_figs.add_vrect(
            x0=250,
            x1=500,
            fillcolor="green",
            opacity=0.1
        )

        all_figs.add_vrect(
            x0=500,
            x1=750,
            fillcolor="yellow",
            opacity=0.1
        )

        all_figs.add_vrect(
            x0=750,
            x1=1000,
            # annotation_text="sprint",
            # annotation_position="top left",
            fillcolor="maroon",
            opacity=0.1
        )

        speed_max_dist = df_temp.loc[df_temp.speed_gps == df_temp.speed_gps.max(), :].distance_gps.values
        speed_max = [df_temp.speed_gps.max()] * len(speed_max_dist)
        all_figs.add_trace(go.Scatter(
            x=speed_max_dist,
            y=speed_max,
            mode="markers+text",
            name="Fastest",
            text="Fastest",
            textposition="top center"
        ))

        speed_min_dist = df_temp.loc[df_temp.speed_gps == df_temp.speed_gps.min(), :].distance_gps.values
        speed_min = [df_temp.speed_gps.min()] * len(speed_min_dist)
        all_figs.add_trace(go.Scatter(
            x=speed_min_dist,
            y=speed_min,
            mode="markers+text",
            name="Slowest",
            text="Slowest",
            textposition="bottom center"
        ))

    all_figs.update_xaxes(range=[0, 1000])
    return all_figs

def createLinePlotSpeedColoredStrokeRate(
        df_temp: pd.DataFrame,
        strokes_to_ignore_temp: int = 5):
    if strokes_to_ignore_temp > 0:
        df_temp = df_temp.loc[df_temp.total_strokes > strokes_to_ignore_temp, :]
    fig = px.scatter(
        df_temp,
        x="distance_gps",
        y="speed_gps",
        color="stroke_rate",
        # color="distance_per_stroke_gps",
        hover_data=["elapsed_time", "stroke_rate", "distance_per_stroke_gps", "total_strokes"],
        labels={
            "speed_gps": "Speed (m/s)",
            "distance_gps": "Distance (m)",
            "stroke_rate": "Stroke Rate",
            "elapsed_time": "Time",
            "distance_per_stroke_gps": "Meters per Stroke",
            "total_strokes": "Stroke Count"
        },
        # color_continuous_scale='aggrnyl',
        # title=f"{file_name.split('.')[0]}: WSRC {sheet_name}",
        # width=1600,
        # height=800
    )

    fig.update_xaxes(range=[0, 1000])
    return fig

def createLinePlotStrokeRateColoredSpeed(
        df_temp: pd.DataFrame,
        strokes_to_ignore_temp: int = 5):
    if strokes_to_ignore_temp > 0:
        df_temp = df_temp.loc[df_temp.total_stroke > strokes_to_ignore_temp, :]
    fig = px.scatter(
        df_temp,
        x="distance_gps",
        y="stroke_rate",
        # color="distance_per_stroke_gps",
        color="speed_gps",
        hover_data=["elapsed_time", "stroke_rate", "distance_per_stroke_gps", "total_strokes"],
        labels={
            "speed_gps": "Speed (m/s)",
            "distance_gps": "Distance (m)",
            "stroke_rate": "Stroke Rate",
            "elapsed_time": "Time",
            "distance_per_stroke_gps": "Meters per Stroke",
            "total_strokes": "Stroke Count"
        },
        color_continuous_scale='aggrnyl',
        # title=f"{file_name.split('.')[0]}: WSRC {sheet_name}",
        # width=1600,
        # height=800
    )

    fig.update_xaxes(range=[0, 1000])
    return fig

def createBoxPlotStrokeRateSpeed(
        df_temp: pd.DataFrame,
        strokes_to_ignore_temp: int = 5):
    if strokes_to_ignore_temp > 0:
        df_temp = df_temp.loc[df_temp.total_strokes > strokes_to_ignore_temp, :]
    fig = px.box(
        df_temp,
        x="stroke_rate",
        y="speed_gps",
        labels={
            "speed_gps": "Speed (m/s)",
            "distance_gps": "Distance (m)",
            "stroke_rate": "Stroke Rate",
            "elapsed_time": "Time",
            "distance_per_stroke_gps": "Meters per Stroke",
            "total_strokes": "Stroke Count"
        },
    )

    return fig


file_name = "2025 Biernacki.xlsx"
sheet_name = "Men's 8+"
# sheet_name = "Men's 4+ (D)"

st.markdown(
    "<h1 style='text-align: center;'> WSRC Race Results </h1>",
    unsafe_allow_html=True
)
st.divider()

col_race, col_breakdown = st.columns(2)

with col_race:
    wb = fetchExcelFile(file_name)
    race_choices = wb.sheetnames
    race_choice = st.selectbox(
        "Choose a race",
        options=race_choices,
        index=None
    )
with col_breakdown:
    breakdown = st.checkbox(
        "Breakdown",
        value=False
    )
    show_start = st.checkbox(
        "Show starting strokes",
        value=False
    )
    strokes_to_ignore = 0 if show_start else 5

if race_choice:
    df = loadDataFrame(wb, race_choice)
    df = cleanDataFrame(df)
    fig1 = createLinePlotSpeedStrokeRate(df, strokes_to_ignore=strokes_to_ignore, breakdown=breakdown)
    st.plotly_chart(fig1)

    fig2 = createLinePlotSpeedColoredStrokeRate(df, strokes_to_ignore=strokes_to_ignore)
    st.plotly_chart(fig2)

    fig3 = createLinePlotStrokeRateColoredSpeed(df, strokes_to_ignore=strokes_to_ignore)
    st.plotly_chart(fig3)

    fig4 = createBoxPlotStrokeRateSpeed(df, strokes_to_ignore=strokes_to_ignore)
    st.plotly_chart(fig4)

    df_course = df[["gps_lat", "gps_lon"]]
    df_course.columns = ["latitude", "longitude"]
    st.map(
        df_course,
        size=1,
        zoom=14
    )


