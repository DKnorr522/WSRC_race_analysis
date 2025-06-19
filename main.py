import streamlit as st
import pandas as pd
from matplotlib import pyplot as plt
import seaborn as sns
from plotly import express as px, graph_objects as go, figure_factory as ff
from plotly.subplots import make_subplots
from math import ceil, floor
import os
import openpyxl


def fetchExcelFile(file_name: str):
    try:
        wb = openpyxl.load_workbook(file_name)
    except FileNotFoundError as err:
        print(f"File not found: {err}")
        wb = None
    except InvalidFileException as err:
        print(f"File must be .xlsx: {err}")
    
    return wb

def loadDataFrame(wb, event_name: str) -> pd.DataFrame:
    try:
        df: pd.DataFrame = pd.DataFrame(wb[event_name].values)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    except KeyError as err:
        print(f"Invalid worksheet name: {err}")
        df = None
    return df

def cleanDataFrame(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
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
        for col in df.columns
    ]

    df["elapsed_time_sec"] = df.elapsed_time.apply(
        lambda x: x.hour*3600 + x.minute*60 + x.second + x.microsecond*1e-6
    )
    df["split_sec"] = df.split_gps.apply(
        lambda x: x.hour*3600 + x.minute*60 + x.second + x.microsecond*1e-6
    )

    return df

def createLinePlotSpeedStrokeRate(
        df: pd.DataFrame,
        strokes_to_ignore: int = 5,
        breakdown: bool = False):# -> go._figure.Figure:
    """
    Adapted secondary y-axis from:
    https://stackoverflow.com/questions/62853539/how-to-plot-on-secondary-y-axis-with-plotly-express
    User: derflo
    On: 6/14/25
    """

    if strokes_to_ignore > 0:
        df = df.loc[df.total_strokes > strokes_to_ignore, :]

    all_figs = make_subplots(specs=[[{"secondary_y": True}]])

    fig1 = px.line(
        df,
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
    fig1.update_traces({'name': "Speed"})

    fig2 = px.line(
        df,
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
    fig2.update_traces({'name': "Stroke Rate"}, yaxis="y2")

    all_figs.add_traces(fig1.data + fig2.data)
    all_figs.layout.xaxis.title="Distance (m)"
    all_figs.layout.yaxis.title="Speed (m/s)"
    all_figs.layout.yaxis2.title="Stroke Rate"
    # all_figs.layout.title=f"{file_name.split('.')[0]}: WSRC {sheet_name}"

    all_figs.for_each_trace(lambda t: t.update(
        line=dict(color=t.marker.color),
        showlegend=True
    ))

    if breakdown:
        speed_lim_max = 5 * floor((500 / df.loc[df.total_strokes > 5, :].speed_gps.max())/5)
        speed_lim_min = 5 * ceil((500 / df.loc[df.total_strokes > 5, :].speed_gps.min())/5)

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

        high_strokes_first, *_, high_strokes_last = df.loc[
            (df.total_strokes >= 6) & (df.total_strokes <= 10), :
        ].distance_gps.values

        all_figs.add_vrect(
            x0=high_strokes_first,
            x1=high_strokes_last,
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
            fillcolor="maroon",
            opacity=0.1
        )

        speed_max_dist = df.loc[df.speed_gps == df.speed_gps.max(), :].distance_gps.values
        speed_max = [df.speed_gps.max()] * len(speed_max_dist)
        all_figs.add_trace(go.Scatter(
            x=speed_max_dist,
            y=speed_max,
            mode="markers+text",
            name="Fastest",
            text="Fastest",
            textposition="top center"
        ))

        speed_min_dist = df.loc[df.speed_gps == df.speed_gps.min(), :].distance_gps.values
        speed_min = [df.speed_gps.min()] * len(speed_min_dist)
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
        df: pd.DataFrame,
        strokes_to_ignore: int = 5):
    if strokes_to_ignore > 0:
        df = df.loc[df.total_strokes > strokes_to_ignore, :]
    df.astype({'stroke_rate': 'int32'})
    fig = px.scatter(
        df,
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
        color_continuous_scale='aggrnyl',
        # title=f"{file_name.split('.')[0]}: WSRC {sheet_name}",
    )

    fig.update_xaxes(range=[0, 1000])
    return fig

def createLinePlotStrokeRateColoredSpeed(
        df: pd.DataFrame,
        strokes_to_ignore: int = 5):
    if strokes_to_ignore > 0:
        df = df.loc[df.total_strokes > strokes_to_ignore, :]
    fig = px.scatter(
        df,
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
        df: pd.DataFrame,
        strokes_to_ignore: int = 5):
    if strokes_to_ignore > 0:
        df = df.loc[df.total_strokes > strokes_to_ignore, :]
    fig = px.box(
        df,
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

def plotCourseMap(df: pd.DataFrame, size: float=0.1, zoom: int=14) -> None:
    st.map(
        df,
        latitude="gps_lat",
        longitude="gps_lon",
        size=size,
        zoom=zoom
    )
    return


file_name = "2025 Biernacki.xlsx"

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

    plotCourseMap(df, 0.1, 14)

