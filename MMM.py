import subprocess
import sys

# Function to install missing packages
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Try to import packages, and install them if not found
try:
    import streamlit as st
except ImportError:
    install("streamlit")
    import streamlit as st

try:
    import pandas as pd
except ImportError:
    install("pandas")
    import pandas as pd

try:
    import matplotlib.pyplot as plt
except ImportError:
    install("matplotlib")
    import matplotlib.pyplot as plt

try:
    from datetime import datetime
except ImportError:
    install("datetime")  # Note: `datetime` is built into Python, so this is just a fallback
    from datetime import datetime

try:
    import altair as alt
except ImportError:
    install("altair")
    import altair as alt

try:
    from streamlit_echarts import JsCode
except ImportError:
    install("streamlit-echarts")
    from streamlit_echarts import JsCode

try:
    from streamlit_echarts import st_echarts
except ImportError:
    install("streamlit-echarts")
    from streamlit_echarts import st_echarts

# Set wide layout for Streamlit
st.set_page_config(layout="wide")

def remove_POs(df_shift_all):
    pattern = "PO#|PO #|INC #|INC#"
    df_filtered_bad = df_shift_all[df_shift_all['Issue'].str.contains(pattern, regex=True)]
    df_filtered_good = df_shift_all[~df_shift_all['Issue'].str.contains(pattern, regex=True)]

    return df_filtered_good, df_filtered_bad

def get_file_as_dataframe(filename):
    try:
        df = pd.read_excel(filename, sheet_name="End Of Shift Report")
    except:
        print(f"Error reading worksheet = 'End Of Shift Report' from file: {filename}")

    df.columns = df.iloc[0, :]
    df_clean = df.iloc[1:, :]

    df_process = df_clean.dropna(subset = ["Process"])

    # Set Date/Month with previous information
    df_process["Date/Month"] = df_process["Date/Month"].fillna(method='ffill')

    df_process.columns = ["Date/Month","Pending Action","Shift1_Process","Shift1_Issue","Shift1_Action Taken","Shift2_Process","Shift2_Issue","Shift2_Action Taken","Shift3_Process","Shift3_Issue","Shift3_Action Taken","NaN"]

    df_process = df_process[["Date/Month","Shift1_Process","Shift1_Issue","Shift1_Action Taken","Shift2_Process","Shift2_Issue","Shift2_Action Taken","Shift3_Process","Shift3_Issue","Shift3_Action Taken"]]

    return df_process

df_process = get_file_as_dataframe("90 TNB_RPAAMS_End of Shift Report & Support Guide Year 4-2024-Q4.xlsx")


def seperate_shift_df(df_process):
    df_shift1 = df_process[["Date/Month","Shift1_Process","Shift1_Issue","Shift1_Action Taken"]]
    df_shift1.columns = ["Date/Month","Process","Issue","Action Taken"]
    df_shift1["Date/Month"] = pd.to_datetime(df_shift1["Date/Month"])

    df_shift2 = df_process[["Date/Month","Shift2_Process","Shift2_Issue","Shift2_Action Taken"]]
    df_shift2.columns = ["Date/Month","Process","Issue","Action Taken"]
    df_shift2["Date/Month"] = pd.to_datetime(df_shift2["Date/Month"])

    df_shift3 = df_process[["Date/Month","Shift3_Process","Shift3_Issue","Shift3_Action Taken"]]
    df_shift3.columns = ["Date/Month","Process","Issue","Action Taken"]
    df_shift3["Date/Month"] = pd.to_datetime(df_shift3["Date/Month"])

    df_shift_all = pd.concat([df_shift1, df_shift2, df_shift3], axis = 0)
    df_shift_all["Date/Month"] = pd.to_datetime(df_shift_all["Date/Month"])
    df_shift_all_good, df_shift_all_bad = remove_POs(df_shift_all)

    return df_shift1, df_shift2, df_shift3, df_shift_all_good, df_shift_all_bad

# Streamlit app title
st.title("End Of Shift Report Analysis")

# File uploader widget to upload a CSV file
uploaded_file = st.file_uploader("Upload your End Of Shift Report file", type=["xlsx"])


if uploaded_file is not None:
    # Read the uploaded CSV file into a pandas DataFrame    
    process_df = get_file_as_dataframe(uploaded_file)

    df_shift1, df_shift2, df_shift3, df_shift_all, df_shift_all_bad = seperate_shift_df(process_df)

    st.write(f"Removed {len(df_shift_all_bad)} Tickets from calculations, remaining: {len(df_shift_all)} issues")
    st.dataframe(df_shift_all_bad)

    if 'Process' in df_shift_all.columns:

        start_date, end_date = st.date_input(
            "Choose a date range", 
            [df_shift_all["Date/Month"].min(), df_shift_all["Date/Month"].max()],
            min_value = df_shift_all["Date/Month"].min(),
            max_value= df_shift_all["Date/Month"].max()
        )

        col1, col2 = st.columns([0.5, 1.5])

        start_date_str = start_date.strftime("%Y-%m-%d")  # Correct format string
        end_date_str = end_date.strftime("%Y-%m-%d")

        process_counts = df_shift_all[(df_shift_all["Date/Month"] >= pd.to_datetime(start_date)) & 
    (df_shift_all["Date/Month"] <= pd.to_datetime(end_date))]

        process_counts_to_display = process_counts.groupby(by=["Process"]).size().reset_index(name='ProcessCount')

        # Display the counts in the app
        with col1:
            st.write("Count of each unique Process:")
            st.dataframe(process_counts_to_display[["Process", "ProcessCount"]], use_container_width=True)  # Display the DataFrame with renamed columns

        # Plot the count of 'Process' values as a bar chart
        with col2: 
            st.write(f"Visualizing the count of each Process, {start_date_str} - {end_date_str}:")

            option = {
            "tooltip": {
                "trigger": 'axis',
                "axisPointer": {      
                "type": 'shadow'      
                }
            },
            "xAxis": {
                "type": 'category',
                "data": list(process_counts_to_display.Process),
                "axisLabel": {
                "rotate": 90 
                }
            },
            "yAxis": {
                "type": 'value'
            },
            "series": [
                {
                "data": list(process_counts_to_display.ProcessCount),
                "type": 'bar',
                'itemStyle': {
                'color': 'red'  # Set the color of the line to red
            }
                }
            ]}

            
            clicked_label = st_echarts(option,
            height = "500px",
            events = {"click": "function(params) {return params.name}"})

    else:
        st.error("'Process' column not found in the uploaded data")

    clicked_process = process_counts[process_counts["Process"] == clicked_label][["Date/Month","Process","Issue","Action Taken"]]

    st.write(f"{clicked_label}")

    process_counts_to_display = clicked_process[["Date/Month"]].groupby(by=["Date/Month"]).size().reset_index(name='ProcessCount')

    option = {
                "tooltip": {
                    "trigger": 'axis',
                    "axisPointer": {      
                    "type": 'shadow'      
                    }
                },
        'xAxis': {
            'type': 'category',
            'data': list(process_counts_to_display["Date/Month"].dt.strftime('%Y-%m-%dT%H:%M:%S'))
        },
        'yAxis': {
            'type': 'value'
        },
        'series': [
            {
                'data': list(process_counts_to_display["ProcessCount"]),
                'type': 'line',
                'itemStyle': {
                'color': 'red'  # Set the color of the line to red
            }
            }
        ]
    }


    secondary_clicked_label = st_echarts(option,
        height = "300px",
        events = {"click": "function(params) {return params.name}"})
    
    seconday_clicked_process = process_counts[(process_counts["Date/Month"] == secondary_clicked_label) & (process_counts["Process"] == clicked_label)][["Date/Month","Process","Issue","Action Taken"]]
    st.dataframe(seconday_clicked_process, use_container_width = True)

