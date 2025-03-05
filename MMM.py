import subprocess
import sys
import tempfile

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

import re

try:
    import pyperclip
except ImportError:
    install("pyperclip")
    import pyperclip

# Set wide layout for Streamlit
st.set_page_config(layout="wide")



def PR2PO(filepath):
    df = pd.read_excel(filepath, sheet_name="Master Report")
    df.columns = df.iloc[0, :]
    df = df.iloc[1:, :]

    # Find all the positions of 'Status' columns
    status_columns = [col for col in df.columns if col == 'PSS Status']

    # If there are exactly two such columns, rename them
    if len(status_columns) == 2:
        # Create a new list of column names with unique names for 'Status' columns
        new_columns = []
        status_count = 1
        for col in df.columns:
            if col == 'PSS Status':
                if status_count == 1:
                    new_columns.append(f'PSS Status')
                else:
                    new_columns.append(f'PSS Status_{status_count}')
                status_count += 1
            else:
                new_columns.append(col)

        df.columns = new_columns

    pr2po_processed = len(df) - len(df[df["Status"].isna()])
    pr2po_convert_po = len(df[df["Status"] == "Convert to PO"])
    pr2po_pss_status = len(df) - len(df[df["PSS Status"].isna()])
    pr2po_total = len(df)
    pr2po_error = len(df[df["Status"] == "PO Not Released"])
    error_dataframe = df[df["Status"] == "PO Not Released"]

    return pr2po_processed, pr2po_convert_po, pr2po_pss_status, pr2po_total, pr2po_error, error_dataframe[["Status", 'PR#', 'PGr', 'OA#', 'Vendor#', 'PR value']]

def PO_Exception(filepath):
    df_po_exception = pd.read_excel(filepath, sheet_name = "Report")

    df_po_exception_processed = len(df_po_exception) - len(df_po_exception[df_po_exception["Status"].isna()])
    df_po_exception_convert = len(df_po_exception[df_po_exception["Status"] == "Convert to PO"])
    
    return df_po_exception_convert, df_po_exception_processed, len(df_po_exception)

def Reaward_PO(filepath):
    df_po_exception = pd.read_excel(filepath, sheet_name = "Report")

    df_po_exception.columns = df_po_exception.iloc[0,:]
    df_po_exception = df_po_exception.iloc[1:, :]
    df_UC57_convert = len(df_po_exception[df_po_exception["Status"] == "Completed"])
    df_UC57_manual = len(df_po_exception[df_po_exception["Status"] != ""])
    df_UC57_except = len(df_po_exception[df_po_exception["Status"] == "PR Exceptioned"])
    df_UC57_total = len(df_po_exception)
    df_uc57_error_table = df_po_exception[df_po_exception["Status"] == "PR Exceptioned"]

    return df_UC57_manual, df_UC57_convert, df_UC57_total, df_UC57_except, df_uc57_error_table
                
def Vendor(filepath):
    df2 = pd.read_excel(filepath, sheet_name="Report")

    vendor_processed = len(df2) - len(df2[df2["Status"].isna()])
    vendor_convert_po = len(df2[df2["Status"] == "Convert to PO"])
    vendor_total = len(df2)

    return vendor_processed, vendor_convert_po, vendor_total



col1, col2 = st.columns([0.5, 1.5])

with col1:
    st.markdown("### Drop Files")
    uploaded_file_list = st.file_uploader("Upload Reports", type=["xlsx", "xlsm"], accept_multiple_files = True)

with col2:
    if uploaded_file_list is not None:
        for uploaded_file in uploaded_file_list:
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                temp_file_path = tmp_file.name
            
            # Get the last modified time (mod time)
            mod_time = os.path.getmtime(temp_file_path)
            
            # Convert mod_time to a readable format
            import datetime
            mod_time_str = datetime.datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
            
            st.write(f"File last modified on: {mod_time_str}")
            
            # Optionally delete the temporary file if needed
            os.remove(temp_file_path)
            
            if re.findall("PR to PO", uploaded_file.name):
                st.write(uploaded_file.name)
                PR2PO_df, PR2PO_graph = st.columns([1, 1])
                pr2po_processed, pr2po_convert_po, pr2po_pss_status, pr2po_total, pr2po_error, error_dataframe = PR2PO(uploaded_file)
                
                with PR2PO_df:
                    # Create DataFrame for PR2PO
                    df_pr2po = pd.DataFrame({
                        "PR2PO Processed": [pr2po_processed],
                        "PR2PO Convert to PO": [pr2po_convert_po],
                        "PSS Status": [pr2po_pss_status],
                        "PR2PO Total": [pr2po_total],
                        "PR2PO Error": [pr2po_error]
                    })
                    st.dataframe(df_pr2po.set_index(df_pr2po.columns[0]))

                    if pr2po_error != 0:
                        st.write(f"{pr2po_error} Errors Found")
                        st.dataframe(error_dataframe, use_container_width= True)

                with PR2PO_graph:
                    df_T = df_pr2po.T.reset_index()
                    df_T.columns = ["Category", "Value"]

                    option = {
                        "tooltip": {
                            "trigger": 'axis',
                            "axisPointer": {      
                            "type": 'shadow'      
                            }
                        },
                        "xAxis": {
                            "type": 'category',
                            "data": list(df_T[df_T.columns[0]]),
                            "axisLabel": {
                            "rotate": 90 
                            }
                        },
                        "yAxis": {
                            "type": 'value'
                        },
                        "series": [
                            {
                            "data": list(df_T[df_T.columns[1]]),
                            "type": 'bar',
                            'itemStyle': {
                            'color': 'red'  # Set the color of the line to red
                        }
                            }
                        ]}
                    
                    st_echarts(option, height = "400px")

                st.write("____")

            elif re.findall("PO Exception Report", uploaded_file.name):
                st.write(uploaded_file.name)
                PO_df, PO_graph = st.columns([1, 1])
                with PO_df:
                    df_po_exception_convert, df_po_exception_processed, df_po_exception_total = PO_Exception(uploaded_file)
                    
                    # Create DataFrame for PO Exception
                    df_po_exception = pd.DataFrame({
                        "PO Exception Processed": [df_po_exception_processed],
                        "PO Exception Convert": [df_po_exception_convert],
                        "PO Exception Total": [df_po_exception_total]
                    })
                    st.dataframe(df_po_exception.set_index(df_po_exception.columns[0]), use_container_width=True)

                with PO_graph:
                    df_T = df_po_exception.T.reset_index()
                    df_T.columns = ["Category", "Value"]

                    option = {
                        "tooltip": {
                            "trigger": 'axis',
                            "axisPointer": {      
                            "type": 'shadow'      
                            }
                        },
                        "xAxis": {
                            "type": 'category',
                            "data": list(df_T[df_T.columns[0]]),
                            "axisLabel": {
                            "rotate": 90 
                            }
                        },
                        "yAxis": {
                            "type": 'value'
                        },
                        "series": [
                            {
                            "data": list(df_T[df_T.columns[1]]),
                            "type": 'bar',
                            'itemStyle': {
                            'color': 'blue'  # Set the color of the line to red
                        }
                            }
                        ]}
                    
                    st_echarts(option, height = "400px")

                st.write("____")

            elif re.findall("PO Reassignment", uploaded_file.name):
                st.write(uploaded_file.name)
                UC57_df, UC57_graph = st.columns([1, 1])

                with UC57_df:
                    df_UC57_manual, df_UC57_convert, df_UC57_total, df_UC57_except, df_UC57_error_table = Reaward_PO(uploaded_file)
                    
                    # Create DataFrame for Reaward PO
                    df_reaward_po = pd.DataFrame({
                        "Reaward PO Manual": [df_UC57_manual],
                        "Reaward PO Convert": [df_UC57_convert],
                        "Reaward PO Total": [df_UC57_total],
                        "Reaward PO Exception": [df_UC57_except]
                    })
                    st.dataframe(df_reaward_po.set_index(df_reaward_po.columns[0]), use_container_width=True)

                    if df_UC57_except != 0:
                        st.write(f"{df_UC57_except} Errors Found")
                        st.dataframe(df_UC57_error_table, use_container_width= True)

                with UC57_graph:
                    df_T = df_reaward_po.T.reset_index()
                    df_T.columns = ["Category", "Value"]

                    option = {
                        "tooltip": {
                            "trigger": 'axis',
                            "axisPointer": {      
                            "type": 'shadow'      
                            }
                        },
                        "xAxis": {
                            "type": 'category',
                            "data": list(df_T[df_T.columns[0]]),
                            "axisLabel": {
                            "rotate": 90 
                            }
                        },
                        "yAxis": {
                            "type": 'value'
                        },
                        "series": [
                            {
                            "data": list(df_T[df_T.columns[1]]),
                            "type": 'bar',
                            'itemStyle': {
                            'color': 'orange'  # Set the color of the line to red
                        }
                            }
                        ]}
                    
                    st_echarts(option, height = "400px")

                st.write("____")

            elif re.findall("Vendor", uploaded_file.name):
                st.write(uploaded_file.name)
                vendor_df, vendor_graph = st.columns([1, 1])

                with vendor_df:
                    vendor_processed, vendor_convert_po, vendor_total = Vendor(uploaded_file)
                    
                    # Create DataFrame for Vendor
                    df_vendor = pd.DataFrame({
                        "Vendor Processed": [vendor_processed],
                        "Vendor Convert to PO": [vendor_convert_po],
                        "Vendor Total": [vendor_total]
                    })
                    st.dataframe(df_vendor.set_index(df_vendor.columns[0]), use_container_width=True)

                with vendor_graph:
                    df_T = df_vendor.T.reset_index()
                    df_T.columns = ["Category", "Value"]

                    option = {
                        "tooltip": {
                            "trigger": 'axis',
                            "axisPointer": {      
                            "type": 'shadow'      
                            }
                        },
                        "xAxis": {
                            "type": 'category',
                            "data": list(df_T[df_T.columns[0]]),
                            "axisLabel": {
                            "rotate": 90 
                            }
                        },
                        "yAxis": {
                            "type": 'value'
                        },
                        "series": [
                            {
                            "data": list(df_T[df_T.columns[1]]),
                            "type": 'bar',
                            'itemStyle': {
                            'color': 'purple'  # Set the color of the line to red
                        }
                            }
                        ]}
                    
                    st_echarts(option, height = "400px")
                
                st.write("____")
