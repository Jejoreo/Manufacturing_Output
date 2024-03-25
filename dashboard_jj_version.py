import pandas as pd
import streamlit as st
import plotly.express as px
import statsmodels.api as sm
import plotly.graph_objects as go

st.set_page_config(
    page_title= "Manufacturing Output Viewer",
    page_icon="📈",
    layout="wide")

st.title("PowerJJ Hahaha 😆")
st.header("SAC Chart Summarizer 2024 🤯", divider="rainbow")
st.markdown('''
So this was made ASAP by considering the bare minium request specs. Drag, drop, or click the file uploader widget on
the left sidebar and make sure you have the sheets with these names:
1️⃣ :rainbow[flashings],
2️⃣ :rainbow[Deliveries],
3️⃣ :rainbow[CNC],
4️⃣ :rainbow[Jobbing], &
5️⃣ :rainbow[Quadro].
Note that the order doesn't matter as you still need to choose the sheets to get the chart. Then, just select the year and months
of the desired data. You could also put the chart title name since I have a hard time listing all the names (plus making it more
flexible for the user). Enjoy!!! 🥹                  
''')
st.caption("Please note that this is a prototype!")

@st.cache_data
def load_data(path: str, sheet_name: str):
    data = pd.read_excel(path, sheet_name=sheet_name)
    return data


with st.sidebar:
    st.header("Configuration & Options")
    file_path = st.file_uploader("Choose a file", type=["xlsx"])

if file_path is None:
    st.info("Please upload an .xlsx file ", icon="😔")
    st.stop()

try:
    sheet_names = pd.ExcelFile(file_path).sheet_names
except Exception as e:
    st.error(f"An error occurred: {e}")
    st.stop()


selected_sheet = st.selectbox("Select Sheet", sheet_names)

df = load_data(file_path, selected_sheet)

column_config_list = {
    "Year": st.column_config.NumberColumn(format="%d"),
    "Date": st.column_config.DateColumn(
        format="D MMMM YYYY"
    ),
    "WE": st.column_config.DateColumn(
        format="D MMMM YYYY"
    )    
}

def filter_data(df):
    # Change the date format
    df['Date'] = pd.to_datetime(df['Date'])

    #* Several Filterings
    '''
    Note: Several of the used filterings items found throughout the graphs from the given task includes:
    1. Years extracted from data set
    2. Months extracted from data set
    3. Material Name as in Options
    '''
    years = df["Date"].dt.year.unique()
    selected_years = st.multiselect("Select Year", years)

    months = df["Date"].dt.month_name().unique()
    selected_months = st.multiselect("Select Month(s):", months)

    # Check if 'MaterialName' column exists in the DataFrame
    if 'MaterialName' in df.columns:
        materials = df['MaterialName'].unique()
        selected_materials = st.multiselect("Select Material (Optional)", materials)
    else:
        selected_materials = []


    # Check if 'Range' column exists in the df
    if 'Range' in df.columns:
        mat_type = df["Range"].unique()
        selected_ranges = st.multiselect("Select Material(s) (Optional):", mat_type)
    else:
        selected_ranges = []


    try:
        # Filter the DataFrame based on selected year, month, and material
        if selected_materials:
            filtered_df = df[
                (df["Date"].dt.year.isin(selected_years)) &
                (df["Date"].dt.month_name().isin(selected_months)) &
                (df['MaterialName'].isin(selected_materials))
            ]
        elif selected_ranges:
            filtered_df = df[
                (df["Date"].dt.year.isin(selected_years)) &
                (df["Date"].dt.month_name().isin(selected_months)) &
                (df['Range'].isin(selected_ranges))
            ]       
        else:
            filtered_df = df[
                (df["Date"].dt.year.isin(selected_years)) &
                (df["Date"].dt.month_name().isin(selected_months))
            ]
    except Exception as e:
        st.error(f"Selection Incomplete")
        st.stop()

    return filtered_df

def generate_plot(df, plot_type, data_mean, y_code:str, plot_title:str, data_target:int, x_code="Date"):
    if plot_type == "line":
        # Create scatter plot with trendline
        fig_scatter = px.scatter(df, x=x_code, y=y_code, title=plot_title, trendline='ols', 
                                trendline_color_override="black", color=y_code, 
                                color_continuous_scale="Viridis", color_continuous_midpoint=8, render_mode="svg", 
                                hover_name=x_code)

        # Create line plot with horizontal lines
        fig_line = px.line(df, x=x_code, y=y_code, title=plot_title, markers=False, color_discrete_sequence=["royalblue"], width= 10)        
        fig_line.add_hline(y=data_mean, line_dash="dash", line_color="red", annotation_text=f'Average: {data_mean}', annotation_position="top right", annotation_font_color="black")
        fig_line.add_hline(y=data_target, line_dash="dash", line_color="green", annotation_text=f"Target: {data_target}", annotation_position="top right", annotation_font_color="black")
        fig_line.update_layout(xaxis=dict(tickmode='auto', dtick='M1', tickformat='%Y-%m-%d'))

        # Extract the traces from fig_line and convert them to tuples
        line_traces = tuple(fig_line.data)

        # Create a trendline trace from fig_scatter
        trendline_trace = fig_scatter.data[1]

        # Concatenate the traces and create a new figure
        fig = go.Figure(data=line_traces + (trendline_trace,), layout=fig_line.layout)
    elif plot_type == "bar":
        fig = px.bar(df, x=x_code, y=y_code, title=plot_title, color_discrete_sequence=["lightpink"])
        fig.add_hline(y=data_mean, line_dash="dash", line_color="red", annotation_text=f'Average: {data_mean}', annotation_position="top right", annotation_font_color="black")
        fig.add_hline(y=data_target, line_dash="dash", line_color="green", annotation_text=f"Target: {data_target}", annotation_position="top right", annotation_font_color="black")
    elif plot_type == "box":
        fig = px.box(df, x=x_code, y=y_code, title=plot_title)
    elif plot_type == "histogram":
        fig = px.histogram(df, x=y_code, title=plot_title)
    elif plot_type == "scatter":
        fig = px.scatter(df, x=x_code, y=y_code, title=plot_title, trendline='ols', trendline_color_override="red")
    else:
        fig = None
        st.warning("No data to display.")
    return fig



if selected_sheet == "flashing":

    filtered_df = filter_data(df)
    
    # Create the pivot table
    pivot_table = filtered_df.groupby(['Date',])['Total Folds'].sum().reset_index()

    # Reordering stuff
    pivot_table = pivot_table[['Date', 'Total Folds']]

    # Assign the pivot table back to df
    df = pivot_table.sort_values(by=['Date'])

    data_mean = 0
    try:
        data_mean = round(df['Total Folds'].mean())
    except Exception as e:
        # Handle the exception gracefully
        st.error(f"An error occurred while calculating the mean: No Data Selected")

    # Plot type selection
    plot_type = st.selectbox("Select Plot Type", ["line", "bar", "box", "histogram", "scatter"])

    chart_title = st.text_input("Enter Chart Title:")
    
    flashings_target = 650

    # plot based on conditions
    plot_figure = generate_plot(df, plot_type, data_mean, y_code="Total Folds", plot_title=chart_title, data_target=flashings_target)

    st.markdown(f"Average Fold Per Day: {data_mean}")

    col1, col2 = st.columns(2)

    with col1:
        st.dataframe(df, column_config=column_config_list, use_container_width=True)
    with col2:
        st.plotly_chart(plot_figure, use_container_width=True)

    st.caption('Note: Trendline was assumed to be OLS')



elif selected_sheet == "Deliveries":

    filtered_df = filter_data(df)

    #* Pivot Table 1: Daily Delivery Average Price
    pivot_table_1  = filtered_df.groupby(['Date'])['Total Cost'].sum().reset_index()

    # Assign the pivot table back to df
    df_1 = pivot_table_1.sort_values(by=['Date'])

    df_1['Date'] = pd.to_datetime(df_1['Date'], format='%d %B %Y')

    data_mean_1 = 0
    try:
        data_mean_1 =round(df_1['Total Cost'].mean())
    except Exception as e:
        # Handle the exception gracefully
        st.error(f"An error occurred while calculating the mean: No Data Selected")

    
    # Plot type selection
    plot_type_1 = st.selectbox("Select Plot Type", ["line", "bar", "box", "histogram", "scatter"], key="type1")

    # Chart Component 1:
    chart_title_1 = st.text_input("Enter Chart 1 Title:")

    daily_charge_target = 5000

    plot_figure_1 = generate_plot(df_1, plot_type_1, data_mean_1, y_code="Total Cost", plot_title=chart_title_1, data_target=5000)
    

    st.header("Daily Delivery Charges", divider="rainbow")
    st.write(f"Average Daily Delivery Charges: $ {data_mean_1}")
    col1, col2 = st.columns(2)

    with col1:
        st.dataframe(df_1, column_config=column_config_list, use_container_width=True)
    with col2:
        st.plotly_chart(plot_figure_1, use_container_width=True)


    #Pivot Table 2
    pivot_table_2 = filtered_df.groupby(['Date'])['Total Cost'].count().reset_index()

    # Assign the pivot table back to df
    df_2 = pivot_table_2.sort_values(by=['Date'])

    data_mean_2 = 0
    try:
        data_mean_2 =round(df_2['Total Cost'].mean())
    except Exception as e:
        st.error(f"An error occurred while calculating the mean: No Data Selected")

    # Plot type selection
    plot_type_2 = st.selectbox("Select Plot Type", ["line", "bar", "box", "histogram", "scatter"], key="type2")


    # Chart Component 2:
    chart_title_2 = st.text_input("Enter Chart 2 Title:")

    daily_delivery_target = 24

    plot_figure_2 = generate_plot(df_2, plot_type_2, data_mean_2, y_code="Total Cost", plot_title=chart_title_2, data_target=daily_delivery_target)
    plot_figure_2.update_yaxes(title_text="Deliveries Quantity")

    st.header("Number of Delivery Statistics", divider = "rainbow")
    st.write(f"Average Daily Deliveries: {data_mean_2} Trip")
    col3, col4 = st.columns(2)

    with col3:
        st.dataframe(df_2, column_config=column_config_list, use_container_width=True)
    with col4:
        st.plotly_chart(plot_figure_2, use_container_width=True)

    st.caption('Note: Both Trendline was assumed to be OLS')


elif selected_sheet == "CNC":
    filtered_df = filter_data(df)

    if not filtered_df.empty:  # Check if DataFrame is not empty before performing operations
        pivot_table = filtered_df.groupby(["Date"])["Number of Finished Panels"].sum().reset_index()

        # Assign the pivot table back to df
        df = pivot_table.sort_values(by=['Date'])

        df["Date"] = df["Date"].dt.strftime('%d %B %Y')
        df["Date"] = pd.to_datetime(df["Date"], format='%d %B %Y')

        data_mean = 0
        try:
            data_mean = round(df["Number of Finished Panels"].mean(), 2)  # Calculate mean using mean() function
        except Exception as e:
            st.error(f"An error occurred while calculating the mean: {e}")

        # Plot type selection
        plot_type = st.selectbox("Select Plot Type", ["line", "bar", "box", "histogram", "scatter"])

        # Chart Component 3:
        chart_title = st.text_input("Enter Chart Title:")

        finish_panel_target = 400

        plot_figure = generate_plot(df, plot_type, data_mean, y_code="Number of Finished Panels", plot_title=chart_title, data_target=finish_panel_target)
        

        st.markdown(f"Average Finished Panel (weekly): {data_mean}")    
        col1, col2 = st.columns(2)
        with col1:
            st.dataframe(df, column_config=column_config_list, use_container_width=True)
        with col2:
            st.plotly_chart(plot_figure, use_container_width=True)
        
        st.caption('Note: Trendline was assumed to be OLS')    


elif selected_sheet == "Jobbing":
    filtered_df = filter_data(df)


    pivot_table = filtered_df.groupby(["WE"])["Sale Price"].sum().reset_index()

    # Assign the pivot table back to df
    df = pivot_table.sort_values(by=['WE'])

    df["WE"] = df["WE"].dt.strftime('%d %B %Y')
    df["WE"] = pd.to_datetime(df["WE"], format='%d %B %Y')

    data_mean = 0
    try:
        data_mean = round(df["Sale Price"].mean(), 2)  # Calculate mean using mean() function
    except Exception as e:
        st.error(f"An error occurred while calculating the mean: {e}")

    # Plot type selection
    plot_type = st.selectbox("Select Plot Type", ["line", "bar", "box", "histogram", "scatter"])

    # Chart Component 4:
    chart_title = st.text_input("Enter Chart Title:")

    finish_panel_target = 2000

    plot_figure = generate_plot(df, plot_type, data_mean=data_mean, y_code="Sale Price", plot_title=chart_title, data_target=finish_panel_target, x_code="WE")
        
            
    st.markdown(f"Average Weekly Sales: $ {data_mean}")    
    col1, col2 = st.columns(2)
    with col1:
        st.dataframe(df, column_config=column_config_list, use_container_width=True)
    with col2:
        st.plotly_chart(plot_figure, use_container_width=True)
        
    st.caption('Note: Trendline was assumed to be OLS')   


elif selected_sheet == "Quadro":
    filtered_df = filter_data(df)

    pivot_table = filtered_df.groupby(["WE"])["Sale Price"].sum().reset_index()

    # Assign the pivot table back to df
    df = pivot_table.sort_values(by=['WE'])

    df["WE"] = df["WE"].dt.strftime('%d %B %Y')
    df["WE"] = pd.to_datetime(df["WE"], format='%d %B %Y')

    data_mean = 0
    try:
        data_mean = round(df["Sale Price"].mean(), 2)  # Calculate mean using mean() function
    except Exception as e:
        st.error(f"An error occurred while calculating the mean: {e}")

    # Plot type selection
    plot_type = st.selectbox("Select Plot Type", ["line", "bar", "box", "histogram", "scatter"])

    # Chart Component 5:
    chart_title = st.text_input("Enter Chart Title:")

    sum_sale_target = 20000

    plot_figure = generate_plot(df, plot_type, data_mean=data_mean, y_code="Sale Price", plot_title=chart_title, data_target=sum_sale_target, x_code="WE")
               
    st.markdown(f"Weekly Average Sales: $ {data_mean}")    
    col1, col2 = st.columns(2)
    with col1:
        st.dataframe(df, column_config=column_config_list, use_container_width=True)
    with col2:
        st.plotly_chart(plot_figure, use_container_width=True)
        
    st.caption('Note: Trendline was assumed to be OLS')       

else:
    st.dataframe(df)
