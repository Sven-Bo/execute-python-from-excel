from pathlib import Path  # Python Standard Library

import pandas as pd  # pip install pandas openpyxl
from pyecharts import options as opts
from pyecharts.charts import Bar


this_dir = Path(__file__).parent
excel_path = this_dir / 'RunPython_Example.xlsx'

df = pd.read_excel(
    io=excel_path,
    engine='openpyxl',
    sheet_name='Financial Report',
    skiprows=3,
    usecols='B:Q',
    nrows=525,
)

df = df.groupby(by="Month Number").sum()[['Sales', 'Profit']]

bar_chart = (
    Bar()
    .add_xaxis(df.index.to_list())
    .add_yaxis("Sales", df["Sales"].round(0).tolist())
    .add_yaxis("Profit", df["Profit"].round(0).tolist())
    .set_global_opts(
        title_opts=opts.TitleOpts(title="Sales & Profit by month", subtitle="in USD"),
        datazoom_opts=opts.DataZoomOpts(),
    )
    .render("Financial_Data.html")    
)
