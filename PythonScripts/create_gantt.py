from pathlib import Path

import pandas as pd
import plotly.express as px

print("Creating Gantt chart...")
excel_file = Path(__file__).parents[1] / "RunPython_Example.xlsx"
df = pd.read_excel(excel_file, sheet_name="Gantt")

# Assign Columns to variables
tasks = df["Task"]
start = df["Start"]
finish = df["Finish"]
complete = df["Complete in %"]

# Create Gantt Chart
fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, color=complete, title="Task Overview")

# Upade/Change Layout
fig.update_yaxes(autorange="reversed")
fig.update_layout(title_font_size=42, font_size=18, title_font_family="Arial")

# Save Gantt and Export to HTML
output_path = str(Path(__file__).parents[1] / "Task_Overview_Gantt.html")
fig.write_html(output_path)
print(f"Gantt Chart has been saved here: {output_path}")
