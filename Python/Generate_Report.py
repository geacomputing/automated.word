import matplotlib.pyplot as plt
import numpy as np
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
from datetime import datetime




import platform
import getpass




# -------------------------
# User-configurable variables
# -------------------------
output_folder = "../reports"
ndays = 10  # Number of days to fetch
label_step = 20  # Show x-axis label every N points to reduce clutter
lastN = 10  # Last N rows for the table

# -------------------------
# Create output folder
# -------------------------
os.makedirs(output_folder, exist_ok=True)

# Word document path
docx_filename = "dynamic_report.docx"
docx_path = os.path.join(output_folder, docx_filename)

# Temporary plot path
plot_path = os.path.join(output_folder, "temp_plot.png")

# -------------------------
# Fetch real-time data from API
# -------------------------
url = "https://api.coingecko.com/api/v3/coins/bitcoin/market_chart"
params = {"vs_currency": "usd", "days": f"{ndays}"}
response = requests.get(url, params=params)
data = response.json()

timestamps = [ts[0] for ts in data["prices"]]
prices = [ts[1] for ts in data["prices"]]
dates = [datetime.fromtimestamp(ts/1000).strftime("%m-%d %H") for ts in timestamps]

# -------------------------
# Generate plot
# -------------------------
plt.figure(figsize=(10, 5))
plt.plot(dates, prices, marker='o', linestyle='-', color='orange', linewidth=2)
plt.xticks(dates[::label_step], rotation=45)
plt.title(f"Bitcoin Price (Last {ndays} Days)", fontsize=16)
plt.xlabel("Date", fontsize=12)
plt.ylabel("Price (USD)", fontsize=12)
plt.grid(True, linestyle="--", alpha=0.5)

plt.tight_layout()
plt.savefig(plot_path, dpi=200)
plt.close()

# -------------------------
# Build Word Document
# -------------------------
doc = Document()
doc.add_heading("Dynamic Real-Time Report", level=0)
doc.add_paragraph(f"Report generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")



# Current user
username = getpass.getuser()
doc.add_paragraph(f"User: {username}")

# Operating system
os_name = platform.system()
os_version = platform.version()
doc.add_paragraph(f"Operating System: {os_name} {os_version}")

# Machine name / hostname
machine_name = platform.node()
doc.add_paragraph(f"Machine Name: {machine_name}")

# Architecture
architecture = platform.machine()
doc.add_paragraph(f"Architecture: {architecture}")

# Python version
python_version = platform.python_version()
doc.add_paragraph(f"Python Version: {python_version}")

# Optional: CPU info
cpu_info = platform.processor()
doc.add_paragraph(f"Processor: {cpu_info}")



# -------------------------
# Add page break before table
# -------------------------
doc.add_page_break()

# Section 1: Plot
doc.add_heading("Section 1: Cryptocurrency Data", level=1)
doc.add_paragraph(f"This section presents Bitcoin price trends over the last {ndays} days fetched in real-time from CoinGecko API.")

# Insert plot centered
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture(plot_path, width=Inches(6))

# Plot caption
caption = doc.add_paragraph(f"Figure 1: Real-time Bitcoin prices (USD) over the last {ndays} days.")
caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
caption.style = "Caption"



# Section 2: Table
doc.add_heading(f"Section 2: Latest Prices (last {lastN})", level=1)
table = doc.add_table(rows=1, cols=2)
table.style = "Table Grid"
table.rows[0].cells[0].text = "Date"
table.rows[0].cells[1].text = "Price (USD)"

for date, price in zip(dates[-lastN:], prices[-lastN:]):
    row_cells = table.add_row().cells
    row_cells[0].text = date
    row_cells[1].text = f"{price:,.2f}"

# Table caption
table_caption = doc.add_paragraph(f"Table 1: Last {lastN} Bitcoin prices (USD).")
table_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
table_caption.style = "Caption"

# Save Word document
doc.save(docx_path)

# Remove temporary plot
os.remove(plot_path)

print(f"Dynamic report created successfully: {docx_path}")
