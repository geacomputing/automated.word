import matplotlib.pyplot as plt
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
import subprocess
import os
from datetime import datetime
import platform
import getpass

# -------------------------
# User-configurable variables
# -------------------------
output_folder = "../reports"
ndays = 10       # Number of days for crypto data
label_step = 10   # Reduce x-axis labels
lastN_crypto = 10   # Last N rows for crypto table
lastN_git = 100      # Last N git commits
repo_path = "../"    # Path to your local git repo

# -------------------------
# Create output folder
# -------------------------
os.makedirs(output_folder, exist_ok=True)

# Paths
docx_filename = "dynamic_report.docx"
docx_path = os.path.join(output_folder, docx_filename)

images_folder = "../images"
os.makedirs(images_folder, exist_ok=True)


plot_filename = f"bitcoin_last{ndays}.png"
plot_path = os.path.join(images_folder, plot_filename)

# -------------------------
# Fetch crypto data
# -------------------------
url = "https://api.coingecko.com/api/v3/coins/bitcoin/market_chart"
params = {"vs_currency": "usd", "days": f"{ndays}"}
response = requests.get(url, params=params)
data = response.json()

timestamps = [ts[0] for ts in data["prices"]]
prices = [ts[1] for ts in data["prices"]]
dates = [datetime.fromtimestamp(ts/1000).strftime("%m-%d %H") for ts in timestamps]

# -------------------------
# Generate crypto plot
# -------------------------
plt.figure(figsize=(10,5))
plt.plot(dates, prices, marker='.', linestyle='-', color='orange', linewidth=1)
plt.xticks(dates[::label_step], rotation=45)
plt.title(f"Bitcoin Price (Last {ndays} Days)", fontsize=16)
plt.xlabel("Date", fontsize=12)
plt.ylabel("Price (USD)", fontsize=12)
plt.grid(True, linestyle="--", alpha=0.5)
plt.tight_layout()
plt.savefig(plot_path, dpi=200)
plt.close()

# -------------------------
# Fetch Git log
# -------------------------
git_format = "--pretty=format:%h|%an|%ad|%s"
result = subprocess.run(
    ["git", "log", f"-n {lastN_git}", git_format],
    cwd=repo_path,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True
)

if result.returncode != 0:
    raise Exception(f"Git error: {result.stderr}")

git_log_lines = result.stdout.strip().split("\n")
commits = []
for line in git_log_lines:
    parts = line.split("|", maxsplit=3)
    if len(parts) == 4:
        commits.append({
            "hash": parts[0],
            "author": parts[1],
            "date": parts[2],
            "message": parts[3]
        })

# -------------------------
# Build Word document
# -------------------------
# -------------------------
# Build Word document
# -------------------------
doc = Document()
doc.add_heading("Dynamic Real-Time Report", level=0)
doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# System info
doc.add_heading("System Information", level=1)
doc.add_paragraph(f"User: {getpass.getuser()}")
doc.add_paragraph(f"Operating System: {platform.system()} {platform.version()}")
doc.add_paragraph(f"Machine Name: {platform.node()}")
doc.add_paragraph(f"Architecture: {platform.machine()}")
doc.add_paragraph(f"Python Version: {platform.python_version()}")
doc.add_paragraph(f"Processor: {platform.processor()}")

doc.add_page_break()
# -------------------------
# Git log table on first page
# -------------------------
doc.add_heading("Git Repository Log", level=1)
doc.add_paragraph(
    f"This section contains the document versioning log captured from GitHub. "
    f"All entries are current and complete as of the report generation date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}."
)

doc.add_heading(f"Last {lastN_git} commits from repository '{repo_path}'", level=2)

git_table = doc.add_table(rows=1, cols=3)
git_table.style = "Table Grid"

git_table.rows[0].cells[0].text = "Author"
git_table.rows[0].cells[1].text = "Date"
git_table.rows[0].cells[2].text = "Message"

for c in commits:
    row_cells = git_table.add_row().cells
    row_cells[0].text = c["author"]
    row_cells[1].text = c["date"]
    row_cells[2].text = c["message"]

git_caption = doc.add_paragraph(f"Table 1: Last {lastN_git} commits in repository '{repo_path}'")
git_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
git_caption.style = "Caption"

# -------------------------
# Page break before Bitcoin section
# -------------------------
doc.add_page_break()

# Section 1: Crypto plot
doc.add_heading("Section 1: Cryptocurrency Data", level=1)
doc.add_paragraph(f"Bitcoin price trends over the last {ndays} days.")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture(plot_path, width=Inches(6))

caption = doc.add_paragraph(f"Figure 1: Real-time Bitcoin prices (USD) over the last {ndays} days.")
caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
caption.style = "Caption"

# Crypto table
doc.add_heading(f"Latest Prices (last {lastN_crypto})", level=2)
table = doc.add_table(rows=1, cols=2)
table.style = "Table Grid"
table.rows[0].cells[0].text = "Date"
table.rows[0].cells[1].text = "Price (USD)"

for date, price in zip(dates[-lastN_crypto:], prices[-lastN_crypto:]):
    row_cells = table.add_row().cells
    row_cells[0].text = date
    row_cells[1].text = f"{price:,.2f}"

table_caption = doc.add_paragraph(f"Table 2: Last {lastN_crypto} Bitcoin prices.")
table_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
table_caption.style = "Caption"

# -------------------------
# Save document and cleanup
# -------------------------
doc.save(docx_path)

print(f"Dynamic report created successfully: {docx_path}")

