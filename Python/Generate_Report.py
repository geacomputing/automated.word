import matplotlib.pyplot as plt         # For creating plots and charts
import requests                          # To fetch data from web APIs
from docx import Document                 # To create and manipulate Word documents
from docx.enum.text import WD_ALIGN_PARAGRAPH  # For paragraph alignment options
from docx.shared import Inches            # To specify image sizes in Word
import subprocess                         # To run external commands (here, git)
import os                                 # For file system operations
from datetime import datetime             # To get current date/time
import platform                            # To get OS and system info
import getpass                             # To get current user name

# ---------------------------------------------------------
# User-configurable variables
# ---------------------------------------------------------
output_folder = "../reports"               # Folder where the final Word report will be saved
ndays = 10                                 # Number of past days for cryptocurrency price data
label_step = 10                            # Show every Nth label on x-axis of plots to avoid clutter
lastN_crypto = 10                          # Number of last entries to show in Bitcoin table
lastN_git = 100                            # Number of latest git commits to include in Git table
repo_path = "../"                          # Path to your local Git repository (for git log)

# ---------------------------------------------------------
# Create output folders if they do not exist
# ---------------------------------------------------------
os.makedirs(output_folder, exist_ok=True)  # Create reports folder
images_folder = "../images"                # Folder to store plots
os.makedirs(images_folder, exist_ok=True)  # Create images folder if it doesn't exist

# ---------------------------------------------------------
# File paths for Word document and images
# ---------------------------------------------------------
docx_filename = "dynamic_report.docx"
docx_path = os.path.join(output_folder, docx_filename)  # Full path for final Word report
plot_filename = f"bitcoin_last{ndays}.png"
plot_path = os.path.join(images_folder, plot_filename)  # Full path for the Bitcoin plot image

# ---------------------------------------------------------
# Fetch cryptocurrency data (Bitcoin) from CoinGecko API
# ---------------------------------------------------------
url = "https://api.coingecko.com/api/v3/coins/bitcoin/market_chart"
params = {"vs_currency": "usd", "days": f"{ndays}"}
response = requests.get(url, params=params)              # GET request to CoinGecko
data = response.json()                                   # Parse JSON response

# Extract timestamps and prices from API response
timestamps = [ts[0] for ts in data["prices"]]
prices = [ts[1] for ts in data["prices"]]

# Convert timestamps to human-readable datetime strings
dates = [datetime.fromtimestamp(ts/1000).strftime("%m-%d %H") for ts in timestamps]

# ---------------------------------------------------------
# Generate cryptocurrency price plot using Matplotlib
# ---------------------------------------------------------
plt.figure(figsize=(10,5))                           # Figure size in inches
plt.plot(dates, prices, marker='.', linestyle='-', color='orange', linewidth=1)
plt.xticks(dates[::label_step], rotation=45)         # Reduce x-axis label density for readability
plt.title(f"Bitcoin Price (Last {ndays} Days)", fontsize=16)
plt.xlabel("Date", fontsize=12)
plt.ylabel("Price (USD)", fontsize=12)
plt.grid(True, linestyle="--", alpha=0.5)           # Add light grid for easier reading
plt.tight_layout()                                   # Adjust layout to avoid clipping labels
plt.savefig(plot_path, dpi=200)                      # Save plot to the images folder
plt.close()                                          # Close figure to free memory

# ---------------------------------------------------------
# Fetch Git log from local repository
# ---------------------------------------------------------
git_format = "--pretty=format:%h|%an|%ad|%s"         # Format: hash|author|date|message
result = subprocess.run(
    ["git", "log", f"-n {lastN_git}", git_format],
    cwd=repo_path,                                    # Run in the specified repository
    stdout=subprocess.PIPE,                           # Capture standard output
    stderr=subprocess.PIPE,                           # Capture standard error
    text=True                                         # Decode output as text (not bytes)
)

# Raise exception if git command failed
if result.returncode != 0:
    raise Exception(f"Git error: {result.stderr}")

# Split git log into lines
git_log_lines = result.stdout.strip().split("\n")

# Parse each line into a structured dictionary
commits = []
for line in git_log_lines:
    parts = line.split("|", maxsplit=3)              # Split line by '|' into 4 fields
    if len(parts) == 4:
        commits.append({
            "hash": parts[0],
            "author": parts[1],
            "date": parts[2],
            "message": parts[3]
        })

# ---------------------------------------------------------
# Build Word document
# ---------------------------------------------------------
doc = Document()                                     # Create a new Word document

# Add main title
doc.add_heading("Dynamic Real-Time Report", level=0)
doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ---------------------------------------------------------
# Add system information
# ---------------------------------------------------------
doc.add_heading("System Information", level=1)
doc.add_paragraph(f"User: {getpass.getuser()}")
doc.add_paragraph(f"Operating System: {platform.system()} {platform.version()}")
doc.add_paragraph(f"Machine Name: {platform.node()}")
doc.add_paragraph(f"Architecture: {platform.machine()}")
doc.add_paragraph(f"Python Version: {platform.python_version()}")
doc.add_paragraph(f"Processor: {platform.processor()}")

# ---------------------------------------------------------
# Add a page break before Git log table
# ---------------------------------------------------------
doc.add_page_break()

# ---------------------------------------------------------
# Git log section
# ---------------------------------------------------------
doc.add_heading("Git Repository Log", level=1)
doc.add_paragraph(
    f"This section contains the document versioning log captured from GitHub. "
    f"All entries are current and complete as of the report generation date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}."
)
doc.add_heading(f"Last {lastN_git} commits from repository '{repo_path}'", level=2)

# Create table with 3 columns for Author, Date, Message
git_table = doc.add_table(rows=1, cols=3)
git_table.style = "Table Grid"

# Add table header
git_table.rows[0].cells[0].text = "Author"
git_table.rows[0].cells[1].text = "Date"
git_table.rows[0].cells[2].text = "Message"

# Populate table rows with commit data
for c in commits:
    row_cells = git_table.add_row().cells
    row_cells[0].text = c["author"]
    row_cells[1].text = c["date"]
    row_cells[2].text = c["message"]

# Add centered table caption
git_caption = doc.add_paragraph(f"Table 1: Last {lastN_git} commits in repository '{repo_path}'")
git_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
git_caption.style = "Caption"

# ---------------------------------------------------------
# Page break before cryptocurrency section
# ---------------------------------------------------------
doc.add_page_break()

# ---------------------------------------------------------
# Cryptocurrency section with plot and table
# ---------------------------------------------------------
doc.add_heading("Section 1: Cryptocurrency Data", level=1)
doc.add_paragraph(f"Bitcoin price trends over the last {ndays} days.")

# Insert plot centered
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture(plot_path, width=Inches(6))

# Add caption for plot
caption = doc.add_paragraph(f"Figure 1: Real-time Bitcoin prices (USD) over the last {ndays} days.")
caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
caption.style = "Caption"

# Add table for latest Bitcoin prices
doc.add_heading(f"Latest Prices (last {lastN_crypto})", level=2)
table = doc.add_table(rows=1, cols=2)
table.style = "Table Grid"

# Table header
table.rows[0].cells[0].text = "Date"
table.rows[0].cells[1].text = "Price (USD)"

# Add last N crypto entries
for date, price in zip(dates[-lastN_crypto:], prices[-lastN_crypto:]):
    row_cells = table.add_row().cells
    row_cells[0].text = date
    row_cells[1].text = f"{price:,.2f}"

# Add table caption
table_caption = doc.add_paragraph(f"Table 2: Last {lastN_crypto} Bitcoin prices.")
table_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
table_caption.style = "Caption"

# ---------------------------------------------------------
# Save the Word document
# ---------------------------------------------------------
doc.save(docx_path)
print(f"Dynamic report created successfully: {docx_path}")
