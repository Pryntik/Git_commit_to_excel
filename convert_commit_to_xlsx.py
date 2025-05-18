import re
import openpyxl
import os
import pandas as pd
from io import StringIO
from datetime import datetime

# Assets
root_dir = os.path.dirname(os.path.abspath(__file__))
asset_dir = os.path.join(root_dir, 'assets')
os.makedirs(asset_dir, exist_ok=True)

# Files
file_name = 'commits'
file_input = f'{file_name}.txt'
file_output = f'{file_name}.xlsx'

# Paths
path_input = os.path.join(asset_dir, file_input)
path_output =  os.path.join(asset_dir, file_output)

# Console command in input file.txt :
# git log --all --decorate --graph
with open(path_input, 'r', encoding='utf-8') as f:
    commit_text = f.read()

def month_short_to_full(short_month):
    months = {
        'Jan': 'January',
        'Feb': 'February',
        'Mar': 'March',
        'Apr': 'April',
        'May': 'May',
        'Jun': 'June',
        'Jul': 'July',
        'Aug': 'August',
        'Sep': 'September',
        'Oct': 'October',
        'Nov': 'November',
        'Dec': 'December'
    }
    return months.get(short_month, short_month)

# Cutting of Committee blocks
commit_blocks = re.split(r"^\s*[\s|/\\]*\*\s*[\s|/\\]*commit\s+", commit_text, flags=re.MULTILINE)[1:]

# Store the commits grouped by year-house
grouped_commits = {}

for block in commit_blocks:
    # Clean : strip and '|'
    block_lines = block.strip().split('\n')
    processed_lines = []
    for current_line in block_lines:
        # Remove all '|', '\', '/' at start and strip the line
        cleaned_line = re.sub(r"^[|\s/\\]*", "", current_line).strip()
        if cleaned_line:
            processed_lines.append(cleaned_line)
    lines = processed_lines

    commit_id = lines[0] if lines else ''

    # Author
    author_line = next((line for line in lines if line.startswith('Author:')), '')
    author_match = re.search(r'Author: (.+?) <(.+?)>', author_line)
    if author_match:
        author_name = author_match.group(1)
        author_email = author_match.group(2)
    else:
        author_name, author_email = '', ''

    # Date
    date_line = next((line for line in lines if line.startswith('Date:')), '')
    date_match = re.search(r'(\w{3}) (\w{3}) (\d{1,2}) (\d{2}:\d{2}:\d{2}) (\d{4})(?: ([\+\-]\d{4}))?', date_line)
    if date_match:
        year = date_match.group(5)
        month = month_short_to_full(date_match.group(2))
    else:
        year, month = '', ''

    message_lines = [
        line for i, line in enumerate(lines)
        if i != 0 and not (line.startswith('Author:') or line.startswith('Date:')) and line
    ]
    title = message_lines[0] if message_lines else ''
    description = '\n'.join(message_lines[1:]) if len(message_lines) > 1 else ''

    key = f"{year}-{month}"
    grouped_commits.setdefault(key, []).append({
        'Year': year,
        'Month': month,
        'Title': title,
        'Description': description,
        'Commit ID': commit_id,
        'Author': author_name,
    })

# Generating rows for the xlsx table
rows = []
previous_year, previous_month = '', ''
for key in sorted(grouped_commits.keys()):
    for commit in grouped_commits[key]:
        if commit['Year'] != previous_year or year == '':
            rows.append({'Title': commit['Year'], 'Description': '', 'Commit ID': '', 'Author': ''})
            previous_year = commit['Year']
        if commit['Month'] != previous_month or month == '':
            rows.append({'Title': commit['Month'], 'Description': '', 'Commit ID': '', 'Author': ''})
            previous_month = commit['Month']
        rows.append({
            'Title': commit['Title'],
            'Commit ID': commit['Commit ID'],
            'Description': commit['Description'],
            'Author': commit['Author']
        })

# Backup in an xlsx file
df = pd.DataFrame(rows)
df.to_excel(path_output, index=False)

# Adjust the width of each column according to the maximum content
wb = openpyxl.load_workbook(path_output)
ws = wb.active
for column_cells in ws.columns:
    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
    ws.column_dimensions[column_cells[0].column_letter].width = length

# Back up the modified file
wb.save(path_output)
print(f"file xlsx generated successfully at the location : {path_output}")