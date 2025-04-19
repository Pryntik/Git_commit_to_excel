import re
import openpyxl
import os
import pandas as pd
from io import StringIO
from datetime import datetime

# File Name
file_name = 'commits'

# Content exemple
commit_text = """
* commit id_commit_1
| Author: FirstName LastName <name@example.com>
| Date:   Mon Jan 01 01:01:01 2000 +0200
|
|     Title of the commit 1
|
* commit id_commit_2
| Author: FirstName LastName <name@example.com>
| Date:   Mon Feb 01 01:01:01 2000 +0200
|
|     Title of the commit 2
|
|     Description line 1 of the commit.
|     Description line 2 of the commit.
"""

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
commit_blocks = re.split(r'\* commit ', commit_text)[1:]

# Store the commits grouped by year-house
grouped_commits = {}

for block in commit_blocks:
    # Clean : strip and '|'
    block_lines = block.strip().split('\n')
    lines = [
        line.strip().lstrip('|').strip() for line in block_lines
        if line.strip() != '|'
    ]

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

# Creation of dataframe
df = pd.DataFrame(rows)

# Get the path of the script file
script_dir = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(script_dir, f'{file_name}.xlsx')

# Backup in an xlsx file
df.to_excel(output_path, index=False)

# Adjust the width of each column according to the maximum content
wb = openpyxl.load_workbook(output_path)
ws = wb.active
for column_cells in ws.columns:
    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
    ws.column_dimensions[column_cells[0].column_letter].width = length

# Back up the modified file
wb.save(output_path)
print(f"file xlsx generated successfully at the location : {output_path}")