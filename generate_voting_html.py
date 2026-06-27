import os
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from jinja2 import Template

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

target_files = [
    v for v in os.listdir(BASE_DIR) if v.split('.')[-1] == 'xlsx' and '結果' not in v
    ]
if len(target_files) != 1:
    raise Exception('There are multiple xlsx files!')
target_file = target_files[0]
target_year = target_file.split('_')[0]

df = pd.read_excel(target_file, index_col=None)
header_values = df.columns
list_from_df = df.values.tolist()

data = {
    "storage_data_id": f"{target_year}ConferenceVoteResult", # このIDをキーにブラウザーにデータが保持される。他のものと混じらないようなユニークな名前を毎回定義すること。例えば開催年。
    "header_values": header_values,
    "presentations": list_from_df,
    "due": "〆切は9/20（金）14:30です。",
}

with open(BASE_DIR / "template.html.jinja", "r", encoding="utf-8") as f:
    template_content = f.read()
template = Template(template_content)
result = template.render(data)
output_filename = f"{data['storage_data_id']}.html"
with open(BASE_DIR / output_filename, "w", encoding="utf-8") as f:
    f.write(result)
print(f"Finish to generate {output_filename}.")
