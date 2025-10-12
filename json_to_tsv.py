# json_to_tsv.py
import json
import csv
import sys
import os

if len(sys.argv) != 2:
    print("Использование: python json_to_tsv.py файл.json")
    sys.exit(1)

json_file = sys.argv[1]
tsv_file = os.path.splitext(json_file)[0] + '.tsv'

fieldnames = [
    "event",
    "status",
    "psychophysiological_state",
    "circumstances",
    "pattern_of_behavior",
    "_source_file",
    "_row_number"
]

with open(json_file, 'r', encoding='utf-8') as f:
    data = json.load(f)

if not data:
    print(f"⚠️ Пропущен (пустой): {json_file}")
    exit()

with open(tsv_file, 'w', encoding='utf-8', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter='\t')
    writer.writeheader()
    writer.writerows(data)

print(f"✅ {tsv_file}")
