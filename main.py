from collections import defaultdict
import re
import csv
import sys
from tika import parser
import datetime
import pprint

DATE_REGEX = r'\d{2}-\d{2}-\d{4}'
RATA_REGEX = r'Rata \d* din \d*'

raw: dict[str, str] = parser.from_file(sys.argv[1])
lines = raw['content'].split("\n")
fd = csv.writer(open(sys.argv[2], "w", newline=''))

i = 0
DATA_TRANZACTIEI, DETALII, RATA, MAGAZIN, NR_TRANZACTIE, SUMA = range(0,6)
fd.writerow(["Data", "Detalii", "Rata", "Magazin", "Nr tranzactie", "Suma"])
content = []
while i < len(lines):
    line = lines[i]
    if not re.match(r'\d{2}-\d{2}-\d{4}', line):
        i += 1
        continue
    
    count = 1
    i += 1
    data: list[str] = [line]

    while count < 4:
        if lines[i] == "":
            i += 1
            continue
        start = i
        while lines[i] != "":
            i += 1
        data.append(" ".join(lines[start:i]))
        count += 1

    data[DETALII] = data[DETALII].replace(";", " ")
    rata = re.findall(RATA_REGEX, data[1])
    data.insert(RATA, rata[0] if len(rata) >= 1 else "")

    try:
        payee = re.split(DATE_REGEX, data[1])[2].strip()
        payee = re.sub(r'\d', '', payee)
    except:
        payee = ""
    data.insert(MAGAZIN, payee)
    data[SUMA] = data[5].replace(".", "").replace(',', ".")
    if float(data[5]) > 0:
        continue
    data = [i.strip() for i in data]
    data[DATA_TRANZACTIEI] = datetime.datetime.strptime(data[DATA_TRANZACTIEI], "%d-%m-%Y")
    fd.writerow(data)
    content.append(data)


rate_buckets = defaultdict(int)
cheltuieli = 0
for row in content:
    rata = row[RATA]
    amount = row[SUMA]
    if not re.match(RATA_REGEX, rata):
        cheltuieli += float(amount)
        continue

    rata = rata.split(" ")
    rate_buckets[int(rata[3]) - int(rata[1])] += float(amount)

print("Suma ratelor finalizate in x luni:")
pprint.pprint(dict(rate_buckets))
print(f"Cheltuieli: {cheltuieli}")

