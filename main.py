from collections import defaultdict
import re
import sys
from tika import parser
import datetime
import xlsxwriter

DATE_REGEX = r"\d{2}-\d{2}-\d{4}"
RATA_REGEX = r"Rata \d* din \d*"

raw: dict[str, str] = parser.from_file(sys.argv[1])
lines = raw["content"].split("\n")

workbook = xlsxwriter.Workbook(sys.argv[2])
worksheet = workbook.add_worksheet()

header = [
    "Data",
    "Detalii",
    "Rata",
    "Magazin",
    "Nr tranzactie",
    "Total tranzactie",
    "Suma de returnat",
]
for col_num, data in enumerate(header):
    worksheet.write(0, col_num, data)

i = 0
DATA_TRANZACTIEI, DETALII, RATA, MAGAZIN, NR_TRANZACTIE, TOTAL_TRANZACTIE, SUMA = range(
    7
)
content = []
row_num = 1  # Start from the second row (first row is header)
while i < len(lines):
    line = lines[i]
    if not re.match(DATE_REGEX, line):
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
        payee = re.sub(r"\d", "", payee)
    except:
        payee = ""

    data.insert(MAGAZIN, payee)
    data.insert(SUMA, data[5].replace(".", "").replace(",", "."))
    if float(data[SUMA]) > 0:
        continue

    if "comerciant" in data[DETALII]:
        data[TOTAL_TRANZACTIE] = (
            data[DETALII].split("comerciant ")[1].split(" RON")[0].replace(" ", "")
        )
    else:
        data[TOTAL_TRANZACTIE] = "4"

    data = [i.strip() for i in data]
    data[SUMA] = float(data[SUMA])
    data[TOTAL_TRANZACTIE] = -float(data[TOTAL_TRANZACTIE])

    data[DATA_TRANZACTIEI] = datetime.datetime.strptime(
        data[DATA_TRANZACTIEI], "%d-%m-%Y"
    ).strftime("%Y-%m-%d")

    for col_num, cell_data in enumerate(data):
        worksheet.write(row_num, col_num, cell_data)
    row_num += 1
    content.append(data)


rate_buckets = defaultdict(int)
cheltuieli = 0
rate_noi = 0
for row in content:
    rata = row[RATA]
    amount = row[SUMA]
    if not re.match(RATA_REGEX, rata):
        cheltuieli += float(amount)
        continue

    rata = rata.split(" ")
    numarul_ratei = int(rata[1])
    numarul_de_rate = int(rata[3])

    rate_buckets[numarul_de_rate - numarul_ratei] += float(amount)

    if numarul_ratei == 1:
        rate_noi += row[TOTAL_TRANZACTIE]


worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Suma ratelor ce vor disparea")
header = ["Peste X luni", "Suma"]
for col_num, data in enumerate(header):
    worksheet.write(1, col_num, data)

row_num = 2
for luna, suma in sorted(rate_buckets.items(), key=lambda x: x[0]):
    worksheet.write(row_num, 0, int(luna))
    worksheet.write(row_num, 1, float(suma))
    row_num += 1


print(f"Cheltuieli in rate {rate_noi}")
print(f"Cheltuit total: {rate_noi + cheltuieli}")

workbook.close()