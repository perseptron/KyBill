import sys
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
from pathlib import Path


def main():
    if len(sys.argv) <= 1:
        print("enter full path to xml file, add '-d' for detailed expense")
        sys.exit(1)

    fields = {'Ос. номер': 'Invoice/Customer/BillingAccount',
              'Номер тел.': 'Invoice/Customer/CustomerPhone',
              'Тарифний план': 'Invoice/Contract/ContractDetail/ContractType',
              'Початкова дата': 'Invoice/Header/BillingPeriod/BeginDate',
              'Кінцева дата': 'Invoice/Header/BillingPeriod/EndDate',
              'Баланс поч.': 'Invoice/InvoiceAmount/AmountDetail/BalBeginMonth',
              'знак балансу': 'Invoice/InvoiceAmount/AmountDetail/BalBeginMonthText',
              'Баланс кін.': 'Invoice/InvoiceAmount/AmountDetail/BalEndMonth',
              'знак балансу.': 'Invoice/InvoiceAmount/AmountDetail/BalBeginEndText',
              'Поповн.': 'Invoice/InvoiceAmount/AmountDetail/Payments/PaymentsBank',
              'Рекоменд': 'Invoice/InvoiceAmount/AmountDetail/RecommendedPayment',
              'Деталізація': 'Invoice/Summary/SummaryRow/RowDetail//Text',
              'Витрати': 'Invoice/Summary/SummaryRow/RowDetail//AmountExclTax',
              'ПДВ': 'Invoice/Summary/SummaryRow/RowDetail/TaxAmount[@Type="VAT"]',
              'Пенс. фонд': 'Invoice/Summary/SummaryRow/RowDetail/TaxAmount[@Type="PF"]',
              'Разом': 'Invoice/Summary/SummaryRow/RowDetail/Amount'
              }

    xml_file = sys.argv[1]
    wb = Workbook()
    ws = wb.active
    make_header(ws, list(fields.keys()), True)
    hide_col(ws, 'G')
    hide_col(ws, 'I')
    data = parse_file(xml_file, fields.values())
    write_cells(ws, data)
    wb.save(Path(xml_file).stem + '.xlsx')


def float_safe(string):
    try:
        return float(string)
    except (TypeError, ValueError):
        return string


def make_header(work_sheet, head_list, freeze):
    i = 1
    for col in head_list:
        cell = work_sheet.cell(1, i)
        cell.value = col
        # make title bold
        cell.font = Font(bold=True)
        # resize column
        col_letter = cell.column_letter
        work_sheet.column_dimensions[col_letter].width = len(cell.value) + 2
        i = i + 1
        # freeze title row
    if freeze:
        work_sheet.freeze_panes = "A2"


def hide_col(work_sheet, col):
    column_dimension = work_sheet.column_dimensions[col]
    column_dimension.hidden = True


def parse_file(xml_file, tag_list):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data = []
    for account in root:
        row = []
        for tag in tag_list:
            if len(sys.argv) > 2 and sys.argv[2] == '-d':
                row.append(account.findall(tag))
            else:
                row.append([account.find(tag)])
        data.append(row)
    return data


def write_cells(ws, data):
    r = 2
    r_max = 0
    for row in data:
        r = r + r_max
        r_max = 0
        c = 0
        for column in row:
            c = c + 1
            t = 0
            for elem in column:
                try:
                    val = float_safe(elem.text)
                except AttributeError:
                    continue
                # ugly
                if c == 6 and row[6][0].text == 'заборгованiсть':
                    val = -val
                if c == 8 and row[8][0].text == 'заборгованiсть':
                    val = -val

                ws.cell(row=r + t, column=c, value=val)
                t = t + 1
                if t > r_max:
                    r_max = t


if __name__ == '__main__':
    main()
