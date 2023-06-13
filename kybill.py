import sys
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from pathlib import Path


def float_safe(string):
    try:
        return float(string)
    except TypeError:
        return ""


if len(sys.argv) <= 1:
    print("enter full path to xml file, add '-d' for detailed expense")
    sys.exit(1)

file = sys.argv[1]
tree = ET.parse(file)
root = tree.getroot()
wb = Workbook()
ws = wb.active
ws.freeze_panes = "A2"
ws['A1'] = "Особ. номер"
ws['B1'] = "Номер"
ws['C1'] = "Тариф"
ws['D1'] = "Дата початку"
ws['E1'] = "Дата закінчення"
ws['F1'] = "Баланс поч."
ws['G1'] = "Баланс кін."
ws['H1'] = "Надходження"
ws['I1'] = "Рекоменд."
ws['J1'] = "Сума"
ws['K1'] = "ПДВ"
ws['L1'] = "Пенс. фонд"
ws['M1'] = "Разом"

i = 2

for account in root:
    begin_date = account.findtext('.//{}'.format('BeginDate'))
    end_date = account.findtext('.//{}'.format('EndDate'))
    customer_phone = account.findtext('.//{}'.format('CustomerPhone'))
    billing_account = account.findtext('.//{}'.format('BillingAccount'))
    amount_exclTax = account.findtext('.//{}'.format('AmountExclTax'))
    tax_amount_vat = account.findtext('.//{}[@{}="{}"]'.format('TaxAmount', 'Type', 'VAT'))
    tax_amount_pf = account.findtext('.//{}[@{}="{}"]'.format('TaxAmount', 'Type', 'PF'))
    amount = account.findtext('.//{}[@{}]'.format('Amount', 'Header'))
    balance_begin = account.findtext('.//{}'.format('BalBeginMonth'))
    balance_begin_text = account.findtext('.//{}'.format('BalBeginMonthText'))
    balance_end = account.findtext('.//{}'.format('BalEndMonth'))
    balance_end_text = account.findtext('.//{}'.format('BalBeginEndText'))
    payments_total_amount = account.findtext('.//{}'.format('PaymentsTotalAmount'))
    recommended_payment = account.findtext('.//{}'.format('RecommendedPayment'))
    contract_type = account.findtext('.//{}'.format('ContractType'))
    # payment = account.findtext('.//{}'.format('PaymentsBank'))

    details = account.find('.//{}[@{}="{}"]'.format('Summary', 'Type', 'BIGroup'))

    ws.cell(row=i, column=1).value = billing_account
    ws.cell(row=i, column=2).value = customer_phone
    ws.cell(row=i, column=3).value = contract_type
    ws.cell(row=i, column=4).value = begin_date
    ws.cell(row=i, column=5).value = end_date
    if balance_begin_text == "заборгованiсть":
        balance_begin = 0 - float_safe(balance_begin)
    ws.cell(row=i, column=6).value = float_safe(balance_begin)
    if balance_end_text == "заборгованiсть":
        balance_end = 0 - float_safe(balance_end)
    ws.cell(row=i, column=7).value = float_safe(balance_end)
    ws.cell(row=i, column=8).value = float_safe(payments_total_amount)
    # ws.cell(row=i, column=13).value = payment
    ws.cell(row=i, column=9).value = float_safe(recommended_payment)

    ws.cell(row=i, column=10).value = float_safe(amount_exclTax)
    ws.cell(row=i, column=11).value = float_safe(tax_amount_vat)
    ws.cell(row=i, column=12).value = float_safe(tax_amount_pf)
    ws.cell(row=i, column=13).value = float_safe(amount)

    if len(sys.argv) > 2 and sys.argv[2] == "-d":
        try:
            text = details.findall('.//{}'.format('Text'))
            amount_exclTax = details.findall('.//{}'.format('AmountExclTax'))
        except AttributeError:
            text = []
            amount_exclTax = []
        j = 0
        while j < len(text):
            i = i + 1

            ws.cell(row=i, column=5).value = text[j].text
            ws.cell(row=i, column=10).value = float_safe(amount_exclTax[j].text)
            j = j + 1
        i = i + 1

    i = i + 1
file = Path(file).stem

wb.save(file + '.xlsx')
