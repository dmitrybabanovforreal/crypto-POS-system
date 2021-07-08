import csv, sys, os, json, qrcode, time, openpyxl, datetime
from decimal import Decimal
from pycoingecko import CoinGeckoAPI
from cv2 import QRCodeDetector
from cv2 import VideoCapture
from cv2 import CAP_DSHOW
from cv2 import destroyAllWindows


if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(__file__)
user_path = os.path.join(os.path.split(os.path.abspath(application_path))[0])

cryptoCurrencies = {}
wallets = {}
for row in list(csv.reader(open(os.path.join(user_path,'supported currencies.csv')))):
    if len(row) != 3:
        continue
    if not row[0] or not row[1] or not row[2]:
        continue
    cryptoCurrencies[row[0]] = row[1]
    wallets[row[0]] = row[2]

parameters = json.load(open(os.path.join(user_path, 'parameters.txt')))

# Select "buy crypto from the customer" or "sell crypto to the customer"
print('Hi, please select the type the number of transaction to do (1/2) and press Enter:')
print('1. Buy crypto from the customer')
print('2. Sell crypto to the customer')
while True:
    try:
        answer = int(input().strip())
        if answer == 1:
            transactionType = 'buy'
            break
        elif answer == 2:
            transactionType = 'sell'
            break
        else:
            raise Exception('other value')
    except:
        print('No, please type either "1" or "2" and press enter')
print(f'Thank you, you selected to {transactionType} crypto')

while True:
    transactionData = []
    # Input the amount (either the one to be paid by the customer or the one to be received).
    print('Please enter the amount and currency code for the transaction and press Enter to calculate the opposite amount.')
    print('Example: "0.0012 BTC" or "1200 USD"')
    while True:
        text = input()
        currency = text.split(' ')[-1].upper()
        if currency not in ['USD'] + list(cryptoCurrencies.keys()):
            print(f'Currency code {currency} was not recognized or is not supported, please try again')
            continue
        try:
            amount = Decimal(text.upper().split(currency)[0].strip())
        except:
            print(f'Amount {text.upper().split(currency)[0].strip()} was not recognized, please try again')
            continue
        break

    if currency == 'USD':
        print('Please specify to crypto currency to convert to (like ETH or BTC) and press Enter:')
        while True:
            cryptoCurrency = input().strip().upper()
            if cryptoCurrency not in list(cryptoCurrencies.keys()):
                print(f'Currency code {cryptoCurrency} was not recognized or is not supported, please try again and press Enter:')
                continue
            break
    else:
        cryptoCurrency = currency

    print('Getting the market data...')
    cg = CoinGeckoAPI()
    rateData = cg.get_price(ids=cryptoCurrencies[cryptoCurrency], vs_currencies='usd')
    rate = Decimal(str(list(rateData.values())[0]['usd']))
    fee = Decimal(parameters['fee'])

    # calculate the transaction breakdown
    print('\n=======================\n')
    if currency == 'USD' and transactionType == 'buy':
        print(f'Customer would like to pay in {cryptoCurrency} to receive {amount} USD')

        if rate > 1:
            rounding = 0
            oneCent = Decimal('0.01') / rate
            while oneCent < 1:
                rounding += 1
                oneCent *= 10
            convertedAmount = round(amount / rate, rounding)
            feeAmount = round(convertedAmount * fee, rounding)
            amountToPay = round(convertedAmount + feeAmount, rounding)
        else:
            convertedAmount = amount / rate
            feeAmount = convertedAmount * fee
            amountToPay = convertedAmount + feeAmount

        print(f'{convertedAmount} {cryptoCurrency} is {amount} USD')
        print(f'Plus {feeAmount} {cryptoCurrency} - transaction fee ({str(fee * 100)}%)')
        print(f'{amountToPay} {cryptoCurrency} to be paid by the customer')

        transactionData = [amountToPay, cryptoCurrency, amount, 'USD', round(amount * fee, 2)]
    if currency == 'USD' and transactionType == 'sell':
        print(f'Customer would like to pay {amount} USD to receive in {cryptoCurrency}')

        if rate > 1:
            rounding = 0
            oneCent = Decimal('0.01') / rate
            while oneCent < 1:
                rounding += 1
                oneCent *= 10
            convertedAmount = round(amount / rate, rounding)
            feeAmount = round(convertedAmount - (convertedAmount / (1 + fee)), rounding)
            amountToReceive = round(convertedAmount - feeAmount, rounding)
        else:
            convertedAmount = amount / rate
            feeAmount = convertedAmount - (convertedAmount / (1 + fee))
            amountToReceive = convertedAmount - feeAmount

        print(f'{str(amount)} USD is {str(convertedAmount)} {cryptoCurrency}')
        print(f'Minus {feeAmount} {cryptoCurrency} - transaction fee ({str(fee * 100)}%)')
        print(f'Client receives {str(amountToReceive)} {cryptoCurrency}')

        transactionData = [amount, 'USD', amountToReceive, cryptoCurrency, round(amount - (amount / (1 + fee)), 2)]
    if currency != 'USD' and transactionType == 'buy':
        print(f'Customer would like to pay {amount} {cryptoCurrency} to receive in USD')

        convertedAmount = round(amount * rate, 2)
        feeAmount = round(convertedAmount - (convertedAmount / (1 + fee)), 2)
        amountToReceive = round(convertedAmount - feeAmount, 2)

        print(f'{str(amount)} {cryptoCurrency} is {str(convertedAmount)} USD')
        print(f'Minus {feeAmount} USD - transaction fee ({str(fee * 100)}%)')
        print(f'Client receives {str(amountToReceive)} USD')

        transactionData = [amount, cryptoCurrency, amountToReceive, 'USD', feeAmount]
    if currency != 'USD' and transactionType == 'sell':
        print(f'Customer would like to pay in USD to receive {amount} {cryptoCurrency}')

        convertedAmount = round(amount * rate, 2)
        feeAmount = round(convertedAmount * fee, 2)
        amountToPay = round(convertedAmount + feeAmount, 2)

        print(f'{amount} {cryptoCurrency} is {convertedAmount} USD')
        print(f'Plus {feeAmount} USD - transaction fee ({str(fee * 100)}%)')
        print(f'{amountToPay} USD to be paid by the customer')

        transactionData = [amountToPay, 'USD', amount, cryptoCurrency, feeAmount]

    print('\n=======================\n')

    print('Please press Enter to continue')
    print('Type "1" and press Enter to enter a different sum')
    text = input().strip()
    if text == '1':
        continue
    else:
        break


if transactionType == 'buy':
    # show the wallet address of the selected currency
    walletAddress = wallets[cryptoCurrency]
    qrcode.make().show(walletAddress)
    print('Please press Enter to save the transaction')
    print('Type "1" and press Enter to cancel and close the app')
    text = input().strip()
    if text == '1':
        quit()
    else:
        pass
if transactionType == 'sell':
    # scan QR code from webcam
    print('Please scan the wallet address QR code of the customer')
    camera = VideoCapture(0, CAP_DSHOW)
    while True:
        time.sleep(0.5)
        image = camera.read()[1]
        walletAddress = QRCodeDetector().detectAndDecode(image)[0]
        if walletAddress:
            print('\n=======================\n')
            print('Wallet address:')
            print(walletAddress)
            print('\n=======================\n')

            print('Please press Enter to save the transaction')
            print('Type "1" and press Enter to enter to scan the wallet address again')
            text = input().strip()
            if text == '1':
                image = camera.read()[1]
                walletAddress = ''
                continue
            else:
                camera.release()
                destroyAllWindows()
                break

# 6. The app saves all the transaction data into a new row in the Excel file: type of transaction "buy/sell", date and
# time, payed and received sum, fee, currency rate at that moment. Anything else?

# Save the transaction data
excelFileName = 'transactions.xlsx'
if excelFileName not in os.listdir(user_path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    headers = [
        'Date',
        'Time',
        'Type',
        'Payed',
        'Payed (currency)',
        'Received',
        'Received (currency)',
        'Transaction fee (in USD)',
        'Rate (1 coin in USD)',
        'Wallet used'
    ]
    for i, item in enumerate(headers):
        sheet.cell(row=1, column=i + 1).value = item
else:
    wb = openpyxl.load_workbook(excelFileName)
    sheet = wb.active

maxRow = sheet.max_row
sheet.cell(row=maxRow + 1, column=1).value = datetime.datetime.now().strftime('%m/%d/%Y')
sheet.cell(row=maxRow + 1, column=2).value = datetime.datetime.now().strftime('%H:%M')
sheet.cell(row=maxRow + 1, column=3).value = transactionType
for i, item in enumerate(transactionData):
    sheet.cell(row=maxRow + 1, column=4 + i).value = item
sheet.cell(row=maxRow + 1, column=9).value = rate
sheet.cell(row=maxRow + 1, column=10).value = walletAddress

wb.save(os.path.join(user_path, excelFileName))

print('The transaction is saved. You can press Enter to exit or close the window')
input()
