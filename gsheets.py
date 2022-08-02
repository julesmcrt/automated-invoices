import gspread
from datetime import datetime



def getData(sheet_url, worksheet_name, credentials_filepath):
    """Retrieve the whole worksheet in order to make the invoices
    Output is a list of dicts, each line is a dict except for the first one (it's used as the keys for all the other lines)
    Note: it uses a Service Account in Google Cloud API (not OAuth like the Google Docs part of this project)
    """

    gc = gspread.service_account(filename=credentials_filepath)
    sh = gc.open_by_url(sheet_url)
    wks = sh.worksheet(worksheet_name)
    return wks.get_all_records()



def calcMissingFields(data):
    """Computes every missing field
    """

    for line in data:
        line['date_invoice'] = str(datetime.today().strftime('%d/%m/%Y')) # needs to be in global script for optimization purposes
        line['price_total_ht'] = 0

        for i in range(1, 5): # max number of lines in table
            if line[f'detail{i}_info'] != '' and line[f'detail{i}_info'] != '-':

                # Sanitize inputs to enable conversion to float
                if type(line[f'detail{i}_unit_price']) == str:
                    line[f'detail{i}_unit_price'] = float(line[f'detail{i}_unit_price'].replace(',', '.').replace(' ', '').replace('€', ''))
                if type(line[f'detail{i}_quantity']) == str:
                    line[f'detail{i}_unit_price'] = float(line[f'detail{i}_unit_price'].replace(',', '.').replace(' ', ''))
                
                # Checks for 'Acompte' keyword: the price must be negative
                if 'Acompte' in line[f'detail{i}_info']:
                    line[f'detail{i}_unit_price'] *= -1 * int(line[f'detail{i}_unit_price'] >= 0)
                    
                line[f'detail{i}_price'] = float(line[f'detail{i}_quantity']) * float(line[f'detail{i}_unit_price'])
                line['price_total_ht'] += line[f'detail{i}_price']
            else:
                line[f'detail{i}_price'] = ''

        line['price_tva'] = line[f'price_total_ht'] * 0.2
        line['price_total_ttc'] = line[f'price_total_ht'] * 1.2
        line['price_to_pay'] = line['price_total_ttc']

        # Splits detail1_info which contains a line in italic
        line['detail1_info'], line['detail1_meals_info'] = line['detail1_info'].split('\n')
    
    return data



def formatData(data):
    """Clears up the data (empty cells) and format everything to look nice
    """

    for line in data:
        for key in line.keys():
            if line[key] == '-':
                line[key] = ''

            if type(line[key]) == int and not (key == 'order_id' or 'quantity' in key):
                line[key] = float(line[key])

            if type(line[key]) == float:
                main, decimal = str(line[key]).split('.')
                if len(main) >= 4: # adds a space if the number is bigger than 1 000,00€
                    main = main[:-3] + ' ' + main[-3:]
                if len(decimal) >= 2: # removes unnecessary digits
                    decimal = str(round(float('0.' + decimal), 2))[2:]
                if len(decimal) <= 2: # adds back zeros to have 2 digits after the decimal point
                    decimal += '0' * (2-len(decimal))
                line[key] = str(main) + ',' + str(decimal) + '€'
            
            else:
                line[key] = str(line[key])
    return data