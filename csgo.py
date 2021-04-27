import pandas as pd
import math
from xlsxwriter.utility import xl_rowcol_to_cell
from steam_community_market import Market, AppID

CURRENCY = "BRL"
FILE_NAME = "csgo_Prices.xlsx"
game = AppID.CSGO # RUST, STEAM, DOTA2, TF2

def main():
    df = pd.read_excel(FILE_NAME, sheet_name="Prices Tracker")
    number_rows = len(df.index)
    names = df['Names']
    current_prices = df['Current Prices']
    prices_paid = df['Price Paid']

    prices_list = []
    profit_list = []
    percent_list = []
    for i in range(len(names)):
        current_price = market.get_lowest_price(str(names.iloc[i]), game)
        if type(names.iloc[i]) == float:
                break
        if current_price is not None:
            profit = current_price - prices_paid.iloc[i]
        else:
            current_price = 0
            profit = 0

        prices_list.append(current_price)
        profit_list.append(profit)
        percent_list.append(profit / prices_paid.iloc[i])
        print(f"{names.iloc[i]}, {current_price}")

    df['Current Prices'] = prices_list
    df['Profit'] = profit_list
    df['%'] = percent_list

    writer = pd.ExcelWriter(FILE_NAME, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name="Prices Tracker")
    workbook = writer.book
    worksheet = writer.sheets["Prices Tracker"]

    total_fmt = workbook.add_format({'align': 'center', 'num_format': 44,'bold': True}) # 44 is the accounting excel id
    money_fmt = workbook.add_format({'align': 'center', 'num_format': 44}) 
    percent_fmt = workbook.add_format({'align': 'center', 'num_format': '0.0%'})

    worksheet.set_column('A:A', 45)
    worksheet.set_column('B:D', 14, money_fmt)
    worksheet.set_column('E:E', 14, percent_fmt)
    worksheet.set_column('F:F', 3)
    worksheet.set_column('G:I', 14, money_fmt)

    for column in range(1, 4):
        cell_location = xl_rowcol_to_cell(1, column+5)
        start_range = xl_rowcol_to_cell(1, column)
        end_range = xl_rowcol_to_cell(number_rows, column)
        formula = "=SUM({:s}:{:s})".format(start_range, end_range)
        worksheet.write_formula(cell_location, formula, total_fmt)

    color_range = f"D2:D{number_rows+1}"
    green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})    
    red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    worksheet.conditional_format(color_range, {'type': 'cell', 'criteria': '>', 'value': '0', 'format': green})
    worksheet.conditional_format(color_range, {'type': 'cell', 'criteria': '<', 'value': '0', 'format': red})

    writer.save()
    
if __name__ == '__main__':
    market = Market(CURRENCY)
    main()
    