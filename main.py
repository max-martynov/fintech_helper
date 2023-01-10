import yfinance as yf
import openpyxl 
from dataclasses import dataclass
import argparse

# some types of this class are missing
# TODO: use the spreadsheet and assign a type to each field of the class
@dataclass
class CompanyFinanceData:
    name: str
    ticker: TODO
    market_cap: float
    enterprise_value: float
    current_price: TODO
    week_high: TODO 
    revenue_growth: TODO
    profit_margins: TODO
    enterprise_to_revenue: TODO
    enterprise_to_ebitda: TODO
    trailing_pe: TODO
    peg_ratio: TODO
    price_to_book: TODO
    trailing_eps: TODO


class ExcelSheet:

    def __init__(self, path_to_file) -> None:
        self.path_to_file = path_to_file
        self.workbook = openpyxl.load_workbook(path_to_file)
        self.sheet = self.workbook.active

    def load_company(self, company_ticker, row_number):
        company_finance_data = self.get_company_finance_data(company_ticker)
        self.write_company(company_finance_data, row_number)

    def save_sheet(self):
        self.workbook.save(self.path_to_file)

    def get_company_finance_data(self, company_ticker):
        data = yf.Ticker(company_ticker).info # here we use external library to retrieve financial data from yahoo
        # data is a regular python dictionary
        # TODO: use it to fill all fields of the class
        # hint: use print(json.dumps(data, sort_keys=True, indent=4)) to find the right keys
        return CompanyFinanceData(
            name=data['longName'],
            ticker=company_ticker,
            market_cap= data['marketCap'],
            enterprise_value= TODO,
            current_price=data, 
            week_high=data['currentPrice'] / data['fiftyTwoWeekHigh'] * 100, 
            revenue_growth= TODO,
            profit_margins= TODO, 
            enterprise_to_revenue= TODO,
            enterprise_to_ebitda= TODO,
            trailing_pe= TODO,
            peg_ratio= TODO,
            price_to_book= TODO,
            trailing_eps= TODO
        )

    def write_company(self, cfd: CompanyFinanceData, row_number):
        attrs = vars(cfd)
        for field, value in attrs.items():
            # TODO: if the field equals to 'market_cap' or 'enterprise_value' we should convert it to billions.
            # hint: you can use function convert_to_billions() below

            # TODO: for some fields we need to add percent sign
            # figure out for which fields percent sign is needed (using spreadsheet)
            # hint: you can use function add_percent_sign() below
           
            self.write_cell(row_number, ExcelSheet.property_to_column[field], value)

    def convert_to_billions(self, value):
        return value / 1000000000

    def add_percent_sign(self, value):
        return str(round(value, 2)) + '%'

    def write_cell(self, row_number, column_number, value):
        self.sheet.cell(row_number, column_number).value = value

    # TODO: complete the dictionary
    # using spreadsheet figure out which column corresponds to each field
    field_to_column = {
        'name': 3,
        'ticker': 4,
        'market_cap': 5,
        'enterprise_value': ,
        'current_price': ,
        'week_high': ,
        'revenue_growth': ,
        'profit_margins': ,
        'enterprise_to_revenue': ,
        'enterprise_to_ebitda': ,
        'trailing_pe': ,
        'peg_ratio': ,
        'price_to_book': ,
        'trailing_eps': 
    }
    

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('excel_file', type=str,
                    help='Path to excel file')
    parser.add_argument('ticker', type=str,
                    help='Company ticker to add')
    parser.add_argument('row', type=int,
                    help='Row number where information will be stored')
    args = parser.parse_args()
    excel_sheet = ExcelSheet(args.excel_file)
    excel_sheet.load_company(args.ticker, args.row)
    excel_sheet.save_sheet()

if __name__ == "__main__":
    main()