import xlsxwriter

class Workbook():
    def __init__(self):
        self.workbook = xlsxwriter.Workbook('Board directors.xlsx')

    def createWorksheet(self):
        self.worksheet = self.workbook.add_worksheet()

    def create_worksheet_headings(self):
        col = 0
        format = self.workbook.add_format({'bold': True, 'font_size': 14})
        headings = ["Company Name", "ASX Code", "Board director Name", "Board Director Title", "pay", "exercised", 
        "year born", "gender", "Sector", "Industry", "Company Description", "Adresss", "Postcode"]
        for heading in headings:
            self.worksheet.write(0, col, heading, format)
            col+=1

    def close_spreadsheet(self):
        self.workbook.close()


i = Workbook()
i.createWorksheet()
i.create_worksheet_headings()
i.close_spreadsheet()