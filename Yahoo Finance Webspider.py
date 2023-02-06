from bs4 import BeautifulSoup
import requests
import xlsxwriter


class Scraper():

    def __init__(self, listdir):
        # directory that contains the ASX codes to search
        self.listdir = listdir
        # list that the ASX codes that are searched are placed
        self.ASXCodes = []
        # list of titles used on the Yahoo Finance website
        self.titles = ["Sir", "Dr.", "Mr.", "Ms."]
        # headers for HTML request that emulate a browser
        self.headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
        #Counts the rows of the current worksheet
        self.row= 1
        #name of the excel spreadsheet
        self.workbook = xlsxwriter.Workbook('Board directors.xlsx')

    def createWorksheet(self):
        self.worksheet = self.workbook.add_worksheet()

    def formatWebsite(self, asx_code):
        # Formats the URL to return the profile \
        # of a company's Yahoo Finance pageformatted
        formatted = f"https://finance.yahoo.com/quote/{asx_code}.AX/profile?p={asx_code}.AX"
        return str(formatted)

    def requestPage(self, ASX, HTMLtag, HTML_class=False):
        # function to return a page request targeting a specific HTML tag. Can specify a class for specific result
        # Uses the Yahoo Finance website request with a given ASX code
        URL = self.formatWebsite(ASX)
        page = requests.get(URL, headers=self.headers)
        soup = BeautifulSoup(page.content, 'lxml')
        #If a HTML class is specified, will use the find() function instead and return specific result, otherwise uses findAll
        if HTML_class:
            results = soup.find(HTMLtag, class_=HTML_class)
            return results
        else:
            results = soup.findAll(HTMLtag)
            return results

    def gender(self, name):
        # Detects the title of a director and returns their gender
        gender = str()
        if "Mr." or "Sir" in name:
            return "Male"
        elif "Ms." or "Mrs." in name:
            return "Female"
        else:
            return "Undefined"

    def findASXValue(self, ASX, value):
        # Finds a desired field in Yahoo Finance information table\
        # from given argument "value"
        results = self.requestPage(ASX, 'span')
        for i in range(len(results)):
            if value in results[i].get_text():
                print("Found {0}'s value for {1}".format(ASX, value))
                return results[i+1].get_text()

    def findCompanyAddress(self, ASX):
        # Finds and returns the company's address.
        # Required a different method than findASXValue function
        print("Finding {0}'s address".format(ASX))
        company_address_extract = self.requestPage(ASX, 'p', "D(ib) W(47.727%) Pend(40px)")
        company_address_list = [text for text in company_address_extract.stripped_strings]

        print(f'Found address for {ASX}: {company_address_list}')
        return company_address_list

    def findCompanyDesc(self, ASX):
        # find company description
        results = self.requestPage(ASX, 'p')
        description = results[3].get_text()
        return description

    # Function for writing which ASX codes had an error scraping
    def writeErrors(self, ASX):
        with open('Errors.txt', 'a', encoding='utf-8') as fd:
            fd.write("Error finding information for: {}\n".format(ASX))

    #function not implemented yet. when implemented, take out of returnSearchASX()
    def hasASXpage(self, results, counter, ASX):
        # checks to see if the supplied ASX code has a Yahoo Finance page
        if not results:
            print("No information found for: {0}".format(ASX))
            self.writeErrors(ASX)
            return False
        # test to see if ASX code is available
        if not any(word in results[counter].get_text() for word in self.titles):
            print("Checking if {0} is in {1}".format(results[counter].get_text(), self.titles))
            print("Tried to pull data for {0}, instead got:\n".format(ASX))
            print(results[counter].get_text())
            self.writeErrors(ASX)
            return False
        return True

    def write_to_excel(self, ASX):
        # Gets ASX director's and puts them in the CSV
        # Focuses on table containing board director information
        results = self.requestPage(ASX, 'td')
        board_mem_num = int(len(results)/5)
        counter = 0 #counter for the 'td' table results. Each row has 5 columns

        #if no 'td' tags on page, means company doesn't have Yahoo Finance page
        if not self.hasASXpage(results, counter, ASX):
            return


        for i in range(board_mem_num):
            col = 0 #counter for column

            #Currently the index error is how end of 'td' table is detected
            try:
                #All data for one row of spreadsheet
                company_name = self.requestPage(ASX, 'h1')[0].get_text()
                board_dir_name = results[counter].get_text()
                board_dir_title = results[counter+1].get_text()
                board_dir_pay = results[counter+2].get_text()
                board_dir_exercised = results[counter+3].get_text()
                board_dir_DOB = results[counter+4].get_text()
                board_dir_gender = self.gender(results[counter].get_text())
                company_sector = self.findASXValue(ASX, "Sector(s)")
                company_industry = self.findASXValue(ASX, "Industry")
                company_desc = self.findCompanyDesc(ASX)
                company_address = self.findCompanyAddress(ASX)
                
                #Puts the above in a list
                row_of_data = [company_name, ASX, board_dir_name, board_dir_title, board_dir_pay,
                board_dir_exercised, board_dir_DOB, board_dir_gender, company_sector, company_industry, company_desc]


                for c in row_of_data:
                    self.worksheet.write(self.row, col, c)
                    col+=1
                for a in company_address:
                    self.worksheet.write(self.row, col, a)
                    col+=1
                self.row+=1
                print(f"Board member {board_dir_name} is being written.\n{i+1} out of {board_mem_num} board members")
                counter = counter + 5
            except IndexError:
                print("An unknown index error occurred!!)
                continue
    
    def create_worksheet_headings(self):
        col = 0
        format = self.workbook.add_format({'bold': True, 'font_size': 14})
        headings = ["Company Name", "ASX Code", "Board director Name", "Board Director Title", "pay", "exercised", 
        "year born", "gender", "Sector", "Industry", "Company Description", "Adresss", "Postcode"]
        for heading in headings:
            self.worksheet.write(0, col, heading, format)
            col+=1

    def itemiseList(self):
        # Creates a list for the ASX codes
        print(self.listdir)
        with open(self.listdir, 'r') as f:
            self.ASXCodes = [line.strip() for line in f]

    def mainloop(self):
        # main loop - lazy code
        self.itemiseList()
        self.createWorksheet()
        self.create_worksheet_headings()

        #for i in range(0, len(self.ASXCodes)):
        for i in range(0, 2):
            print("Searching {0}: {1} out of {2}".format(self.ASXCodes[i], i, len(self.ASXCodes)))
            self.write_to_excel(self.ASXCodes[i])
        self.workbook.close()



s = Scraper('List of ASX codes.txt')
s.mainloop()
print("Program finished successfully")
