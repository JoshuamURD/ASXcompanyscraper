from bs4 import BeautifulSoup
import requests


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

    def formatWebsite(self, asxcode):
        # Formats the URL to return the profile \
        # of a company's Yahoo Finance pageformatted
        formatted = 'https://au.finance.yahoo.com/quote/' +\
                asxcode + '.AX/profile?p=' + asxcode + ".AX"
        return str(formatted)

    def formattingCommas(self, text):
        # Gets rid of commas so CSV works, consider switching to xlsx
        try:
            # Lazy error exception that just tells me where the program failed
            # if there is a comma in the text replaces with a dash
            if ',' in text:
                formatted = text.replace(',', ' - ')
                return formatted
            else:
                # print('No formatting required')
                return text
        except TypeError:
            print("Couldn't convert {}".format(text))
            # Returns N/A if couldn't convert for some reason
            return "N/A"

    def requestPage(self, ASX, HTMLtag):
        # function to return a page request targeting a specific HTML tag
        # Uses the Yahoo Finance website request with a given ASX code
        URL = self.formatWebsite(ASX)
        page = requests.get(URL, headers=self.headers).text
        soup = BeautifulSoup(page, 'lxml')
        results = soup.findAll(HTMLtag)
        # returns the result of the targeted HTML tag Beautiful soup'd
        return results

    def gender(self, name):
        # Detects the title of a director and returns their gender
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
        results = self.requestPage(ASX, 'p')
        ausindex = results[1].get_text().find("Australia")

        #Checks to see if foudn "Australia" in ausindex
        if ausindex == -1:
            print("Not located in Australia")
            return "N/A"

        try:
            formattedResults = results[1].get_text()
        except IndexError:
            print("Couldn't find address data for: {}".format(ASX))
            return

        if len(formattedResults) == 9:
            print("only know that {} is in Australia".format(ASX))
            return formattedResults[:ausindex]

        # if the function found a correct Australian address, returns the address
        return formattedResults[:ausindex]

    def findCompanyDesc(self, ASX):
        # find company description
        results = self.requestPage(ASX, 'p')
        formatted = self.formattingCommas(results[3].get_text())
        return formatted

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

    def returnSearchASX(self, ASX):
        # Gets ASX director's and puts them in the CSV
        # NOTE: consider seperating the consolidating of information
        # and putting it into CSV file.FileExistsError
        counter = 0
        # creates a result page to check
        results = self.requestPage(ASX, 'td')

        if not self.hasASXpage(results, counter, ASX):
            return

        companyName = self.requestPage(ASX, 'h1')[0].get_text()
        sector = self.findASXValue(ASX, "Sector(s)")
        industry = self.findASXValue(ASX, "Industry")

        with open('Board Directors.csv', 'a', encoding="utf-8") as fd:
            fd.write("\n")
            for i in range(len(results)):
               
                try:
                    fd.write(
                            "{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}\n".format(
                                companyName,
                                ASX,
                                self.formattingCommas(results[counter].get_text()),
                                self.formattingCommas(results[counter+1].get_text()),
                                self.formattingCommas(results[counter+2].get_text()),
                                self.formattingCommas(results[counter+3].get_text()),
                                self.formattingCommas(results[counter+4].get_text()),
                                self.gender(results[counter].get_text()),
                                self.formattingCommas(sector),
                                self.formattingCommas(industry),
                                self.findCompanyDesc(ASX),
                                self.findCompanyAddress(ASX)
                                ))
                    print("Board member {0} is being written.\n{1} out of {2}".format(results[counter].get_text(), i, len(results))) 
                    counter = counter + 5
                except IndexError:
                    print("Reached end of board members")
                    continue

    def openCSVFile(self):
        # creates CV file that stores director's of companies
        with open('Board Directors.csv', 'w+') as fd:
            fd.write("Company Name, ASX Code, Board director Name, Board Director Title, pay, exercised, year born, gender, Sector, Industry, Company Decsription, Adresss, Postcode")

    def itemiseList(self):
        # Creates a list for the ASX codes
        print(self.listdir)
        with open(self.listdir, 'r') as f:
            self.ASXCodes = [line.strip() for line in f]

    def mainloop(self):
        # main loop - lazy code
        self.itemiseList()
        self.openCSVFile()
        for i in range(0, len(self.ASXCodes)):
            print("Searching {0}: {1} out of {2}".format(self.ASXCodes[i], i, len(self.ASXCodes)))
            self.returnSearchASX(self.ASXCodes[i])


s = Scraper('List of ASX codes.txt')
s.mainloop()
print("Done")
