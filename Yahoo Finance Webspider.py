from bs4 import BeautifulSoup
import requests
class Scraper():

	def __init__(self, listdir):
		self.listdir = listdir #Directory that contains the ASX codes to search
		self.ASXCodes = [] #List that the ASX codes that are searched are placed
		self.titles = ["Sir", "Dr.", "Mr.", "Ms."] #List of titles used on the Yahoo Finance website
		self.headers = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36' } #headers for HTML request that emulate a browser

	def formatWebsite(self, asxcode): #Formats the URL to return the profile page of a company's Yahoo Finance page
		formatted = 'https://au.finance.yahoo.com/quote/' + asxcode + '/profile?p=' + asxcode
		return str(formatted)

	def formattingCommas(self, text): #Gets rid of commas so CSV works, consider switching to xlsx
		try: #Lazy error exception that just tells me where the program failed 
			if ',' in text: # if there is a comma in the text replaces with a dash
				#print("Formatted")
				formatted = text.replace(',', ' - ')
				return formatted
			else:
				#print('No formatting required')
				return text	
		except TypeError:
			print("Couldn't convert {}".format(text))
			return "N/A" # Returns N/A if couldn't convert for some reason

	def requestPage(self, ASX, HTMLtag): #function to return a page request targeting a specific HTML tag
		URL = self.formatWebsite(ASX) #Uses the Yahoo Finance website request with a given ASX code
		page = requests.get(URL, headers=self.headers).text
		soup = BeautifulSoup(page, 'lxml')
		results = soup.findAll(HTMLtag)
		return results #returns the result of the targeted HTML tag Beautiful soup'd

	def gender(self, name): #Detects the title of a director and returns their gender
		if "Mr." or "Sir" in name:
			return "Man"
		elif "Ms." in name:
			return "Female"
		else: 
			return "Undefined"

	def findASXValue(self, ASX, value): #Finds a desired field in the Yahoo Finance information table from given argument "value"
		results = self.requestPage(ASX, 'span')
		for i in range(len(results)):
			if value in results[i].get_text():
				print(results[i+1].get_text())
				return results[i+1].get_text()

	def findCompanyAddress(self, ASX): #Finds and returns the company's address. Required a different method than findASXValue function
		counter = 0
		results = self.requestPage(ASX, 'p')
		try:
			formattedResults = results[1].get_text()
		except IndexError:
			print("Couldn't find address data for: {}".format(ASX))
			return
		ausindex = results[1].get_text().find("Australia") + 9
		if len(formattedResults) == 9:
			print("only know that {} is in Australia".format(ASX))
			return formattedResults[:ausindex]
		return formattedResults[:ausindex]

	def findCompanyDesc(self, ASX): 	#find company description
		results = self.requestPage(ASX, 'p')
		formatted = self.formattingCommas(results[3].get_text())
		return formatted

	def returnSearchASX(self, ASX): #Gets ASX director's and puts them in the CSV NOTE: consider seperating the consolidating of information and putting it into CSV file.
		counter = 0 #counter for putting the results in the CSV document
		results = self.requestPage(ASX, 'td') #creates a result page to check 

		#Sometimes website has no data despite having the company. Checks to see if company's page is empty. Consider putting this in seperate function.
		if not results:
			print("No information found for: {0}".format(ASX))
			return
		#test to see if ASX code is available
		if not any(word in results[counter].get_text() for word in self.titles):
			print("Checking if {0} is in {1}".format(results[counter].get_text(), self.titles))
			print("Tried to pull data for {0}, instead got:\n".format(ASX))
			print(results[counter].get_text())
			return
		########################################################

		name = soup.findAll('h1')[0].get_text()
		sector = self.findASXValue(ASX, "Sector(s)")
		industry = self.findASXValue(ASX, "Industry")
		with open('Board Directors.csv', 'a', encoding="utf-8") as fd:
			fd.write("\n")
			for i in range(len(results)):
				try:
					fd.write(
						"{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}\n".format(
						name,
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
					
					counter = counter + 5
				except IndexError:
					continue
   	
	def openCSVFile(self): #creates CV file that stores director's of companies
   		with open('Board Directors.csv', 'w+') as fd:
   			fd.write("Company Name, ASX Code, Board director Name, Board Director Title, pay, exercised, year born, gender, Sector, Industry, Company Decsription, Adresss, Postcode")

	def itemiseList(self): #Creates a list for the ASX codes
		print(self.listdir)
		with open(self.listdir, 'r') as f:
			self.ASXCodes = [line.strip() for line in f]

	def mainloop(self): #main loop - lazy code
		self.itemiseList()
		self.openCSVFile()
		for i in range(0, len(self.ASXCodes)):
			print("Searching {0}: {1} out of {2}".format(self.ASXCodes[i], i, len(self.ASXCodes)))
			self.returnSearchASX(self.ASXCodes[i])

s = Scraper(r'C:\Users\clerk\Desktop\List of ASX codes.txt')
s.mainloop()
print("Done")