# ASXcompanyscraper
A python script that inputs a list of ASX codes, scrapes Yahoo Finance and outputs a spreadsheet with relevant information. This would be useful for a person who was attmepting to do initial due diligence on a handful of companies.  

# Dependencies
- Requests
- Beautifulsoup4
- xlsxwriter

# Work to be done
- [ ] Code needs cleaning up. Certain functions currently do more than one thing.
- [ ] Code needs to have it's dependencies added to a folder. Perhaps a setup.py?
- [ ] GUI would make it easier to interact with the program.
- [ ] Consider switching the output from CSV to an xlsx spreadsheet using Pandas as the component.
- [ ] Consider making certain outputs optional like the market cap, board members, etc.
- [ ] Allow user to choose an output directory.
- [ ] Allow the user to choose the list location.
- [ ] Create a tester program that makes sure each module is working correctly.

# Issues
- Currently has terrible exception handling
- Currently has no explanation if a company page cannot be found
