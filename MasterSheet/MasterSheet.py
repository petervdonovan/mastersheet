from XlsxMake import XlsxMake
spreadsheetBuilder = XlsxMake() # Create the object that will make the workbook
spreadsheetBuilder.getUserInput()   # Ask the user for the URL of the Google sheet and 
                                    # the number of programs for each school
spreadsheetBuilder.save() # Create and save the workbook.