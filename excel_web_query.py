################################################################################
##
##  File           : excel_web_query.py
##  Description    : Takes input url and attempts to gather relevant data from
##                   embedded tables and format that data in Excel.
##
##   Author        : *** Andrew Millward ***
##   Last Modified : *** 02/27/2020 ***
##

## Import Files
import requests, xlsxwriter, webbrowser, re, os

################################################################################
##
## Function     : format
## Description  : Conditionally formats all columns of an input spreadsheet.
##
## Inputs       : worksheet - xlsxwriter spreadsheet to be formatted.
##                row_count - Number of rows to apply formatting to.
## Outputs      : 0 if successful, -1 if failure.

def format(worksheet, row_count):

    try:

        # Specify predetermined rows of table to format.
        cells_to_format = 'D3:D' + row_count
        cells_to_format2 = 'E3:E' + row_count
        cells_to_format3 = 'F3:F' + row_count
        cells_to_format4 = 'G3:G' + row_count
        cells_to_format5 = 'H3:H' + row_count
        cells_to_format6 = 'I3:I' + row_count
        cells_to_format7 = 'J3:J' + row_count

        # Apply 3_color_scale formatting to those aforementioned rows.
        worksheet.conditional_format(cells_to_format, {'type': '3_color_scale'})
        worksheet.conditional_format(cells_to_format2, {'type': '3_color_scale'})
        worksheet.conditional_format(cells_to_format3, {'type': '3_color_scale'})
        worksheet.conditional_format(cells_to_format4, {'type': '3_color_scale'})
        worksheet.conditional_format(cells_to_format5, {'type': '3_color_scale'})
        worksheet.conditional_format(cells_to_format6, {'type': '3_color_scale'})
        worksheet.conditional_format(cells_to_format7, {'type': '3_color_scale'})

        return ( 0 ) # Success

    except:
        return ( -1 ) # Failure

################################################################################
##
## Function     : write_spreadsheet
## Description  : Takes input data and exports it to a new spreadsheet.
##
## Inputs       : heading_list - List of table headings for columns.
##                row_contents_array - Contains all data for each array row.
##                rows - Stores the total number of rows to write to.
## Outputs      : 0 if successful, -1 if failure.

def write_spreadsheet(heading_list, row_contents_array, rows):

    try:

        # Create a new workbook with a worksheet.
        workbook = xlsxwriter.Workbook('DataQuery.xlsx') 
        worksheet = workbook.add_worksheet()

        # Write each heading across the top of the worksheet.
        col = 0
        for heading in heading_list:
            worksheet.write(0, col, heading)
            col += 1

        # Write content to each cell of table.
        row = 1
        col = 0
        for row_index in row_contents_array:
            
            for item in row_index:
                
                # Attempt to treat data as a float to handle decimal inputs.
                try: 
                    item2 = float(re.compile(r'[^\d.]+').sub('',item))
                    worksheet.write(row, col, item2)

                # If it is not a float, an exception will occur.
                except:

                    # Now try treating it as an integer for decimal numbers.
                    try:
                        item2 = int(re.sub("[^0-9]", "",item))
                        worksheet.write(row, col, item2)

                    # If that doesn't work, treat it as a general case of strings.
                    except:
                        item2 = str(item)
                        worksheet.write(row, col, item2)

                col += 1
            row += 1
            col = 0

        # Format the worksheet.
        row_count = str(len(rows)+1)
        format(worksheet, row_count)

        # Close the workbook from editing to commit changes.
        workbook.close()

        return ( 0 ) # Success

    except:
        return ( -1 ) # Failure

################################################################################
##
## Function     : open_spreadsheet
## Description  : Launches a spreadsheet application for the user to view data.
##
## Inputs       : None
## Outputs      : 0 if successful, -1 if failure.

def open_spreadsheet():
    
    # Use OS module to open spreadsheet by name.
    try:
        path = os.path.abspath('DataQuery.xlsx')
        webbrowser.open(path)

        return ( 0 ) # Success
    
    except:
        return ( -1 ) # Failure

################################################################################
##
## Function     : DataQuery
## Description  : Read data from online database and exports it to Excel
##
## Inputs       : url - URL with valid table to scan.
## Outputs      : 0 if successful, -1 if failure.

def DataQuery(url):

    try:
    
        # Specialized local import
        from bs4 import BeautifulSoup as soup
        
        # Create a request to the given URL and store data.
        request = requests.get(url)
        data = request.text
        soup = soup(data, "html.parser")

        # Search for a table with specified attributes.
        table = soup.findAll("table",
            {"class", "display sortable_datatable fixed-headers"})[0]

        # Gather all header cell tags and write them to a list.
        heading_list = (table.thead.tr.findAll("th"))
        for i in range(len(heading_list)):
            heading_list[i] = heading_list[i].text

        # Gather all body data points and write them to list of lists.
        rows = table.tbody.findAll("tr")
        row_contents_array = []
        for i in range(len(rows)):
            row_contents = rows[i].findAll("td")
            for j in range(len(row_contents)):
                row_contents[j] = (row_contents[j].text)
            row_contents_array.append(row_contents)

        # Write data to spreadsheet.
        write_spreadsheet(heading_list, row_contents_array, rows)

        # Launch Excel and open spreadsheet.
        open_spreadsheet()

        return ( 0 ) # Success

    # Exception handling for various cases of use.
    except xlsxwriter.exceptions.FileCreateError: 
        print("Please close existing Excel spreadsheet and try again.")
        return ( -1 ) # Failure
    except requests.exceptions.ConnectionError:
        print("Make sure you are using a valid URL.")
        return ( -1 ) # Failure
    except IndexError:
        print("Verify the URL you are using has a suitable table.")
        return ( -1 ) # Failure
    except:
        print("Unknown error, try again.")
        return ( -1 ) # Failure

# Run
DataQuery("https://www.bls.gov/oes/current/oes_nat.htm#15-0000")