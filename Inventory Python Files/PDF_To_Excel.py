from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.chart.axis import DateAxis

from PDF_Class import PDF

def excel_wb_maker():

    # makes the excel workbook
    wb = Workbook()

    #sets the relevant sheet of the Excel document that will be edited
    ws = wb.active

    #gives the name for the resulting saved Excel file
    ws.title = "Inventory Analytics.xlsx"

    #this sheet stores the graphs made from the pdf data
    wb.create_sheet("Graphs")

    return wb, ws



def excel_pdf_vertical_copier(ws):
    #allows the rows to be placed where they need to be, and allows for space in between pasted pdfs in Excel
    pdf_offset = 1

    #moves the dates where each pdf's data is collected to start at 'I2' and continue rightward
    date_offset = 9

    #iterates through every pdf in formatted_data
    for pdf in PDF.pdf_list:
        row_count = 1

        for row in pdf.data:
            #adds to the rows in Excel without header entries like SCDB, FTM, etc.
            if len(row) < 6:
                row.insert(0, "")

            col_count = 1

            for row_data in row:
                #gives the column letter and advances it forward for the next cell in a pdf's given row
                char = get_column_letter(col_count)
                col_count += 1

                #this line is for debug purposes, but the one below it takes the data from a row on the formatted pdf data and puts it in Excel.
                row_cell = char + str(row_count + pdf_offset)
                ws[row_cell] = row_data

            row_count += 1

        pdf_offset += row_count + 2

        #adds the filename above each pdf's pasted data.
        title_cell = "A" + str(pdf_offset - 34)
        ws[title_cell] = pdf.filename

        #adds the dates that all the inventory will be compared against
        date_cell = get_column_letter(date_offset) + "2"
        ws[date_cell] = pdf.filename[:8]

        if date_cell == "CA":
            continue

        date_offset += 1


#works from the first pdf, assuming it was formatted correctly and identically to all other pdfs
#frameshift issues occur when this assumption is not met
def excel_media_type_adder(ws):

    #the first two ranges handle the cases where SCDB and FTM have to be added to the media type titles
    for i in range(3, 11):
        ws["H" + str(i)] = ws["A3"].value + " " + ws["B" + str(i)].value

    for i in range(11, 19):
        ws["H" + str(i)] = ws["A11"].value + " " + ws["B" + str(i)].value

    for i in range(19, 33):
        ws["H" + str(i)] = ws["A" + str(i)].value + " " + ws["B" + str(i)].value



#this can be re-written using pdf.inventory_list, which would greatly reduce errors and enhance readability
def excel_pdf_inventory_copier(ws):

    #this is where the data starts to be printed
    col_data_offset = 9

    for pdf in PDF.pdf_list:
        row_data_offset = 3
        col_let = get_column_letter(col_data_offset)

        inv_list = pdf.inventory_list

        for inv_val in inv_list:
            ws[col_let + str(row_data_offset)] = inv_val

            #the row changes to allow each inventory value for a pdf to be printed
            row_data_offset += 1
        
        #the column changes once each pdf has their values pasted
        col_data_offset += 1



#this is still in the testing phases, as specifying data for use in the graphs has poor documentation for openpyxl
def excel_graph_maker(wb, ws):
    
    #dateaxis might be an import from openpyxl.chart.axis
    #I think with dateaxis your 'dates' cells have to be in a formatting it can turn into a true date
    #this would format and scale better as Excel is treating the axis as one with special date properties.
    for i in range(1, PDF.pdf_id):
        line_chart = LineChart()
        line_chart.title = "Test Inventory"
        line_chart.x_axis.title = "Date"
        line_chart.y_axis.title = "Inventory"

        #experiment with what this 'crossAx' thing is...
        #line_chart.x_axis = DateAxis()
        #line_chart.x_axis = DateAxis(crossAx=100)

        #this mcustom number format could break things
        #line_chart.x_axis.number_format = "yy-mm-dd"

        #line_chart.x_axis.majorTimeUnit = "days"

        #PDF.pdf_id might have to be +8 or something else to work
        #data
        y_values = Reference(ws, min_col = 9, min_row = i + 2, max_col = PDF.pdf_id + 7, max_row = i + 2)

        #categories
        x_values = Reference(ws, min_col = 9, min_row = 2, max_col = PDF.pdf_id + 7, max_row = 2)

        #experiment with what 'titles_from_data' does
        #this should add the y-values
        #if it's from_rows that's the answer I'm gonna be so mad
        line_chart.set_categories(x_values)

        #this should add the x-values
        #line_chart.set_categories(x_values)
        line_chart.add_data(y_values, from_rows = True)

        ws.add_chart(line_chart, "H34")
        break


    #Chat-GPT says that set_categories has to take in numbers, and not text...


    #IT TRUNCATES THEM WHEN IT DETECTS IT CAN BE ONE REFERENCE
    #series Reference for SCDB100 looks like this
    #'Inventory Analytics.xlsx'!$I$3:$CG$3

    #series Reference for SCDB400 looks like this
    #'Inventory Analytics.xlsx'!$I$4:$CG$4

    #SEE THIS?, THIS IS THE AXIS LABEL RANGE, THIS IS YOUR X VALUES. SERIES VALUES ARE THE Y VALUES
    #AHA
    #='Inventory Analytics.xlsx'!$I$2:$CG$2



#runs all excel related functions. data_scraper() has to be run first to create the PDF class with all it's PDF objects
def excel_batch_processor():

    wb, ws = excel_wb_maker()

    excel_pdf_vertical_copier(ws)

    excel_media_type_adder(ws)

    excel_pdf_inventory_copier(ws)

    #could switch worksheets for this by this point

    #adding wb to help with debugging
    excel_graph_maker(wb, ws)

    wb.save(ws.title)

    print("Excel file complete!")

    #figure out how to remove the pdf files from the github repo
