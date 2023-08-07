from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
import math

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
    date_offset = 3

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
        date_cell = "H" + str(date_offset)
        ws[date_cell] = pdf.filename[:8]

        if date_cell == "CA":
            continue

        date_offset += 1
    
    #adds the word "Date" above its respective column
    ws["H2"] = "Date"



#works from the first pdf, assuming it was formatted correctly and identically to all other pdfs
#frameshift issues occur when this assumption is not met
def excel_media_type_adder(ws):

    #start at I2 and move to J2, K2, etc.

    #the first two ranges handle the cases where SCDB and FTM have to be added to the media type titles
    #there is an assumption that nothing new will be added into SCDB and FTM, built into the lengths of each section being hardcoded
    #this allows for the minimum columns to exist alongside the data for each media type
    min_offset = 0
    for i in range(3, 11):
        ws[get_column_letter(i + 6 + min_offset) + "2"] = ws["A3"].value + " " + ws["B" + str(i)].value
        PDF.pdf_media_type_list.append(ws["A3"].value + " " + ws["B" + str(i)].value)
        min_offset += 1

    for i in range(11, 19):
        ws[get_column_letter(i + 6 + min_offset) + "2"] = ws["A11"].value + " " + ws["B" + str(i)].value
        PDF.pdf_media_type_list.append(ws["A11"].value + " " + ws["B" + str(i)].value)
        min_offset += 1

    #the end is dynamically encoded to match the length of each PDF
    for i in range(19, PDF.pdf_length + 2):
        ws[get_column_letter(i + 6 + min_offset) + "2"] = ws["A" + str(i)].value + " " + ws["B" + str(i)].value
        PDF.pdf_media_type_list.append(ws["A" + str(i)].value + " " + ws["B" + str(i)].value)
        min_offset += 1

    #adds the words "Media Type" and "Minimum" above their respective column
    for i in range(len(PDF.pdf_media_type_list)):
        inv_title_col_let = get_column_letter(9 + (2 * i))
        med_title_col_let = get_column_letter(10 + (2 * i))

        ws[inv_title_col_let + "1"] = "Media Type"
        ws[med_title_col_let + "2"] = "Minimum"



#this can be re-written using pdf.inventory_list, which would greatly reduce errors and enhance readability
def excel_pdf_inventory_copier(ws):

    #this is where the data starts to be printed
    row_data_offset = 3

    for pdf in PDF.pdf_list:
        col_data_offset = 9

        inv_list = pdf.inventory_list
        min_list = pdf.minimum_list

        if len(inv_list) != len(min_list):
            raise Exception("inv_list and min_list for " + str(pdf.filename) + " are not of the same length.")

        #inv_list and minimum_list should have the same length
        for i in range(len(inv_list)):
            #the column letters have to be offset for the minimum to be one ahead of the column letters for the inventory
            inv_col_let = get_column_letter(col_data_offset)
            min_col_let = get_column_letter(col_data_offset)
            ws[inv_col_let + str(row_data_offset)] = inv_list[i]

            #the row changes to allow each inventory value for a pdf to be printed
            #the += 2 accounts for the fact that the minimum values need to be pasted alongside the inventory values
            min_col_let = get_column_letter(col_data_offset + 1)
            try:
                ws[min_col_let + str(row_data_offset)] = int(min_list[i])
            except:
                ws[min_col_let + str(row_data_offset)] = ""
            col_data_offset += 2
        
        #the column changes once each pdf has their values pasted
        row_data_offset += 1



#this is still in the testing phases, as specifying data for use in the graphs has poor documentation for openpyxl
def excel_graph_maker(wb, ws):

    #starts at the eighth column, or H
    graph_col_offset = 8
    graph_row_offset = 0
    inv_min_offset = 0
    #dateaxis might be an import from openpyxl.chart.axis
    #I think with dateaxis your 'dates' cells have to be in a formatting it can turn into a true date
    #this would format and scale better as Excel is treating the axis as one with special date properties.

    #PDF.length - 1 is equivalent to the amount of different media types
    for i in range(1, PDF.pdf_length):
        line_chart = LineChart()
        #takes the title from the list of media types given as a class attribute of PDF, dynamically
        line_chart.title = PDF.pdf_media_type_list[i - 1]
        line_chart.x_axis.title = "Date"
        line_chart.y_axis.title = "Inventory"

        #experiment with what this 'crossAx' thing is...
        #line_chart.x_axis = DateAxis(crossAx=100)

        #this custom number format could break things
        #line_chart.x_axis.number_format = "yy-mm-dd"

        #line_chart.x_axis.majorTimeUnit = "days"


        #data
        #added one more column to max_col to see if it would include the minimum values as its own series
        y_values = Reference(ws, min_col = 8 + i + inv_min_offset, min_row = 2, max_col = 9 + i + inv_min_offset, max_row = PDF.pdf_id + 1)
        inv_min_offset += 1

        #categories
        x_values = Reference(ws, min_col = 8, min_row = 3, max_col = 8, max_row = PDF.pdf_id + 1)

        #this should add the x-values
        line_chart.add_data(y_values, titles_from_data = True)
        line_chart.set_categories(x_values)
        
        #makes each individual chart and offsets the position for the next one
        #the +3 is for the top of the chart, I believe (check)
        graph_anchor = get_column_letter(graph_col_offset) + str(PDF.pdf_id + 3 + graph_row_offset)

        ws.add_chart(line_chart, graph_anchor)

        graph_metadata_adder(ws, graph_anchor, i - 1)

        #the +14 is to get the data below the graph
        ws[get_column_letter(graph_col_offset) + str(PDF.pdf_id + 3 + graph_row_offset + 14)] = "test"


        graph_col_offset += 10
        #starts drawing graphs lower on the screen instead of more horizontally
        if graph_col_offset > 28:
            graph_col_offset = 8
            #the '16' should be the height of each graph
            graph_row_offset += 17


    #final graph of all the data
    final_line_chart = LineChart()
    final_line_chart.title = "Total Inventory"
    final_line_chart.x_axis.title = "Date"
    final_line_chart.y_axis.title = "Inventory"

    y_values = Reference(ws, min_col = 9, min_row = 2, max_col = 9 + (PDF.pdf_id * 2), max_row = PDF.pdf_id + 1)
    x_values = Reference(ws, min_col = 8, min_row = 3, max_col = 8, max_row = PDF.pdf_id + 1)
    final_line_chart.add_data(y_values, titles_from_data = True)
    final_line_chart.set_categories(x_values)
    ws.add_chart(final_line_chart, get_column_letter(graph_col_offset) + str(PDF.pdf_id + 3 + graph_row_offset))



#pdf_monthly_inv_ratio is a list of lists, where each list has the averages for every media type for a given month
#pdf_monthly_inv_ratio[0] is January's averages, [2] is March, etc.


#adding the x_shift makes the most sense here since the number code can be returned with the addition or subtraction
def let_to_base_26(letters, x_shift = 0):
    #this is neither the prettiest, nor fastest, implementation. But it is one still
    let_list = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    high_exp = len(letters) - 1

    base_26_num = 0
    for let in letters.upper():
        mult_val = let_list.index(let)
        base_26_num += mult_val * (26 ** high_exp)
        high_exp -= 1

    #check if this works with negative x_shift values
    base_26_num += x_shift

    return base_26_num



def base_26_to_let(base_26_num):
    quotient_int = base_26_num
    code_list = []
    loop_logic = True
    let_list = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    while loop_logic:
        quotient = quotient_int / 26
        quotient_int = math.floor(quotient)
        #float stuff might cause issues here
        remainder = round((quotient - quotient_int) * 26)
        code_list.insert(0, remainder)

        if quotient_int < 26:
            code_list.insert(0, quotient_int)
            loop_logic = False

    letters = ""
    for code in code_list:
        letters += let_list[code]

    return letters

#returns a cell shifted up, down, left, or right, according to the shift parameters
#+3 in x_shift moves the column 3 to the right, while -2 in y_shift moves the row 2 down.
def excel_cell_shifter(cell, x_shift = 0, y_shift = 0):

    #determines where the numbers start in the cell code, considers cases like "A2" as well as "AA2", or even "AAAAA2"
    num_start_ind = -1
    letters = ""
    for char in cell:
        num_start_ind += 1
        try:
            int(char)
            break

        except:
            letters += char
            continue
    
    numbers = int(cell[num_start_ind:])

    if y_shift != 0:
        numbers = numbers + y_shift

    #runs the functions needed to advance the letters
    if x_shift != 0:
        #shifting letters is not easy, but entirely feasible by converting to a base 10 number code, assuming letters are a base 26 system
        letters = base_26_to_let(let_to_base_26(letters, x_shift))

    return letters + str(numbers)




#factoring in graph_length and width is a bit too hard since Excel doesn't by default make graphs full cells in width and length
def graph_metadata_adder(ws, graph_anchor, media_type_indice, graph_length=None, graph_width=None):
    #now, how to take the graph anchor, displace by things like graph_length or graph_width
    #test

    ZAD = base_26_to_let(let_to_base_26("A", 16903))
    print(ZAD)

    metadata_start_point = excel_cell_shifter(graph_anchor, y_shift = 14)
    should_be_same = graph_anchor[0] + str(int(graph_anchor[1:]) + 14)
    print("are these the same?")

#should only iterate for as many months there are
    for i in range(0, len(PDF.pdf_month_list)):
        #so each time we want to iterate 
        cur_month_ratio = PDF.pdf_monthly_inv_ratios[i][media_type_indice]

        #testing variables
        #NOW TO CHANGE THIS USING OUR NEW FUNCTIONS
        inv_min_loc = ws[metadata_start_point[0] + str(int(graph_anchor[1:]) + 1)]
        mon_rat_loc = ws[metadata_start_point[0] + str(int(graph_anchor[1:]) + 2)]
        print("hmm")
        ws[metadata_start_point[0] + str(int(graph_anchor[1:]) + 1)] = "Inv/Min"
        ws[metadata_start_point[0] + str(int(graph_anchor[1:]) + 2)] = PDF.pdf_monthly_inv_ratios[i][media_type_indice]

        if PDF.pdf_month_list[i] != "January":
            try:
                ws[metadata_start_point[0] + str(int(graph_anchor[1:]) + 3)] = round(((cur_month_ratio / pre_month_ratio) * 100), 2)
            except:
                #test
                ws[metadata_start_point[0] + str(int(graph_anchor[1:]) + 3)] = "NA"

        pre_month_ratio = PDF.pdf_monthly_inv_ratios[i][media_type_indice]



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
