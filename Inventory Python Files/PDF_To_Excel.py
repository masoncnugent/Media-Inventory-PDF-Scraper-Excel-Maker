from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import Font, PatternFill
import math

from PDF_Class import PDF

#utilizes openpyxl to make an Excel workbook that can hold all the pdf data and graphs
def excel_wb_maker():

    # makes the excel workbook
    wb = Workbook()

    #sets the relevant sheet of the Excel document that will be edited
    ws = wb.active

    #gives the name for the resulting saved Excel file
    ws.title = "Inventory Analytics.xlsx"

    #this sheet stores the graphs made from the pdf data
    #wb.create_sheet("Graphs")

    return wb, ws



#copies each pdfs self.data line by line on the Excel file, with a filename above each one. Includes added media types if they weren't in the original pdf
def excel_pdf_vertical_copier(ws):
    #allows the rows for each pdf to be placed where they need to be, and allows for space in between pasted pdfs in Excel
    pdf_offset = 1

    #moves the dates where each pdf's data is collected to start at 'I2' and continue rightward
    date_offset = 3

    #iterates through every pdf in formatted_data
    for pdf in PDF.pdf_list:
        row_count = 1

        for row in pdf.data:

            col_count = 1

            for row_data in row:
                #gives the column letter and advances it forward for the next cell in a pdf's given row
                char = get_column_letter(col_count)
                col_count += 1

                #this line is for debug purposes, but the one below it takes the data from a row on the formatted pdf data and puts it in Excel.
                row_cell = char + str(row_count + pdf_offset)

                #this is so the data in Excel is not inputted as a string if it can be an integer, solely to remove the green triangles in Excel.
                try:
                    ws[row_cell] = int(row_data)
                
                except:
                    ws[row_cell] = row_data

            row_count += 1

        pdf_offset += row_count + 2

        #adds the filename above each pdf's pasted data.
        title_cell = "A" + str(pdf_offset - 34)
        ws[title_cell] = pdf.filename

        #adds the dates that all the inventory will be compared against
        date_cell = "I" + str(date_offset)
        ws[date_cell] = pdf.filename[:8]

        if date_cell == "CA":
            continue

        date_offset += 1
    
    #adds the word "Date" above its respective column
    ws["I2"] = "Date"
    ws["I2"].font = Font(bold = True)



#works from the first pdf, assuming it was formatted correctly and identically to all other pdfs
#frameshift issues occur when this assumption is not met
def excel_media_type_adder(ws):

    #start at I2 and move to J2, K2, etc.

    #the first two ranges handle the cases where SCDB and FTM have to be added to the media type titles
    #there is an assumption that nothing new will be added into SCDB and FTM, built into the lengths of each section being hardcoded
    #this allows for the minimum columns to exist alongside the data for each media type
    min_offset = 0
    for i in range(3, 11):
        ws[excel_cell_shifter("G2", x_shift = i + min_offset)] = ws["A3"].value + " " + str(ws["B" + str(i)].value)
        PDF.pdf_media_type_list.append(ws["A3"].value + " " + str(ws["B" + str(i)].value))
        min_offset += 1

    for i in range(11, 19):
        ws[excel_cell_shifter("G2", x_shift = i + min_offset)] = ws["A11"].value + " " + str(ws["B" + str(i)].value)
        PDF.pdf_media_type_list.append(ws["A11"].value + " " + str(ws["B" + str(i)].value))
        min_offset += 1

    #the end is dynamically encoded to match the length of each PDF
    for i in range(19, PDF.pdf_length + 2):
        ws[excel_cell_shifter("G2", x_shift = i + min_offset)] = ws["A" + str(i)].value + " " + str(ws["B" + str(i)].value)
        PDF.pdf_media_type_list.append(ws["A" + str(i)].value + " " + str(ws["B" + str(i)].value))
        min_offset += 1

    #adds the words "Media Type" and "Minimum" above their respective column
    for i in range(len(PDF.pdf_media_type_list)):
        inv_title_col_let = excel_cell_shifter("A1", x_shift = 9 + (2 * i))
        med_title_col_let = excel_cell_shifter("A1", x_shift = 10 + (2 * i), y_shift = 1)

        ws[inv_title_col_let] = "Media Type / Inv"
        ws[med_title_col_let] = "Minimum"
        ws[inv_title_col_let].font = Font(bold = True)
        ws[med_title_col_let].font = Font(bold = True)



#this can be re-written using pdf.inventory_list, which would greatly reduce errors and enhance readability
def excel_pdf_inventory_copier(ws):

    #this is where the data starts to be printed
    row_data_offset = 3

    for pdf in PDF.pdf_list:
        col_data_offset = 10

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



def excel_pdf_lots_over_min_copier(ws):

    #determines the start_point for all the other data
    lots_over_min_start_point = "I" + str(PDF.pdf_id + 3)

    month_header_loc = excel_cell_shifter(lots_over_min_start_point, y_shift = 1)
    ws[month_header_loc] = "Month"
    ws[month_header_loc].font = Font(bold = True)

    media_offset = 1
    for media_type in PDF.pdf_media_type_list:
        #prints "Media Type"
        media_type_title_cell = excel_cell_shifter(lots_over_min_start_point, x_shift = media_offset)
        ws[media_type_title_cell] = "Media Type"
        ws[media_type_title_cell].font = Font(bold = True)

        media_offset += 2
        #prints the actual media type
        media_type_proper_cell = excel_cell_shifter(media_type_title_cell, y_shift = 1)
        ws[media_type_proper_cell] = media_type


    month_offset = 1
    for month in PDF.pdf_month_list:
        #prints the months so they move down below the "Month" header
        cur_month_loc = excel_cell_shifter(month_header_loc, y_shift = month_offset)
        ws[cur_month_loc] = month
        month_offset += 1
    

        #test
        month_row_offset = 0
    for inv_over_min_month in PDF.pdf_lots_over_min:
        month_col_offset = 0
        
        for i in range(0, len(PDF.pdf_media_type_list)):

            inv_over_min_cell = excel_cell_shifter(lots_over_min_start_point, x_shift = 1 + month_col_offset, y_shift = 2 + month_row_offset)

            ws[inv_over_min_cell] = inv_over_min_month[i]

            month_col_offset += 2
        
        month_row_offset += 1
    
    #testing purposes
    #the y_shift is compensatory
    average_cell_header = excel_cell_shifter(inv_over_min_cell, x_shift = 2, y_shift = -6)
    ws[average_cell_header] = "Averages"
    ws[average_cell_header].font = Font(bold = True)

    for i in range(0, len(PDF.pdf_month_list)):
        average_data_cell = excel_cell_shifter(average_cell_header, y_shift = i + 1)
        ws[average_data_cell] = PDF.pdf_lots_over_min_averages[i]






#this is still in the testing phases, as specifying data for use in the graphs has poor documentation for openpyxl
def excel_graph_maker(ws):

    #starts at the eighth column, or H
    graph_col_offset = 8
    graph_row_offset = 0
    inv_min_offset = 0

    #where the min rows for the y values start in line_chart_lots_over_min
    min_row_y_start = PDF.pdf_id + 4


    #PDF.length - 1 is equivalent to the amount of different media types

    for i in range(1, PDF.pdf_length):
        line_chart_inv_min = LineChart()

        line_chart_lots_over_min = LineChart()
        line_chart_lots_over_min.title = PDF.pdf_media_type_list[i - 1] + " Lots Above Min"
        line_chart_lots_over_min.x_axis.title = "Month"
        line_chart_lots_over_min.y_axis.title = "Lots Above Min"

        lots_over_min_y_values = Reference(ws, min_col = 8 + (2 * i), min_row = min_row_y_start, max_col = 8 + (2 * i), max_row = min_row_y_start + len(PDF.pdf_month_list))

        lots_over_min_x_values = Reference(ws, min_col = 9, min_row = min_row_y_start + 1, max_col = 9, max_row = min_row_y_start + len(PDF.pdf_month_list))


        line_chart_lots_over_min.add_data(lots_over_min_y_values, titles_from_data = True)
        line_chart_lots_over_min.set_categories(lots_over_min_x_values)


        #test the 33 part
        graph_lots_over_min_anchor = excel_cell_shifter("A1", x_shift = graph_col_offset + 32, y_shift = PDF.pdf_id + len(PDF.pdf_month_list) + 5 + graph_row_offset)


        #scaling test
        #line_chart_lots_over_min.y_axis.scaling.min = -3
        #line_chart_lots_over_min.y_axis.scaling.max = 3

        ws.add_chart(line_chart_lots_over_min, graph_lots_over_min_anchor)


        graph_metadata_adder(ws, graph_lots_over_min_anchor, i - 1)


        







        #takes the title from the list of media types given as a class attribute of PDF, dynamically
        line_chart_inv_min.title = PDF.pdf_media_type_list[i - 1] + " Inventory"
        line_chart_inv_min.x_axis.title = "Date"
        line_chart_inv_min.y_axis.title = "Inventory"


        #added one more column to max_col to see if it would include the minimum values as its own series
        inv_min_y_values = Reference(ws, min_col = 8 + (2 * i), min_row = 2, max_col = 9 + (2 * i), max_row = PDF.pdf_id + 1)

        #categories
        inv_min_x_values = Reference(ws, min_col = 9, min_row = 3, max_col = 9, max_row = PDF.pdf_id + 1)

        #this should add the x-values
        line_chart_inv_min.add_data(inv_min_y_values, titles_from_data = True)
        line_chart_inv_min.set_categories(inv_min_x_values)
        
        #makes each individual chart and offsets the position for the next one
        graph_inv_min_anchor = excel_cell_shifter("A1", x_shift = graph_col_offset, y_shift = PDF.pdf_id + len(PDF.pdf_month_list) + 5 + graph_row_offset)

        ws.add_chart(line_chart_inv_min, graph_inv_min_anchor)

        #utilizes new functions in progress
        #graph_metadata_adder(ws, graph_inv_min_anchor, i - 1)


        #starts drawing graphs lower on the screen instead of more horizontally
        graph_col_offset += 10
        if graph_col_offset > 28:
            graph_col_offset = 8

            #the '17' should be the height of each graph
            #graphs should be offset more or less depending on how much metadata is beneath them
            if len(PDF.pdf_month_list) > 9: 
                graph_row_offset += 23
            else:
                graph_row_offset += 20

    line_chart_total_lots_over_min = LineChart()
    line_chart_total_lots_over_min.title = "Total Inventory Lots Over Min"
    line_chart_total_lots_over_min.x_axis.title = "Month"
    line_chart_total_lots_over_min.y_axis.title = "Lots Above Min"

    #test to add the second lots_over_min charts
    #play around with the range
    valid_y_data = [col_idx for col_idx in range(10, 10 + ((PDF.pdf_length) * 2), 2) if ws.cell(row = min_row_y_start, column = col_idx).value]
    print(valid_y_data)
    y_references = [Reference(ws, min_col=col_idx, max_col=col_idx, min_row = min_row_y_start, max_row = min_row_y_start + len(PDF.pdf_month_list)) for col_idx in valid_y_data]
    print(y_references)

    for y_ref in y_references:
        line_chart_total_lots_over_min.add_data(y_ref, titles_from_data = True)

    #total_lots_over_min_y_values = Reference(ws, min_col = 10, min_row = PDF.pdf_id + 4, max_col = 10 + inv_min_offset + len(PDF.pdf_media_type_list), max_row = PDF.pdf_id + len(PDF.pdf_month_list) + 4)

    total_lots_over_min_x_values = Reference(ws, min_col = 9, min_row = PDF.pdf_id + 5, max_col = 9, max_row = PDF.pdf_id + len(PDF.pdf_month_list) + 4)

    #line_chart_total_lots_over_min.add_data(total_lots_over_min_y_values, titles_from_data = True)
    line_chart_total_lots_over_min.set_categories(total_lots_over_min_x_values)


    graph_total_lots_over_min_anchor = excel_cell_shifter("A1", x_shift = graph_col_offset + 32, y_shift = PDF.pdf_id + len(PDF.pdf_month_list) + 5 + graph_row_offset)


    #scaling test
    #line_chart_lots_over_min.y_axis.scaling.min = -3
    #line_chart_lots_over_min.y_axis.scaling.max = 3

    ws.add_chart(line_chart_total_lots_over_min, graph_total_lots_over_min_anchor)


#pdf_monthly_inv_ratio is a list of lists, where each list has the averages for every media type for a given month
#pdf_monthly_inv_ratio[0] is January's averages, [2] is March, etc.

#adding the x_shift makes the most sense here since the number code can be returned with the addition or subtraction
def let_to_base_26(letters, x_shift = 0):
    offset = 64
    base_26_num = 0

    for i in range(len(letters)):
        base_26_num += (ord(letters[len(letters) - i - 1]) - offset) * (26 ** i)

    return base_26_num + x_shift



def base_26_to_let(base_26_num):
    offset = 64
    quotient = base_26_num
    letters = ""
    while quotient > 0:
        quotient, remainder = divmod(quotient - 1, 26)

        letters = chr(remainder + offset + 1) + letters

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
    #tested for negative numbers
    if x_shift != 0:
        #shifting letters is not easy, but entirely feasible by converting to a base 10 number code, assuming letters are a base 26 system
        letters = base_26_to_let(let_to_base_26(letters, x_shift))

    return letters + str(numbers)




#factoring in graph_length and width is a bit too hard since Excel doesn't by default make graphs full cells in width and length
def graph_metadata_adder(ws, graph_anchor, media_type_indice, graph_length=None, graph_width=None):
    #now, how to take the graph anchor, displace by things like graph_length or 
#temp variable names
    y_val = 0
    x_val = 0

    ratio_start_point = "H" + str(PDF.pdf_id + 3)
    ratio_row_offset = 0
    ratio_col_offset = 0

    #metadata starts below each graph, and from the top left anchor of a given graph the first available cell below the graph is 13 cells down
    metadata_start_point = excel_cell_shifter(graph_anchor, x_shift = y_val, y_shift = 14 + x_val)

#should only iterate for as many months there are
    for i in range(0, len(PDF.pdf_month_list)):
        #so each time we want to iterate 
        #changed to lots / min
        cur_month_ratio = PDF.pdf_lots_over_min[i][media_type_indice]

        #test to make graphs with all the ratios
        ratio_graph_loc = excel_cell_shifter(ratio_start_point, x_shift=ratio_col_offset, y_shift=ratio_row_offset)
        ws[ratio_graph_loc] = cur_month_ratio
        ratio_row_offset += 1



        #these use excel_cell_shifter() to determine where to be relative to the start of where metadata should be printed below each graph
        mon_loc = excel_cell_shifter(metadata_start_point, x_shift = x_val, y_shift = y_val)


        inv_min_loc = excel_cell_shifter(metadata_start_point, x_shift = x_val, y_shift = 1 + y_val)
        mon_rat_loc = excel_cell_shifter(metadata_start_point, x_shift = x_val, y_shift = 2 + y_val)

        ws[mon_loc] = PDF.pdf_month_list[i]
        ws[mon_loc].font = Font(bold = True)
        ws[inv_min_loc] = "Lots / Min"


        #this should make the percentage differences not have rounding errors
        try:
            ws[mon_rat_loc] = round(PDF.pdf_lots_over_min[i][media_type_indice], 2)
        except:
            ws[mon_rat_loc] = "NA"

        if PDF.pdf_month_list[i] != "January":
            #the cell name for where the percentage change against last month will be mapped to
            ratio_dif_loc = excel_cell_shifter(metadata_start_point, x_shift = x_val, y_shift = 3 + y_val)
            try:
                lot_num_change = round((pre_month_ratio - cur_month_ratio) * -1, 2)


                ws[ratio_dif_loc] = lot_num_change

                red_color = "FF8888"
                green_color = "88FF88"
                blue_color = "8888FF"

                #WAIT SUPER COOL IDEA, turn the lot_num_change into a gradient color 

                if lot_num_change < 0:
                    ws[ratio_dif_loc].fill = PatternFill(start_color = red_color, end_color = red_color, fill_type = "lightGray")

                elif lot_num_change > 0:
                    ws[ratio_dif_loc].fill = PatternFill(start_color = green_color, end_color = green_color, fill_type = "lightGray")

                else:
                    ws[ratio_dif_loc].fill = PatternFill(start_color = blue_color, end_color = blue_color, fill_type = "lightGray")

            except:
                #this would occur if one of the month ratio's involves no media recorded
                ws[ratio_dif_loc] = "NA"           

        #cur_month_ratio will be ahead of pre_month_ratio since this is declared here
        pre_month_ratio = PDF.pdf_lots_over_min[i][media_type_indice]

        x_val += 1
        #should allow for the metadata to not stack up
        if x_val == 9:
            y_val += 4
            x_val = 0



#runs all excel related functions. data_scraper() has to be run first to create the PDF class with all it's PDF objects
def excel_batch_processor():

    wb, ws = excel_wb_maker()

    excel_pdf_vertical_copier(ws)

    excel_media_type_adder(ws)

    excel_pdf_inventory_copier(ws)

    excel_pdf_lots_over_min_copier(ws)

    #could switch worksheets for this by this point

    #adding wb to help with debugging
    excel_graph_maker(ws)

    wb.save(ws.title)

    print("Excel file complete!")

    #figure out how to remove the pdf files from the github repos
