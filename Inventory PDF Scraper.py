from PyPDF2 import PdfReader
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference



#special case list 1 and 2 form the bulk of the cases where smart_splitter() should add the read_ahead to a given phrase
special_case_list1 = ["o", "h", "%"]

special_case_list2 = ["pe", "ys"]


#changes the directory to where inventory pdfs are stored
def directory_changer():
    print("current path:")
    print(os.getcwd())
    os.chdir(input("Where are the inventory files stored?\n"))
    #os.chdir(r"C:\Users\Mason\Documents\NAMSA Test\Inventory")



#helper function for smart_splitter() that gives it the read_ahead to add to phrase to make conjoined_phrase. This allows for Excel cells to have meaningful 'categories' in cells as opposed to the raw data all being in separate cells, which would ruin formatting
def read_ahead(list, r_i, read=None):
    read = ""

    #r_i is used in place of i because it needs to progress through the list of scraped data independently from it
    r_i += 1

    #continues as long as the read index, or r_i, is less than the length of the 'list', (now just text from the scraped PDFs)
    while r_i < len(list):
        
        #creates the initial read without reading ahead for special cases that should be added onto the read by smart_splitter()
        if list[r_i] != " ":
            read += list[r_i]

            r_i += 1

        #reading ahead is done whenever a space is encountered, since it could indicate a meaningful addition to phrase
        elif list[r_i] == " ":
            if (
                #covers the two special case categories where read_ahead should be added to phrase
                read[-1] in special_case_list1
                or read[-2:] in special_case_list2

                #adds 'trays (x)' to the list of special cases (but wouldn't special_case_list2's 'ys' cover this? Is this not some other exception?)
                or len(read) > 4
                and read[-5] == "s"

                #adds Saline P80 to the list of special cases, since having '80' be a special case would cause issues
                or len(read) == 3
                and read == "P80"
                or (len(read) == 3 and read[-1] == ")")
                or (len(read) == 4 and read[-1] == ")")
            ):
                
                #recursively calls read_ahead() so long as there is a special case in front of read
                future_read = read_ahead(list, r_i, read)

                #each call will get progressively longer so long as special cases are in front of read
                if future_read != None:
                    future_read = read + " " + future_read
                    return future_read
                    
                #cuts off the read at the end of the scraped data for this line (check) 
                else:
                    return read
            #returns None if a special read is not seen 
            else:
                return None
                
    #this might not have to be written explicitly
    return None



#each pdf that is scraped is a list containing each line of the pdf, which is also a list, hence the input parameter
def smart_splitter(list_of_lists):
    final_lists = []

    for list in list_of_lists:
        split_list = []
        #'phrase' is used as the word to designate what is read, as it could be multiple words long
        phrase = ""
        delay = 0

        #each list at this point is a line of text as one whole string, so this iterates over every character of that line string
        for i in range(len(list)):

            #causes i to not iterate over what's already been read by read_ahead()
            if delay > 0:
                delay -= 1

                #resets phrase when reading ahead has been done
                phrase = ""

            elif list[i] != " ":
                phrase += list[i]

            #read_ahead() is called when a space is encountered, as it may designate a phrase made up of multiple words
            elif list[i] == " ":
                text_ahead = read_ahead(list, i)

                #adds the text ahead of a given word, should that text be a special case. read_ahead() is recursively called so long as special cases are encountered, so the conjoined_phrase is as long as possible
                if text_ahead != None:
                    conjoined_phrase = phrase + " " + text_ahead

                    split_list.append(conjoined_phrase)

                    #delays i beyond the full extent of what's already been read
                    delay = len(text_ahead) + 1

                #simply adds phrase to split_list, should no special cases be found ahead of it
                else:
                    split_list.append(phrase)
                    phrase = ""

        #this line accounts for the final phrase, if it wasn't already part of a special case
        split_list.append(phrase)

        #adds each list to the final_lists, which is a list containing each line list with its multi-word phrases
        final_lists.append(split_list)

    return final_lists



#will add every media type to each pdf's Excel representation, even if none were recorded for the date. Takes from the data processed by smart_splitter(). Only adds 1000 FTM and DFD because too much more would vastly overcomplicate this. These are the only types that are missing in 2023
def data_formatter(data_list):
    
    #debug tool
    check_against_formatted_list = data_list.copy()
    #a shallow copy was needed to prevent the loop running indefinitely
    formatted_list = data_list.copy()

    #number of pdfs examined
    doc_num = 0

    for pdf_data in data_list:

        #knowing where 1000 FTM and 1000 DFD should be on an ideally formatted pdf is what is used to reference where they should be added. Sublist_count tracks what line of the pdf we should add to on the Excel representation
        sublist_count = 0
        sublist_mod = 0

        FTM_400S_Needed = False
        FTM_1000_Needed = False
        DFD_1000_Needed = False
        REMOVE_OD = False

        #pdf_data is indexed at [1] because [0] contains the filename, while [1] is a list of each line, which is also a list of all the functional phrases
        for sub_list in pdf_data[1]:
            sublist_count += 1

            #adds 400-S FTM
            if sublist_count == 12 - sublist_mod:
                if sub_list[0] == "400-S":
                    sublist_count += 1
                
                else:
                    FTM_400S_Needed = True
                    sublist_count += 1
                    sublist_mod += 1


            elif sublist_count == 15 - sublist_mod:
                if sub_list[0] == "1000":
                    sublist_count += 1
                    
                else:
                    FTM_1000_Needed = True
                    sublist_count += 1
                    sublist_mod += 1


            
            #adds 1000 DFD
            #was 27
            elif sublist_count == 28 - sublist_mod:
                #the or condition is given because the minimum for DFD was changed on 06/26/23 from 136 to 99
                if sub_list[3] == "136" or sub_list[3] == "99":
                    sublist_count += 1

                else:
                    DFD_1000_Needed = True
                    sublist_count += 1
                    sublist_mod == 1
            
            #the large degree of ambiguity on the sublist_count comes from the fact that with FTM1000 or DFD1000 the end of the PDF could come at different times. Some of this ambiguity could be needed in the DFD1000 check...
            #CHANDED THIS TO SUBLIST_MOD
            elif sublist_count == 34 - sublist_mod:
                if sub_list[0] == "OD=":
                    REMOVE_OD = True
        

        #runs after all the other data has been added, so that the .insert() method knows where to add 1000 FTM and 1000 DFD based on where they would be in an ideally formatted pdf
        #the final value added, "", indicates that no media for this media type was recorded on this date
        if FTM_400S_Needed:
            formatted_list[doc_num][1].insert(11, ["400-S", "9 trays/lot", "4", "9 trays (2)", ""])

        if FTM_1000_Needed:
            #experiment with removing different indices, or maybe saving which exact indice to add to specific to each unique case instead of hard-coding it
            formatted_list[doc_num][1].insert(13, ["1000", "17-20/lot", "36", "90 (5)", ""])

        if DFD_1000_Needed:
            formatted_list[doc_num][1].insert(25, ["DFD","1000", "34 bottles/lot", "99", "136", ""])

        if REMOVE_OD:
            formatted_list[doc_num][1].pop(-1)

        #THIS USED TO BE AT THE START OF THE FUNCTION THIS IS A TEST
        doc_num += 1

    return formatted_list



def data_scraper():
    all_data = []
    pdf_count = 0

    #looks for pdfs in the current directory given to find media inventory pdfs
    for filename in os.listdir(os.getcwd()):
        if filename[-4:] == ".pdf":
            pdf_count += 1

            #uses imported functions from PyPDF2 to read from pdfs
            reader = PdfReader(open(filename, "rb"))

            #this can be removed in the final version
            info = reader.metadata

            #this is the raw line by line data scraped from each pdf as a list containing each line, which is also a list
            unformatted_data = []

            #every media pdf should be one page long, this could cause issues should that requirement not be met. (UPDATE)
            for i in range(0, len(reader.pages)):
                selected_page = reader.pages[i]

                text = selected_page.extract_text()

                unformatted_data += text.splitlines()

            #data that is split into functional phrases without 1000 FTM and 1000 DFD added to the Excel pdf representation
            spaced_data = smart_splitter(unformatted_data)

            all_data.append([[filename], spaced_data])

    return data_formatter(all_data), pdf_count



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

#update / make large changes to what's below this



def excel_pdf_paster(formatted_data, ws):
    #allows the rows to be placed where they need to be, and allows for space in between pasted pdfs in Excel
    pdf_offset = 1

    #moves the dates where each pdf's data is collected to start at 'I2' and continue rightward
    date_offset = 9

    #iterates through every pdf in formatted_data
    for data_list in formatted_data:
        row_count = 1

        for row in data_list[1]:
            #adds to the rows in Excel without SCDB, FTM, etc.
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

        #adds the filename above each pdf's pasted data. title_cell and date_cell are used for debugging
        title_cell = "A" + str(pdf_offset - 34)
        ws[title_cell] = data_list[0][0]

        date_cell = get_column_letter(date_offset) + "2"
        ws[date_cell] = data_list[0][0][:8]

        if date_cell == "CA":
            continue

        date_offset += 1



def excel_media_type_adder(ws):

    # makes the relevant data cell of the resulting Excel document
    #media_type_list = []

    for i in range(3, 11):
        ws["H" + str(i)] = ws["A3"].value + " " + ws["B" + str(i)].value

        #something something is not subscriptable... (FIX)
        #media_type_list.append[ws["H" + str(i)].value]

    for i in range(11, 19):
        ws["H" + str(i)] = ws["A11"].value + " " + ws["B" + str(i)].value
        
        #(FIX)
        #media_type_list.append[ws["H" + str(i)].value]

    for i in range(19, 33):
        ws["H" + str(i)] = ws["A" + str(i)].value + " " + ws["B" + str(i)].value
        
        #(FIX)
        #media_type_list.append[ws["H" + str(i)].value]



def excel_data_mover(ws, pdf_count):
    gap = 3
    data_offset = 0

    for col in range(9, pdf_count + 9):
        col_char = get_column_letter(col)

        for row in range(3, 33):

            gap += 1

            cell_recieve = col_char + str(row)

            cell_give = "F" + str(data_offset + row)

            #conditions for ignoring weirdly formatted cells
            if ws[cell_give].value == "OD":
                continue

            #None can be replaced with "" to make something an Excel graph might prefer more
            elif ws[cell_give].value == None:
                continue

            else:
                try:
                    ws[cell_recieve] = int(ws[cell_give].value)

                except:
                    ws[cell_recieve] = ws[cell_give].value
        
        data_offset += 34
        gap = data_offset


#this is a test
def excel_graph_maker(wb, ws):
    SCDB100_chart = LineChart()
    SCDB100_chart.title = "SCDB 100 Inventory"
    SCDB100_chart.x_axis.title = "Time"
    SCDB100_chart.y_axis.title = "Inventory"

    #max_col should be taken from the data (UPDATE)

    #this should work but does not
    SCDB100_data = Reference(ws, min_col = 9, min_row = 3, max_col = 82, max_row = 3)
    SCDB100_categories = Reference(ws, min_col = 9, min_row = 2, max_col = 82, max_row = 2)

    #categories should be the x axis with data on the y, but instead categories has to be the series name and data has to be both x and y, which have to be flipped in Excel.
    #this also wouldn't work for media types other than SCDB 100, since their data is separated by gaps

    #this works but has backwards axes
    SCDB100_data = Reference(ws, min_col = 9, min_row = 2, max_col = 82, max_row = 3)
    SCDB100_categories = Reference(ws, min_col = 8, min_row = 2, max_col = 8, max_row = 2)

    SCDB100_chart.add_data(SCDB100_data, titles_from_data = True)
    SCDB100_chart.set_categories(SCDB100_categories)

    ws.add_chart(SCDB100_chart, "H34")

    #test of iteratively adding all the charts
    """
    #32 should be programmatically found as where the last media type row is located (UPDATE)
    for row in range(3, 32):
        #this part has yet to have a function
        cell = "G" + str(row)

        chart = LineChart()

        chart.title = ws[cell].value + " Inventory"

        chart.x_axis.title = "Date"
        chart.y_axis.title = "Inventory"

        chart_data = Reference(ws, min_col = 9, min_row = 2, max_col = 82, max_row = 3)
        chart_categories = Reference(ws, min_col = 8, min_row = 3, max_col = 8, max_row = 3)
    """

    return ":)"



def run_program():
    directory_changer()

    #formatted_data is twice the length it should be (FIX)
    formatted_data, pdf_count = data_scraper()

    workbook, worksheet = excel_wb_maker()

    #maybe reduce what's below this to a macro function
    excel_pdf_paster(formatted_data, worksheet)
    excel_media_type_adder(worksheet)
    excel_data_mover(worksheet, pdf_count)

    #test
    excel_graph_maker(workbook, worksheet)

    workbook.save(worksheet.title)

    print("Excel file complete!")

run_program()
