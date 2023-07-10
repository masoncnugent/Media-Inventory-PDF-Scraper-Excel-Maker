from PyPDF2 import PdfReader
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


print("current path:")
print(os.getcwd())
# os.chdir(input("Where are the inventory files stored?"))
os.chdir(r"C:\Users\Mason\Documents\NAMSA Test\Inventory")
print("path should be changed:")
print(os.getcwd())

#special case list 1 and 2 form the bulk of the cases where smart_splitter() should add the read_ahead to a given phrase
special_case_list1 = ["o", "h", "%"]

special_case_list2 = ["pe", "ys"]

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

                #adds Saline P80 to the list of special cases, since having '80' be a special case would case issues
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
    
    #a shallow copy was needed to prevent the loop running indefinitely
    formatted_list = data_list.copy()

    #number of pdfs examined
    doc_num = 0

    for pdf_data in data_list:
        doc_num += 1

        #knowing where 1000 FTM and 1000 DFD should be on an ideally formatted pdf is what is used to reference where they should be added. Sublist_count tracks what line of the pdf we should add to on the Excel representation
        sublist_count = 0

        FTM_1000_Needed = False
        DFD_1000_Needed = False

        #pdf_data is indexed at [1] because [0] contains the filename, while [1] is a list of each line, which is also a list of all the functional phrases
        for sub_list in pdf_data[1]:
            sublist_count += 1

            #adds 1000 FTM
            if sublist_count == 14:
                if sub_list[0] == "1000":
                    sublist_count += 1
                    
                else:
                    FTM_1000_Needed = True
                    print("FTM1000 needed for " + str(pdf_data[0]))
                    sublist_count += 1
            
            #adds 1000 DFD
            elif sublist_count == 27:
                if sub_list[4] == "136":
                    sublist_count += 1

                else:
                    DFD_1000_Needed = True
                    print("DFD1000 needed for " + str(pdf_data[0]))
                    sublist_count += 1


        formatted_list.append(pdf_data)

        #runs after all the other data has been added, so that the .insert() method knows where to add 1000 FTM and 1000 DFD based on where they would be in an ideally formatted pdf
        #the final value added, "", indicates that no media for this media type was recorded on this date
        if FTM_1000_Needed:
            formatted_list[doc_num][1].insert(12, ["1000", "17-20/lot", "36", "90 (5)", ""])

        if DFD_1000_Needed:

            formatted_list[doc_num][1].insert(25, ["DFD","1000", "34 bottles/lot", "99", "136", ""])

    return formatted_list

#this should be put in a function
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


formatted_data = data_formatter(all_data)


# makes the excel workbook
wb = Workbook()

#sets the relevant sheet of the Excel document that will be edited
ws = wb.active

#gives the name for the resulting saved Excel file
ws.title = "Inventory Analytics.xlsx"

#this sheet stores the graphs made from the pdf data
wb.create_sheet("Graphs")


#BELOW THIS SHOULD BE MADE NOT DISPICABLE
row_offset = 1

temp = 2

for data_list in all_data:
    row_count = 1

    for row in data_list[1]:
        # adds to the rows without SCDB, FTM, etc.
        if len(row) < 6:
            row.insert(0, "")

        col_count = 1

        for cell_data in row:
            char = get_column_letter(col_count)
            col_count += 1

            ws[char + str(row_count + row_offset)] = cell_data

        row_count += 1

    row_offset += row_count + 2

    ws["A" + str(row_offset - 33)] = data_list[0][0]

    ws[get_column_letter(temp + 6) + "2"] = data_list[0][0][:8]

    temp += 1


# makes the relevant data cell of the resulting Excel document
#media_type_list = []

for i in range(3, 11):
    ws["H" + str(i)] = ws["A3"].value + " " + ws["B" + str(i)].value

    #something something is not subscriptable... (FIX)
    #media_type_list.append[ws["H" + str(i)].value]

for i in range(11, 19):
    ws["H" + str(i)] = ws["A11"].value + " " + ws["B" + str(i)].value
    
    (FIX)
    #media_type_list.append[ws["H" + str(i)].value]

for i in range(19, 33):
    ws["H" + str(i)] = ws["A" + str(i)].value + " " + ws["B" + str(i)].value
    
    (FIX)
    #media_type_list.append[ws["H" + str(i)].value]


#old implementation (UPDATE)
amogus = 3
data_offset = 0


for col in range(9, pdf_count + 9):
    col_char = get_column_letter(col)


    for row in range(3, 32):

        print("worked " + str(amogus))
        amogus += 1

        print("row " + str(row))

        print("data_offset + row " + str(data_offset + row))

        cell_recieve = col_char + str(row)

        cell_give = "F" + str(data_offset + row)

        #conditions for ignoring weirdly formatted cells
        if ws[cell_give].value == "OD":
            print("code run!")
            continue

        elif ws[cell_give].value == None:
            continue

        else:
            try:
                ws[cell_recieve] = int(ws[cell_give].value)
            except:

                ws[cell_recieve] = ws[cell_give].value
    
    data_offset += 33 + 3
    amogus = data_offset

wb.save(ws.title)
