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


special_case_list1 = ["o", "h", "%"]

special_case_list2 = ["pe", "ys"]


def read_ahead(list, r_i, read=None):
    read = ""
    r_i += 1

    while r_i < len(list):
        if list[r_i] != " ":
            read += list[r_i]

            r_i += 1

        elif list[r_i] == " ":
            if (
                read[-1] in special_case_list1
                or read[-2:] in special_case_list2
                or len(read) > 4
                and read[-5] == "s"
                or len(read) == 3
                and read == "P80"
                or (len(read) == 3 and read[-1] == ")")
                or (len(read) == 4 and read[-1] == ")")
            ):
                future_read = read_ahead(list, r_i, read)

                if future_read != None:
                    future_read = read + " " + future_read
                    return future_read

                else:
                    return read

            else:
                return None

    return None


def smart_splitter(list_of_lists):
    final_lists = []

    for list in list_of_lists:
        split_list = []
        phrase = ""
        delay = 0

        for i in range(len(list)):

            # causes i to not iterate over what's already been read by read_ahead
            if delay > 0:
                delay -= 1

                phrase = ""

            elif list[i] != " ":
                phrase += list[i]

            elif list[i] == " ":
                text_ahead = read_ahead(list, i)

                if text_ahead != None:
                    conjoined_phrase = phrase + " " + text_ahead

                    split_list.append(conjoined_phrase)

                    delay = len(text_ahead) + 1

                else:
                    split_list.append(phrase)
                    phrase = ""

        split_list.append(phrase)
        final_lists.append(split_list)

    return final_lists



#will add every media type to each pdf, only considering adding 1000 FTM and DFD because too much more would vastly overcomplicate this
def data_formatter(data_list):
    formatted_list = data_list.copy()
    doc_num = 0

    for pdf_data in data_list:
        doc_num += 1

        print(doc_num)

        #print(pdf_data)

        sublist_count = 0

        FTM_1000_Needed = False
        DFD_1000_Needed = False

        for sub_list in pdf_data[1]:
            #print(sub_list)
            #print("")
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

                    print(pdf_data)
                    print("")

                    print("DFD1000 needed for " + str(pdf_data[0]))


                    sublist_count += 1


        formatted_list.append(pdf_data)

        #print("dfsfdsfsdf")
        #print(formatted_list)
        #print(" ")

        if FTM_1000_Needed:
            formatted_list[doc_num][1].insert(12, ["1000", "17-20/lot", "36", "90 (5)", ""])

            #print(formatted_list)[doc_num]


        if DFD_1000_Needed:
            #print(len(formatted_list[doc_num][1]))

            #print("uu")

            #print(formatted_list[doc_num])

            #print("xx")

            formatted_list[doc_num][1].insert(25, ["DFD","1000", "34 bottles/lot", "99", "136", ""])

            #print(formatted_list[doc_num])

            #print(len(formatted_list[doc_num][1]))

    for sample in formatted_list:
        print(sample)
        print("")

    return formatted_list

#need to iterate one more list lol


all_data = []
pdf_count = 0

for filename in os.listdir(os.getcwd()):
    if filename[-4:] == ".pdf":
        pdf_count += 1

        reader = PdfReader(open(filename, "rb"))
        info = reader.metadata

        unformatted_data = []

        for i in range(0, len(reader.pages)):
            selected_page = reader.pages[i]

            text = selected_page.extract_text()

            unformatted_data += text.splitlines()

        spaced_data = smart_splitter(unformatted_data)

        all_data.append([[filename], spaced_data])


#print(all_data)

#test

for smars in all_data:
    print(smars)
    print("")

print(all_data)

formatted_data = data_formatter(all_data)

#for test in formatted_data:
    #print(test)
    #print("")


# makes the excel workbook
wb = Workbook()

ws = wb.active
ws.title = "Inventory Analytics.xlsx"

wb.create_sheet("Graphs")

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








# makes the relevant data cells
#media_type_list = []

for i in range(3, 11):
    ws["H" + str(i)] = ws["A3"].value + " " + ws["B" + str(i)].value

    #something something is not subscriptable...
    #media_type_list.append[ws["H" + str(i)].value]

for i in range(11, 19):
    ws["H" + str(i)] = ws["A11"].value + " " + ws["B" + str(i)].value
    #media_type_list.append[ws["H" + str(i)].value]

for i in range(19, 33):
    ws["H" + str(i)] = ws["A" + str(i)].value + " " + ws["B" + str(i)].value
    #media_type_list.append[ws["H" + str(i)].value]

ws["H2"] = ws["A2"].value


#old implementation
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


#test of new implementation, get rid of the old one if this works




wb.save(ws.title)
