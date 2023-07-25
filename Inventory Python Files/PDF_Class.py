import os
from PyPDF2 import PdfReader
class PDF:
    #class variables
    pdf_id = 1
    pdf_list = []


    def __init__(self, unformatted_data, filename):
        #instance variables
        self.id = PDF.pdf_id
        #filename portion is the 0th index as a string
        self.filename = filename

        self.FTM_400S_Needed = False
        self.FTM_1000_Needed = False
        self.DFD_1000_Needed = False
        self.Remove_OD = False
        #data portion is the 1st index as a list of all the lines
        self.data = self.data_correcter(unformatted_data)
        self.inventory_list = self.inventory_list_maker()
        self.length = self.length_checker()

        PDF.pdf_list.append(self)

        PDF.pdf_id += 1

        #useful for gauging the speed of the work pc
        print(str(PDF.pdf_id) + " pdf's processed")

    #formats the data to have functional phrases with missing media types added
    def data_correcter(self, unformatted_data):
        spaced_data = smart_splitter(unformatted_data)

        formatted_data = data_formatter(self, spaced_data)

        return formatted_data


    #returns a particular line of the pdf
    def line_retrieval(self, line_num):

        #the '-1' turns the line_num into a list indice
        return self.data[line_num - 1]


    #gets the length of the pdf based on its self.data instance variable
    def length_checker(self):
        length = 0

        for line in self.data:
            length += 1
        
        return length
    
    def inventory_list_maker(self):
        inventory_list = []
        for pdf_line in self.data:

            try:
                #adds to the inventory list only if the line has a integer at the end
                inventory_list.append(int(pdf_line[-1]))
            
            except:
                if pdf_line[-1] == "OD":
                    inventory_list.append("")
                continue
        
        return inventory_list
    


    #helper function for smart_splitter() that gives it the read_ahead to add to phrase to make conjoined_phrase. This allows for Excel cells to have meaningful 'categories' in cells as opposed to the raw data all being in separate cells, which would ruin formatting
def read_ahead(list, r_i, read=None):

    #special case list 1 and 2 form the bulk of the cases where smart_splitter() should add the read_ahead to a given phrase
    special_case_list1 = ["o", "h", "%"]

    special_case_list2 = ["pe", "ys"]

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



#turns whole lines of the pdf into a list of separate strings, each of which should hold an appropriate phrase. EX: '45 trays (5)' is a phrase, as are '165(5)' and '100'
#reading ahead of a phrase is needed to determine how long each phrase should be
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



#will add every media type to each pdf's Excel representation, even if no media of that type were listed on the pdf for the date. Takes from the data processed by smart_splitter(). Only adds 400-S FTM, 1000 FTM, and DFD as well as removing OD lines because too much more would vastly overcomplicate this.
#since so much of this is hard-coded, there is the assumption that the pdf will remain untouched in the future. In the event that it isn't, this is where changes should most likely be made first.
def data_formatter(pdf, spaced_data):
    
    #a shallow copy was needed to prevent the loop running indefinitely
    spaced_data_copy = spaced_data.copy()

    #knowing where 1000 FTM and 1000 DFD should be on an ideally formatted pdf is what is used to reference where they should be added. Sublist_count tracks what line of the pdf we should add to on the Excel representation
    sublist_count = 0
    sublist_mod = 0

    for sub_list in spaced_data_copy:
        sublist_count += 1

        #adds 400-S FTM
        if sublist_count == 12 - sublist_mod:
            if sub_list[0] == "400-S":
                sublist_count += 1
            
            else:
                pdf.FTM_400S_Needed = True
                sublist_count += 1
                sublist_mod += 1

        #adds 1000 FTM
        elif sublist_count == 15 - sublist_mod:
            if sub_list[0] == "1000":
                sublist_count += 1
                
            else:
                pdf.FTM_1000_Needed = True
                sublist_count += 1
                sublist_mod += 1
        
        #adds 1000 DFD
        elif sublist_count == 28 - sublist_mod:
            #the or condition is given because the minimum for DFD was changed on 06/26/23 from 136 to 99
            if sub_list[3] == "136" or sub_list[3] == "99":
                sublist_count += 1

            else:
                pdf.DFD_1000_Needed = True
                sublist_count += 1
                sublist_mod == 1
        
        #removes the 'OD' or 'On Demand' line from previous pdfs
        elif sublist_count == 34 - sublist_mod:
            if sub_list[0] == "OD=":
                pdf.Remove_OD = True
    

    #runs after all the other data has been added, so that the .insert() method knows where to add 1000 FTM and 1000 DFD based on where they would be in an ideally formatted pdf
    #the final value added, "", indicates that no media for this media type was recorded on this date
    if pdf.FTM_400S_Needed:
        spaced_data_copy.insert(11, ["400-S", "9 trays/lot", "4", "9 trays (2)", ""])

    if pdf.FTM_1000_Needed:
        spaced_data_copy.insert(13, ["1000", "17-20/lot", "36", "90 (5)", ""])

    if pdf.DFD_1000_Needed:
        spaced_data_copy.insert(25, ["DFD","1000", "34 bottles/lot", "99", "136", ""])

    if pdf.Remove_OD:
        spaced_data_copy.pop(-1)

    return spaced_data_copy



def data_scraper(pdf_location):
    pdf_count = 0

    #looks for pdfs in the current directory given to find media inventory pdfs
    for filename in os.listdir(pdf_location):
        if filename[-4:] == ".pdf":
            pdf_count += 1

            #uses imported functions from PyPDF2 to read from pdfs
            reader = PdfReader(open(filename, "rb"))

            #this is the raw line by line data scraped from each pdf as a list containing each line, which is also a list
            unformatted_data = []

            #every media pdf should be one page long, this could cause issues should that requirement not be met. (UPDATE)
            for i in range(0, len(reader.pages)):
                selected_page = reader.pages[i]

                raw_text = selected_page.extract_text()

                #takes each line of the pdf as a separate string, not separated by meaningful phrases
                unformatted_data += raw_text.splitlines()

            PDF(unformatted_data, filename)