import os
import calendar
import copy
from PyPDF2 import PdfReader
class PDF:
    #class variables
    pdf_id = 1
    pdf_list = []
    pdf_length = 0
    #each index will be a inv/min ratio for each month brought down to 2 decimal places
    pdf_lots_over_min = []
    pdf_month_list = []
    pdf_media_type_list = []

    #instance variables and functions
    def __init__(self, unformatted_data, filename):
        #used once in end_functions() in reference to PDF.pdf_id, since the final pdf will have a self.id value one less than PDF.pdf_id
        self.id = PDF.pdf_id
        #filename portion is the 0th index of the pdf's data, stored here as a string
        self.filename = filename

        #determines the month the pdf applies to. This is used for analytics on each month
        self.month = self.month_determiner()

        #variables that store what the pdf is missing, in reference to an ideal pdf with all media types listed
        #these are used to determine what lines need to be added to each pdf instance, so it's data can be reliably printed onto the resulting Excel document
        self.FTM_400S_Needed = False
        self.FTM_1000_Needed = False
        self.DFD_1000_Needed = False
        self.Remove_OD = False
        #data portion is the 1st index of the formatted pdf as a list of all the lines
        self.data = self.data_correcter(unformatted_data)
        #if every pdf has the same length, this value can be used to increment over self.data, self.inventory_list, self.minimum_list, and self.inv_ratio_list_full_float
        PDF.pdf_length = self.length_checker()
        #solely used to check for cases where a given pdf does not have the same number of lines as other pdfs, which throws an error
        pdf_length = self.length_checker()

        #nothing formally stores the meaning of where each recorded value is put in self.inventory_list, but it is built from first media type, SCDB 100, to last, Saline P80
        #a dictionary might be a better implementation, unless the types included aren't known...
        self.inventory_list = self.inventory_list_maker()

        #like for self.inventory_list, there is nothing that formally stores what each entry in self.minimum_list is in reference to
        self.minimum_list = self.minimum_list_maker()

        #test for retrieving the number of inventory units per lot
        self.lot_size_list = self.lot_size_list_maker()

        #gets ratios in float form for accuracy, the rounding is done later after all floated values for each month are added together and divided by the number of recorded values for each month
        self.lots_above_min_float_list = self.lots_above_min_float_list_maker()
        
        #the PDF class has a list which stores every pdf instance
        PDF.pdf_list.append(self)

        PDF.pdf_id += 1

        #useful for approximating the speed of the work pc
        print(str(PDF.pdf_id) + " pdf's processed")


    #determines the month from the filename, assuming standardized yy-mm-dd format
    def month_determiner(self):
        try:
            month_name = calendar.month_name[int(self.filename[3:5])]
            #adds the months name to the PDF list of month names if it isn't already in there
            if month_name not in PDF.pdf_month_list:
                PDF.pdf_month_list.append(month_name)
            return month_name
        except:
            raise Exception("PDF " + str(self.filename) + " has an improperly formatted filename, should be 'yy-mm-dd Media Inventory.pdf'")


       
    #formats the data to have functional phrases with missing media types added
    def data_correcter(self, unformatted_data):
        return self.data_formatter(self.smart_splitter(unformatted_data))



    #returns a particular line of the pdf
    def line_retrieval(self, line_num):

        #the '-1' turns the line_num into a list indice
        return self.data[line_num - 1]



    #gets the length of the pdf based on its self.data instance variable
    def length_checker(self):
        self_length = 0

        for line in self.data:
            self_length += 1

        if PDF.pdf_length != 0:
            if PDF.pdf_length != self_length:
                #every pdf should have the same length so the data printed into Excel is properly referenced with respect to every media type
                raise Exception("PDF " + str(self.filename) + " has a formatting which gave it a length of " + str(self_length) + " while the others had a length of " + str(PDF.pdf_length))

        return self_length
    


    def inventory_list_maker(self):
        inventory_list = []

        #the first line of each pdf will have "Inv" at the end of it. Cases where "Inv" is at the end of another line, for whatever reason that might be, should still be handled as errors
        #first_line was made in place of making "Inv" a condition in the Except if case for this reason
        first_line = True
        for pdf_line in self.data:

            try:
                #adds to the inventory list only if the line has a integer at the end, which would be the inventory on an ideally formatted pdf
                inventory_list.append(int(pdf_line[-1]))
            
            #lines without inventory values at the end should have a string added to inventory list in it's place, to denote that no media was recorded. This is different from adding 0, since this would suggest that no media was present, but it was recorded
            #lines without a recognized replacement for an inventory value at the end, however, should raise an error as the validity of their addition to the list of cases where an empty string should be added is ambiguous
            except:
                if first_line:
                    first_line = False
                    continue

                elif pdf_line[-1] == "" or pdf_line[-1] == "OD":
                    inventory_list.append("")
                
                else:
                    raise Exception("PDF " + str(self.filename) + " has a formatting which gave it a line without an inventory value stored at the end of it.\nThis line is " + str(pdf_line))
      
            first_line = False

        return inventory_list
    


    #makes a list for the minimum inventory values given for the pdf running the function
    def minimum_list_maker(self):
        minimum_list = []
        first_line = True
        for pdf_line in self.data:
            try:
            #turns the minimum value into an integer if there is a minimum value present
                minimum_list.append(int(pdf_line[-3]))
            
            except:
                #doesn't add the first line of each pdf as it has no inventory
                if first_line:
                    first_line = False
                    continue
                #not adding this empty string to minimum_list would cause minimum_list to be indexed improperly in Excel related functions
                if pdf_line[-3] == "":
                    minimum_list.append("")
                else:
                    raise Exception("PDF " + str(self.filename) + " has a formatting which didn't give it a minimum value 2 places to the left of the end of the line.\nThis line is " + str(pdf_line))
            
            first_line = False

        return minimum_list
    


    def lot_size_list_maker(self):
        lot_size_list = []
        first_line = True
        for pdf_line in self.data:

            if first_line:
                first_line = False
                continue

            lot_size = ""
            for char in pdf_line[2]:
                try:
                    int(char)
                    lot_size += char
                
                except:
                    lot_size_list.append(int(lot_size))
                    break
        return lot_size_list
    


    #determines the number of lots above minimum for each media type in the pdf
    def lots_above_min_float_list_maker(self):

        if len(self.inventory_list) != len(self.minimum_list):
            raise Exception("The length of self.inventory_list is " + str(len(self.inventory_list)) + " while self.minimum_list is " + str(len(self.minimum_list)) + " items long. These do not match.")

        #utilizes floats
        lots_above_min_float_list = []
        #this indexing works because self.inventory_list is the same length as self.minimum_list
        for i in range(0, len(self.inventory_list)):

            try:
                lots_above_min_float_list.append((self.inventory_list[i] - self.minimum_list[i]) / self.lot_size_list[i])
            
            except:
                lots_above_min_float_list.append("")

        return lots_above_min_float_list



    #adds float ratios for each month's data and rounds to two decimal places
    #could also do other things, with the current stuff restricted to one function, of many, that end_functions() calls
    @classmethod
    def end_functions(cls):
        inv_over_min_float_list = []
        #starts the old_pdf_month as 'January' for the first pdf so it has the variable declared
        old_pdf_month = PDF.pdf_list[0].month
        for pdf in PDF.pdf_list:
            cur_pdf_month = pdf.month

            #runs once a change in month is detected, or if the pdf is the last of the made pdfs
            if cur_pdf_month != old_pdf_month or pdf.id == PDF.pdf_id - 1:

                #holds the ratios for each media type rounded to two decimals for a given month
                media_monthly_inv_over_min_list = []

                #the length of inv_over_min_float_list[0] is the num of different media types
                #the length of inv_over_min_float_list is num of pdfs for a given month
                for i in range(0, len(inv_over_min_float_list[0])):
                    exp = PDF.pdf_list[0].data[i]
                    media_monthly_sum = 0

                    #recorded_num does not iterate from "" entries, and doesn't use them for the denominator in the average
                    recorded_num = 0
                    non_record_count = 0
                    for daily_ratio in inv_over_min_float_list:
                        if daily_ratio[i] != "":
                            media_monthly_sum += daily_ratio[i]
                            recorded_num += 1
                        else:
                            non_record_count += 1

                            print(str(non_record_count) + " pdfs without data for " + str(pdf.data[i]) + " in the month of " + old_pdf_month)

                    #prevents the averaging of media types with no data for them
                    #does work, however, with media types with some pdfs with data and others without
                    if recorded_num != 0:
                        media_monthly_inv_over_min_list.append(round(media_monthly_sum / recorded_num, 2))

                    elif recorded_num == 0:
                        media_monthly_inv_over_min_list.append("")
                
                PDF.pdf_lots_over_min.append(media_monthly_inv_over_min_list)
            
            elif cur_pdf_month == old_pdf_month:
                inv_over_min_float_list.append(pdf.lots_above_min_float_list)

            #add a condition for the last pdf, however you want to determine which is the last pdf, where it still runs the cur_pdf_month != old_pdf_month code
            #otherwise the final month will not be added to PDf.pdf_monthly_inv_ratios

            old_pdf_month = pdf.month

    """
    #attempting a re-write
    def end_functions(cls):
        inv_over_min_float_list = []

        old_pdf_month = PDF.pdf_month_list[0]
        for pdf in PDF.pdf_list:
            cur_pdf_month = pdf.month

            if cur_pdf_month != old_pdf_month or pdf.id == PDF.pdf_id - 1:
                media_monthly_inv_over_min_list = []
                #will look like [0, .2, .3, as long as there are media types]








            elif cur_pdf_month == old_pdf_month:
                inv_over_min_float_list.append(pdf.lots_above_min_float_list)

    """





    #helper function for smart_splitter() that gives it the read_ahead to add to phrase to make conjoined_phrase. This allows for Excel cells to have meaningful 'categories' in cells as opposed to the raw data all being in separate cells, which would ruin formatting
    def read_ahead(self, list, r_i, read=None):

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

            #the read is cut off and considered to be a special case once the word is completed, as denoted by a " " following it
            #since so much of this is hard-coded, there is the assumption that the ideal pdf will remain untouched in the future. In the event that it isn't, this is where changes would need to be made for adding new functional phrase cases that future pdfs might have
            elif list[r_i] == " ":
                if (
                    #covers the two special case lists where read should be added to phrase
                    read[-1] in special_case_list1
                    or read[-2:] in special_case_list2

                    #adds 'trays (x)' to the list of special cases (but wouldn't special_case_list2's 'ys' cover this? Is this not some other exception?)
                    or len(read) > 4
                    and read[-5] == "s"

                    #adds Saline P80 to the list of special cases, since having '80' be a special case would cause issues
                    or len(read) == 3
                    and read == "P80"

                    #these appear to be for additional special cases... find out what they're for
                    or (len(read) == 3 and read[-1] == ")")
                    or (len(read) == 4 and read[-1] == ")")
                ):
                    
                    #recursively calls read_ahead() so long as there is a special case in front of read
                    future_read = self.read_ahead(list, r_i, read)

                    #each call will get progressively longer so long as special cases are in front of read
                    if future_read != None:
                        future_read = read + " " + future_read
                        return future_read
                        
                    #returns the read when it is a special case but future_read was not, denoting that the read plus the initial phrase in smart_splitter() make up the whole conjoined_phrase
                    else:
                        return read
                #returns None if a special read is not seen after the phrase from smart_splitter()
                else:
                    return None
                    


    #turns whole lines of the pdf into a list of separate strings, each of which should hold an appropriate phrase. EX: '45 trays (5)' is a phrase, as are '165(5)' and '100'
    #as can be seen by these examples, a 'phrase' is of variable length, which is why this function needs rules for determining what constitutes a phrase
    #but why phrases in the first place? Phrases were made so each line takes up the same number of Excel cells, so things like inventory and minimum values can be accessed correctly both in Python and in Excel graph cell references
    #reading ahead of a given phrase is needed to determine how long each phrase should be. Once you find something ahead of the existing phrase that would does not meet phrase criteria, return None. Until then, recursively call the same function, indexed ahead of the space following the existing phrase
    #each pdf that is scraped is a list containing each line of the pdf, which holds each line as a separate string
    def smart_splitter(self, unformatted_data):
        final_lists = []

        for list in unformatted_data:
            split_list = []
            #'phrase' is used as the word to designate what is read, as it could be multiple words long
            phrase = ""
            delay = 0

            #each list at this point is a line of text as one whole string, so this iterates over every character of that line string
            for i in range(len(list)):

                #causes i to not iterate over what's already been read by read_ahead()
                #delay is only incremented when read_ahead() has added the characters ahead of a phrase into that same phrase
                #characters read but not added to the phrase are not added to delay, as they could still be part of their own phrase and must have read_ahead() called on them
                if delay > 0:
                    delay -= 1

                    #resets phrase when reading ahead has been done, so the next phrase can start as an empty string
                    phrase = ""

                elif list[i] != " ":
                    phrase += list[i]

                #read_ahead() is called when a space is encountered, as it may designate a phrase made up of multiple words
                elif list[i] == " ":
                    text_ahead = self.read_ahead(list, i)

                    #adds the text ahead of a given phrase, should that text be a special case that designates that it should be added to the phrase
                    if text_ahead != None:
                        conjoined_phrase = phrase + " " + text_ahead

                        split_list.append(conjoined_phrase)

                        #delays i beyond the full extent of what's already been read
                        delay = len(text_ahead) + 1

                    #simply adds phrase to split_list, should no special cases be found ahead of it. Every phrase, except for the last, ends with a None return by read_ahead() and this else block being executed
                    else:
                        split_list.append(phrase)
                        phrase = ""

            #this line accounts for the final phrase, if it wasn't already part of a special case that incremented delay far enough where read_ahead() wasn't called again
            #wait if it was part of a special case would read_ahead()'s attempt to increment past it break things???
            split_list.append(phrase)

            #adds each list to the final_lists, which is a list containing each line list with its multi-word phrases
            final_lists.append(split_list)

        return final_lists



#one potential optimization
#when read_ahead() reads ahead and finds that the word ahead of the phrase is not a special case, this word can still be returned to smart_splitter as phrase = read with phrase = "" bypassed and delay incremented appropriately
#the only issue is that bypassing phrase = "" might introduce additional time in the consideration of when to do so. You could test different implementations



#also might want to separate the terms PDF as the class representation and the actual pdfs you're working with



    #at this point each line of the data from a pdf contains functional phrases, but many pdfs are missing certain media types.
    #data_formatter() adds every media type to each pdf's data, even if no media of that type were listed on the pdf for a given date. Takes from the data processed by smart_splitter(). Only adds 400-S FTM, 1000 FTM, and DFD as well as removing OD lines because too much more would vastly overcomplicate this
    #since so much of this is hard-coded, there is the assumption that the ideal pdf will remain untouched in the future. In the event that it isn't, this is where changes would need to be made for adding new media types to each pdf's self.pdf_data
    #identifying ways to bypass this function would save some time, for pdfs with already ideal formatting. Length can't be used as a bypass condition, as an incorrectly formatted pdf could deceptively have the same number of lines as an ideally formatted one
    def data_formatter(self, spaced_data):
        
        #a deep copy is needed so changes can be made to spaced_data_copy, while iterating through spaced_data, without affecting spaced data in any way
        spaced_data_copy = copy.deepcopy(spaced_data)

        #knowing where 400-S FTM, 1000 FTM, and 1000 DFD should be on an ideally formatted pdf is what is used to reference where they should be added. Sublist_count tracks makes sure comparisons for whether a media type is or isn't present is always done on the appropriate line, by adjusting for pdfs without 
        #space_offset adjusts for the indexing of each line on spaced_data changing on pdfs with multiple missing media types
        #the 'ideally formatted pdf' comes from observing where each media type is placed in pdfs that have the given media type. The placements never vary, and shouldn't in the future
        sublist_count = 0
        space_offset = 1


        #this will require a hefty re-write, but let's attempt to have every 'Media Type' column filled in with "" if nothing is listed


        #as strange as it is, a copy of spaced_data_copy needed to be made to add to spaced_data_copy without causing an infinite loop
        #if you update the original it updates the copy, you have to update the copy
        for sub_list in spaced_data:
            sublist_count += 1
            #print(sublist_count)

            if len(sub_list) < 6:
                try:
                    spaced_data_copy[sublist_count - space_offset].insert(0, "")
                except:
                    #in the case of the "OD" line on the bottom of some pdfs, this is run
                    pass

            #adds 400-S FTM
            if sublist_count == 12:
                #since this sub_list is from a copy of spaced_data_copy, it doesn't have "" at its first index
                if sub_list[0] == "400-S":
                    pass
                
                else:
                    self.FTM_400S_Needed = True
                    sublist_count += 1
                    space_offset += 1

            #adds 1000 FTM
            elif sublist_count == 14:
                if sub_list[0] == "1000":
                    pass
                    
                else:
                    self.FTM_1000_Needed = True
                    sublist_count += 1
                    space_offset += 1
            
            #adds 1000 DFD
            elif sublist_count == 26:
                #the or condition is given because the minimum for DFD was changed on 06/26/23 from 136 to 99
                #volume can't be referenced instead, because DFD 1000 is surrounded by other 1000mL media types
                if sub_list[0] == "DFD":
                    pass

                else:
                    self.DFD_1000_Needed = True
                    sublist_count += 1
                    space_offset += 1
            
            #removes the 'OD,' or 'On Demand,' line at the end of older pdfs
            #each pdf is assumed to be a certain length at this point, but PDF.pdf_length cannot be made yet to check if each pdf has the correct number of lines
            #check if sublist_count ever reaches 34 for a normal pdf...
            elif sublist_count == 32:
                if sub_list[0] == "OD=":
                    self.Remove_OD = True

        #runs after all the other data has been added, so that the .insert() method knows where to add 400-S FTM, 1000 FTM, and 1000 DFD based on where they would be in an ideally formatted pdf
        #the final value added, "", indicates that no media for this media type was recorded on this date

        #ADDED THE FIRST QUOTATIONS
        if self.FTM_400S_Needed:
            spaced_data_copy.insert(11, ["", "400-S", "", "", "", ""])

        if self.FTM_1000_Needed:
            spaced_data_copy.insert(13, ["", "1000", "", "", "", ""])

        if self.DFD_1000_Needed:
            spaced_data_copy.insert(25, ["DFD", "1000", "", "", "", ""])

        if self.Remove_OD:
            spaced_data_copy.pop(-1)

        return spaced_data_copy



#reads the actual pdfs in the folder where they were specified to exist in directory_changer()
def data_scraper(pdf_location):
    pdf_count = 0

    for filename in os.listdir(pdf_location):
        if filename[-4:] == ".pdf":
            pdf_count += 1

            #uses imported functions from PyPDF2 to read from pdfs
            reader = PdfReader(open(filename, "rb"))

            #this is the raw line by line data scraped from each pdf as a list containing each line, each of which is a string containing all the info with spaces in between words
            unformatted_data = []

            if len(reader.pages) > 1:
                Exception("PDF " + filename + " contains more than one page of data. It should be only one page.")

            selected_page = reader.pages[0]

            raw_data = selected_page.extract_text()

            #makes each line of the pdf into a separate string, not separated by meaningful phrases, and appends them to unformatted_data
            #the fact that words haven't been combined into phrases is what makes this data 'unformatted.'
            unformatted_data += raw_data.splitlines()

            PDF(unformatted_data, filename)
    
    #runs functions related to when all PDF data is inputted
    #these functions collect metadata about the whole of the pdf data
    PDF.end_functions()
