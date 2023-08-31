import os
from datetime import datetime
from PDF_Class import data_scraper
from PDF_To_Excel import excel_batch_processor

#changes the directory to where inventory pdfs are stored
def directory_changer():
    print("current path:")
    print(os.getcwd())

    #manual version
    return os.chdir(input("Where are the inventory files stored?\n"))

    #work version H drive
    #return os.chdir(r"P:\Public\Microbiology\Media Prep\Media Inventory\2023")

    #work version upstairs
    #return os.chdir(r"C:\Users\MCN2226\inventory speed test\Inventory Files")

    #return os.chdir(r"C:\Users\MCN2226\Documents\inventory test")

    #home desktop version
    #return os.chdir(r"C:\Users\Mason\Documents\NAMSA Test\Inventory\Inventory Files")

    #home laptop version
    #return os.chdir(r"C:\Users\mason\OneDrive\Documents\Python Projects\Inventory Files")



#data_scraper() calls ---> smart_splitter() whose product calls ---> read_ahead()

def run_program():
    pdf_location = directory_changer()
    print("")

    #compares the start and end time of running operations to determine how long the program took
    start_time = datetime.now().time()
    start_seconds = (start_time.hour * 60 + start_time.minute) * 60 + start_time.second

    data_scraper(pdf_location)

    excel_batch_processor()

    end_time = datetime.now().time()
    end_seconds = (end_time.hour * 60 + end_time.minute) * 60 + end_time.second

    print("\n~" + str(end_seconds- start_seconds) + " seconds of compute")
    print("If times are unsatisfactory, move network inventory files to a local drive")

run_program()

#maybe make a list of key words for the comments for clarity. EX: 'Line' should be replaced with something easily understandable that couldn't be confused with something else.
