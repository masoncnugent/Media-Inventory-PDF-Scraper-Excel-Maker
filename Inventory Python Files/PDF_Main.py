import os
from PDF_Class import data_scraper
from PDF_To_Excel import excel_batch_processor

#changes the directory to where inventory pdfs are stored
def directory_changer():
    print("current path:")
    print(os.getcwd())

    #manual version
    #return os.chdir(input("Where are the inventory files stored?\n"))

    #work version
    #return os.chdir(r"P:\Public\Microbiology\Media Prep\Media Inventory\2023")

    #home desktop version
    #return os.chdir(r"C:\Users\Mason\Documents\NAMSA Test\Inventory")

    #home laptop version
    return os.chdir(r"C:\Users\mason\OneDrive\Documents\Python Projects\Inventory Files")



#data_scraper calls ---> smart_splitter whose product calls ---> read_ahead

def run_program():
    pdf_location = directory_changer()

    data_scraper(pdf_location)

    excel_batch_processor()

run_program()

#One assumption is that the first pdf is formatted ideally, since it's media types are used to format the rest of the data into cells for use in graphs
