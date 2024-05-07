import pdfplumber
import re
import os


# Path to your PDF file
capture_data = False
test_folder = 'test'

table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_y_tolerance": 5,
    "intersection_x_tolerance": 15
}

"""Adderer opp alle antall på den varartiklen jeg definerer."""
def utstyrsteller():
    for (root, dirs, files) in os.walk('test'):
        for file in files:
            if str(file).endswith('pdf'): #Sjekker at filen er en .pdf før jeg åpner den.
                print(f"Starter med ISO-Tegning: {file}")
                with pdfplumber.open(os.path.join(test_folder, file)) as pdf:  # Åpner PDF-en.
                    page_counter = 1
                    for page in pdf.pages: #Looper gjennom sidene i pdf'en.

                        im = page.to_image()
                        pikk = im.reset().debug_tablefinder(table_settings)
                        image_path = os.path.join(root)
                        pikk.save('C:\\Users\\haakon.granheim\\PycharmProjects\\MTO_sammenlignef\\test\\wolla.jpg')
                        print("shit saved")

                        """https://github.com/jsvine/pdfplumber/blob/stable/examples/notebooks/extract-table-nics.ipynb"""

utstyrsteller()