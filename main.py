"""Jeg må ha mulighet til å sammenligne både FRA Excel TIL PDF, men også FRA PDF og TIL EXCEL!
Dette er viktig for nå fant jeg nettopp ett eksempel hvor noe var I PDF, men ikke i EXCEL."""

"""Jeg begynner først med å sammenligne antall jeg. M-313-001 så har de gjort endringer i krager og boltesett...Må sjekke antall på kragene og rørlengder"""

import pandas as pd
import pdfplumber
import re
import os



# Path to your PDF file
check_folder_path = 'check_folder'
capture_data = False


"""Adderer opp alle antall på den varartiklen jeg definerer."""
def utstyrsteller():
    for (root, dirs, files) in os.walk('check_folder'):
        for file in files:
            if str(file).endswith('pdf'): #Sjekker at filen er en .pdf før jeg åpner den.
                print(f"Starter med ISO-Tegning: {file}")
                with pdfplumber.open(os.path.join(check_folder_path, file)) as pdf:  # Åpner PDF-en.
                    page_counter = 1
                    for page in pdf.pages: #Looper gjennom sidene i pdf'en.


                        # Get dimensions of the page
                        width = page.width
                        height = page.height
                        # Skaffer coordinatene i hjørnet
                        x0 = width * 0.70
                        y0 = 0
                        x1 = width
                        y1 = height

                        bbox = (x0,y0,x1,y1)

                        cropped_page = page.crop(bbox)


                        # Extract the text from the cropped area
                        page_text = cropped_page.extract_text(layout=False)
                        # Define the regex pattern to capture groups: ID, DN, TYPE
                        pattern = r'^(\d+)\s+(\d+)\s+(\w+(?:\w+\s)*)$'


                        # List to hold all extracted records
                        ids = []
                        dns = []
                        types = []
                        document_names = []
                        material_kvalitet = []
                        sveisemetode = []
                        wps = []
                        pre = []

                        if page_text:
                            lines = page_text.split('\n')
                            for line in lines:
                                if 'SVEISELISTE' in line:
                                    capture_data = True
                                    continue

                                if capture_data:
                                    match = re.search(pattern, line.strip())
                                    if match:
                                        ids.append(match.group(1))
                                        dns.append(match.group(2))
                                        types.append(match.group(3))
                                        document_names.append(f"{file} s.{page_counter}")
                                        material_kvalitet.append("AISI 304L")
                                        sveisemetode.append("141")
                                        wps.append("18_02")
                                        pre.append("PRE")

                            capture_data = False