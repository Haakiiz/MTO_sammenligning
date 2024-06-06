import pdfplumber
import re
import os
import pandas as pd
from collections import defaultdict


test_folder = 'test'

table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines", #lines
    "snap_y_tolerance": 7,
    "intersection_x_tolerance": 10,
}

"""Adderer opp alle antall på den varartiklen jeg definerer."""
def utstyrsteller():
    article_totals = defaultdict(float)
    for (root, dirs, files) in os.walk('test'):
        for file in files:
            if str(file).endswith('pdf'): #Sjekker at filen er en .pdf før jeg åpner den.
                print(f"Starter med ISO-Tegning: {file}")
                with pdfplumber.open(os.path.join(test_folder, file)) as pdf:  # Åpner PDF-en.
                    page_counter = 1
                    for page in pdf.pages: #Looper gjennom sidene i pdf'en

                        # Get dimensions of the page
                        width = page.width
                        height = page.height
                        # Skaffer coordinatene i hjørnet
                        x0 = width * 0.70
                        y0 = 0
                        x1 = width
                        y1 = height * 0.30

                        bbox = (x0,y0,x1,y1)

                        cropped_page = page.crop(bbox)
                        p0 = cropped_page


                        #Debbug for å finne hva slags rader og kolonner den finner.
                        im = p0.to_image()
                        im = im.reset().debug_tablefinder(table_settings)
                        image_path = os.path.join(root)
                        im.save('C:\\Users\\haakon.granheim\\PycharmProjects\\MTO_sammenlignef\\test\\Debug.jpg')
                        print(f"Debug image saved to {test_folder}")


                        tables = p0.extract_tables(table_settings)
                        for table in tables:
                            for row in table:
                                if not row or not row[0]:
                                    continue

                                mengde = row[1]
                                beskrivelse = row[4]

                                if str.lower('KAPPLISTE') in str.lower(row[0]):
                                    break

                                # Check if 'mengde' is a valid number
                                if mengde and 'M' in mengde:
                                    try:
                                        antall_type = float(mengde.replace('M', '').replace(',', '.'))
                                    except ValueError:
                                        continue
                                elif mengde:
                                    try:
                                        antall_type = float(mengde.replace(',', '.'))
                                    except ValueError:
                                        continue
                                else:
                                    antall_type = 0.0

                                # Accumulate quantity for the description
                                if beskrivelse:
                                    article_totals[beskrivelse] += antall_type

                # Print the summarized results
                for article, total in article_totals.items():
                    print(f"Antall av ---{article}---: {total}")



utstyrsteller()


