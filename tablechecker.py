import pdfplumber
import os
import pandas as pd
from collections import defaultdict
import tabulate as tabulate

###Viktige ting det er viktig å endre på:
    #Excel filnavn i excel_path, og sjekk om det er .xls eller .xlsx
    #sjekk skiprows= stemmer hvor man åpner excelfila i compare funksjonen
#Viktig å endre excel_path!!#

test_folder = 'check_folder'
excel_path = 'check_folder/NVO-AE1-30-MU-005-0_03E_E1 - VBA. Mengdeliste MOD samlestokker fra SPV-basseng.xls'

table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines", #lines
    "snap_y_tolerance": 8, #var tidligere 7
    "intersection_x_tolerance": 10,
}



def normaliserer_beskrivelser(description):
    """
    Normaliserer og rengjør beskrivelsene....en stor kilde til feil
    """
    return str(description).replace('\n', ' ').replace('\t', ' ').strip().upper()

"""Adderer opp alle antall på den varartiklen jeg definerer."""
def utstyrsteller():
    article_totals = defaultdict(float)
    for (root, dirs, files) in os.walk(test_folder):
        for file in files:
            if str(file).endswith('pdf'): #Sjekker at filen er en .pdf før jeg åpner den.
                print(f"ISO-Tegning: {file}")
                with pdfplumber.open(os.path.join(test_folder, file)) as pdf:  # Åpner PDF-en.
                    page_counter = 1
                    for page in pdf.pages: #Looper gjennom sidene i pdf'en

                        # Get dimensions of the page
                        width = page.width
                        height = page.height

                        # Skaffer coordinatene i hjørnet
                        crop_margin = 28.35 * 5.5
                        x0 = width * 0.740
                        y0 = 0
                        x1 = width - crop_margin
                        y1 = height * 0.30

                        bbox = (x0,y0,x1,y1)

                        #Cropper shitten
                        cropped_page = page.crop(bbox)
                        p0 = cropped_page

                        #henter tabellen
                        tables = p0.extract_tables(table_settings)
                        for table in tables:
                            for row in table:
                                if not row or not row[0]:
                                    continue

                                mengde = row[1]
                                beskrivelse = row[-1] #enten 3 eller 4

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

                                # Accumulate quantity for the description. 'Beskrivelse' er Key
                                if beskrivelse:
                                    rengjort_beskrivelse = normaliserer_beskrivelser(beskrivelse)
                                    article_totals[rengjort_beskrivelse] += antall_type


                            # Print the summarized results
        for article, total in article_totals.items():
            print(f"Antall av ---{article}---: {total}")

    return article_totals

def compare_with_excel(article_totals, output_excel_path = 'forskjellsrapport.xlsx'):
    all_differences = []
    # Load the Excel file
    for (root, dirs, files) in os.walk('check_folder'):
        for file in files:
            if str(file).endswith('.xls'): #Sjekker at filen er excel før jeg åpner den.
                print(f"Sammenligner med excelfil {file}")
                df = pd.read_excel(excel_path, sheet_name='Innhold', skiprows=6)

                # For excel kolonnene 'Beskrivelse' og 'Mengde'
                forskjeller = []

                for index, row in df.iterrows():
                    if pd.isna(row['Beskrivelse']):
                        continue
                    description = normaliserer_beskrivelser(row['Beskrivelse'])

                    if pd.isna(row['Mengde']):
                        excel_quantity = 0.0
                    else:
                        excel_quantity = row['Mengde']

                    # Get the quantity from the article_totals dictionary
                    pdf_quantity = article_totals.get(description, 0)

                    # Compare the quantities
                    if pdf_quantity != excel_quantity:
                        forskjeller.append({
                            'Kilde': 'Excel',
                            'Beskrivelse': description,
                            'Excel mengde': excel_quantity,
                            'PDF mengde': pdf_quantity
                        })

                #List of descriptions from Excel
                excel_descriptions = df['Beskrivelse'].dropna().apply(normaliserer_beskrivelser).tolist()

                #Check for items in PDFs but not in excel
                for description in article_totals.keys():
                    if description not in excel_descriptions:
                        forskjeller.append({
                            'Kilde': 'PDF',
                            'Beskrivelse': description,
                            'Excel mengde': 0.0,
                            'PDF mengde': article_totals[description]
                        })

                all_differences.extend(forskjeller)

                if forskjeller:
                    print("Forskjeller funnet:")
                    table = [
                        [diff['Kilde'], diff['Beskrivelse'], diff['Excel mengde'], diff['PDF mengde']]
                        for diff in forskjeller
                    ]
                    print(tabulate.tabulate(table, headers=["Kilde", "Beskrivelse", "Excel Mengde", "PDF Mengde"]))
                else:
                    print("Ingen forskjeller funnet. Mengden matcher.")

    # Save differences to an Excel file
    if all_differences:
        df_differences = pd.DataFrame(all_differences)
        df_differences.to_excel(output_excel_path, index=False)
        print(f"Differences saved to {output_excel_path}")
    else:
        print("Ingen forskjeller funnet. Mengden matcher.")


def debug():
    for (root, dirs, files) in os.walk(test_folder):
        for file in files:
            if str(file).endswith('pdf'):  # Sjekker at filen er en .pdf før jeg åpner den.
                print(f"ISO-Tegning: {file}")
                with pdfplumber.open(os.path.join(test_folder, file)) as pdf:  # Åpner PDF-en.
                    page_counter = 1
                    for page in pdf.pages:  # Looper gjennom sidene i pdf'en

                        # Get dimensions of the page
                        width = page.width
                        height = page.height
                        # Skaffer coordinatene i hjørnet

                        crop_margin = 28.35 * 5.5
                        x0 = width * 0.740
                        y0 = 0
                        x1 = width - crop_margin
                        y1 = height * 0.30

                        bbox = (x0, y0, x1, y1)

                        cropped_page = page.crop(bbox)
                        p0 = cropped_page

                        # Debbug for å finne hva slags rader og kolonner den finner.
                        im = p0.to_image()
                        im = im.reset().debug_tablefinder(table_settings)
                        image_path = os.path.join(root)
                        im.save(f'C:\\Users\\haakon.granheim\\PycharmProjects\\MTO_sammenlignef\\debug\\{file[:20]}_debug.jpg')
                        print(f"Debug image saved to debug folder")


# Kaller funksjonene
article_totals = utstyrsteller()
compare_with_excel(article_totals)


"""Debug om jeg skulle trenge det. Sjekker hvordan scripted leser pdf-filene"""
#debug()


