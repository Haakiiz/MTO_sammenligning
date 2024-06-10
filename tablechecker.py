import pdfplumber
import os
import pandas as pd
from collections import defaultdict
import tabulate as tabulate

###Viktige ting det er viktig å endre på:
    #Excel filnavn i excel_path, og sjekk om det er .xls eller .xlsx
    #sjekk skiprows= stemmer hvor man åpner excelfila i compare funksjonen
    #endre .endswith i compare funksjonen

test_folder = 'check_folder'

"""Disse må endres på før du kjører scriptet"""
excel_path = 'check_folder/NVO-AE1-30-MU-002-0_03E_E1 - VBA. Mengdeliste Slamlager (SLL).xls'
skipped_rows = 6
sheeeet_name = 'Sheet1_A3'
excelending = '.xls'

table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines", #lines
    "snap_y_tolerance": 8, #var tidligere 7
    "intersection_x_tolerance": 10,
}

def remove_none_types(x):
    [item for item in x if item is not None]


def normaliserer_beskrivelser(description):
    """
    Normaliserer og rengjør beskrivelsene....en stor kilde til feil
    """
    return str(description).replace('\n', ' ').replace('\t', ' ').strip().upper()

def format_iso_quantities(iso_files):
    """
    Format the quantities from different ISO files into a single string with line breaks.
    """
    formatted = "\n".join([f"{qty} fra {iso[:20]}..." for iso, qty in iso_files.items()])
    return formatted

def flatten_row(row):
    """
    Flattens a nested list into a single list.
    """
    flat_row = []
    for item in row:
        if isinstance(item, list):
            flat_row.extend(item)
        else:
            flat_row.append(item)
    return flat_row

"""Adderer opp alle antall på den varartiklen jeg definerer."""
def utstyrsteller():
    article_totals = defaultdict(lambda:defaultdict(float))
    for (root, dirs, files) in os.walk(test_folder):
        for file in files:
            if str(file).endswith('pdf'): #Sjekker at filen er en .pdf før jeg åpner den.
                print(f"ISO-Tegning: {file}")
                with pdfplumber.open(os.path.join(test_folder, file)) as pdf:  # Åpner PDF-en.
                    for page in pdf.pages: #Looper gjennom sidene i pdf'en
                        expected_id = 1

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
                                if not row or not row[0]: #var på 0
                                    continue

                                #row = flatten_row(row)

                                #Teller og går gjennom ID, at de kommer etterhverandre
                                try:
                                    current_id = int(row[0])
                                except ValueError:
                                    continue

                                if current_id != expected_id:
                                    print(f"Expected ID {expected_id} but found {current_id}. Skipping row: {row}")
                                    continue

                                expected_id += 1


                                mengde = row[1]
                                beskrivelse = row[-1] #enten 3 eller 4

                                if str.lower('KAPPLISTE') in str.lower(row[0]): #var før 0
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
                                    article_totals[rengjort_beskrivelse][file[:20]] += antall_type


                            # Print the summarized results
        for article, total in article_totals.items():
            print(f"Antall av ---{article}---: {total}")

    return article_totals

def compare_with_excel(article_totals, output_excel_path='forskjellsrapport.xlsx'):
    """
    Sammenligner PDF-data med Excel-data og lagrer resultatene i en Excel-fil.
    """
    all_differences = []
    all_articles = []
    # Load the Excel file
    for (root, dirs, files) in os.walk('check_folder'):
        for file in files:
            if str(file).endswith(excelending):  # Sjekker at filen er excel før jeg åpner den.
                print(f"Sammenligner med excelfil {file}")
                df = pd.read_excel(excel_path, sheet_name=sheeeet_name, skiprows=skipped_rows)

                for index, row in df.iterrows():
                    if pd.isna(row['Beskrivelse']):
                        continue
                    description = normaliserer_beskrivelser(row['Beskrivelse'])

                    if pd.isna(row['Mengde']):
                        excel_quantity = 0.0
                    else:
                        excel_quantity = row['Mengde']

                    # Get the quantity from the article_totals dictionary
                    pdf_quantities = article_totals.get(description, {})

                    formatted_iso_quantities = format_iso_quantities(pdf_quantities)

                    total_pdf_quantity = sum(pdf_quantities.values())

                    all_articles.append({
                        'Kilde': 'Excel',
                        'Beskrivelse': description,
                        'Excel mengde [mm]': excel_quantity,
                        'PDF mengde [m]': total_pdf_quantity,
                        "ISO fil": formatted_iso_quantities
                    })

                    # Compare the quantities
                    if total_pdf_quantity != excel_quantity:
                        all_differences.append({
                            'Kilde': 'Excel',
                            'Beskrivelse': description,
                            'Excel mengde [mm]': excel_quantity,
                            'PDF mengde [m]': total_pdf_quantity,
                            "ISO fil": formatted_iso_quantities
                        })

                # List of descriptions from Excel
                excel_descriptions = df['Beskrivelse'].dropna().apply(normaliserer_beskrivelser).tolist()

                # Check for items in PDFs but not in Excel
                for description, iso_files in article_totals.items():
                    if description not in excel_descriptions:
                        formatted_iso_quantities = format_iso_quantities(iso_files)
                        total_pdf_quantity = sum(iso_files.values())
                        all_articles.append({
                            'Kilde': 'PDF',
                            'Beskrivelse': description,
                            'Excel mengde [mm]': 0.0,
                            'PDF mengde [m]': total_pdf_quantity,
                            'ISO fil': formatted_iso_quantities
                        })

                        all_differences.append({
                            'Kilde': 'PDF',
                            'Beskrivelse': description,
                            'Excel mengde [mm]': 0.0,
                            'PDF mengde [m]': total_pdf_quantity,
                            'ISO fil': formatted_iso_quantities
                        })

    if all_articles:
        print("Alle artikler:")
        table = [
            [article['Kilde'], article['Beskrivelse'], article['Excel mengde [mm]'], article['PDF mengde [m]']]
            for article in all_articles
        ]
        print(tabulate.tabulate(table, headers=["Kilde", "Beskrivelse", "Excel Mengde [mm]", "PDF Mengde [m]"]))

        df_artikler = pd.DataFrame(all_articles)
        df_artikler.to_excel(output_excel_path, index=False)
        print(f"Artikler lagret på {output_excel_path}")




# Kaller funksjonene
utstyrsteller()
#article_totals = utstyrsteller()
#compare_with_excel(article_totals)


"""Debug om jeg skulle trenge det. Sjekker hvordan scripted leser pdf-filene"""


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

#debug()