from bs4 import BeautifulSoup as Bullshit
from os import chdir
import requests, csv
import xlwings as xw


def get_last_idx(sheet):
    cell_idx = 1
    is_empty = False

    while is_empty == False:
        cell = get_excel_cell(cell_idx, "B")
        if sheet.range(cell).value == None:
            is_empty = True
            cell_idx -= 1
        else:
            cell_idx += 1

    return cell_idx


def get_cell_value(excel_idx, column_letter):
    cell_name = get_excel_cell(excel_idx, column_letter)
    cell_value = sht.range(cell_name).value

    return cell_value

def get_excel_idx(python_idx):
    excel_idx = python_idx + 1 

    return excel_idx

def get_excel_cell(excel_idx, column_letter):
    excel_cell = "{}{}".format(column_letter, excel_idx)
    return excel_cell

def set_status(sheet, excel_idx, message):
    excel_cell = get_excel_cell(excel_idx, "D")
    sht.range(excel_cell).value = message
    return True

def append_csv_row(excel_idx, mp3_filename):
    article = get_cell_value(excel_idx, "A")
    french_word = get_cell_value(excel_idx, "B")
    translation = get_cell_value(excel_idx, "C")
    full_french_word = article + " " + french_word

    mp3_containter = " [sound:" + mp3_filename + "]"
    word_and_sound = full_french_word + mp3_containter

    csv_row = [word_and_sound, translation]
    all_words.append(csv_row)

    return csv_row

def create_csv(word_list):
    chdir(r'C:\Users\bartosz.gajewski\Desktop\Anki_decks')

    with open('anki_deck.csv', 'w', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile, delimiter="\t", quoting=csv.QUOTE_MINIMAL)

        for row in word_list:
            writer.writerow(row)

all_words = []

wb = xw.Book('Zeszyt1.xlsx')
sht = wb.sheets("Arkusz1")
last_idx = get_last_idx(sht)
excel_cell = get_excel_cell(last_idx, "B")
words = sht.range("B1:B29").value

print("Ready to debug!")

for idx, word in enumerate(words):
    excel_idx = get_excel_idx(idx)

    mp3_filename = word + ".mp3"

    word_link = "https://www.wordreference.com/fren/" + word
    whole_site = requests.get(word_link)
    site_html = whole_site.text


    site_soup = Bullshit(site_html, 'html.parser')
    audio_header = {"type": "audio/mpeg"}
    audio_container = site_soup.find('source', audio_header)

    try:
        audio_link = audio_container["src"]
        full_link = "https://www.wordreference.com" + audio_link

        file = requests.get(full_link, allow_redirects=True)

        chdir(r'C:\Users\bartosz.gajewski\AppData\Roaming\Anki2\Użytkownik 1\collection.media')

        set_status(sht, excel_idx, "done")
        open(mp3_filename, 'wb').write(file.content)

        append_csv_row(excel_idx, mp3_filename)


    except TypeError:
        set_status(sht, excel_idx, "ERROR")

create_csv(all_words)