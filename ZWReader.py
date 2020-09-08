# Author: Jaime García Chaparro
# Version: 6.4

# Intended for personal use only.
# I do not own the content of the dictionaries used in the code.
# I am not responsable for any irregular use of the code or the files used.

#-----------------------------------
# Modify basic parameters:

output_file_name = 'ZHWikiReader_output_file' # Do not add the file extension.
always_slice = False # If true, gives the definition of the individual characters of the word.
traditional = True # False for using simplified in the output file.

translation = True

smart_slicing = True # If true, slices characters with a frequency lower than ss_threshold
ss_threshold = 20
ss_decrease = 1
ss_minumum = 0

#-----------------------------------
# DO NOT MODIFY THE CODE BELOW THIS
# LINE UNLESS IT IS NECESSARY
#-----------------------------------

# coding=utf-8

import re
import jieba
import openpyxl
from openpyxl.styles import Alignment
import requests
import re
import sys
import os
from bs4 import BeautifulSoup
import pyperclip as p
import pandas as pd
import pyautogui
from googletrans import Translator

#Basic counters

counter = 2
procedence_counter = 5
last_procedence = 6

script_dir = os.path.dirname(__file__)
jieba.set_dictionary(os.path.join(script_dir, 'Files', 'dict.txt.big.txt'))

#Define functions

    # General functions

def generate_lists():
    """Creates a list of the words in the dicitonary."""

    print('Generating list and data frame.')

    global df_words
    global full_dic
    global full_dic_simp
    df_words = pd.read_csv(os.path.join(script_dir, 'Files', 'Dictionary 3.2.csv'), sep='\\', encoding='utf-8')
    df_words = df_words.sort_values(by=['freq', 'pinyin'], ascending=[False, False])
    full_dic = df_words.loc[:,'trad'].tolist()
    full_dic_simp = df_words.loc[:,'simp'].tolist()

    print('Lists and data frames generated.')

def obtain_raw_text(raw_url):
    """Extracts the raw HTML text from Wikipedia."""
    
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
    }

    print('Obtaining text.')

    if 'https://zh.wikipedia.org/' in raw_url:
        raw_url = 'https://zh.wikipedia.org/zh-tw/' + raw_url.split('/')[-1]

    if raw_url[-4:].lower() == '.pdf':
        pars = extract_from_pdf(raw_url)
    else:
        with requests.get(raw_url, headers=headers) as page:
            soup = BeautifulSoup(page.content, 'html.parser')
            pars = soup.find_all("p")

    print('Text extracted.')

    return pars

def extract_from_pdf(raw_url):
    from tika import parser
    pattern = re.compile('https?://')
    if bool(pattern.match(raw_url)):
        PDF_arch = requests.get(raw_url)
        with open(os.path.join(script_dir, 'Files', 'ZWR_PDF.pdf'), 'wb') as PDF_file:
            PDF_file.write(PDF_arch.content)
        PDF_parse = parser.from_file(os.path.join(script_dir, 'Files', 'ZWR_PDF.pdf'))
    else:
        PDF_parse = parser.from_file(raw_url[8:].replace('%20', ' '))

    PDF_text = PDF_parse['content']
    pars = [sentence.replace('\n', '').strip() for sentence in PDF_text.split('\n') if len(sentence.replace('\n', '').strip()) > 0]

    return pars

def translate(pars):
    t_counter = 2
    translator = Translator()
    n_pars = len(pars)

    for par in pars:
        i = t_counter - 1
        print(f'Translating paragraph {i} out of {n_pars}.')
        try:
            par = par.text
        except:
            pass

        translation = translator.translate(par).text


        c_original = 'A' + str(t_counter)
        c_translation = 'B' + str(t_counter)

        zh_cell = trans[c_original]
        eng_cell = trans[c_translation]

        zh_cell.value = par
        zh_cell.alignment =  Alignment(wrapText=True, vertical = 'center')
        eng_cell.value = translation
        eng_cell.alignment = Alignment(wrapText=True, vertical = 'center')
        t_counter += 1     

def clean_and_slice(pars):
    """Removes punctuation characters, as well as other symbols."""

    global words

    for par in pars:
        try:
            raw_1 = re.sub(r"\[[A-z]*\]", "", par.text.strip()) #Removes anotations.
        except:
            raw_1 = re.sub(r"\[[A-z]*\]", "", par.strip()) # For PDFs
        raw_2 = re.sub(r"[A-z]*", "", raw_1) #Removes Latin characters.

        to_clean = ["|", "‧", "．", "○", "…", '\"', "─", "〉", "〈", " ", "﻿", "#", ":", "％", "~", ")", "(", ";", "/", "′", "°", "《", "》", "%", "%", "-", "“", "”", "－", "·", "\\", " ", ".", "︰", "（", "）", "；", "、", ",", "："]
        clean_text = raw_2
        for symbol in to_clean:
            if symbol in raw_2:
                clean_text = clean_text.replace(symbol, "")
        
        clean_text += '\n'
        print('Text cleaned.')

        # Extract words using Jieba algorithm.

        cut_data = jieba.cut(clean_text)
        cut_words = " ".join(cut_data).split(" ")
        
        for word in cut_words:
            words.append(word)
        
    print('Words cut.')

def detect_simp(words):
    sample_size = 50
    simp_threshold = 0.2

    words_to_evaluate = [word for word in words[:sample_size] if word != '\n']
    sample_size = len(words_to_evaluate)
    total_zis = 0
    total_errors = 0

    for i in range(sample_size):
        for i2 in range(len(words_to_evaluate[i])):
            if type(words_to_evaluate[i][i2]) == str:
                total_zis += 1
                if words_to_evaluate[i][i2] not in full_dic:
                    total_errors += 1
    
    errors_to_zis = total_errors/total_zis

    if errors_to_zis > simp_threshold and 'https://zh.wikipedia.org/' not in p.paste():
        print('Switching to simplified output.')
        switch_to_simp()

def switch_to_simp():
    jieba.set_dictionary(os.path.join(script_dir, 'Files', 'dict.txt.small.txt'))
    global full_dic
    global full_dic_simp
    global traditional
    full_dic = full_dic_simp
    traditional = False

def main():
    """Search the words or characters in the lists."""
    
    detect_simp(words)

    n_words = len(words)

    global counter
    i = 1

    for word in words:
        print(f'Processing word {i} out of {n_words}')
        i += 1
        if word == '\n':
            if last_procedence not in [6, 10, 11]:
                add_to_excel('', '', '', procedence = 6)
            elif last_procedence == 10:
                counter -= 1
                add_to_excel('', '', '', procedence = 6)
            elif last_procedence == 11 and temp['G' + str(counter - 2)].value == 10:
                counter -= 2
                add_to_excel('', '', '', procedence = 6)
            
        elif word in ['「', '」', '，', '。']:
            if word == '，':
               #add_to_excel('', '', '', procedence = 9)
               pass
            elif word == '。':
                add_to_excel('', '', '', procedence = 10) 
            elif last_procedence != 11:
                add_to_excel('', '', '', procedence = 11)
        elif re.search('[0-9]', word) != None:
            if last_procedence != 7:
                add_to_excel(int(word[:4]), '', '', procedence = 7)
        elif word != '':
                process(word)

    # -----------------------------------------------------------------------------------
    # Word functions
    # -----------------------------------------------------------------------------------

#-------------------------

def process(word, is_zi = False, procedence = 2):
    if len(word) == 1:
        if is_zi == True:
            if procedence != 3:
                procedence = ''
        else:   # Set procedence to 1 if non-zi word is 1 char long.
            procedence = 1

    if word in current_words:
        retrieve_from_current(word, is_zi, procedence)

    elif in_ignore_words(word, is_zi):
        pass

    elif word in ignores_dont_try:
        if is_zi == False:
            #print('In not-to-try list. Trying to rescue.')
            rescue_word(word)
        else:
            #print('Character in ignore list.')
            add_to_excel(word, 'X', 'Character in ignore list', procedence = 3)

    elif word in full_dic:
        retrieve_from_dictionary(word, is_zi, procedence)
    
    else:
        out_of_dictionary(word, is_zi)

def extract_info(index):
    if traditional == True:
        word = df_words.iloc[index, 0]
    else:
        word = df_words.iloc[index, 1]

    pinyin = df_words.iloc[index, 2]
    defs = df_words.iloc[index, 3]
    
    return word, pinyin, defs

def retrieve_from_current(word, is_zi = False, procedence = 2):
        i = current_indices[current_words.index(word)]
        #print('Retrieving from current.')

        word, pinyin, defs = extract_info(i)
        add_to_excel(word, pinyin, defs, index = i, is_zi = is_zi, procedence = procedence)

        if always_slice == True and is_zi == False and len(word) > 1:
            slice_into_zis(word, procedence = "")

def in_ignore_words(word, is_zi = False):
    if is_zi == False and word in ignores_words:
        #print('In ignored words.')
        return True
    elif is_zi == True and word in ignores_zi:
        #print('In ignored characters.')
        return True
    else:
        return False

def check_if_in_dont_try(word, is_zi = False, procedence = 2):
    ignores_dont_try.index(word)
    if is_zi == False:
        #print('In not-to-try list. Trying to rescue.')
        rescue_word(word)
    else:
        #print('Character in ignore list.')
        add_to_excel(word, 'X', 'Character in ignore list', 3)

def rescue_word(word):
    if len(word) == 2:
        slice_into_zis(word, procedence = 3)

    elif len(word) == 3:
        if word[:2] in full_dic:
            process(word[:2])
            process(word[2])
            #print('Rescued.')
        elif word[1:] in full_dic:
            process(word[0])
            process(word[1:])
            #print('Rescued.')
        else:
            slice_into_zis(word, procedence = 3)

    elif len(word) == 4:
        if word[:2] in full_dic and word[2:] in full_dic:
            process(word[:2])
            process(word[2:])
            #print('Rescued.')
        elif word[:2] in full_dic and word[2:] not in full_dic:
            process(word[:2])
            combined_pinyin = ''
            for zi in word[2:]:
                try:
                    combined_pinyin += (df_words.iloc[(full_dic.index(zi)), 2] + ' ')
                except:
                    combined_pinyin += ' X '
            add_to_excel(word[2:], combined_pinyin, 'X', procedence = 3)
            #print('Partially rescued.')
        elif word[:2] not in full_dic and word[2:] in full_dic:
            combined_pinyin = ''
            for zi in word[:2]:
                try:
                    combined_pinyin += (df_words.iloc[(full_dic.index(zi)), 2] + ' ')
                except:
                    combined_pinyin += ' X '
            add_to_excel(word[:2], combined_pinyin, 'X', procedence = 3)
            process(word[2:])
            #print('Partially rescued.')
        elif word[:-1] in full_dic:
            process(word[:-1])
            process(word[-1])
            #print('Rescued.')
        elif word[1:] in full_dic:
            process(word[0])
            process(word[1:])
            #print('Rescued.')
        else:
            slice_into_zis(word, procedence = 3)
    
    else:
        combined_pinyin = ''
        for zi in word:
            try:
                combined_pinyin += (df_words.iloc[(full_dic.index(zi)), 2] + ' ')
            except:
                combined_pinyin += ' X '
        add_to_excel(word, combined_pinyin, 'X', procedence = 3)

def retrieve_from_dictionary(word, is_zi = False, procedence = 2):
    df_i = full_dic.index(word)
    #print('Retrieving from dictionary.')

    word, pinyin, defs = extract_info(df_i)

    if is_zi == False:
        add_to_excel(word, pinyin, defs, index = df_i, procedence = procedence)
    else:
        add_to_excel(word, pinyin, defs, procedence = procedence, is_zi = True)

    add_to_current(word, df_i)

    if always_slice == True and is_zi == False and len(word) > 1:
        slice_into_zis(word, procedence = "")

def slice_into_zis(word, procedence = ''):
    for i in range(0, len(word)):
        process(word[i], is_zi = True, procedence = procedence)

def out_of_dictionary(word, is_zi = False):
    if is_zi == False:
        #print('Not in dictionary. Slicing.')
        rescue_word(word)

    else:
        #print('Character not in dictionary.')
        add_to_excel(word, 'X', 'Character not in dictionary.', procedence = 3)

# Add to excel functions

def add_to_excel(word, pinyin, defs, index = None, is_zi = False, procedence = 2):

    #Prepare cells id

    c_zh = 'A' + str(counter)
    c_pin = 'B' + str(counter)
    c_en = 'C' + str(counter)
    c_proc = 'G' + str(counter)

    #Add values to cells

    temp[c_zh].value = word
    temp[c_pin].value = pinyin
    temp[c_en].value = defs
    temp[c_proc].value = procedence

    add_counter()
    global last_procedence
    last_procedence = procedence

    if type(word) == str and is_zi == False:
        if smart_slicing == True and len(word) > 1 and defs != 'X':
            smart_slice(word)

    if index != None:
        add_count(word, index)

def add_to_current(word, index):
    current_words.append(word)
    current_indices.append(index)

def add_count(word, index):
    try:
        df_words.iloc[index, 4] += 1
    except:
        pass

    if len(word) > 1:
        for i in range(len(word)):
            try:
                df_words.iloc[full_dic.index(word[i]), 4] += 1
            except:
                pass

def smart_slice(word):
    ac_freq = df_words.iloc[full_dic.index(word), 4]
    if len(word) > 1 and ac_freq < ss_threshold:
        for i in range(len(word)):
            zi_ac_freq = df_words.iloc[full_dic.index(word[i]), 4]
            if zi_ac_freq < ss_threshold:
                process(word[i], is_zi = True, procedence = '')

def add_counter():
    global counter
    counter += 1

#------------------------
#   MAIN CODE
#-------------------------

#Load working predefined lists

generate_lists()

    # Ignore lists

with open(os.path.join(script_dir, 'Files', 'ignore.txt'), "r", encoding="utf-8") as ignores_txt:
    ignores_words = ignores_txt.read().split(" ")

with open(os.path.join(script_dir, 'Files', 'ignore_zi.txt'), "r", encoding="utf-8") as ignores_zi_file:
    ignores_zi = ignores_zi_file.read().split(" ")

with open(os.path.join(script_dir, 'Files', 'ignore_not_in_dic_wordlist.txt'), "r", encoding="utf-8") as dont_try_file:
    ignores_dont_try = dont_try_file.read().split(" ")

    # Create empty lists for the current article words, zis, pinyins and definitions

current_words, current_indices = [], []

#Load template

temp_file = openpyxl.load_workbook(os.path.join(script_dir, 'Files', 'template.xlsx'))
temp = temp_file['ZH']
trans = temp_file['Translation']


#Extract article text

words = []

pars = obtain_raw_text(p.paste())
clean_and_slice(pars)
if translation == True:
    translate(pars)

#Execute

main()

# Create output file

full_output_filename = output_file_name + '.xlsx'
temp_file.save(os.path.join(script_dir, 'Output_file', full_output_filename))

# Add changes to dataframe

df_words.freq -= ss_decrease
df_words.loc[df_words.loc[:,'freq'] <= ss_minumum, 'freq'] = ss_minumum
df_words.to_csv(os.path.join(script_dir, 'Files', 'Dictionary 3.2.csv'), sep='\\', encoding='utf-8', index=False)
print('Changes to csv made.')

pyautogui.hotkey('shift', 'alt', 'e')