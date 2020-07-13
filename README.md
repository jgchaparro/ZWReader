# ZWReader

ZWReader is a Python script that allows the user to extract and translate in advance every word of a Chinese Wikipedia article.
Its aim is to help students learning Chinese to read Wikipedia articles without having to check up every word. 
It is **not intended to work as a translator**, but as a learning tool. The code relies on the Jieba algorithm to identify words,  which may cause incorrect splitting.

## Features

• Translation for every word and character in the article before even reading it.
• Translation for every paragraph.
• Separate translation of the characters of the word.
• Allows dynamic selection of characters to translate. 
• Color codes to help searching words.  
• Ignore desired words or characters by adding them to a .txt file.

## Basic functionting

1) Clone the repository. Make sure the script and the folders are in the same directory.
2) Copy to the clipboard the full URL of the page you want to translate. **Do not alter the clipboard**
3) Run the script. It will create an Excel file with the translated words.
4) Open the output file localed in the Output_file folder.

## Modules required

• jieba ver. 0.42.1  
• openpyxl ver. 3.0.3  
• requests ver. 2.22.0  
• re ver. 2.2.1  
• os  
• bs4 ver. 4.8.2  
• pyperclip ver. 1.8.0  
• pandas ver. 1.0.1
• pyautogui
• googletrans  

## Color codes

None, white:  
1, light orange: single character.  
2, blue: two or more characters with a common translation.  
3, light red: word not present as a whole in the dictionary and divided into it's individual characters.  
6, black: end of a paragraph.  
7, yellow: number. Used as a reference to find words faster in both the Excel and the article.
10, dark blue: dot. Used as a reference.
11, purple: Chinese brackets. Used as a reference.

# ZWReader_2
