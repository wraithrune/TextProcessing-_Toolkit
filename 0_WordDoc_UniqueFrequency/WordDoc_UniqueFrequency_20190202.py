# -*- coding: utf-8 -*-
"""

Created on Sat Feb  2 06:17:30 2019

@author: Ho Wei Jing

Description: Document loader to extract unique word tokens and frequency count of each unique word tokens

Objectives:
    1) Load word document
    2) Split text into tokens via white space
    3) Remove tokens that are punctuations
    4) Find unique word tokens
    5) Find frequency count of each unique word tokens

"""

from docx import Document
import pandas as pd

def fWordDocumentLoad_v1(vInput):
    # Step 1: Setup variable paths for loading files
    vStopWordsPath = "Dictionaries/Dictionary_Stopwords.txt"
    vPunctuationsPath = "Dictionaries/Dictionary_Punctuations.txt"
    vEncodingPath = "Dictionaries/Dictionary_Encodings.txt"
    vOutputPath = "Output/UniqueWordsFrequency.csv"
    
    # Step 2: Load files from variable paths
    vLoadDocument = Document(vInput)
    vStopWords = open(vStopWordsPath, "r", encoding="utf-8")
    vPunctuations = open(vPunctuationsPath, "r", encoding="utf-8")
    vEncodings = open(vEncodingPath, "r", encoding="utf-8")
    
    # Step 3: Prepare stopwords, punctuations and encoding lists
    vStopWordsArray = vStopWords.read().splitlines()
    vStopWords.close()
    
    vPunctuationsArray = vPunctuations.read().splitlines()
    vPunctuations.close()
    
    vEncodingArray = vEncodings.read().splitlines()
    vEncodings.close()
    
    # Step 4: Remove stopwords, punctuations and encodings from text and sort unique words with their frequency counts
    vDocText = []
    vUniqueWords = {}
    
    for p in vLoadDocument.paragraphs:
        vDocText.append(p.text)
    
    for a in range(0, len(vDocText)):
        vTempString = str(vDocText[a])
    
        # 4.1 Replace words that are in punctuations list
        for i in range(0, len(vPunctuationsArray)):
            vTempPunctuation = str(vPunctuationsArray[i])
    
            vTempString = vTempString.replace(vTempPunctuation, ' ')
            vTempString = vTempString.replace('  ', ' ')
        
        # 4.2 Replace words that are in encodings list
        for i in range(0, len(vEncodingArray)):
            vTempEncoding = str(vEncodingArray[i])
    
            vTempString = vTempString.replace(vTempEncoding, '')
            vTempString = vTempString.replace('  ', ' ')
    
        # 4.3 Store all unique words
        vTempArray = vTempString.split(' ')
        for j in range(0, len(vTempArray)):
            
            # 4.3.1 Replace words that are in stopwords list
            for k in range(0, len(vStopWordsArray)):
                if str(vStopWordsArray[k]).lower() == str(vTempArray[j]).lower():
                    vTempArray[j] = ""
                    
            if str(vTempArray[j]) != "":
                vCount = 1
                # 4.3.2 If key (unique word in lower case) exists in dictionary, else...
                if str(vTempArray[j]).lower() in vUniqueWords:
                    vCount = vUniqueWords[str(vTempArray[j]).lower()]
                    vCount = vCount + 1
                    vUniqueWords[str(vTempArray[j]).lower()] = vCount
                else:
                    vUniqueWords[str(vTempArray[j]).lower()] = vCount
    
    # Step 5: Fit dictionary into dataframe
    # 5.1 Convert dictionary into arrays
    vListforDictionary_Key = []
    vListforDictionary_Value = []
    
    for key in vUniqueWords:
        vListforDictionary_Key.append(key)
        vListforDictionary_Value.append(vUniqueWords[key])
        
    # 5.2 Fit arrays into dataframe
    vExportCSV = pd.DataFrame()
    vExportCSV["Word"] = vListforDictionary_Key
    vExportCSV["Frequency"] = vListforDictionary_Value

    # 5.3 Export dataframe into csv
    vExportCSV.to_csv(vOutputPath, encoding='utf-8', index=False)
    
fWordDocumentLoad_v1('Input/TNNU Chapter 1 to 14.docx')