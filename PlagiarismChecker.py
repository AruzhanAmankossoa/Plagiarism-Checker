#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os
import docx2txt
import pandas as pd
import numpy as np
import math
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# creating tkinter window
root=tk.Tk()    
root.title("Plagiarism Checker")
root.geometry('400x320') #window with dimensions 400x320
root.configure(bg='#81c0f7')

#Title Plagiarism Checker
LabelTitle =  tk.Label(root,text= "Plagiarism Checker", 
                       font=('Times', 30, 'bold'),
                       padx=20,pady=30, 
                       justify = 'center',
                       bg='#81c0f7',fg='#45433e')
LabelTitle.grid(row = 0) #grid method to arrange labels in respective rows and columns as specified 

#Label "Select a path"
labelPath = tk.Label(root,text="Select a path",
                     font=("Times", 12),
                     bg='#81c0f7')
labelPath.grid(row=1)


list_documents = [] #contains name of documents
documents = [] #contains all the information contained in the documents[list_documents]


entry=tk.Entry(root, width=30, font=("Times", 14, "normal"))

def select():
        a  = filedialog.askdirectory(title='Please select a directory')
        path = "{}".format(a)
        
        entry.insert(0,path)
        
        # Open a file, read only .docx file and append to list_documents
        dirs = os.listdir(path)
        for f_name in dirs:
            if f_name.endswith('.docx'):
                list_documents.append(f_name)
        
        c= len(list_documents)
        for i in range(c):
            b = docx2txt.process(list_documents[i]) #extract text from docx files.
            documents.append(b)
            
entry.grid(row=2)

#Button Browse
buttonBrowse=tk.Button(root,text="Browse",
                       font=('helvetica', 12, 'bold'),
                       bg='#ffe11f',
                       height=1,
                       command=select)

buttonBrowse.grid(row=3, pady=(5,2))

#check for plagiarism and save in excel format
def checkAndsave():
    countVector = CountVectorizer() #convert a collection of text documents to a vector of term/token counts
    sparseMatrix = countVector.fit_transform(documents)
    termMatrix = sparseMatrix.todense()
    
    df = pd.DataFrame(termMatrix, 
                      columns=countVector.get_feature_names(),
                      index= list_documents)
                 
    arr = np.array(cosine_similarity(df, df)) # here the result is without a percentage
    rounded_array = np.around(arr*100, 0)     # therefore I multiplied by 100
    
    df = pd.DataFrame(rounded_array,
                      index = list_documents)
    
    df.columns = pd.RangeIndex(1, len(df.columns)+1) #start columns' index from 1
    
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    df.to_excel(export_file_path, index = True, header=True)

#Label Save
labelResult = tk.Label(root, font = ("Times", 12),
                       text = "Click a button to save the result",
                       bg = '#81c0f7')

labelResult.grid(row=4,pady = (25,2))

#Button Save
saveAsButtonExcel = tk.Button(text='Export a result to Excel', 
                              bg='#ffe11f', fg='black',
                              font=('helvetica', 12, 'bold'),
                              command=checkAndsave)

saveAsButtonExcel.grid(row=5)

root.mainloop()

