from os import listdir
import os
import pandas as pd
from math import ceil
import json
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
import os

#Writing a function for combining multiple CSV files into one file
def combine(path_to_dir):
    #Used Try-except blocks for Error Handling
    try:
        path_to_dir=path_to_dir+'\\'
        suffix=".csv"
        #Extracting all csv files in the selected folder
        filenames= listdir(path_to_dir)
        filenames= [ filename for filename in filenames if filename.endswith( suffix ) ]
        d={}
        #creation of json data
        for i in filenames:
            results = pd.read_csv(path_to_dir+i)
            d[i]=len(results)
        #concating all csvfiles
        combined_csv = pd.concat([pd.read_csv(path_to_dir+f) for f in filenames ])
        #saving into a single CSV file
        combined_csv.to_csv( path_to_dir+"combined_csv.csv", index=False, encoding='utf-8-sig')
        #Converting the csv file into Excel File..
        #this function is working fine you can un comment lines and check...as my system is taking too much time so i commented theese lines during initial testing
        '''writer = pd.ExcelWriter('test.xlsx')
        for i in range (ceil(len(combined_csv)/1000000)):
            print(i)
            data=pd.DataFrame(combined_csv[i*1000000:(i+1)*1000000])
            print(len(data))
            data.to_excel(writer, sheet_name='Sheet'+str(i))
            print(i)
        writer.save()'''
        json_object=json.dumps(d, indent=4)
        return [1, 'Combining CSV files is Successful',json_object]
    except Exception as e:
        return [e]

#Writing a function for merging and grouping the data in a given CSV file
def merge(path_to_file):
    #Used try-except blocks for error Handling
    try:
        df=pd.read_csv(path_to_file)
        path=os.path.abspath(os.path.join(path_to_file, os.pardir)) #Getting Parent Directory to store the newlycreated file in same directory
        path=path+"\\"
        #Groupby Function Usage
        kk=df.groupby(['Doc.Num'],as_index=False).agg({'Itm.no':lambda x: ' '.join(map(str,x)), 'Bill Date': ' '.join,'Currency': lambda x: x.mode(),'Payer City': lambda x: x.mode() ,'Invoice Quantity': 'sum','MRP Value': 'sum', 'Cost Price': 'sum', 'Taxable Amount': 'sum', 'Total Tax': 'sum', 'Total Amount':'sum', 'MRP Price':'sum' })
        #Converting the Final Output to CSV again
        kk.to_csv( path+"final_output.csv", index=False, encoding='utf-8-sig')
        return [1, 'Merging and Grouping the File is Successful']
    except Exception as e:
        return [e]




# Create an instance of tkinter frame
win = Tk()
#Creating a text widget for showing json data of combine function
text = Text(win, height=10,width=40)  
#Creating Scrollbar for text widget
sb = Scrollbar(win,orient=VERTICAL)
# Set the geometry of tkinter frame
win.geometry("700x350")
x=Label(win, text="selected Path : ", font=('Aerial 8'))
x.config( fg= "blue")
x.grid(row=1,column=2)


#Function to Browse a File(Used for Merge function)
def open_file():
   file = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
   if file:
      filepath = os.path.abspath(file.name)
      x.config(fg="red", text="Selected Path : "+str(filepath))
      #Creating a submit button which will be showed after file selection
      ttk.Button(win, text="Submit", command=lambda:[successMessage(merge(filepath))]).grid(row=4,column=4)

#Function to Browse a Folder(Used for Combine Function)
def open_folder():
   file = filedialog.askdirectory()
   if file:
      filepath = os.path.abspath(file)
      x.config(fg="green", text="Selected Path : "+str(filepath))
      #Creating a submit button which will be showed after folder selection
      ttk.Button(win, text="Submit", command=lambda:[successMessage(combine(filepath))]).grid(row=2,column=4)

# Create a function for showing Process Status (Success Message / Error Message)
def successMessage(output):
    print(type(output),output)
    if output[0]==1:
        msg = Label(win, text=output[1], font=('Georgia 9'))
        msg.config(fg='green')
        msg.grid(row=5, column=2)
        if len(output)==3:
            text.insert(END,output[-1])
            text.grid(row=5,column=3)     
            sb.grid(row=5, column=4, sticky=NS)
            text.config(yscrollcommand=sb.set)
            sb.config(command=text.yview)
        else:
            text.destroy()
            sb.destroy()
    else:
        msg = Label(win, text=output, font=('Georgia 9'))
        msg.config(fg='red')
        msg.grid(row=5, column=2)



# Creating Labels
label = Label(win, text="Click the Button to browse the Files", font=('Georgia 13'))
label.grid(row=0, column=2)

label1 = Label(win, text="Combine CSV Files", font=('Georgia 13'))
label1.grid(row=2, column=2)

label2 = Label(win, text="Merge File", font=('Georgia 13'))
label2.grid(row=4, column=2)

# Create a Buttons
ttk.Button(win, text="Browse", command=open_folder).grid(row=2,column=3)  #Browse Button for Combine 
ttk.Button(win, text="Browse", command=open_file).grid(row=4,column=3)   #Browse Button for Merge

win.mainloop()