from tkinter import *
from tkinter import ttk
from  PIL import Image, ImageTk
import os,sys
import pyodbc 
import tkinter.messagebox as tsmg
import datetime as dt
import cx_Oracle 
import sqlite3
import csv
import datetime


class Root_Sql_Editor(Tk):

    def __init__(self) :
         super().__init__()
         self.geometry("800x600")
         self.title("SQL QUERY EXECUTOR v1.1")
         photo = PhotoImage(file = r"C:\temp\sqllogo_ico.ico")
         self.iconphoto(False, photo)
         self.label = Label(self,text="SQL QUERY:")
         self.label.pack(anchor=NW)
        #  self.label = Label(self,text="Please Press ClearAll key before executing the new Query!",fg= "red",font=("callibri", 12))
        #  self.label.pack(anchor=NE)
         self.textBox1 = Text(self,width=200,height=5,background="light grey")
         self.textBox1.pack()
         self.label = Label(self,text="Select Your Data base Type:")
         self.label.pack(anchor=NW)
         self.db_select = StringVar()
         self.db_select.set(" ")
         
         self.radioBtn_sqllite = Radiobutton(self,text="SQLite3",variable=self.db_select,value="SQLite3")
         self.radioBtn_sqllite.pack(anchor=W)
         self.radioBtn_MySql = Radiobutton(self,text="MSSqlServer",variable=self.db_select,value="MSSqlServer")
         self.radioBtn_MySql.pack(anchor=W)
         self.radioBtn_ORACLE = Radiobutton(self,text="Oracle",variable=self.db_select,value="Oracle")
         self.radioBtn_ORACLE.pack(anchor=W)
        #  self.l1=[]

    def execute(self):

        if hasattr(self,'tree') and self.tree:# destroying already created tree if present
              self.tree.destroy()

        # self.dbconnection = DatabaseManager()
        self.logtime = dateTimeSet() #INIT date time CLASS
        try:
            dbManager = DatabaseManager() #INIT DATA BASE MANAGER CLASS
            print(self.db_select.get())
            if self.db_select.get() == "SQLite3":
                dbManager.connect_to_sqlite3()

            elif self.db_select.get() == "MSSqlServer":
                dbManager.connect_to_mssql()

            elif self.db_select.get() == "Oracle":
                dbManager.connect_to_oracle()
            
            else:
                tsmg.showerror("ERROR!", "Please Select your DataBase First.")


            self.query = self.textBox1.get('1.0','end-1c')
            self.result = dbManager.cursor.execute(self.query)
            self.column_list=[i[0] for i in dbManager.cursor.description]
        
            self.tree = ttk.Treeview(self,selectmode="extended",columns=self.column_list,show='headings',height=200)
            self.tree.pack(side=TOP,anchor='center',fill=BOTH,expand=True)
            self.y_scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscroll=self.y_scrollbar.set)
            self.y_scrollbar.pack(side="right", fill="y")
           
            self.x_scrollbar = ttk.Scrollbar(self.tree, orient="horizontal", command=self.tree.xview)
            self.tree.configure(xscroll=self.x_scrollbar.set)
            self.x_scrollbar.pack(side="bottom", fill="x")
            self.result = list(self.result)


            for self.i in self.column_list:
                
                self.tree.column(self.i,anchor='center',width=200)
                self.tree.heading(self.i,text=self.i)

            self.new_response= []
            for row in self.result:
                    self.new_row =[]
                    for item in row:
                        if isinstance(item,datetime.datetime):
                              self.new_row.append(item.strftime('%d/%m%Y%H:%M:%S'))
                        else:
                              self.new_row.append(item)
                    self.new_response.append(self.new_row)

            for row in self.new_response:
                    self.tree.insert("",END,values=row)
                   



        except Exception as ex:
            
            tsmg.showerror("Execution Error:", ex)
            try:
                with open(f"C://temp//ERROR.log_{self.logtime.dateTime()}", "a") as f:
                 f.write(f"{self.logtime.dateTime()}------>{ex}")
            except Exception as e:
                 tsmg.showerror("Error.", str(e) )
                 print(e)
        if self.db_select.get() == "SQLite3":          
            dbManager.close_connection()
        else:
             dbManager.sql_close_connection()
             

    def clearAll(self):
                 
                
                self.textBox1.delete('1.0','end-1c')
                self.db_select.set(" ")
                self.tree.destroy()
                self.y_scrollbar.destroy()
                self.x_scrollbar.destroy()
       
    

    def myButtons(self,text_btname,funcName,side,):
        runbtn = Button(self,text=text_btname,command=funcName,bg="light green",border="2")
        runbtn.pack(side=side,anchor=SW)



    def saveFile(self):
            self.date = dt.datetime.now()
            self.time = self.date.strftime("%d%m%Y%H%M%S")
            with open(f"C://temp//result{self.time}.csv", "w", newline='') as f:
                csv_writer = csv.writer(f)
                csv_writer.writerow(self.column_list)  # Write the column headers
            # giving one line space
                csv_writer.writerow('')
                for item in self.tree.get_children():
                    values = self.tree.item(item, "values")
                    csv_writer.writerow(values)
            tsmg.showinfo("Save Status", f"Data saved to c://temp//result{self.time}.csv")


    def exit_editor(self):
          
          self.destroy()

class DatabaseManager():
    def __init__(self):
        self.connection = None
        self.cursor = None

    def connect_to_sqlite3(self):
        self.connection = sqlite3.connect("mydatabase.db")
        self.cursor = self.connection.cursor()
        tsmg.showinfo("Connection status", "Connection established with SQLite3")

    def connect_to_mssql(self):
        with open ("C:\\ConnectionStrings.txt", "r") as f:
                self.connection_string = f.read().splitlines()  # reading file
                self.oracle_conncetion_string = self.connection_string[0] # reading from line 1 
                # print(self.oracle_conncetion_string)
                self.sqlConnection = pyodbc.connect(self.oracle_conncetion_string)
                self.cursor = self.sqlConnection.cursor()
        tsmg.showinfo("Connection status","Connection esatablished with MSSQL Data base")

    def connect_to_oracle(self):
        with open ("C:\\ConnectionStrings.txt", "r") as f:
                self.connection_string = f.read().splitlines() # reading file 
                self.oracle_conncetion_string = self.connection_string[1]# reading from line 2  
                # print(self.oracle_conncetion_string)
                self.sqlConnection = cx_Oracle.connect(self.oracle_conncetion_string)
                self.cursor = self.sqlConnection.cursor()
                tsmg.showinfo("Connection status","Connection esatablished with Oracle Data base")

 

    def close_connection(self):
        if self.connection:
            self.connection.commit()
            self.connection.close()
    
    def sql_close_connection(self):
        
            self.sqlConnection.commit()
            self.sqlConnection.close()


class dateTimeSet():
     def dateTime(self):
          self.dtime = dt.datetime.now()
          self.time_now = self.dtime.strftime("%d%m%Y%H%M%S")
          return self.time_now
     
    

if __name__== '__main__':
    
    sqlEditor = Root_Sql_Editor()
    sqlEditor.myButtons("Exit Editor",sqlEditor.exit_editor,RIGHT)
    sqlEditor.myButtons('EXECUTE SCRIPT',sqlEditor.execute,RIGHT)
    sqlEditor.myButtons('Save As CSV',sqlEditor.saveFile,LEFT)
    sqlEditor.myButtons("Clear All",sqlEditor.clearAll,LEFT)


    sqlEditor.mainloop()
