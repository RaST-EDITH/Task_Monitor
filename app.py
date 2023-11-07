# Library and Modules used 
import os                                                    # pip install os ( In case if not present )
import numpy as np                                           # pip install numpy ( In case if not present )
import pandas as pd                                          # pip install pandas==1.4.3
import openpyxl as oxl                                       # pip install openpyxl==3.0.10
from tkinter import *                                        # pip install tkinter==8.6
import customtkinter as ctk                                  # pip install customtkinter==4.6.3
from PIL import Image ,ImageTk                               # pip install pillow==9.3.0
import matplotlib.pyplot as plt                              # pip install matplotlib==3.4.3
from tkinter.messagebox import showerror, showinfo

class TaskMonitor :

    def __init__(self) :
        ctk.set_appearance_mode( "dark" )
        ctk.set_default_color_theme( "dark-blue" )
        self.width = 1200
        self.height = 700
        self.root = ctk.CTk()
        self.root.title( "Task Monitor" )
        self.root.geometry( "1200x700+200+80" )
        self.root.resizable( False, False )
        self.path = os.path.join( os.getcwd(), r"Data File\task_file.xlsx")
        xl = pd.ExcelFile( self.path )
        self.all_sheets = xl.sheet_names
    
    def change( self, can, page) :

        # Switching canvas
        can.destroy()
        page()

    def removeTask( self, indx, area, page) :
        
        task_sheet = pd.read_excel( pd.ExcelFile( self.path ), self.all_sheets[0])
        row, col = task_sheet.shape
        
        wb = oxl.load_workbook( self.path )
        sheet_xl = wb[self.all_sheets[0]]
        
        if ( indx != "" ) :

            indx = int(indx)
            if ( indx>0 ) and ( indx<=row ) :

                if ( sheet_xl[f"C{indx+1}"].value != "Done" ) :

                    for i in range( indx+1, row+1 ) :
                        sheet_xl[f"A{i}"].value = sheet_xl[f"A{i+1}"].value - 1
                        sheet_xl[f"B{i}"].value = sheet_xl[f"B{i+1}"].value
                    
                    sheet_xl[f"A{row+1}"].value = None
                    sheet_xl[f"B{row+1}"].value = None

                    try :

                        wb.save( self.path )
                        area.destroy()
                        area = ctk.CTkTextbox( page, 
                                                width = 850, height = 400, 
                                                text_font = ( "Georgia", 20 ), 
                                                state = "disabled"  )
                        area.place( x = 80, y = 230, anchor = "nw")
                        self.insertTask( area )
                        self.insertTaskAnalysis()
                    
                    except :
                        showerror( message = "Close Program related Files", title = "Open File found")

                else :
                    showerror( message = "Marked Task auto Remove Tommorow!", title = "Marked")

    def updateTask( self, indx, task, area, page) :

        task_sheet = pd.read_excel( pd.ExcelFile( self.path ), self.all_sheets[0])
        row, col = task_sheet.shape
        
        wb = oxl.load_workbook( self.path )
        sheet_xl = wb[self.all_sheets[0]]

        if ( indx != "" ) and ( len(task) > 0 ) :
        
            try :
                indx = int(indx)
                if ( indx == 1 ) :
                    for i in range( row+2, 2, -1 ) :
                        sheet_xl[f"A{i}"] = int(sheet_xl[f"A{i-1}"].value) + 1
                        sheet_xl[f"B{i}"] = sheet_xl[f"B{i-1}"].value
                    
                    sheet_xl[f"A{2}"] = indx
                    sheet_xl[f"B{2}"] = task
                
                elif ( indx > 1 )  :
                    for i in range( row+2, indx, -1 ) :
                        sheet_xl[f"A{i}"] = int(sheet_xl[f"A{i-1}"].value) + 1
                        sheet_xl[f"B{i}"] = sheet_xl[f"B{i-1}"].value
                    
                    sheet_xl[f"A{indx+1}"] = indx
                    sheet_xl[f"B{indx+1}"] = task
                
                else :
                    showerror( message = "Invalid Entry!", title = "Invalid")
                
                try :

                    wb.save( self.path )
                    area.destroy()
                    area = ctk.CTkTextbox( page, 
                                            width = 850, height = 400, 
                                             text_font = ( "Georgia", 20 ), 
                                              state = "disabled"  )
                    area.place( x = 80, y = 230, anchor = "nw")
                    self.insertTask( area )
                    self.insertTaskAnalysis()
                
                except :
                    showerror( message = "Close Program related Files", title = "Open File found")
            
            except :
                showerror( message = "Insert values of Valid Type", title = "Invalid value found")
        
        else :
            showerror( message = "Field Empty!", title = "Value Not Found")

    def taskMonitoringPage(self) :

        # Defining Structure
        taskMon_page = Canvas( self.root, 
                                width = self.width, height = self.height, 
                                 bg = "black", highlightcolor = "#3c5390", 
                                  borderwidth = 0 )
        taskMon_page.pack( fill = "both", expand = True )

        # Heading
        taskMon_page.create_text( 530, 130, text = "Task Monitoring", 
                                font = ( "Georgia", 42, "bold" ), fill = "#ec1c24" )

        # Task Index Box
        indx = ctk.CTkEntry( master = taskMon_page, 
                              placeholder_text = "Index", text_font = ( "Georgia", 20 ), 
                               width = 95, height = 30, corner_radius = 14,
                                placeholder_text_color = "#666666", text_color = "#191919", 
                                 fg_color = "#e1f5ff", bg_color = "black", 
                                  border_color = "white", border_width = 3)
        indx_win = taskMon_page.create_window( 200, 320-120, anchor = "nw", window = indx )
        
        # Task Entry Box
        task = ctk.CTkEntry( master = taskMon_page, 
                              placeholder_text = "Enter Task", text_font = ( "Georgia", 20 ), 
                               width = 550, height = 30, corner_radius = 14,
                                placeholder_text_color = "#666666", text_color = "#191919", 
                                 fg_color = "#e1f5ff", bg_color = "black", 
                                  border_color = "white", border_width = 3)
        task_win = taskMon_page.create_window( 325, 320-120, anchor = "nw", window = task )

        task_box = ctk.CTkTextbox( taskMon_page, 
                                    width = 850, height = 400, 
                                     text_font = ( "Georgia", 20  ), 
                                      state = "disabled"  )
        task_box.place( x = 80, y = 230, anchor = "nw")

        self.insertTaskAnalysis()
        self.insertTask( task_box )

        task.bind('<Return>', lambda event = None : self.updateTask( indx.get(), task.get(), task_box, taskMon_page ) )

        # Insert Button
        insert_bt = ctk.CTkButton( master = taskMon_page, 
                                   text = "Insert", text_font = ( "Georgia", 20  ), 
                                    width = 100, height = 40, corner_radius = 18,
                                     bg_color = "black", fg_color = "red", 
                                      hover_color = "#ff5359", border_width = 0, 
                                       command = lambda : self.updateTask( indx.get(), task.get(), task_box, taskMon_page ) )
        insert_bt_win = taskMon_page.create_window( 1030, 320-120, anchor = "nw", window = insert_bt )

        # Remove Task
        taskMon_page.create_text( 1100+20+130+50, 350-10, text = "Remove Task", 
                                font = ( "Tahoma", 18, "italic", "underline" ), fill = "white" )

        # Task Remove Box
        remove = ctk.CTkEntry( master = taskMon_page, 
                                placeholder_text = "Index", text_font = ( "Georgia", 20  ), 
                                 width = 180, height = 30, corner_radius = 14,
                                  placeholder_text_color = "#666666", text_color = "#191919", 
                                   fg_color = "#e1f5ff", bg_color = "black", 
                                    border_color = "white", border_width = 3)
        remove_win = taskMon_page.create_window( 1015+130+50, 370, anchor = "nw", window = remove )

        remove.bind('<Return>', lambda event = None : self.removeTask( remove.get(), task_box, taskMon_page ) )

        # Remove Button
        remove_bt = ctk.CTkButton( master = taskMon_page, 
                                    text = "Remove", text_font = ( "Georgia", 20  ), 
                                     width = 80, height = 40, corner_radius = 18,
                                      bg_color = "black", fg_color = "red", 
                                       hover_color = "#ff5359", border_width = 0, 
                                        command = lambda : self.removeTask( remove.get(), task_box, taskMon_page ) )
        remove_bt_win = taskMon_page.create_window( 1285, 440, anchor = "nw", window = remove_bt )

        # Mark Done Task
        taskMon_page.create_text( 1115+20+130+50, 550+50, text = "Mark Done Task", 
                                font = ( "Tahoma", 18, "italic", "underline" ), fill = "white" )

        # Task Done Box
        tkdone = ctk.CTkEntry( master = taskMon_page, 
                                placeholder_text = "Index", text_font = ( "Georgia", 20 ), 
                                 width = 180, height = 30, corner_radius = 14,
                                  placeholder_text_color = "#666666", text_color = "#191919", 
                                   fg_color = "#e1f5ff", bg_color = "black", 
                                    border_color = "white", border_width = 3)
        tkdone_win = taskMon_page.create_window( 1015+130+50, 580+50, anchor = "nw", window = tkdone )

        tkdone.bind('<Return>', lambda event = None : self.statusTask( tkdone.get(), task_box, taskMon_page ) )
    
        # Done Button
        tkdone_bt = ctk.CTkButton( master = taskMon_page, 
                                   text = "Mark Done", text_font = ( "Georgia", 20 ), 
                                    width = 80, height = 40, corner_radius = 18,
                                     bg_color = "black", fg_color = "red", 
                                      hover_color = "#ff5359", border_width = 0, 
                                       command = lambda : self.statusTask( tkdone.get(), task_box, taskMon_page ) )
        tkdone_bt_win = taskMon_page.create_window( 1250, 650+50, anchor = "nw", window = tkdone_bt )

        # Analysis Button
        analysis_bt = ctk.CTkButton( master = taskMon_page, 
                                      text = "Analysis", text_font = ( "Georgia", 20 ), 
                                       width = 120, height = 40, corner_radius = 18,
                                        bg_color = "black", fg_color = "red", 
                                         hover_color = "#ff5359", border_width = 0, 
                                          command = lambda : self.taskAnalysis() )
        analysis_bt_win = taskMon_page.create_window( 650, 805, anchor = "nw", window = analysis_bt )

        # Return Button
        back_bt = ctk.CTkButton( master = taskMon_page, 
                                  text = "Back", text_font = ( "Georgia", 20 ),  
                                   width = 45, height = 45, corner_radius = 23, 
                                    bg_color = "#fcd7ab", fg_color = "red", 
                                     hover_color = "#ff5359", border_width = 0, 
                                      command = lambda : self.change( taskMon_page, self.firstPage ))
        back_bt_win = taskMon_page.create_window( 30, 20, anchor = "nw", window = back_bt )

        self.root.mainloop()

    def firstPage(self) :

        # Defining Structure
        first_page = Canvas( self.root, 
                              width = self.width, height = self.height, 
                               bg = "black", highlightcolor = "#3c5390", 
                                borderwidth = 0 )
        first_page.pack( fill = "both", expand = True )

        # Heading
        first_page.create_text( 400, 119, text = "Task Monitor", 
                                font = ( "Georgia", 42, "bold" ), fill = "#ec1c24" )

        # Next Page Button
        next_bt = ctk.CTkButton( master = first_page, 
                                  text = "Let's Go ->", text_font = ( "Tahoma", 20 ), 
                                   width = 100, height = 40, corner_radius = 18,
                                    bg_color = "#fecc8f", fg_color = "#ec1c24", 
                                     hover_color = "#ff5359", border_width = 0,
                                      text_color = "white",
                                       command = lambda : self.change( first_page, self.taskMonitoringPage ) )
        next_bt_win = first_page.create_window( 320, 720, anchor = "nw", window = next_bt )

        self.root.mainloop()

if __name__ == "__main__" :

    task_class = TaskMonitor()
    task_class.firstPage()
