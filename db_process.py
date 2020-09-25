"""Author : Allan D Maseghe/ allandavidke@gmail.com"""
"""Databases are walkin.xlsx and website.xlsx"""

from dBModule import *
from openpyxl import load_workbook
import time
import os

def main():
    refresh_home()
    while True:
        menu()
        choice = input("\nMenu option:")

        if choice == "0":
            print("Exiting the program....\n")
            break

        elif choice == "1":
            #Office walkin-in main info
            w_book,w_sheet = CreateVars("walkin.xlsx","main").vars()
            names_list = ListsClass(w_sheet).names()
            first_name = input("Client first name[without the '']:")

            if first_name in names_list:
                print("==============MAIN DETAILS===================")
                DbSearch(first_name,names_list,w_sheet)
            else:
                errors_function() 

        elif choice == "2":
            #Office walkin-in main info
            w_book,w_sheet = CreateVars("walkin.xlsx","services").vars()
            names_list = ListsClass(w_sheet).names()
            first_name = input("Client first name[without the '']:")

            if first_name in names_list:
                print("==============SERVICES DETAILS===================")
                DbSearch(first_name,names_list,w_sheet)
            else:
                errors_function() 

        elif choice == "3":
            #Office walkin-in main info
            w_book,w_sheet = CreateVars("website.xlsx","Sheet1").vars()
            names_list = ListsClass(w_sheet).names()
            first_name = input("Client first name[without the '']:")

            if first_name in names_list:
                print("==============MAIN DETAILS===================")
                DbSearch(first_name,names_list,w_sheet)
            else:
                errors_function() 
            
        elif choice =="4":
            #New client info
            print("\nALERT!:Mistakes will be corrected with Menu Option 5")
            try:
                NewClient("Nothing")

            except:
                """Handle funny exceptions like
                openpyxl.utils.exceptions.IllegalCharacterError"""
                print("\nERROR! Did you enter a special character?")
   
        elif choice=="5":
            #Edit existing client info
            w_book,w_sheet = CreateVars("walkin.xlsx","main").vars()
            names_list = ListsClass(w_sheet).names()
            first_name = input("Client first name[without the '']:")
            
            if first_name in names_list:
                sections_menu()
                section = input("Section to edit[enter 0,1 or 2 from menu]:")
                if section == "0":
                    print("Exiting editing section......")
                       
                elif section == "1":
                    print("==============MAIN DETAILS===================")
                    DbSearch(first_name,names_list,w_sheet)
                    try:
                        EditInfo(first_name,names_list,"main")
                        DuplicateInfo(first_name,names_list,"s_main")
                        #Last parameter is a marker for within the class use.
                    except:
                        """Handle the many errors like ValueError"""
                        errors_function()
                        
                elif section == "2":
                    print("==============SERVICES DETAILS===================")
                    #w_b = w_book & w_s = w_sheet. Changed to conform to PEP8
                    w_b,w_s = CreateVars("walkin.xlsx","services").vars()
                    DbSearch(first_name,names_list,w_s)
                    try:
                        EditInfo(first_name,names_list,"services")
                        DuplicateInfo(first_name,names_list,"s_services")
                        #Last parameter is a marker for within the class use.

                    except:
                        """Handle the many errors like ValueError"""
                        errors_function()

                else:
                    print("ERROR!!Choice should be 0,1 or 2")
                
            else:
                errors_function()
        else:
            errors_function()
    refresh_home()

def menu():
    print("\n==========MAIN MENU=========\n")
    print("0 = EXIT")
    print("1 = VIEW MAIN DETAILS OFFICE WALK-IN DATABASE")
    print("2 = VIEW OFFERED SERVICES & COMMENTS IN OFFICE WALK-IN DATABASE")
    print("3 = VIEW MAIN DETAILS IN WEBSITE DATABASE")
    print("4 = ADD NEW CLIENT INTO WALKIN DATABASE")
    print("5 = UPDATE CLIENT INFORMATION IN OFFICE WALK-IN DATABASE ")
    print("\n===================================\n")

def errors_function():
    """Print search errors message """
    print("\nERROR!Not in the database or MENU!!")
    print("\nError could be because of:\n1)Caps or small letters\
    \n2)Spaces after name ]\n3)A letter entered instead of number")
 
def sections_menu():
    print("\n========Edit Section Menu======")
    print("0: Exit editing ")
    print("1: Edit Main Information Section")
    print("2: Edit Services Information Section\n")
    print("====================================\n")
    
def refresh_home():
    cmd = "cp /opt/python/*xlsx /home/netops"
    os.system(cmd)
    print("\nDB updated for winSCP\n")

main()
