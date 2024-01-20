import sys
import os
import ezgmail
import openpyxl
import string
import bs4
import re
import os
from contextlib import suppress
import argparse
import tkinter as tk
from tkinter import filedialog

class PatchError(Exception): pass

def patch():
	site_packages = sys.path[5]
	ezgmail_code_path = os.path.join(site_packages, "ezgmail", "__init__.py")
	# __version__ = \"2022.10.10\"
	with open(ezgmail_code_path, "r") as f:
		code = f.read()
	if "__version__ = \"2022.10.10\"" in code:
		pass
	elif "__version__ = \"2022.10.10.PATCHED\"" in code:
		return
	else:
		raise PatchError("Unable to apply patch! Invalid version!")
	lines = code.split("\n")
	lines[6] = "__version__ = \"2022.10.10.PATCHED\""
	lines[491] = "#"+lines[491]
	lines[525] = "#"+lines[525]
	code_new = "\n".join(lines)
	with open(ezgmail_code_path, "w") as f:
		f.truncate()
		f.write(code_new)

def prettify_except(soup_obj: bs4.BeautifulSoup, tag_name: str) -> str:
    #https://stackoverflow.com/a/69589000/9091833
    regex_string = "<{0}>.*<\/{0}>".format(tag_name)
    regex = re.compile(regex_string, re.DOTALL)
    replacing_txt = str(getattr(soup_obj, tag_name))
    return re.sub(regex, replacing_txt, soup_obj.prettify())


#batch of 20, complete list, range of rows, start row then quant,
#add sent email date & time
#use sent date to confirm if email needs to be sent again

def mail_from_excel(mail_list, template, sig, _closing, _name):
    wb = openpyxl.load_workbook(mail_list)
    try:
        sheet = wb['Sheet1']
        data = tuple(sheet.rows)
    finally:
        wb.close()

    with open(template,'r') as infile:
            template = infile.read()
            soup = bs4.BeautifulSoup(template, 'html.parser')

    with open(sig,'r') as infile:
            sig = infile.read()

    soup.body.append(bs4.BeautifulSoup(sig, 'html.parser'))
    with open(r'Templates\message.html','w') as outfile:
            outfile.write(prettify_except(soup, 'body'))
    with open(r'Templates\message.html','r') as infile:
            message = infile.read()
            
    field_names = [v[1] for v in string.Formatter().parse(template)]

    print(f"Rows in sheet {sheet.max_row}")
    #Start at 1 to skip header, "name" and "email"
    for i in range (1, sheet.max_row):
            ezgmail.draft(data[i][1].value,f'Mail for {data[i][0].value}', message.format(closing=_closing, name=_name), mimeSubtype='html')

def logout():
    with suppress(OSError):
        os.remove("token.json")

def user_auth():
    while True:
        ezgmail.init()
        print("="*80)
        print("Python Emailer, you're currently logged in as: {email}".format(email=ezgmail.EMAIL_ADDRESS))
        print("="*80+"\n\n")

        #prompt for user change
        while True:
                logout_choice = input("Do you want to logout? ((y)es / (N)o): ").lower()
                if logout_choice in ("yes", "no" , "y", "n", ""):
                        break
                print("invalid input")
                
        if logout_choice in ("yes", "y"):
                logout()
                continue
        else:
                print("Continuing with email program... \n")
                break
        
def prompt_for_file(msg):
        while True:
                file_path = input(f"\n{msg}: ")
                if os.path.isfile(file_path):
                        return file_path
                else:
                        print ("Invalid file path. Please enter a valid file path.")

def select_file_dialog(prompt):
    print(prompt)
    return filedialog.askopenfilename(parent=root, initialdir=cwd)

def main():
    #patch()

    user_auth()
    cwd = os.getcwd()
    print(f"The cwd is: {cwd}")

    parser = argparse.ArgumentParser()
    parser.add_argument("-m","--mail_list", help="The mailing list excel sheet (.xlsx,.xlsm,.xltx,.xltm)")
    args = parser.parse_args()
    #args = parser.parse_args([r"Mail lists\example.xlsx"])

    #Determine if user will use gui tk file prompt or cmdline input
    while True:
        GUI_choice = input("Do you wish to use the GUI File picker? (Y)es/(n)o: ").lower()
        if GUI_choice in ("yes", "no", "y", "n", ""):
                break
        print("Invalid choice, please retry. ")
        
    if GUI_choice in ("yes", "y", ""):
        using_GUI_filedialog = True
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.focus_force()
    else:
        using_GUI_filedialog = False
    print("="*80+"\n")


    #get/prompt for mail list excel sheet (.xlsx,.xlsm,.xltx,.xltm)
    if args.mail_list:
            #print("Sys.argv found!")
            mail_list = args.mail_list
    else:
            if using_GUI_filedialog:
                    mail_list = select_file_dialog("Select a mailing list file")
            else:
                    mail_list = prompt_for_file("Enter mailing list file") 
    print(f"Mailing list file: {mail_list}\n")
    
    #prompt for template
    if using_GUI_filedialog:
            template = select_file_dialog("Select a template file")
    else:
        template = prompt_for_file("Enter template file")
    print(f"Template file: {template}\n")
    
    #prompt for sig name
    if using_GUI_filedialog:
            sig = select_file_dialog("Select a signature file")
    else: 
        sig = prompt_for_file("Enter signature file")
    print(f"Signature file: {sig}\n")
    
    #Prompt for closing
    print("="*80+"\n")
    closing = input("Enter a closing (Best, Best regards, Signed, etc.): ")
    
    #Prompt for name
    name = input("Enter a name to come after the closing: ")

    
    print("\n\n")
    mail_from_excel(mail_list, template, sig, closing, name)
    
    input("Enter to exit: ")
    

if __name__ == "__main__":

            
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    main()

    
	
