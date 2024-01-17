import sys
import os
import ezgmail
import openpyxl
import string
import bs4

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

#batch of 20, complete list, range of rows, start row then quant,
#add sent email date & time
#use sent date to confirm if email needs to be sent again

def mailFromExcel():
    #Assuming that we pass in args for the file
    email_list = sys.argv[1]
    print(email_list)
    
    wb = openpyxl.load_workbook(email_list) #Todo do I need to close?
    sheet = wb['Sheet1']

    #https://stackoverflow.com/questions/23332259/copy-cell-style-openpyxl
    
    org_name_col = "A"
    org_mail_col = "B"
    
    with open(r'Templates\promo.html','r') as infile:
            template = infile.read()
            
    field_names = [v[1] for v in string.Formatter().parse(template)]

##    for row in sheet.iter_rows(min_row=1, max_col=2, max_row=sheet.max_row):
##        for cell in row:
##                print(cell.value)
    data = tuple(sheet.rows)
    #Start at 1 to skip header, "name" and "email"
    for i in range (1, 3):
            ezgmail.draft(data[i][1].value,'for {a}'.format(a=data[i][0].value), template, mimeSubtype='html')
    
def main():
    #patch()
    ezgmail.init()
    
    
    print(sys.argv[1])
    print("Python Emailer, you're currently logged in as: {email}".format(email=ezgmail.EMAIL_ADDRESS))
    mailFromExcel()
    
    input("Enter to exit: ")
    

if __name__ == "__main__":
    
    main()

    
	
