import sys
import os
import ezgmail
import openpyxl
import string
import bs4
import re
import os
from contextlib import suppress

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
            soup = bs4.BeautifulSoup(template, 'html.parser')

    with open(r'Templates\sig.html','r') as infile:
            sig = infile.read()

    soup.body.append(bs4.BeautifulSoup(sig, 'html.parser'))
    with open(r'Templates\combined.html','w') as outfile:
            outfile.write(prettify_except(soup, 'body'))
            
    field_names = [v[1] for v in string.Formatter().parse(template)]

##    for row in sheet.iter_rows(min_row=1, max_col=2, max_row=sheet.max_row):
##        for cell in row:
##                print(cell.value)
    data = tuple(sheet.rows)
    #Start at 1 to skip header, "name" and "email"
    #for i in range (1, 3):
            #ezgmail.draft(data[i][1].value,'for {a}'.format(a=data[i][0].value), template, mimeSubtype='html')

def logout():
    with suppress(OSError):
        os.remove("token.json")

def user_auth():
    while True:
        ezgmail.init()
        print("Python Emailer, you're currently logged in as: {email}".format(email=ezgmail.EMAIL_ADDRESS))
                                    
        #prompt for user change
        while True:
                logout_choice = input("Do you want to logout? ((y)es / (n)o): ").lower()
                if logout_choice in ("yes", "no" , "y", "n"):
                        break
                print("invalid input")
                
        if logout_choice in ("yes", "y"):
                logout()
                continue
        else:
                print("Sending email")
                break
        
def main():
    #patch()

    
    user_auth()
    print("Emailing")

    
    
    #prompt for mail list
    #prompt for template
    #prompt for sig name
    #mailFromExcel()
    
    input("Enter to exit: ")
    

if __name__ == "__main__":

            
    
    main()

    
	
