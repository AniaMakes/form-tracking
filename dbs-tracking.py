__author__ = 'Admin'

from functools import wraps
import re
import ssl
import mechanize
from bs4 import BeautifulSoup
from xlutils.copy import copy
from xlrd import open_workbook # import open workbook ability from read excel module

br = mechanize.Browser()
br.set_handle_robots(False)
br.set_handle_equiv(False)
br.set_handle_refresh(False)
br.addheaders = [('User-agent', 'Firefox')]

def sslwrap(func):
    @wraps(func)
    def bar(*args, **kw):
        kw['ssl_version'] = ssl.PROTOCOL_TLSv1
        return func(*args, **kw)
    return bar

ssl.wrap_socket = sslwrap(ssl.wrap_socket)


book = open_workbook("C:/sandbox/AllSent.xlsx") # define our spreadsheet
sheet = book.sheet_by_index(0) # define the sheet we're working with

# if cell is empty, print end
# if else, run the process

print sheet.nrows

wb = copy(book) # a writable copy (I can't read values out of this, only write to it)


def spr_input(par):
    if par.text.find("cannot remember") != -1:
        return "Error: form not received or incorrect data was input"
    else:
        return par.text.strip()



for x in range(1, sheet.nrows):

    ref = sheet.cell(x,0).value
    # print ref

    dd = sheet.cell(x,1).value
    # print dd

    mm = sheet.cell(x,2).value
    # print mm

    yyyy = sheet.cell(x,3).value
    # print yyyy

    response = br.open("https://secure.crbonline.gov.uk/enquiry/enquirySearch.do")
    # for f in br.forms():
    #     print f

    # for form in br.forms():
      #  print "Form name", form.name
       # print form

    br.select_form(nr=0)
    br.form['applicationNo'] = ref[3:14]
    br.form['dateOfBirthDay'] = [dd,]
    br.form['dateOfBirthMonth'] = [mm,]
    br.form['dateOfBirthYear'] = [yyyy,]

    response = br.submit()
    # print response.read()

    soup = BeautifulSoup(response.read(), "html.parser")
    search_result = (soup.get_text("|", strip=True))
    #print search_result

    paragraphs = soup.find_all('p')   # re.compile("application", re.MULTILINE))
    print paragraphs

    wb.get_sheet(0).write(x,8, spr_input(paragraphs[0]))

#    wb.get_sheet(0).write(x,4, paragraphs)



wb.save("C:/sandbox/ExternalSent2.xls")

#print search_result[600:750]
print

