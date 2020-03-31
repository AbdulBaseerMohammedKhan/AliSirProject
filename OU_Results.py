URL = input('enter the URL of the OU results page')
collegeCode=int(input("enter college code(1604)..."))
yearCode=int(input("enter year code(18)..."))
StreamCode=int(input("enter stream code(733)..."))
LE = int(input("are there any LE students(1 for yes  /  0 for no)"))
extra = int(input("are there any students with roll numbers 121-180(1 for yes  /  0 for no)"))

import bs4, mechanize, urllib
from xlwt import Workbook

wb = Workbook()

sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, 'Roll Number')
sheet1.write(0, 1, "Result(CGPA)")
sheet1.write(0, 2, "Name")


br = mechanize.Browser()
br.set_handle_robots(False)
br.open(URL)

def marksGenerator(i):



    
    br.select_form('FrontPage_Form1')
    br.form['htno'] = str(collegeCode) +str(yearCode) +str(StreamCode) + str(i)

    response1 = br.submit()
    raw = response1.read()
    soup = bs4.BeautifulSoup(raw, "html5lib")
    containers = soup.findAll("td", {"width": "50%"})
    containers2 = soup.findAll("font", {"face": "Arial"})

    marks = containers[-1].text[10:14]
    name1 = str(containers2[8]).replace("</font>", "")
    name = str(name1).replace('<font face="Arial" size="2">', '')


    sheet1.write(int(i), 0, int(i))
    sheet1.write(int(i), 1, marks)
    sheet1.write(int(i), 2, name)

    print(int(i))

for i in range(1,10):
    marksGenerator("00" + str(i))
for i in range(10, 100):
    marksGenerator("0" + str(i))
for i in range(100, 121):
    marksGenerator( str(i))

if extra == 1:
    for i in range(122, 181):
        marksGenerator( str(i))

if LE ==1:
    for i in range(301, 313):
        marksGenerator( str(i))

wb.save('results revised12.xls')



