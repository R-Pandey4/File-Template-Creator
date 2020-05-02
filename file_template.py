from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

print("FILE TEMPLATE CREATOR")
num_prac= int(input("\n\nEnter the number of experiments: "))

document = Document() 

#Index table
table = document.add_table(rows=num_prac+1, cols=4)
table.cell(0,0).text="S.No"
table.cell(0,1).text="Program"
table.cell(0,2).text="Date"
table.cell(0,3).text="Sign"

for i in range(1,num_prac+1):
    table.cell(i,0).text=f"{i}"

document.add_page_break()

#Now moving to the experiments
for i in range(1,num_prac+1):

    para1 = document.add_paragraph() #Aim of the experiment

    para1.alignment= WD_ALIGN_PARAGRAPH.LEFT
    run1 = para1.add_run(f"{i}. Enter the aim of the experiments\n")
    font1 = run1.font
    font1.name = "Arial"
    font1.bold = True
    font1.size = Pt(20)

    para2 = document.add_paragraph()

    para2.alignment= WD_ALIGN_PARAGRAPH.CENTER #The Word "CODE"
    run2 = para2.add_run("CODE\n\n")
    font2 = run2.font
    font2.name = "TimesNewRoman"
    font2.bold = True
    font2.size = Pt(15)

    para3 = document.add_paragraph()

    para3.alignment= WD_ALIGN_PARAGRAPH.CENTER #The Word "OUTPUT"
    run3 = para3.add_run("OUTPUT\n")
    font3 = run3.font
    font3.name = "TimesNewRoman"
    font3.bold = True
    font3.size = Pt(15)

    document.add_page_break()

document.save("practical.docx")
