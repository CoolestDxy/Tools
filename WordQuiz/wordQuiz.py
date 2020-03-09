import tkinter as tk  # 使用Tkinter前需要先导入
import xlrd
import xlwt
chList=[]
enList=[]
wrongList=[]
fileName='./Word.xls'
outputFile='./Word.xls'
style0 = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue')
style1=xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow')
style2=xlwt.easyxf('pattern: pattern solid, fore_colour light_orange')
f1 = xlrd.open_workbook(fileName)
name = f1.sheet_names()[0]
sheet = f1.sheet_by_name(name)
rowl=sheet.nrows
for i in range(0,rowl):
    row=sheet.row_values(i)
    if row[0]=='':
        continue
    chList.append(row[0])
    enList.append(row[1])
    if len(row)==3 and row[2] in ['0','1','2','3']:
        wrongList.append(int(row[2]))
    else:
        wrongList.append(0)
length=len(chList)
n=0
def check(ev=None):
    global n
    if n>length-1:
        var1.set('Quiz Completed!')
        return
    myAnswer=var4.get()
    var4.set('')
    answer=chList[n]+'\nYour answer: '+myAnswer+'\nCorrect Answer: '+enList[n]
    var2.set(answer)
    if myAnswer!=enList[n]:
        result['fg']='red'
        wrongList[n]+=1
        if wrongList[n]>3:
            wrongList[n]=3
    else:
        result['fg'] = 'black'
        wrongList[n]=0
    n += 1
    if var0.get() == 1:
        while n<length:
            if wrongList[n] == 0:
                n += 1
            else:
                break
    if n>length-1:
        var1.set('Quiz Completed!')
        return
    var1.set(chList[n])

def hop():
    if n>length-1:
        var1.set('Quiz Completed!')
        return
    if wrongList[n]>0 and result['fg']=='red':
        wrongList[n]-=1
    result['fg'] = 'black'
def closeWindow():
    workbook  = xlwt.Workbook(encoding = 'utf-8')
    sheet=workbook .add_sheet('WordQuiz')
    for nw in range(0,len(chList)):
        if wrongList[nw]==1:
            sheet.write(nw,0,chList[nw],style0)
            sheet.write(nw,1,enList[nw],style0)
            sheet.write(nw,2,'1',style0)
        if wrongList[nw] == 0:
            sheet.write(nw, 0, chList[nw])
            sheet.write(nw, 1, enList[nw])
            sheet.write(nw, 2, '0')
        if wrongList[nw] == 2:
            sheet.write(nw, 0, chList[nw], style1)
            sheet.write(nw, 1, enList[nw], style1)
            sheet.write(nw, 2, '2', style1)
        if wrongList[nw] == 3:
            sheet.write(nw, 0, chList[nw], style2)
            sheet.write(nw, 1, enList[nw], style2)
            sheet.write(nw, 2, '3', style2)
    workbook.save(outputFile)
    window.destroy()

window = tk.Tk()
window.title('Word Quiz')
window.geometry('400x400')
var0=tk.IntVar()
var1 = tk.StringVar()
var2 = tk.StringVar()
var3=tk.StringVar()
var4=tk.StringVar()
var0.set(0)
var1.set(chList[0])
var2.set('correct answer')
var4.set('')
button=tk.Checkbutton(window, text='only mistakes',variable=var0, onvalue=1, offvalue=0)
wordText = tk.Label(window, textvariable=var1, font=('Arial', 15), width=30, height=5)
input = tk.Entry(window,textvariable=var4,show=None, font=('Arial', 14))
input.bind("<Return>", check)
none=tk.Label(window, textvariable=var3, text='', font=('Arial', 12), width=30, height=1)
result = tk.Label(window, textvariable=var2,font=('Arial', 12), width=30, height=10,wraplength = 200)
button.pack()
wordText.pack()
input.pack()
none.pack()
b = tk.Button(window, text="I'm right!",bd=3, font=('Arial', 12), width=10, height=1, command=hop)
b.pack()
result.pack()
window.protocol('WM_DELETE_WINDOW', closeWindow)
window.mainloop()
