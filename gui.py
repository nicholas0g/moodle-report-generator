from tkinter import *
from tkinter.ttk import *
from datetime import datetime
import moodle_api
import xlsxwriter
from threading import *

from tinydb import TinyDB, Query
db = TinyDB('db.json')

#thread for asinc op
class ops(Thread):
    def run(self):
        moodle_api.URL=url.get()
        moodle_api.KEY=token.get()
        totali=[]
        spec_corsi=[]
        mai_acceduti=0
        utenti_totali=0
        name=datetime.now().strftime("%d-%m-%Y %H.%M.%S")
        workbook = xlsxwriter.Workbook('Report-'+name+".xlsx")
        bold = workbook.add_format({'bold': True})
        red_cell = workbook.add_format()
        red_cell.set_fg_color('red')
        orange_cell = workbook.add_format({'bold': True})
        orange_cell.set_fg_color('orange')
        yellow_cell = workbook.add_format()
        yellow_cell.set_fg_color('yellow')
        worksheet = workbook.add_worksheet()
        worksheet.set_column('E:E', 18)
        worksheet.set_column('G:G', 18)
        worksheet.set_column('H:H', 18)
        worksheet.set_column('A:A', 40)
        row=3
        col=0
        status=0
        logs.insert(0,"Recupero elenco corsi ...")
        try:
            #ottengo elenco di tutti i corsi
            corsi=moodle_api.call('core_course_get_courses')
            #rimuovo il corso con id=1 che corrisponde a "tutti gli iscritti in piattaforma"
            corsi.pop(0)
            logs.insert(0,"Elenco corsi processato. "+str(len(corsi))+" corsi totali")
            status=0
            for k in corsi:
                #sleep(5)
                logs.insert(0,"Recupero iscritti corso "+str(k['fullname'])+" ...")
                iscritti=moodle_api.call('core_enrol_get_enrolled_users',courseid=k['id'])
                logs.insert(0,"Iscritti corso "+str(k['fullname'])+" recuperati")
                status=status+1
                logs.insert(0,"...."+str(round((status*100/len(corsi)),2))+"%....")
                progress['value'] = status*100/len(corsi)
                window.update_idletasks() 
                #logs.insert(0,iscritti[0])
                unit={}
                unit['nomecorso']=k['fullname']
                for i in iscritti:
                    if(mail.get() not in i['email']):
                        if(i['firstaccess']==0):
                            totali.append(k['fullname']+","+i['fullname']+","+i['email']+","+i['roles'][0]['shortname']+",mai")
                            mai_acceduti=mai_acceduti+1
                        else:
                            totali.append(k['fullname']+","+i['fullname']+","+i['email']+","+i['roles'][0]['shortname']+","+str(datetime.fromtimestamp(i['firstaccess'])))
                        if(i['roles'][0]['shortname'] in unit):
                            unit[i['roles'][0]['shortname']]=unit[i['roles'][0]['shortname']]+1
                        else:
                            unit[i['roles'][0]['shortname']]=1
                spec_corsi.append(unit)
            worksheet.merge_range('A1:F1', 'Piattaforma '+url.get(), yellow_cell)
            worksheet.merge_range('A2:F2', 'Report dei singoli corsi al '+name, red_cell)
            worksheet.merge_range('A3:F3',"Questo report gia ignora tutte le utenze con mail contenenti "+mail.get()+" e tutti gli utenti non iscritti ad un corso",red_cell)
            worksheet.write(row, 0,"nome corso",bold)
            worksheet.write(row, 1,"Studenti",bold)
            worksheet.write(row, 2,"Docenti",bold)
            worksheet.write(row, 3,"Manager",bold)
            worksheet.write(row, 4,"Docenti non editor",bold)
            worksheet.write(row, 5,"Totale",bold)
            col=6
            for k in spec_corsi:
                if('student' not in k):
                    k['student']=0
                if('editingteacher' not in k):
                    k['editingteacher']=0
                if('manager' not in k):
                    k['manager']=0
                if('noneditingteacher' not in k):
                    k['noneditingteacher']=0
                tot=k['student']+k['editingteacher']+k['manager']+k['noneditingteacher']
                row=row+1
                worksheet.write(row, 0,k['nomecorso'])
                worksheet.write(row, 1,k['student'])
                worksheet.write(row, 2,k['editingteacher'])
                worksheet.write(row, 3,k['manager'])
                worksheet.write(row, 4,k['noneditingteacher'])
                worksheet.write(row, 5,tot)
                utenti_totali=utenti_totali+tot
            fine_somma=row+1
            row=row+2
            worksheet.merge_range("A"+str(row)+":F"+str(row),'Report complessivo dei singoli utenti iscritti a corsi in data '+name,red_cell)
            row=row+1
            worksheet.merge_range("A"+str(row)+":F"+str(row),"Questo report gia ignora tutte le utenze con mail contenenti "+mail.get()+" e tutti gli utenti non iscritti ad un corso",red_cell)
            worksheet.write(row, 0,"Nome corso",bold)
            worksheet.write(row, 1,"Nominativo",bold)
            worksheet.write(row, 2,"Email",bold)
            worksheet.write(row, 3,"Ruolo",bold)
            worksheet.write(row, 4,"Primo accesso",bold)
            for k in totali:
                row=row+1
                riga=k.split(",")
                col=0
                for i in riga:
                    worksheet.write(row,col,i)
                    col=col+1
                #logs.insert(0,k)

            logs.insert(0,"---------------------------------------------------------")
            logs.insert(0,"Totale attivi:"+str(utenti_totali-mai_acceduti))
            worksheet.write("G1","Utenti totali",orange_cell)
            worksheet.write("G2","=SUM(F4:F"+str(fine_somma)+")")
            worksheet.write("H1","Utenti mai attivi",orange_cell)
            worksheet.write("H2","=COUNTIF(E"+str(fine_somma+3)+":E"+str(row)+',"mai")')
            workbook.close()
            logs.insert(0,"Procedura terminata.")
        except Exception as e: 
            logs.insert(0,"!!!!")
            logs.insert(0,"Si è verificato un errore. Impossibile procedere.")
            logs.insert(0,"!!!!")
            logs.insert(0,str(e))

#global variable
moodle_api.MAX_RETRIES=0

##thread start function
def start():
    db.truncate()
    db.insert({'url':url.get(),'token':token.get(),'mail':mail.get()})
    progress['value'] = 0
    window.update_idletasks()
    logs.delete(0,END)
    p1=ops()
    p1.start()

def end_t():
    progress['value'] = 0
    window.update_idletasks()
    logs.delete(0,END)
    logs.insert(0,"Waiting to start")

window=Tk()

#label variables    
#logs=StringVar()

#color definition
background="orange"
style = Style()
style.configure("BW.TLabel", foreground="black", background=background)
style.configure("BC.TLabel", foreground="white", background="black")

#main
window.title('Moodle global report generator')
window.geometry("400x300+10+20")
window.configure(background=background)
window.resizable(0,0)

#label
Label (window,text="Inserisci url moodle:", style="BW.TLabel", font="none 12").grid(row=1,column=1,padx=5,pady=5,sticky=W)
Label (window,text="Inserisci token:", style="BW.TLabel", font="none 12").grid(row=3,column=1,padx=5,pady=5,sticky=W)
Label (window,text="Email da ignorare:", style="BW.TLabel", font="none 12").grid(row=4,column=1,padx=5,pady=5,sticky=W)

#textbox
tk=StringVar()
ur=StringVar()
ml=StringVar()
url=Entry(window,width=38,textvariable=ur)
url.grid(row=1,column=2,sticky=W,padx=5,pady=5)
token=Entry(window,width=38,textvariable=tk)
token.grid(row=3,column=2,sticky=W,padx=5,pady=5)
mail=Entry(window,width=38,textvariable=ml)
mail.grid(row=4,column=2,sticky=W,padx=5,pady=5)

#button
Button(window,text="Generate xlsx",width=15,command=start).grid(row=6,column=2)
Button(window,text="Reset",width=15,command=end_t).grid(row=6,column=1)

#status
progress=Progressbar (window, orient = HORIZONTAL, length=390, mode = 'determinate') 
progress.grid(row=7,column=1,columnspan=2,padx=5,pady=5,sticky=W)

#listbox
logs =Listbox(window,width=65)
logs.grid(row=8,column=1,columnspan=2,padx=2,pady=2,sticky=W)
logs.insert(0,"Waiting to srart")

#auto-put data into textbox
if len(db.all())!=0:
    ur.set(db.all()[0]['url'])
    tk.set(db.all()[0]['token'])
    ml.set(db.all()[0]['mail'])


window.mainloop()

