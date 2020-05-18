from datetime import datetime
import moodle_api
import xlsxwriter
from time import sleep

from tinydb import TinyDB, Query
db = TinyDB('db.json')


moodle_api.URL = "your_moodle_url"
moodle_api.KEY = "your_moodle_token"
moodle_api.MAX_RETRIES=0
mail_skip="mail_to_skip"

totali=[]
spec_corsi=[]
mai_acceduti=0
utenti_totali=0
name=datetime.now().strftime("%d-%m-%Y %H.%M.%S")
workbook = xlsxwriter.Workbook('Report-'+name+".xlsx")
worksheet = workbook.add_worksheet()
row=2
col=0
status=0
print("Recupero elenco corsi ...")
try:
    #ottengo elenco di tutti i corsi
    corsi=moodle_api.call('core_course_get_courses')
    #rimuovo il corso con id=1 che corrisponde a "tutti gli iscritti in piattaforma"
    corsi.pop(0)
    print("Elenco corsi processato. "+str(len(corsi))+" corsi totali")
    status=status+1
    for k in corsi:
        #sleep(5)
        print("Recupero iscritti corso "+str(k['fullname'])+" ...")
        iscritti=moodle_api.call('core_enrol_get_enrolled_users',courseid=k['id'])
        print("Iscritti corso "+str(k['fullname'])+" recuperati")
        status=status+1
        print("...."+str(round((status*100/len(corsi)),2))+"%....")
        #print(iscritti[0])
        unit={}
        unit['nomecorso']=k['fullname']
        for i in iscritti:
            if(mail_skip not in i['email']):
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
    worksheet.write(0, 0,'Report dei singoli corsi al '+name)
    worksheet.write(1, 0,"Questo report gia ignora tutte le utenze con mail contenenti "+mail_skip+" e tutti gli utenti non iscritti ad un corso")
    worksheet.write(row, 0,"nome corso")
    worksheet.write(row, 1,"Studenti")
    worksheet.write(row, 2,"Docenti")
    worksheet.write(row, 3,"Manager")
    worksheet.write(row, 4,"Docenti non editor")
    worksheet.write(row, 5,"Totale")
    col=6
    for k in spec_corsi:
        chiavi=k.keys()
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
    row=row+1
    worksheet.write(row,0,'Report complessivo dei singoli utenti iscritti a corsi in data '+name)
    row=row+1
    worksheet.write(row,0,"Questo report gia ignora tutte le utenze con mail contenenti "+mail_skip+" e tutti gli utenti non iscritti ad un corso")
    row=row+1
    worksheet.write(row, 0,"Nome corso")
    worksheet.write(row, 1,"Nominativo")
    worksheet.write(row, 2,"Email")
    worksheet.write(row, 3,"Ruolo")
    worksheet.write(row, 4,"Primo accesso")
    for k in totali:
        row=row+1
        riga=k.split(",")
        col=0
        for i in riga:
            worksheet.write(row,col,i)
            col=col+1
        #print(k)

    print("---------------------------------------------------------")
    print("Totale attivi:"+str(utenti_totali-mai_acceduti))
    worksheet.write("K2","=SUM(F4:F"+str(fine_somma)+")")
    workbook.close()
    print("Procedura terminata Premere un tasto per uscire")
    input()
except Exception as e: 
    print("!!!!")
    print("Si è verificato un errore. Impossibile procedere. Premere un tasto per uscire")
    print("!!!!")
    print(str(e))
    input()