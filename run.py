import  openpyxl
import os

PATH = "./table.xlsx"
ROW=2

wb_obj = openpyxl.load_workbook(PATH, read_only=True, data_only=True, keep_vba=False, keep_links=False)
 
sheet_obj = wb_obj.active



## Получаем строку из ячейки ексель
def GetString(row, column)->str:
    text = sheet_obj.cell(row = row, column = column).value
    if text is None or text == '-' or text == '' or text == ' ':
        text = "NULL"
    return text

def SetNULL(s):
    if s is None:
        s="NULL"
    return s

def Append (s, mass):
    if s is not None and s != '-':
        if s not in mass:
            mass.append(s)

def GetID(s, mass):
    if s in mass:
        return mass.index(s) + 1
    else:
        return "NULL"
    
def PrintINSERT(mass, table, f):
    for i, s in enumerate(mass):
        f.write(f"INSERT INTO {table} VALUES({i+1}, '{s}');\n\n")

def CheckMass(s, m)->int:
     if s is not None and s != '-' and s != '' and s != ' ':
        if s not in (m):
            m.append(s)
            for i, x in enumerate(m):
                if x==s:
                     return int(i)
        else:
             for i, x in enumerate(m):
                if x==s:
                     return int(i)
                
def MakeShort(s):
    if s is None:
        return "NULL"
    p = s.split(".")
    r = ".".join(p[:3])
    if len(r) > 255:
        r = str(r)[:255]
    return r

## При слишком большой длмне строки Firebird не засчитывает строку, и делает все остальное в CON
## Эта функция переносит строку на следующую
def WrapString(input_string, max_length):
    if len(input_string) > max_length:
        new_string = ""
        for i in range(0, len(input_string), max_length):
            new_string += input_string[i:i+max_length] + "\n"
        return new_string
    else:
        return input_string

jobs = []
degrees = []
napravs = []

def main():

    os.remove('d.sql')

    with open('d.sql', 'w') as file:

        file.write("create database 'D:\test.fdb' user 'user' password '1234';\n\n")

        file.write("CREATE TABLE job (id INT NOT NULL PRIMARY KEY, name varchar(255));\n\n")
        file.write("CREATE TABLE cathedra (id INT NOT NULL PRIMARY KEY, name varchar(255));\n\n")
        file.write("CREATE TABLE cval (id INT NOT NULL PRIMARY KEY, name varchar(255));\n\n")

        file.write("CREATE TABLE teacher(\n")
        file.write("id INT NOT NULL PRIMARY KEY,\n")

        file.write("last_name varchar(255),\n")
        file.write("name varchar(255),\n")
        file.write("middle_name varchar(255),\n")

        file.write("job_id INT,\n")
        file.write("cathedra_id INT,\n")
        file.write("cval_id INT,\n")

        file.write("uch_stepen varchar(255),\n")
        file.write("uch_zvanie varchar(255),\n")

        file.write("all_stash INT,\n")
        file.write("spec_stash INT,\n")

        file.write("code varchar(255),\n")
        file.write("disciplines varchar(255),\n")

        file.write("FOREIGN KEY (job_id) REFERENCES job(id),\n")
        file.write("FOREIGN KEY (cathedra_id) REFERENCES cathedra(id),\n")
        file.write("FOREIGN KEY (cval_id) REFERENCES cval(id)\n")
        file.write(");\n\n")

        file.write("show table teacher;\n")
        file.write("show table job;\n")
        file.write("show table cathedra;\n")
        file.write("show table cval;\n")

        x=1

        for row in range(2,32):

            id = x
            
            lastName = GetString(row, 1)
            name = GetString(row, 2)
            middleName = GetString(row, 3)

            job = GetString(row, 4)
            degree = GetString(row, 5)
            naprav = GetString(row, 6)

            uch_stepen = GetString(row, 7)
            uch_zvanie = GetString(row, 8)

            # cvalificaition = GetString(row, 8)

            allStahs = GetString(row, 9)
            specStahs = GetString(row, 10)

            code = GetString(row, 11)
            discipline = MakeShort(GetString(row, 12))

            year = "г."
            # if cvalificaition is not None:
            #         index = cvalificaition.find(year)
            #         if index != -1:
            #                 new_cvalification = cvalificaition[:index+len(year)]
            #                 if len(new_cvalification) > 255:
            #                         cvalificaition = str(new_cvalification)[:255]


            jobID = CheckMass(job, jobs)

            dgID = CheckMass(degree, degrees)
            npID = CheckMass(naprav, napravs)

            if jobID is None:
                jobID = "NULL"
            else:
                 jobID +=1

            if dgID is None:
                dgID = "NULL"
            else:
                 dgID +=1

            if npID is None:
                npID = "NULL"
            else:
                 npID +=1


            text = f"INSERT INTO teacher VALUES({id}, '{lastName}', '{name}', '{middleName}', {jobID}, {dgID}, {npID}, '{uch_stepen}', '{uch_zvanie}', {allStahs}, {specStahs}, '{code}', '{discipline}');\n\n"

            fText = WrapString(text, 150)

            file.write(fText)

            x+=1


        PrintINSERT(jobs, "job", file)
        PrintINSERT(degrees, "cathedra", file)
        PrintINSERT(napravs, "cval", file)

        file.write("INSERT INTO teacher VALUES(110, 'X', 'X', 'X', 10, 'X', 'X', 3, 3, 2, 30, 8);  \n\n select * from teacher; \n\nselect * from job; \n\nselect * from degree; \n\nselect * from naprav;")





main()
    