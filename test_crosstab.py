"""" Flexible features: Can have as many number of columns in the input sheet.
                    Can have CHANGE PERCENT column anywhere in the input sheet. It will directly omit this column
                    Columns labels can be different. eg Instead of REFERENCE VALUE column, it can be AVERAGE VALUE column.
                    The PERCENT, REFERENCE VALUE and VALUE can be in any order.
                    Can have as many scenarios in the SCENARIOS column
     Limitations:       The last 3 columns of the input sheet should be PERCENT, REFERENCE VALUE and VALUE in any order (CHANGE PERCENT can be the last column of input sheet)
     Specifications:    Install xlrd library from the Python documentation
     Author: Kushagra Mall """"


import xlrd
import xlsxwriter
import itertools
from tempfile import TemporaryFile

workbook1=xlrd.open_workbook("C:/MultCTG/SampleOutput3.xlsx")  # Defines the complete path of the input excel
sh = workbook1.sheet_by_name("SampleOutput")  # Defines the worksheet in the input excel

#==============================FIRST ROW=====================================

r1 = []   # r1 = First row of output excel
posChangePer = 1000
for j in range (0, sh.ncols):
    if(sh.cell_value(0, j) != 'Change Percent'):
        r1.append(sh.cell_value(0, j))
    else:
        posChangePer = j

posScenario = 0
for j in range (0, sh.ncols):
    if(sh.cell_value(0, j) == 'Scenario'):
        posScenario = j

c1 = []
for i in range(1, sh.nrows):
    c1.append(sh.cell_value(i, posScenario))

output = []  # Non Repeated array of Scenarios
for x in c1:
    if x not in output:
        output.append(x)
        
for i in range(0, (3*len(output))-3):    
    r1.append(0)

val = 0
ref = 0
per = 0
if(sh.cell_value(0, sh.ncols-1) == 'Change Percent'):
    val = sh.cell_value(0, sh.ncols-2)
    ref = sh.cell_value(0, sh.ncols-3)
    per = sh.cell_value(0, sh.ncols-4)
else:
    val = sh.cell_value(0, sh.ncols-1)
    ref = sh.cell_value(0, sh.ncols-2)
    per = sh.cell_value(0, sh.ncols-3)

# Last 3 values of first row     
for i in range(0, len(output)):
    r1[len(r1)-((i*3)+1)] = output[len(output)-(i+1)] + '/' + val
    r1[len(r1)-((i*3)+2)] = output[len(output)-(i+1)] + '/' + ref
    r1[len(r1)-((i*3)+3)] = output[len(output)-(i+1)] + '/' + per   

posLabel = 0  # Column Position of Label
posNumFrom = 0  # Column Position of Number From
posNumTo = 0  # Column Position of Number To
posCircuit = 0  # Column Position of Cicuit Number

for i in range(0, len(r1)):
    if(r1[i] == 'Label'):
        posLabel = i
    if(r1[i] == 'Number From'):
        posNumFrom = i
    if(r1[i] == 'Number To'):
        posNumTo = i
    if(r1[i] == 'Circuit'):
        posCircuit = i

#=====================================================================================        
#===================================OUTPUT EXCEL======================================
             
a = []  # Data Array containing values from input excel

for i in range (1, sh.nrows):
    a.append([])
    for j in range (0, sh.ncols):
        if(j != posChangePer):
            a[i-1].append(sh.cell_value(i,j))

k1 = []  # Data array containing values after applying crosstab query

for i in range(0, len(a)):
    k = []
    for f in range(0, len(r1)-1):
        k.append('')
    for z in range(1, len(a[i])-3):
        k[z-1] = a[i][z]
    for j in range(0, len(a)):
        if(a[i][posLabel] == a[j][posLabel] and a[i][posNumFrom] == a[j][posNumFrom] and a[i][posNumTo] == a[j][posNumTo] and a[i][posCircuit] == a[j][posCircuit]):
            for g in range(0, len(output)):
                if(a[j][0] == output[g]):
                    k[(len(k)-(3*(len(output)-g)))] = a[j][len(a[j])-3]
                    k[(len(k)-(3*(len(output)-g)))+1] = a[j][len(a[j])-2]
                    k[(len(k)-(3*(len(output)-g)))+2] = a[j][len(a[j])-1]
                 
    
    k1.append(k)

k1.sort()
k2 = list(k1 for k1,_ in itertools.groupby(k1))

r2 = r1[1:]  # New updated first row

#==============================================================================

workbook = xlsxwriter.Workbook("C:\Users\Intern\Desktop"+'/Cross_Tab_for_branches_buses_angles.xlsx')  # Complete path of output excel + Name of the output excel
sheet1 = workbook.add_worksheet('Branches')
sheet2 = workbook.add_worksheet('Buses')
sheet3 = workbook.add_worksheet('Angles')


# Exporting first row in the output excel
cat = 0
for i in range(0, len(r2)):
    sheet1.write(0, i, r2[i])
    sheet2.write(0, i, r2[i])
    sheet3.write(0, i, r2[i])
    if(r2[i] == 'Category'):
        cat = i
        

# Exporting values in the output excel
br = 1
bu = 1
an = 1
for i in range(0, len(k2)):
    for j in range(0, len(k2[i])):
        if(k2[i][cat] == 'Branch MVA' or k2[i][cat] == 'Branch Amp'):
            sheet1.write(br, j, k2[i][j])
        elif(k2[i][cat] == 'Change Bus Low Volts' or k2[i][cat] == 'Bus Low Volts' or k2[i][cat] == 'Bus High Volts'):
            sheet2.write(bu, j, k2[i][j])
        elif(k2[i][cat] == 'AngleChangeMonitor'):
            sheet3.write(an, j, k2[i][j])
    if(k2[i][cat] == 'Branch MVA' or k2[i][cat] == 'Branch Amp'):
        br = br + 1
    elif(k2[i][cat] == 'Change Bus Low Volts' or k2[i][cat] == 'Bus Low Volts' or k2[i][cat] == 'Bus High Volts'):
        bu = bu + 1
    elif(k2[i][cat] == 'AngleChangeMonitor'):
        an = an + 1

    
workbook.close()