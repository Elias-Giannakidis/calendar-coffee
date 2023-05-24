
from openpyxl import Workbook
from excelManager import genWorkSheet
from pathManager import getNextMonthPath, getPreviousMonthPath, getThisMonthPath

print('\nΛοιπόν γλυκούλι άκου τι θα πατίσεις <3')
print('---------')
print('Για τον προιγούμενο μήνα πατάς: "last"')
print('Για αυτόν τον μήνα πατάς: "this"')
print('Και για τον επόμενο μήνα πατάς: "next"')
print('---------------')
month = input('Άρα γλυκούλι... Για ποιο μήνα θες να σου φτιάξω το excel-ακι? :) \n')

def getPath(month):
    if month == 'last':
        path = getPreviousMonthPath()
        print('Τέλεια σου φτιάχνω excel για τον προηγούμενο μήνα ;)')
    elif month == 'this':
        path = getThisMonthPath()
        print('Τέλεια σου φτιάχνω excel για αυτόν τον μήνα ;)')
    elif month == 'next':
        path = getNextMonthPath()
        print('Τέλεια σου φτιάχνω excel για τον επόμενο μήνα')
    else:
        month = input('Ρε γλυκούλι κάτι δε πάτησες καλά ξανά πάμε \n')
        path = getPath(month)
    return path

path = getPath(month)

# Define the worksheet
wb = Workbook()

wb = genWorkSheet(wb, 'Main', month)
wb.save(path)

ws = wb['Sheet']
wb.remove(ws)

wb.save(path)



print('\n\n\n')
print('  ***   ***')
print(' *    *    *')
print('* Στεφανία! *')
print(' *         *')
print('   *     *')
print('     * *')
print('      | ')
print('     /')
print('     \ ')
print('      \___ Το Excel-άκι σου είναι έτοιμο :)')




