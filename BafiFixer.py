#this utility gets a csv export from Bafi and create an xls
# it handles the phones that are in one column in the original file


beffiExport="BafiExportCheckifHebisRedabule.csv"
try:
    fileH=open('inputs/'+beffiExport)
except:
    print('no such file'+beffiExport)
    exit()
filh=open('inputs/'+beffiExport,encoding="ISO-8859-8")

#fileH=open('inputs/'+beffiExport)
for line in fileH:
    line=line.strip()
    line=line.split(';')
    print(word)

fileH=open("b.csv",encoding="ISO-8859-8")
#for line in fileH:
#    line=line.strip()
#    print(line)
