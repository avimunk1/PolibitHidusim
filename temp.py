#this utility gets a csv export from Bafi and create an xls
# it handles the phones that are in one column in the original file

status=''
customerName=''
customerID=''
phones=''
Address=''
email=''
policyNumber=''
insuranceType=''
insuranceCompny=''
anafID=''
anafName=''
policyStartDate=''
policyEndDate=''
prmia=''
veicleType=''
SubAnaf=''
carLicenessPlate=''
manifctureYear=''
vehicleModelId=''
vehicleModelDescription=''
allowedToDrive=''
protuction=''
propertyAddress=''
paymentType=''
numofPayments=''
egentName=''
egentId=''
ipdatedBy=''
crmRepName=''
salesRepName=''
comments=''
policyhandkerName=''
count=0

policy=list()
fileH=open("b.csv",encoding="ISO-8859-8")


for line in fileH:
    line=line.strip()
    policy=line.split(',')
    status=policy[1]
    customerName=policy[2]
    customerID=policy[3]
    phones=policy[4]
    Address=policy[5]
    email=policy[6]
    policyNumber=policy[7]
    insuranceType=policy[8]
    insuranceCompny=policy[9]
    anafID=policy[10]
    anafName=policy[11]
    policyStartDate=policy[12]
    policyEndDate=policy[13]
    prmia=policy[14]
    veicleType=policy[15]
    SubAnaf=policy[16]
    carLicenessPlate=policy[17]
    manifctureYear=policy[18]
    vehicleModelId=policy[19]
    vehicleModelDescription=policy[20]
    allowedToDrive=policy[21]
    protuction=policy[22]
    propertyAddress=policy[23]
    paymentType=policy[24]
    numofPayments=policy[25]
    egentName=policy[26]
    egentId=policy[27]
    ipdatedBy=policy[28]
    crmRepName=policy[29]
    salesRepName=policy[30]
    comments=policy[31]
    policyhandkerName=policy[32]
    print(count)
    print(policy[count])
    count=count+1

#    excel style example
    from openpyxl.styles import colors
    from openpyxl.styles import Font, Color
    from openpyxl import Workbook
#    a1 = rws['A1']
#    d4 = rws['D4']
#    ft = Font(color=colors.RED)
#    a1.font = ft
#    d4.font = ft
