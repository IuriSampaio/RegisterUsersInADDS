Option Explicit

Dim objRootLDAP, objContainer, objUser, objShell, objExcel
Dim intRow
Dim strOU, strSheet, strCN, strPWD, strtitle, strTitulo, strMensagem, strcompany

strOU = "OU=Meus usuarios ," 
strcompany = "Fatec"
strSheet = "C:\scripts\Users.xlsx"
intRow = 3 

Set objRootLDAP = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://" & strOU & objRootLDAP.Get("defaultNamingContext")) 
Set objExcel = CreateObject("Excel.Application")

objExcel.Workbooks.Open(strSheet)

Do Until objExcel.Cells(intRow,1).Value = ""

  strCN = Trim(objExcel.Cells(intRow, 2).Value)
  strPWD = Trim(objExcel.Cells(intRow, 6).Value)
  
  Set objUser = objContainer.Create("User", "cn=" & strCN)

  objUser.sAMAccountName = Trim(objExcel.Cells(intRow, 1).Value)
  objUser.givenName = Trim(objExcel.Cells(intRow, 3).Value)
  objUser.initials = Trim(objExcel.Cells(intRow, 4).Value)
  objUser.sn = Trim(objExcel.Cells(intRow, 5).Value)
  objUser.SetInfo

  objUser.physicalDeliveryOfficeName = Trim(objExcel.Cells(intRow, 7).Value)
  objUser.mail = Trim(objExcel.Cells(intRow, 8).Value)
  objUser.userPrincipalName= objUser.sAMAccountName & "@fatec.lab"
  objUser.displayName = strCN
  objUser.title = Trim(objExcel.Cells(intRow, 9).Value)
  objUser.department = Trim(objExcel.Cells(intRow, 10).Value)
  objUser.company = strcompany
  objUser.description = Trim(objExcel.Cells(intRow, 11).Value)
  objUser.telephoneNumber = Trim(objExcel.Cells(intRow, 12).Value)
  objUser.userAccountControl = 512
  objUser.pwdLastSet = 0
  objUser.SetPassword strPWD
  objUser.SetInfo

  intRow = intRow + 1

Loop

objExcel.Quit 


strTitulo = "COMANDO CONCLUIDO!!"
strMensagem = "USUARIO(S) CRIADO(S) COM SUCESSO!"
msgbox strMensagem, 0 + 64, strTitulo


WScript.Quit