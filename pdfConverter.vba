Private Sub Workbook_Open()
 
    Dim adobeReaderPath, pathAndFileName, shellPathName As String
    
   

Set myWorksheet = ActiveWorkbook.Worksheets(1)

Application.ScreenUpdating = False              # it's useful not to see the desktop working

adobeReaderPath = "insert here the .exe adobe path on your computer"
pathAndFileName = "insert here the path of your pdf file you want to pass on excel"
shellPathName = adobeReaderPath & " """ & pathAndFileName & """"

' shell is used to open other applications from excel 

Call Shell(shellPathName, vbNormalFocus)


Application.Wait Now + TimeValue("0:00:03")         # since opening a file can require some seconds, the vba code could continue reading the code so it's useful waiting some time

' SendKeys is the command for the shortcuts 
' You can find more info about the commands here  https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sendkeys-statement

SendKeys "%vpc"
SendKeys "^a"
SendKeys "^c"

Application.Wait Now + TimeValue("0:00:03")         # sendkeys like shell can take some seconds to work

Sheets.Add , Worksheets(1)

Sheets(2).Activate

ActiveSheet.Range("A1").Select
ActiveCell.PasteSpecial

VBA.Shell "TaskKill /F /IM AcroRd32.exe", vbHide    # TaskKill closes an external app 


Sheets.Add , Worksheets(2)



ActiveWorkbook.Sheets(2).Activate
ActiveSheet.Range("A1").Select

Do Until ActiveSheet.Range("A" & i) = ""

' Here you can start working in your excel that contains all the pdf informations

End Sub
