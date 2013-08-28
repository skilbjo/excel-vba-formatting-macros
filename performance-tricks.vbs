'1. Call an Excel Function within VBA

Application.WorksheetFunction.[excel function]

'2. Stop the screen from updating <-- Massively useful

Sub example()'
Application.SreenUpdating = False'

  'code here

Application.ScreenUpdating = True'
End Sub

'3. Stop cells from being recalculated

Sub example()
Application.Calculation = xlCalculationManual'

'code here

Application.Calculation = xlCalculationAutomatic'
End Sub

'4. File size and File Length

Sub fileSizeInfo()'
Dim myFilePath as String
Dim myFileSize as Integer

myFilePath = "C:\Users\jskilbeck\Documents"
myFileSize = FileLen(myFilePath & "workbook.xlsm")

End Sub

'5. Find an object values (eg. rbg, xl, vb)

'Enable the object browser using F2

'6. Exit Sub vs End Sub

'Exit Sub allows you to end the sub but still have code after the Exit Sub. A good use case for this is
'error trapping; eg. you want your errors to be within the same sub procedure but you don't want your
'code / loops to call the error trap if there was no error!
