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

'7. Keep variable global
'Preface the variable with `Static` and leave it undeclared. For example, to normally define a counter you'd 
'do Dim i as Integer. To keep the counter's value after the macro has been run, write: Static i. That's it!

'8. Test if a variable has been initialized or not with the following code (dependant on your variable type)

| Type                                 | Test                            | Test2
| Numeric (Long, Integer, Double etc.) | If obj.Property = 0 Then        | 
| Boolen (True/False)                  | If Not obj.Property Then        | If obj.Property = False Then
| Object                               | If obj.Property Is Nothing Then |
| String                               | If obj.Property = "" Then       | If LenB(obj.Property) = 0 Then
| Variant                              | If obj.Property = Empty Then    |
