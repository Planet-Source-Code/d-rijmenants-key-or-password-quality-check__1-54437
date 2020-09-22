Attribute VB_Name = "modCheckKey"
'-----------------------------------------------------------
'
'     Quality checking for keys, passwords and passfrazes
'
'                  written by D. Rijmenants
'
'-----------------------------------------------------------
'
' This function check the quality from a key, password or
' passfraze, by checking the combination of upper and lower
' cases, figures and signs. A good password should contain
' both lower and upper cases, as well as signsand figures,
' with a lenght of at least five chars, and doen't contain
' any repetitions. Following this guidelines will provide you
' the highest protection level. The function will return an
' integer from 0 (not secure) upto 100 (highest security)
'
' The function:
'
' return = KeyQuality(MyPassword)
'
' Where:
' MyPassword (string) the password or key to check
' return    (integer) integer 0-100 quality rating
'
' Example on using the function with a progressbar,
' to check during key entry in the txtKey textbox:
'
' Private Sub txtKey_Change()
' MyForm.ProgressBar1.Value = KeyQuality(MyForm.txtKey.text)
' End Sub
'
' Note: Set the ProgressBar Max value at 100
'
' Comments and suggestion most welcome at mail: dr.defcom@telenet.be
'
'
Public Function KeyQuality(ByVal aKey As String) As Integer
' returns an integer value (0 to 100) rating the key quality
Dim QC As Integer
Dim LN As Integer
Dim k As Integer
Dim Uc As Boolean
Dim Lc As Boolean
Dim Wid As Integer
Dim ValidKey As Boolean
LN = Len(aKey)
QC = LN * 4
'check key lenght (at least 5 chars!)
If Len(aKey) < 5 Then KeyQuality = 0: Exit Function
' check for repetitions (abcabc, aaaaa, 121212, etc.)
For Wid = 1 To Int(Len(aKey) / 2)
    ValidKey = False
    For k = Wid + 1 To Len(aKey) Step Wid
        If Mid(aKey, 1, Wid) <> Mid(aKey, k, Wid) Then ValidKey = True: Exit For
    Next
If ValidKey = False Then Exit For
Next
If ValidKey = False Then KeyQuality = 0: Exit Function
'check ucases and lcases
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 64 And Asc(Mid(aKey, k, 1)) < 91 Then Uc = True
    If Asc(Mid(aKey, k, 1)) > 96 And Asc(Mid(aKey, k, 1)) < 123 Then Lc = True
Next
If Uc = True And Lc = True Then QC = QC * 1.5
'check numbers
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 47 And Asc(Mid(aKey, k, 1)) < 58 Then
        If Uc = True Or Lc = True Then QC = QC * 1.5
        Exit For
        End If
Next
'check signs
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) < 48 Or Asc(Mid(aKey, k, 1)) > 122 Or (Asc(Mid(aKey, k, 1)) > 57 And Asc(Mid(aKey, k, 1)) < 65) Then QC = QC * 1.5: Exit For
Next
If QC > 100 Then QC = 100
KeyQuality = Int(QC)
End Function

