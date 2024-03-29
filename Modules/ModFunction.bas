Attribute VB_Name = "ModFunction"
Option Explicit
Public valTB() As Double   'used by sum

Public Function TrimSpaces(text As String) As String
    Dim Loop1 As Long, SpaceCheck As String
    Dim FullString As String
    For Loop1 = 1 To Len(text)
        SpaceCheck = Mid(text, Loop1, 1)
        If SpaceCheck <> " " Then
            FullString = FullString & SpaceCheck
        End If
    Next Loop1
    TrimSpaces = FullString
End Function

'Function used to format recordset
'/coded by edwin delos santos
Public Function FormatRS(ByVal srcField As Field) As String
    Dim strRet As String
     With srcField
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,###,##0.00")
        ElseIf srcField.Type = 7 Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        ElseIf srcField.Type = 3 Then
           If IsNumeric(srcField) Then
             strRet = Format$(srcField, "###,##0")
           End If
        ElseIf srcField.Type = 202 Or srcField.Type = 203 Then
            strRet = CStr(srcField)
        End If
    End With
    FormatRS = strRet
    strRet = vbNullString
End Function
'//Function used to display MONTHNAME
Public Function Month_Name(ByVal srcdate As Date) As String
Dim MonthNames As Variant
Dim moName As String
If Not IsDate(srcdate) Then Exit Function
MonthNames = Array("January", "February", "March", "April", "May", "June", _
                  "July", "August", "September", "October", "November", "December")
moName = Month(Format(srcdate, "MM/DD/YYYY"))
Month_Name = MonthNames(moName - 1)
End Function


Public Function WeekDay_Name(ByVal srcdate As Date) As String
Dim daynames() As Variant
If Not IsDate(srcdate) Then Exit Function
daynames = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
Dim wkDay As String
wkDay = Weekday(Format(srcdate, "MM/DD/YYYY"))
WeekDay_Name = daynames(wkDay - 1)
End Function

Public Function SplitString(ByVal strText As String) As String
'label1 = edwin_delos_santos
'<< syntax >>
'Private Sub Command1_Click()
'   Text1.Text = RemoveAllNonAlphaNumeric(Label1.Caption)
'End Sub
'result:  text1 = "edwin delos santos"
    Dim strResult As String
    Dim i As Integer

    For i = 1 To Len(strText)
        Select Case Asc(Mid$(strText, i, 1))
        Case 48 To 57, 65 To 90, 97 To 122 'a digit or Uppercase Alphabet or Lowercase Alphabet
            strResult = strResult & Mid$(strText, i, 1)
        Case Else 'Reject any other key.
            strResult = strResult & Space(1) 'add space
        End Select
    Next i

ExitHere:
    SplitString = strResult
End Function

Public Function splitStr(ByRef str As String, Optional ByRef chr As String = "_") As String
Dim iChar As Integer
Dim mystring
Dim sResult As String
iChar = InStr(1, str, chr, 1)  'search for the char "_"
mystring = Split(str, chr, -1, 1)
If iChar > 0 Then
  sResult = mystring(0) & Space(1) & mystring(1)
Else
  sResult = str
End If
  splitStr = sResult
End Function

'Function used to change the yes/no value
Public Function changeValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeValue = "1"
        Case "N": changeValue = "0"
        Case "1": changeValue = "Y"
        Case "0": changeValue = "N"
    End Select
End Function

