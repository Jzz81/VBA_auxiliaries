Attribute VB_Name = "aux"
Option Explicit
Option Base 0
Option Compare Text

'module with general auxiliary functions and subs. To be used in Excel.
'made by: Joos Dominicus (Jzz)
'email: joos.dominicus@gmail.com
'latest revision date: 20-04-2015

Public Sub send_mail(Address As String, body As String, subject As String, Optional attach_path As String, Optional CC As String)
'sub that will send a mail message via outlook
Dim oApp As Object
Dim msg As Object

'get outlook application, error out if not found
On Error Resume Next
    Set oApp = CreateObject("Outlook.Application")
    If oApp Is Nothing Then
        MsgBox "Outlook not found", vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
On Error GoTo 0

'create mailitem (enum is 0)
Set msg = oApp.CreateItem(0)

'construct mail
With msg
    .To = Address
    If CC <> vbNullString Then
        .CC = CC
    End If
    .subject = subject
    .body = body
    If attach_path <> vbNullString Then
        On Error Resume Next
        .Attachments.Add attach_path
        Do Until Err.Number = 0
            DoEvents
            Err.Clear
            .Attachments.Add attach_path
        Loop
        On Error GoTo 0
    End If
    'display the message for review by user
    .Display
End With

End Sub
Public Function add_tag_to_string(ByRef s As String, Tag As String, val As String) As String
'adds a XML formatted tag with value to string s (seperated by a newline)
s = s & "<" & Tag & ">"
s = s & val
s = s & "</" & Tag & ">"
s = s & vbNewLine
End Function
Public Function get_numeric_value_from_string(s As String) As Long
'gets all numeric digits from a string and returns them as a long value
'example: "test001" will return 1. "01test_2015" will return 12015.
Dim i As Long
Dim n As String
For i = 1 To Len(s)
    If Mid(s, i, 1) Like "#" Then
        n = n & Mid(s, i, 1)
    End If
Next i
get_numeric_value_from_string = val(n)
End Function
Public Function string_is_in_collection(ByRef c As Collection, s As String, Optional leave As Boolean = False) As Boolean
'checks if string s is in collection c. If true, it deletes the string from the collection,
'unless leave is set to true.
'c must be a collections of strings
Dim i As Long
On Error Resume Next
For i = 1 To c.Count
    If c(i) = s Then
        string_is_in_collection = True
        If Not leave Then c.Remove (i)
        Exit For
    End If
Next i
On Error GoTo 0
End Function
Sub WriteLogFileEntry(entry As String, flag As Long)
'make a logfile entry with a flag in a logfile.
Dim ff As Long

ff = FreeFile
Open "<PATH TO LOGFILE>" For Append As #ff
Write #ff, Now & ": " & flag & ", " & entry
Close #ff

End Sub

