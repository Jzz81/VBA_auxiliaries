Attribute VB_Name = "template_load"
Option Explicit
Option Base 0
Option Compare Text

'made by: Joos Dominicus (Jzz)
'email: joos.dominicus@gmail.com
'version: 1_0_0

'module to be use as a template loader. This allows you to have a system of versioning in your
'global templates. This module holds the subs to load the newest version of the templates from
'the directory where they are stored. Particularly handy if you work with network clients. One
'network directory with the newest template files will serve all the clients. This is the only
'module that needs to be installed locally on every client.

'add:
'Private Sub Workbook_Open()
'Call template_load.load_all_templates
'End Sub
'to 'ThisWorkbook' code, to execute the load_all_templates sub on the start of excel.

Const template_store_dir As String = "<<PATH TO NETWORK FOLDER>>\excel_startup_templates\"
Const DEBUG_MODE As Boolean = False

'load_all_templates will look into the template_store_dir path and find all subdirectories there.
'it expects the template files in those subdirectories.
Public Sub load_all_templates()
'finds all folders in the directory
Dim s As String
Dim c As New Collection
Dim i As Long

s = Dir(template_store_dir, vbDirectory)
'load folders in collection (to avoid 'dir()' conflicts later on)
Do Until s = vbNullString
    If s <> "." And s <> ".." Then
        c.Add template_store_dir & s & "\"
    End If
    s = Dir()
Loop

'load templates in each folder
For i = 1 To c.Count
    Call load_template(c.Item(i))
Next i

Set c = Nothing

End Sub

'load_template will take the full path of the directory as an argument (d). In the given
'directory, it will look for the newest version of the files and load that.
Private Sub load_template(d As String)
'find newest template file
Dim s As String
Dim version As Long
Dim latest_version As String
Dim v As Long

s = Dir(d & "*.xlam")
'loop all files in the folder and find the latest version
Do Until s = vbNullString
    v = parseversion(s)
    'if v = 0, then no templates are found (only development versions, and debug_mode is not
    'set)
    If v > version And v > 0 Then
        version = v
        latest_version = s
    End If
    s = Dir()
Loop

If latest_version <> vbNullString Then
    Application.Workbooks.Open (d & latest_version)
End If
    
End Sub
'parseversion will parse the filename. It expects the filename to be constructed
'like: 'somefilename_1_0_0' the underscores are important, for they seperate the
'version filename and the version digits. The DEBUG flag will force the function
'to consider the last number as well (for debugging and development). If it is
'false, it will only consider the version if the last number is 0.
Public Function parseversion(s As String) As Long
'function to parse the version number from the file name
Dim ss() As String
Dim i As Long
Dim vers As String

'split the filename on the dot (cut the extension string) and then split to
'underscores
ss = Split(Split(s, ".")(0), "_")

'if debug_mode is not set, consider only production versions (with a 0 as a last
'number)
If Not DEBUG_MODE And val(ss(UBound(ss))) <> 0 Then Exit Function

'get the version number from the filename
For i = 0 To UBound(ss)
    If IsNumeric(ss(i)) Then
        'construct a string with all version numbers
        vers = vers & Format(ss(i), "00")
    End If
Next i

'cast to a long
parseversion = CLng(vers)
End Function
