Attribute VB_Name = "aux"
Option Explicit
Option Base 0
Option Compare Text

Private Type PLASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

Private Declare Function GetLastInputInfo Lib "user32.dll" (ByRef plii As PLASTINPUTINFO) As Long
'gettickcount will return the number of milliseconds since the computer is turned on
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Purpose   :    Returns the number of milliseconds the system (or computer) has been idle for.
'Inputs    :    N/A
'Outputs   :    Returns the number of milliseconds the system has been idle for.
'Date      :    25/03/2005
'Author    :    Andrew Baker(www.vbusers.com)
'Notes     :
Public Function SystemIdleTime() As Long
    Dim tLastInput As PLASTINPUTINFO

    tLastInput.cbSize = Len(tLastInput)
    
    Call GetLastInputInfo(tLastInput)
    
    SystemIdleTime = GetTickCount - tLastInput.dwTime
End Function

Public Sub ProtectDoc(doc As Document)
'sub to lock the document without password and allow formfields
If doc.ProtectionType = wdNoProtection Then doc.Protect Type:=wdAllowOnlyFormFields, noreset:=True

End Sub
Public Sub UnProtectDoc(doc As Document)
'sub to remove lock, without a password
On Error Resume Next
If doc.ProtectionType <> wdNoProtection Then doc.Unprotect

If Err.Number = 5485 Then
    MsgBox "A password is required."
End If
On Error GoTo 0

End Sub

