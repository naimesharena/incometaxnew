Attribute VB_Name = "PWD"
Sub sbUnProtectAll()
  
    On Error GoTo ErrorOccured
   
    Dim pwd1 As String
    pwd1 = InputBox("Please Enter the password")
    If pwd1 = "" Then Exit Sub
    For Each ws In Worksheets
        ws.Unprotect Password:=pwd1
    Next
    MsgBox "All sheets Unprotected."

    Exit Sub
     
ErrorOccured:
    MsgBox "Sheets could not be UnProtected - Password Incorrect"
    Exit Sub
    
End Sub

Sub sbProtectAll()
  
    On Error GoTo ErrorOccured
   
    Dim pwd1 As String
    pwd1 = InputBox("Please Enter the password")
    If pwd1 = "" Then Exit Sub
    For Each ws In Worksheets
        ws.Protect Password:=pwd1
    Next
    MsgBox "All sheets Protected."

    Exit Sub
     
ErrorOccured:
    MsgBox "Sheets could not be Protected - Password Incorrect"
    Exit Sub
    
End Sub

