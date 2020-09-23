Attribute VB_Name = "MainBas"
Option Explicit

Public gsDatabase As String
Public gConn As ADODB.Connection

Public Sub Main()
    Dim sMsg As String
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    'Check if application is already running
    If App.PrevInstance = True Then
        MsgBox "The CMYK Conversion System is already running.", vbOKOnly + vbExclamation, "Order Entry System"
        End
    End If
    
    'Open application database
    Call OpenDatabase
    
    'Show main form
    frmMain.Show
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub OpenDatabase()
    Dim sMsg As String
    Dim sConnect As String
    
    On Error GoTo OPEN_ERROR
    
    gsDatabase = App.Path & "\PantoneConv.mdb"
    sConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gsDatabase & ";Uid=;Pwd=;"
    Set gConn = New ADODB.Connection
    gConn.ConnectionString = sConnect
    gConn.Open
    
Exit Sub
OPEN_ERROR:
    sMsg = "Error opening database '" & gsDatabase & "'.  Error: " & Err.Description
    MsgBox sMsg, vbOKOnly + vbCritical, "System Error"
    End
End Sub
