VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2400
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar recieved 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "c:\documents and settings\"
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdlisten 
      Caption         =   "Listen for File"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbltime 
      Caption         =   "Time Left"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Caption         =   "No Connection"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Save to folder/Source file:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label thespeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label byteslabel 
      Caption         =   "Bytes Label"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim mainbuffer As String
Dim sendsize As Integer
Dim sendmore As Integer
Dim thename As String
Dim filesize As Long
Dim currentint As Long
Dim rate As Integer
Dim filestart As Long


Private Sub cmdbrowse_Click()
On Error Resume Next
'set directory
CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
    txtfile.Text = CommonDialog1.FileName
    thename = CommonDialog1.FileTitle
End If

End Sub

Private Sub cmdlisten_Click()

Winsock1.Close
Winsock1.LocalPort = 9998
Winsock1.Listen
AddStat "Listening"
End Sub

Private Sub cmdSend_Click()
Dim theip As String
theip = InputBox("Enter Remote IP Address:", "IP Address", IIP)
AddStat "Connecting"
Winsock1.Close
Winsock1.Connect theip, 9998
End Sub

Private Sub Command4_Click()
Me.Hide
ipchat.Show
End Sub

Public Sub AddStat(message As String)

lblstatus.Caption = message


End Sub

Private Sub Form_Load()

sendsize = 1024
End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub Timer1_Timer()
On Error GoTo timeerror


thespeed.Caption = "Speed: " & (rate / 2) & " KB/second (" & ((rate / 2) * 8) & " KBits/second)"

If ((filesize - currentint) / ((rate / 2) * 1024)) <= 60 Then
    lbltime.Caption = "Time left: " & Int((filesize - currentint) / ((rate / 2) * 1024)) & " seconds"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) > 60 And ((filesize - currentint) / ((rate / 2) * 1024)) <= 120 Then
    lbltime.Caption = "Time left: 1 minute"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) >= 120 And Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) < 60 Then
    lbltime.Caption = "Time left: " & Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) & " minutes"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) > 0 Then
    lbltime.Caption = "Time left: " & Int(Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) / 60) & " hours"
End If
rate = 0

Exit Sub
timeerror:
    lbltime.Caption = "Time left: Infinity"
End Sub

Private Sub Winsock1_Close()

Winsock1.Close

AddStat "Conexion Cerrada"
End Sub

Private Sub Winsock1_Connect()

AddStat "Enviando Informacion del Archivo..."
Winsock1.SendData "FILESIZE " & FileLen(txtfile.Text)
DoEvents
Sleep 1000



Winsock1.SendData "SENDNAME " & thename
DoEvents
Sleep 1000



AddStat "Enviando el Archivo..."
SendFile


End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

'accept the request
AddStat "Connected"
'FrmMain.lblport = "Recibiendo Datos..."
FrmMain.lblport = "Aceptando conexion Remota"
Winsock1.Close
Winsock1.Accept requestID

Recepcion.bd_Peticion_Inicio = Now
Recepcion.bd_ComentarioRecepcion = "Inicio del Envio del Archivo"
'Recepcion.bd_StatusRecepcion = "1"


'//para el registro detallado en la tabla body

Recepcion.bd_RemoteHost = Winsock1.RemoteHost
Recepcion.bd_RemoteIp = Winsock1.RemoteHostIP
Recepcion.bd_RemotePort = Winsock1.RemotePort

Debug.Print "REMOTE HOST : " & Recepcion.bd_RemoteHost
Debug.Print "REMOTE ip : " & Recepcion.bd_RemoteIp
Debug.Print "REMOTE PORT : " & Recepcion.bd_RemotePort




'Dim OSF As Object 'Objeto para manipular los archivos
'Set OSF = CreateObject("scripting.filesystemobject")

'Dim Carpeta As Folder
'Set Carpeta = OSF.GetFolder(Opciones.Rutacarpeta_Recibidos)

' Carpeta.Files.Count

'    If Carpeta.Files.Count > 0 Then
    ' se borra todo lo que hay hay
'         Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"
    
'    End If


'FrmRecepcionFile.Show

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next
Dim temprecieve As String
DoEvents
Winsock1.GetData temprecieve

If InStr(1, temprecieve, "FILESIZE ") <> 0 Then
    filesize = Mid(temprecieve, 10, sendsize)
    
ElseIf InStr(1, temprecieve, "SENDNAME ") <> 0 Then
    
    thename = Mid(temprecieve, 10, sendsize)
    currentint = 0
    currentint = FileLen(txtfile.Text & thename)
    
    If currentint >= filesize Then
        Winsock1.SendData "ALLDONE"
        DoEvents
        Close #2
        Winsock1.Close
        AddStat "DONE!"
        Beep
        Exit Sub
    Else
        
        currentint = currentint + 1
        Winsock1.SendData "RESUME " & currentint
    End If
    Close #2
    
    Open txtfile.Text & thename For Binary Access Write As #2
    
ElseIf InStr(1, temprecieve, "THEDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    currentint = currentint + sendsize
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recieved " & Int((currentint / filesize) * 100) & "%"
    Me.Caption = "File Transfer - " & Int((currentint / filesize) * 100) & "%"
    
    Put #2, , temprecieve
    Winsock1.SendData "SENDMORE"
    
    rate = rate + 1
    byteslabel = "Recieved " & currentint & " of " & filesize & " bytes"
ElseIf InStr(1, temprecieve, "SRTDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recieved " & Int((currentint / filesize) * 100) & "%"
    Me.Caption = "File Transfer - " & Int((currentint / filesize) * 100) & "%"
    
    Put #2, currentint, temprecieve
    Winsock1.SendData "SENDMORE"
    
    rate = rate + 1
    byteslabel = "Recieved " & currentint & " of " & filesize & " bytes"
    currentint = currentint + sendsize
ElseIf InStr(1, temprecieve, "RESUME ") <> 0 Then
    
    currentint = Mid(temprecieve, 8, 20)
    filestart = currentint
    sendmore = 1
ElseIf temprecieve = "SENDMORE" Then
    
    sendmore = 1
ElseIf temprecieve = "ALLDONE" Then
    
    Close #2
    Winsock1.Close
    AddStat "DONE!"
    Beep
End If

End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'error
Winsock1.Close
AddStat "Error Transfering"
End Sub

Public Sub SendFile()
'send the puppy
On Error GoTo tranyerror
Dim tempbuffer As String

sendmore = 0
currentint = 0
filestart = 0

Do Until sendmore = 1
DoEvents
Loop

filesize = FileLen(txtfile.Text)
'open file to get info
Close #1
Open txtfile.Text For Binary Access Read As #1

tempbuffer = Space$(sendsize)

Get #1, filestart, tempbuffer

Winsock1.SendData "SRTDATA " & tempbuffer
sendmore = 0


Do Until EOF(1)

Do Until sendmore = 1
DoEvents
Loop


tempbuffer = Space$(sendsize)

Get #1, , tempbuffer



currentint = currentint + sendsize
recieved.Value = Int(currentint / filesize * 100)
AddStat "Sent " & Int((currentint / filesize) * 100) & "%"
Me.Caption = "File Transfer - " & Int((currentint / filesize) * 100) & "%"
byteslabel = "Sent " & currentint & " of " & filesize & " bytes"
rate = rate + 1


Winsock1.SendData "THEDATA " & tempbuffer
sendmore = 0





Loop

On Error Resume Next
Sleep 500

Close #1
DoEvents
Winsock1.SendData "ALLDONE"
DoEvents
Sleep 500
DoEvents
Winsock1.Close
DoEvents
AddStat "DONE!"
Exit Sub

tranyerror:

MsgBox "Transmission Error"
Winsock1.Close
AddStat "Error Transfering"
Exit Sub

End Sub
