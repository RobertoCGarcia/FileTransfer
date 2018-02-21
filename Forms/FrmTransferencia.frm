VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRecepcion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recepcion de Archivo"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4440
      Top             =   3480
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5040
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_salir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar recieved 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar Pb_Extraccion 
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Extrayendo:"
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   4200
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tiempo Restante:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Velocidad:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Porcentaje:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label lbltime 
      Caption         =   "Time "
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Conexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label thespeed 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label byteslabel 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   4455
   End
End
Attribute VB_Name = "FrmRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mainbuffer As String
Dim sendsize As Integer
Dim sendmore As Integer
Dim thename As String
Dim filesize As Long
Dim currentint As Long
Dim rate As Integer
Dim filestart As Long


Private Sub cmd_salir_Click()

Me.Hide
End Sub

Public Sub AddStat(message As String)

lblstatus.Caption = message

'FrmMain.lblstatusRecep = message


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


'Debug.Print "resultado: " & ((filesize - currentint) / ((rate / 2) * 1024))
rate = 0

Exit Sub

timeerror:
   lbltime.Caption = "Time left: Infinity"

End Sub

Private Sub Winsock1_Close()

Winsock1.Close

AddStat "Conexion cerrada"


End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)


'accept the request





AddStat "Conectado"
Winsock1.Close
Winsock1.Accept requestID

Me.Show




End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next
Dim temprecieve As String
DoEvents
Winsock1.GetData temprecieve





If InStr(1, temprecieve, "FILESIZE ") <> 0 Then
    filesize_Recep = Mid(temprecieve, 10, sendsize)
    
ElseIf InStr(1, temprecieve, "SENDNAME ") <> 0 Then
    
    thename = Mid(temprecieve, 10, sendsize)
    currentint = 0
    currentint = FileLen(Opciones.Rutacarpeta_Recibidos & "\" & thename)
    
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
    
    Open Opciones.Rutacarpeta_Recibidos & "\" & thename For Binary Access Write As #2
    
ElseIf InStr(1, temprecieve, "THEDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    currentint = currentint + sendsize
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recibido " & Int((currentint / filesize) * 100) & "%"
    Me.Caption = "Transferencia de Archivo - " & Int((currentint / filesize) * 100) & "%"
    
    Put #2, , temprecieve
    Winsock1.SendData "SENDMORE"
    
    rate = rate + 1
    byteslabel = "Recibido " & currentint & " of " & filesize & " bytes"
ElseIf InStr(1, temprecieve, "SRTDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recibido " & Int((currentint / filesize) * 100) & "%"
    Me.Caption = "Transferencia de Archivo - " & Int((currentint / filesize) * 100) & "%"
    
    Put #2, currentint, temprecieve
    Winsock1.SendData "SENDMORE"
    
    rate = rate + 1
    byteslabel = "Recibido " & currentint & " de " & filesize & " bytes"
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
    
    AddStat "HECHO!"
    
    'Llamo al vento para poder interpretar el archivo que llego y extraer el contenido del mismo
        
    Beep
    recieved.Value = 0
    byteslabel = ""
    thespeed = ""
    lbltime = ""
   
    
    ' DoEvents
    ' Sleep 1000
    
   '   Call UpdateFile(Opciones.Rutacarpeta_Generados & "\" & Mat_Envio(c, 1), "ENV")
    
     Call UpdateFile(Opciones.Rutacarpeta_Recibidos & "\" & thename, "RCP")
     
 
 ' Se envia un mensaje de que existe un archivo pendiente por extraer
 
''     Call Extraer_Contenido(thename_Recep)
     Call Extraer_Archivos(thename)
    


ElseIf InStr(1, temprecieve, "READY ") <> 0 Then 'Para saber si esta activo
 Winsock1.SendData "OK" & Opciones.PuertoRecepcion
  
End If

End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'error
Winsock1.Close
AddStat "Error de Transferencia:   " & Number & "    " & Description

End Sub
