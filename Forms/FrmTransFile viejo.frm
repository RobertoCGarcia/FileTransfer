VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRecepcionFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transferencia de Archivo"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Extraer 
      Caption         =   "&Extraer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2760
      Top             =   1560
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_Salir 
      Caption         =   "&Salir"
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
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar recieved 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblEnt 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folio Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label lbltime 
      Caption         =   "Tiempo Restante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Caption         =   "Sin Conexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label thespeed 
      Caption         =   "Velocidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label byteslabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "FrmRecepcionFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)


Dim mainbuffer As String ''
Dim sendsize As Integer ''
Dim sendmore As Integer ''
Dim thename As String ''
Dim filesize As Long ''
Dim currentint As Long ''
Dim rate As Integer ''
Dim filestart As Long ''



Public Sub AddStat(message As String)

lblstatus.Caption = message


End Sub


Private Sub cmd_Extraer_Click()

Dim NombreArch As String


NombreArch = InputBox("Ruta del Archivo .Dpk a Extraer:", "Extraer")


'If Len(NombreArch) <> 0 Then

 'Call Extraer_Archivos(thename)

Call ListFileContents(Opciones.Rutacarpeta_Recibidos & "\" & NombreArch)
'Ruta carpeta Extraidos sacar todos los archivos en esta carpeta
DoRestore Opciones.Rutacarpeta_Ensamble & "\"

MsgBox "Extraido con exito", vbInformation, "Extraido"
Exit Sub
'End If

End Sub

Private Sub cmd_Salir_Click()


Unload Me

End Sub

Private Sub Form_Load()

'Winsock1.Close
'Winsock1.LocalPort = 6002
'Winsock1.Listen
'AddStat "Listening"
lblEnt.Caption = Usuario.Entradas



sendsize = 1024


End Sub




Private Sub Timer1_Timer()

On Error GoTo timeerror


thespeed.Caption = "Velocidad: " & (rate / 2) & " KB/second (" & ((rate / 2) * 8) & " KBits/second)"

'Debug.Print "RESULTADO " & ((filesize - currentint) / ((rate / 2) * 1024))


If ((filesize - currentint) / ((rate / 2) * 1024)) <= 60 Then
    lbltime.Caption = "Tiempo Restante: " & Int((filesize - currentint) / ((rate / 2) * 1024)) & " segundos"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) > 60 And ((filesize - currentint) / ((rate / 2) * 1024)) <= 120 Then
    lbltime.Caption = "Tiempo Restante: 1 minuto"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) >= 120 And Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) < 60 Then
    lbltime.Caption = "Tiempo Restante: " & Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) & " minutos"
ElseIf ((filesize - currentint) / ((rate / 2) * 1024)) > 0 Then
    lbltime.Caption = "Tiempo Restante: " & Int(Int(Int((filesize - currentint) / ((rate / 2) * 1024)) / 60) / 60) & " horas"
End If


'Debug.Print "resultado: " & ((filesize - currentint) / ((rate / 2) * 1024))
rate = 0



Exit Sub
timeerror:
    lbltime.Caption = "Tiempo Restante: Indefinido"
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

'Debug.Print "REMOTE HOST : " & Recepcion.bd_RemoteHost
'Debug.Print "REMOTE ip : " & Recepcion.bd_RemoteIp
'Debug.Print "REMOTE PORT : " & Recepcion.bd_RemotePort



End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)




On Error Resume Next


'On Error GoTo tranyerror


MsgBox "Transmission Error"
Winsock1.Close
AddStat "Error Transfering"
Exit Sub

Dim temprecieve As String
DoEvents
Winsock1.GetData temprecieve





If InStr(1, temprecieve, "FILESIZE ") <> 0 Then
    filesize = Mid(temprecieve, 10, sendsize)
    
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
    FrmMain.lblport = "Recibido " & Int((currentint / filesize) * 100) & "%"
    
    Me.Caption = "Transferencia de Archivo - " & Int((currentint / filesize) * 100) & "%"
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Recibiendo Archivo: " & Int((currentint / filesize) * 100) & "%"
    
    Put #2, , temprecieve
    Winsock1.SendData "SENDMORE"
    
    rate = rate + 1
    byteslabel = "Recibido " & currentint & " of " & filesize & " bytes"
ElseIf InStr(1, temprecieve, "SRTDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    recieved.Value = Int(currentint / filesize * 100)
    AddStat "Recibido " & Int((currentint / filesize) * 100) & "%"
    FrmMain.lblport = "Recibido " & Int((currentint / filesize) * 100) & "%"
    
    Me.Caption = "Transferencia de Archivo - " & Int((currentint / filesize) * 100) & "%"
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Recibiendo Archivo: " & Int((currentint / filesize) * 100) & "%"
    
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
    FrmMain.lblport = "Terminado de Recibir OK!!"
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Recepcion Terminada!"
    Beep
    recieved.Value = 0
    byteslabel = ""
    thespeed = ""
    lbltime = ""
    FrmMain.lblport = ""
    
Recepcion.bd_Peticion_Final = Now
Recepcion.bd_ComentarioRecepcion = "Recepcion Terminada con exito"
Recepcion.bd_StatusRecepcion = "1"
    
    
    
    ' DoEvents

'Se espera 6.5 segundos para despues mandar extraer el contenido del archivo
Sleep 6525
 
FrmMain.lblport = "Archivo Recibido Adecuadamente"
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Archivo Recibido Adecuadamente"


Call Extraer_Archivos(thename)
  

  
End If





tranyerror:

MsgBox "Error en la Recepcion del Archivo, Presione Reset para Volver a Condiciones Normales", vbCritical, "Error en la Recepcion del Archivo"
'Para Evitar Errores con lo relacionado a los Archivos que llegan
Kill Opciones.Rutacarpeta_Ensamble & "\" & "*.*"
Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"
Exit Sub


End Sub


Public Sub SendFile()
'send the puppy
On Error GoTo tranyerror
'Para Evitar Errores con lo relacionado a los Archivos que llegan
Kill Opciones.Rutacarpeta_Ensamble & "\" & "*.*"
Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"


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

On Error GoTo tranyerror
'Para Evitar Errores con lo relacionado a los Archivos que llegan
Kill Opciones.Rutacarpeta_Ensamble & "\" & "*.*"
Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"



tranyerror:

MsgBox "Transmission Error"
Winsock1.Close
AddStat "Error Transfering"
Exit Sub

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'error
Winsock1.Close
AddStat "Error de Transferencia!!!"


Recepcion.bd_ComentarioRecepcion = "Error en recepcion:  " & Description & "   Numero: " & Number
Recepcion.bd_StatusRecepcion = "0"

'Para Evitar Errores con lo relacionado a los Archivos que llegan
Kill Opciones.Rutacarpeta_Ensamble & "\" & "*.*"
Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"




End Sub
