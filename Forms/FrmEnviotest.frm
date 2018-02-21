VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmInfoEnvio 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progreso de Envio "
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Trans 
      Caption         =   "&Transmitir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   17
      Top             =   2400
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Sck_Envio 
      Left            =   8280
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Tim_Sock 
      Interval        =   2000
      Left            =   9480
      Top             =   3000
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Timer Tim_State 
      Interval        =   1000
      Left            =   9000
      Top             =   3000
   End
   Begin MSComctlLib.ProgressBar recieved 
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ListView Lv_Info 
      Height          =   1452
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   8532
      _ExtentX        =   15055
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Conexion"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Destino"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bytes Enviados"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Velocidad"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tiempo restante"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblReady 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1680
      TabIndex        =   20
      Top             =   2640
      Width           =   75
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Host Remoto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1920
      TabIndex        =   18
      Top             =   3120
      Width           =   105
   End
   Begin VB.Label lbltime 
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Caption         =   "No hay Conexion"
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label thespeed 
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label byteslabel 
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Conexion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   12
      Top             =   1920
      Width           =   1740
   End
   Begin VB.Label lblport 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label lbIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label lblArch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo Restante:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad Envio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes Enviados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto Remoto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion Remota:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje de Envio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2940
   End
End
Attribute VB_Name = "FrmInfoEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public EquipoRemotoListo As Boolean




Private Sub cmd_cancel_Click()

Sck_Envio.Close
'lblstatus1 = "Cancelado Todo"
Unload Me

End Sub
Private Sub cmd_Trans_Click()
cmd_Trans.Enabled = False


Call EnvioProgramado


End Sub



Private Sub Form_Load()

Dim i As Integer
Dim x As ListItem




'If UBound(Mat_Envio) = 0 Then
'Exit Sub
'End If


''''''''''resumen

For i = 1 To FrmEnvio!Lv_User.ListItems.Count

 
           Set x = Lv_Info.ListItems.Add(, , i)
           x.Tag = "e"
           
           x.SubItems(1) = "???" 'destino
           x.SubItems(2) = "0%" 'bytes enviados
           x.SubItems(3) = " " ' velocidad
           x.SubItems(4) = " "  ' tiempo restante
           x.SubItems(5) = "No enviado" ' status


Next i


End Sub



Private Sub Sck_Envio_Connect()



'TerminoProceso = False



''Debug.Print "Intentando conectarme..."


 'cmd_Inicio.Enabled = False
 
 

 ' lblstatus1 = "Enviando Informacion del Archivo"
 '              Wsk_Envio(Index).SendData "INFOFILE"
 
 
              
              
              Sck_Envio.SendData "FILESIZE " & FileLen(RutArchElegido)
              DoEvents
              Sleep 2000
              
              ' Debug.Print "Evento connect : " & "SENDNAME " & NombreArchivoGenerado & ".dpk"

              Sck_Envio.SendData "SENDNAME " & NombreArchivoGenerado & ".dpk"

              DoEvents
              Sleep 2000
    
             ' lblstatus1 = "Enviando Archivo"
            '  FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(Index).SubItems(7) = "Enviando Archivo..."
              
'             Call UpdateFile(Opciones.Rutacarpeta_Generados & "\" & Mat_Envio(C, 1), "ENV")
              Call SendFile(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk", 1)
     
           '  MsgBox "Terminado el Envio!!!", vbInformation, "Operacion Exitosa"
             
           '  Unload FrmInfoEnvio


             
             ' Sck_Envio(Index).SendData "ARCHIVE " & "1010101010101010101010101010101010101010000101010101010101010101001010"
             ' DoEvents
             ' Sleep 1000



    recieved.Value = 0
    byteslabel = ""
    thespeed = ""
    lbltime = ""





End Sub

'dentro del socket los indices del array son deacuerdo a los

Private Sub Sck_Envio_DataArrival(ByVal bytesTotal As Long)



'TerminoProceso = False




'cmd_Inicio.Enabled = False
 
On Error Resume Next
Dim temprecieve As String
DoEvents



Sck_Envio.GetData temprecieve
'Debug.Print "Recibido: " & temprecieve



If InStr(1, temprecieve, "FILESIZE ") <> 0 Then
    filesize(1) = Mid(temprecieve, 10, sendsize)
    
    
ElseIf InStr(1, temprecieve, "SENDNAME ") <> 0 Then
    
    thename = Mid(temprecieve, 10, sendsize)
    currentint(1) = 0
    currentint(1) = FileLen(Opciones.Rutacarpeta_Recibidos & thename)
    
    If currentint(1) >= filesize(1) Then
        Sck_Envio.SendData "ALLDONE"
        DoEvents
        Close #2
        Sck_Envio.Close
        AddStat "DONE!"
        Beep
        Exit Sub
    Else
        
        currentint(1) = currentint(1) + 1
        Sck_Envio.SendData "RESUME " & currentint(1)
    End If
    Close #2
    
    Open Opciones.Rutacarpeta_Recibidos & thename For Binary Access Write As #2
    
ElseIf InStr(1, temprecieve, "THEDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    currentint(1) = currentint(1) + sendsize
    recieved.Value = Int(currentint(1) / filesize(1) * 100)
    AddStat "Recieved " & Int((currentint(1) / filesize(1)) * 100) & "%"
    Me.Caption = "File Transfer - " & Int((currentint(1) / filesize(1)) * 100) & "%"
    
    Put #2, , temprecieve
    Sck_Envio.SendData "SENDMORE"
    
    rate(1) = rate(1) + 1
    byteslabel = "Recieved " & currentint(1) & " of " & filesize(1) & " bytes"
ElseIf InStr(1, temprecieve, "SRTDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    recieved.Value = Int(currentint(1) / filesize(1) * 100)
    AddStat "Recieved " & Int((currentint(1) / filesize(1)) * 100) & "%"
    FrmInfoEnvio.Caption = "File Transfer - " & Int((currentint(1) / filesize(1)) * 100) & "%"
    
    Put #2, currentint(1), temprecieve
    Sck_Envio.SendData "SENDMORE"
    
    rate(1) = rate(1) + 1
    byteslabel = "Recieved " & currentint(1) & " of " & filesize(1) & " bytes"
    currentint(1) = currentint(1) + sendsize
ElseIf InStr(1, temprecieve, "RESUME ") <> 0 Then
    
    currentint(1) = Mid(temprecieve, 8, 20)
    filestart(1) = currentint(1)
    sendmore(1) = 1
ElseIf temprecieve = "SENDMORE" Then
    
    sendmore(1) = 1
    
ElseIf temprecieve = "ALLDONE" Then
    
    Close #2
    Sck_Envio.Close
    AddStat "HECHO !"
    Beep
    
    
    
    
'ElseIf InStr(1, temprecieve, "OK ") <> 0 Then

    
End If




End Sub

Private Sub Sck_Envio_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)


'TerminoProceso = True



'cmd_Inicio.Enabled = False

MsgBox "Error de transferencia : " & Description & "  " & Number, vbCritical, "Error de Transferencia"
lblstatus = "Error de transferencia :  " & Description & " " & Number

'FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(Inx).SubItems(7) = "No Enviado Error..."
'FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(Inx).SubItems(7) = "Enviando Archivo..."
Sck_Envio.Close

'Mat_Envio(C, 4) = "No Enviado Error" 'Indicar que el archivo no se envio por algun error x




End Sub

Private Sub Tim_Envio_Timer()



     


















End Sub








'Timer del socket 1




Private Sub Tim_sock_Timer()


On Error GoTo timeerror


thespeed.Caption = "Velocidad: " & (rate(1) / 2) & " KB/segundo (" & ((rate(1) / 2) * 8) & " KBits/segundo)"

If ((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) <= 60 Then
    lbltime.Caption = "Tiempo Restante: " & Int((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) & " seconds"
ElseIf ((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) > 60 And ((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) <= 120 Then
    lbltime.Caption = "Tiempo Restante: 1 minute"
ElseIf ((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) >= 120 And Int(Int((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) / 60) < 60 Then
    lbltime.Caption = "Tiempo Restante: " & Int(Int((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) / 60) & " minutes"
ElseIf ((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) > 0 Then
    lbltime.Caption = "Tiempo Restante: " & Int(Int(Int((filesize(1) - currentint(1)) / ((rate(1) / 2) * 1024)) / 60) / 60) & " hours"
End If
rate(1) = 0

Exit Sub
timeerror:
    lbltime.Caption = "Tiempo Restante: Indefinido"


End Sub

Private Sub Tim_State_Timer()


Select Case Sck_Envio.State

 
 
Case Is = sckClosed

''Debug.Print "Socket Cerrado"

Case Is = sckOpen

''Debug.Print "Socket Abierto"


Case Is = sckListening

''Debug.Print "Escuchando"

Case Is = sckConnectionPending

''Debug.Print "Conexion Pendiente"

Case Is = sckResolvingHost

''Debug.Print "Resolviendo Host"

Case Is = sckHostResolved
''Debug.Print "Host Resuelto"

Case Is = sckConnecting
''lblstatusSck = "Conectando"

Case Is = sckConnected
''Debug.Print "Conectado"

Case Is = sckClosing
''Debug.Print "Cerrando Socket"

Case Is = sckError
''Debug.Print "Error en Socket"
 
End Select




End Sub



Public Sub AddStat(message As String)

'lblstatus1.Caption = message


End Sub



