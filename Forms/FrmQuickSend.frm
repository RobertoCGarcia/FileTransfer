VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuickSend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quick Send"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pbx_Info 
      Height          =   2652
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   6075
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   6132
      Begin VB.TextBox txtComent 
         Height          =   852
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   360
         Width           =   5772
      End
      Begin VB.CommandButton cmd_OKInfo 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5040
         TabIndex        =   15
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   192
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   36
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Informacion del Favorito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   2016
      End
   End
   Begin VB.PictureBox Pbx_usuarios 
      Height          =   2652
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   6075
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   6132
      Begin VB.CommandButton cmd_OKUsers 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5040
         TabIndex        =   11
         Top             =   2160
         Width           =   852
      End
      Begin MSComctlLib.ListView LV_Usuarios 
         Height          =   1812
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5892
         _ExtentX        =   10398
         _ExtentY        =   3201
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario Destino"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Folio Salida"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IP"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Registro"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Lista de Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1572
      End
   End
   Begin VB.PictureBox Pbx_Archivos 
      Height          =   2652
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   6132
      Begin VB.CommandButton cmd_OKArch 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5040
         TabIndex        =   8
         Top             =   2160
         Width           =   852
      End
      Begin MSComctlLib.ListView LV_Archivos 
         Height          =   1812
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5892
         _ExtentX        =   10398
         _ExtentY        =   3201
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tamaño bytes"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ruta Origen"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ruta Destino"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Lista de Archivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1572
      End
   End
   Begin VB.CommandButton cmd_Info 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      MouseIcon       =   "FrmQuickSend.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmQuickSend.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Informacion acerca del Favorito Elegido"
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton cmd_usuarios 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      MouseIcon       =   "FrmQuickSend.frx":0594
      MousePointer    =   99  'Custom
      Picture         =   "FrmQuickSend.frx":06E6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Muestra la lista de usuarios que se le envian los archivos"
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton cmd_Archivos 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      MouseIcon       =   "FrmQuickSend.frx":0B28
      MousePointer    =   99  'Custom
      Picture         =   "FrmQuickSend.frx":0C7A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Muestra la lista de Archivos asocidos al Favorito"
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      MouseIcon       =   "FrmQuickSend.frx":10BC
      MousePointer    =   99  'Custom
      Picture         =   "FrmQuickSend.frx":120E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sale de QuickSend"
      Top             =   720
      Width           =   1092
   End
   Begin VB.ComboBox cb_Favoritos 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4572
   End
   Begin VB.CommandButton cmd_Iniciar 
      Caption         =   "&Iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MouseIcon       =   "FrmQuickSend.frx":1650
      MousePointer    =   99  'Custom
      Picture         =   "FrmQuickSend.frx":17A2
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sale de QuickSend"
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Favoritos Disponibles"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "FrmQuickSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFavorit As ADODB.Recordset
Dim RstUsers As ADODB.Recordset
Dim RstArchivos As ADODB.Recordset
Dim RstHeader As ADODB.Recordset
Dim strSql As String
Dim Mat_Info() As String
Dim elementoX As ListItem
Public Coment As String


Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF


Private Declare Function GetWindowsDirectory _
    Lib "KERNEL32" _
    Alias "GetWindowsDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long _
    ) As Long


Private Declare Function OpenProcess _
    Lib "KERNEL32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long _
    ) As Long


Private Declare Function WaitForSingleObject _
    Lib "KERNEL32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long _
    ) As Long


Private Declare Function GetExitCodeProcess _
    Lib "KERNEL32" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long _
    ) As Long


Private Declare Function CloseHandle _
    Lib "KERNEL32" ( _
    ByVal hObject As Long _
    ) As Long



Private Sub cb_Favoritos_Click()
Me.Height = 2200
Pbx_Archivos.Visible = False
Pbx_usuarios.Visible = False
Pbx_Info.Visible = False
FrmQuickSend.LV_Usuarios.ListItems.Clear
FrmQuickSend.LV_Archivos.ListItems.Clear

ObtenArchivos (Mat_Info(((cb_Favoritos.ListIndex) + 1), 2))
ObtenUsuarios (Mat_Info(((cb_Favoritos.ListIndex) + 1), 2))


End Sub

Private Sub cmd_Archivos_Click()

If Len(cb_Favoritos.Text) = 0 Then
    MsgBox "Elija un Favorito de manera adecuada!!!", vbCritical, "Dato requerido"
    cb_Favoritos.SetFocus
    Exit Sub
End If

'ObtenArchivos (Mat_Info(((cb_Favoritos.ListIndex) + 1), 2))

Me.Height = 5052
Pbx_Archivos.Visible = True
Me.Pbx_usuarios.Visible = False
Me.Pbx_Info.Visible = False








End Sub

Private Sub cmd_Cancel_Click()
Unload Me

End Sub


Private Sub cmd_Info_Click()


Dim TxtInfo As String


If Len(cb_Favoritos.Text) = 0 Then
    MsgBox "Elija un Favorito de manera adecuada!!!", vbCritical, "Dato requerido"
     cb_Favoritos.SetFocus
    Exit Sub
End If


txtcoment.Text = Mat_Info(((cb_Favoritos.ListIndex) + 1), 3) & vbCrLf

TxtInfo = TxtInfo & "USUARIO CREO: " & Mat_Info(((cb_Favoritos.ListIndex) + 1), 4) & vbCrLf
TxtInfo = TxtInfo & "FECHA CREACION: " & Mat_Info(((cb_Favoritos.ListIndex) + 1), 5) & vbCrLf
TxtInfo = TxtInfo & "USUARIO ACTUALIZO: " & Mat_Info(((cb_Favoritos.ListIndex) + 1), 6) & vbCrLf
TxtInfo = TxtInfo & "FECHA ACTUALIZACION: " & Mat_Info(((cb_Favoritos.ListIndex) + 1), 7) & vbCrLf

lblInfo = TxtInfo


Me.Height = 5052
Pbx_Archivos.Visible = False
Me.Pbx_usuarios.Visible = False
Me.Pbx_Info.Visible = True



End Sub

Private Sub cmd_Iniciar_Click()
Dim IdEnvio As String

If Len(cb_Favoritos.Text) = 0 Then
    MsgBox "Elija un Favorito de manera adecuada!!!", vbCritical, "Dato requerido"
    cb_Favoritos.SetFocus
    Exit Sub
End If


'Call ArchivoOid(2)


NombreArchivoGenerado = Usuario.Nick & Mid(CStr(Rnd), 3, 5)
'IdEnvio = NombreArchivoGenerado & "_" & Mid(CStr(Rnd), 3, 8)
Envio.OidEnv = Mat_Info(((cb_Favoritos.ListIndex) + 1), 2)

Screen.MousePointer = vbHourglass
Call Archivo_Enviar
Call DoBackup(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk", 1, RstArchivos)
Call ActualizaHeaderEnvio
Screen.MousePointer = vbDefault
Call ArchivoOid(2)
Unload Me
Call EnvioControlado
           



End Sub

Private Sub cmd_OKArch_Click()
Me.Height = 2200
Pbx_Archivos.Visible = False
Pbx_usuarios.Visible = False
Me.Pbx_Info.Visible = False

End Sub

Private Sub cmd_OKInfo_Click()
Me.Height = 2200
Pbx_Archivos.Visible = False
Pbx_usuarios.Visible = False
Me.Pbx_Info.Visible = False

End Sub

Private Sub cmd_OKUsers_Click()

Me.Height = 2200
Pbx_Archivos.Visible = False
Pbx_usuarios.Visible = False
Me.Pbx_Info.Visible = False



End Sub

Private Sub cmd_usuarios_Click()

If Len(cb_Favoritos.Text) = 0 Then
    MsgBox "Elija un Favorito de manera adecuada!!!", vbCritical, "Dato requerido"
    cb_Favoritos.SetFocus
    Exit Sub
End If


Me.Height = 5052
Pbx_Archivos.Visible = False
Me.Pbx_usuarios.Visible = True
Me.Pbx_Info.Visible = False






End Sub

Private Sub Form_Load()
Dim Inx As Integer

Me.Height = 2200
Pbx_Archivos.Visible = False


strSql = "SELECT * From FAVORITOS WHERE STATUS='OK';"
''Debug.Print "Cadena sql de envio pendientes consulta: " & strSql
'Debug.Print strSql


    Set RstFavorit = New ADODB.Recordset
    RstFavorit.CursorLocation = adUseClient
    RstFavorit.CursorType = adOpenDynamic
    RstFavorit.LockType = adLockPessimistic
    RstFavorit.Open strSql, CadenaCnx


    With RstFavorit
    
    If .RecordCount = 0 Then
      cb_Favoritos.Enabled = False
      Me.cmd_Iniciar.Enabled = False
    Else
      cb_Favoritos.Enabled = True
      Me.cmd_Iniciar.Enabled = True
      cmd_Archivos.Enabled = True
      cmd_usuarios.Enabled = True
      cmd_Info.Enabled = True
      
      
      
        .MoveFirst
        ReDim Mat_Info(.RecordCount, 8)
        Inx = 1
        Do While Not .EOF
          cb_Favoritos.AddItem RstFavorit!NOMBRE_FAVORITO
          
          Mat_Info(Inx, 1) = RstFavorit!IDFAVORITO 'idfavorito
          Mat_Info(Inx, 2) = RstFavorit!OidEnv 'oidenvio
          Mat_Info(Inx, 3) = RstFavorit!Comentario 'COMENTARIO
          Mat_Info(Inx, 4) = RstFavorit!USUARIO_CREO 'USUARIO_CREO
          Mat_Info(Inx, 5) = RstFavorit!FECHA_CREACION 'FECHA_CREACION
          Mat_Info(Inx, 6) = RstFavorit!USUARIO_ACTUALIZO 'USUARIO_ACTUALIZO
          Mat_Info(Inx, 7) = RstFavorit!FECHA_ACTUALIZACION 'FECHA_ACTUALIZACION
          Mat_Info(Inx, 8) = RstFavorit!ID_LISTA_USUARIOS 'ID_LISTA_USUARIOS
          
           
        .MoveNext
        Inx = Inx + 1
        Loop
    End If
    
    
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RstFavorit = Nothing
Set RstUsers = Nothing
Set RstArchivos = Nothing
Set RstHeader = Nothing
End Sub



Sub ObtenUsuarios(IdMov As String)


    FrmQuickSend.LV_Usuarios.ListItems.Clear
    strSql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & IdMov & "';"
    Set RstUsers = New ADODB.Recordset
    RstUsers.CursorLocation = adUseClient
    RstUsers.CursorType = adOpenDynamic
    RstUsers.LockType = adLockPessimistic
    RstUsers.Open strSql, CadenaCnx
    
    
    strSql = "SELECT * from HEADERENVIO WHERE OIDENV =  '" & IdMov & "';"
    Set RstHeader = New ADODB.Recordset
    RstHeader.CursorLocation = adUseClient
    RstHeader.CursorType = adOpenDynamic
    RstHeader.LockType = adLockPessimistic
    RstHeader.Open strSql, CadenaCnx
    
    
    RstUsers.MoveFirst
    
    Do While Not RstUsers.EOF
       Set elementoX = FrmQuickSend.LV_Usuarios.ListItems.Add(, , RstUsers!USUARIO_DESTINO)
       elementoX.Tag = RstUsers!OidEnv
       elementoX.SubItems(1) = RstUsers!FOLIO_SALIDA 'FOLIO_SALIDA
       elementoX.SubItems(2) = RstUsers!ip ' ip
       elementoX.SubItems(3) = RstUsers!FECHA_REGISTRO ' FECHA_REGISTRO
       RstUsers.MoveNext
    Loop
    
    Coment = RstHeader!Comentario
    'RstHeader.Close
End Sub



Sub ObtenArchivos(IdMov As String)
     
    
     FrmQuickSend.LV_Archivos.ListItems.Clear
     strSql = "SELECT * from enviados WHERE OIDENV =  '" & IdMov & "';"
     Set RstArchivos = New ADODB.Recordset
     RstArchivos.CursorLocation = adUseClient
     RstArchivos.CursorType = adOpenDynamic
     RstArchivos.LockType = adLockPessimistic
     RstArchivos.Open strSql, CadenaCnx
     
    RstArchivos.MoveFirst
    Do While Not RstArchivos.EOF
       Set elementoX = FrmQuickSend.LV_Archivos.ListItems.Add(, , RstArchivos!NOMBRE_ARCHIVO)
       elementoX.Tag = RstArchivos!OidEnv
       elementoX.SubItems(1) = RstArchivos!Tamaño ' tamaño archivo
       elementoX.SubItems(2) = RstArchivos!RUTA_ORIGEN ' ruta origen
       elementoX.SubItems(3) = RstArchivos!RUTA_DESTINO ' ruta destino
       RstArchivos.MoveNext
    Loop

     
End Sub


Sub Archivo_Enviar()

'Archivo de envio



Dim C As Integer
Dim RutExt As String
Dim NombreArchivoGenerado As String

' Genero el archivo que tiene informacion de todos los archivos elegidos


'Ruta carpeta ensamble



Open Opciones.Rutacarpeta_Ensamble & "\0x00.Raw" For Output As #1
            RstArchivos.MoveFirst
    
            Print #1, "[ENCABEZADO_ARCHIVO]"
            'Print #1, "CLIENTE_FUENTE=" & ArchivoMaestro.CLIENTE_FUENTE
                Print #1, "USUARIO_ORIGEN=" & Usuario.Nombre & " " & Usuario.Apellido
                Print #1, "OID_USUARIO_ORIGEN=" & Opciones.OidUserDefault
                Print #1, "FOLIO_SALIDA=" & Usuario.Salidas
                Print #1, "OID_MOVIMIENTO=" & Mat_Info(((cb_Favoritos.ListIndex) + 1), 2)
                Print #1, "No_ARCHIVOS=" & RstArchivos.RecordCount
                Print #1, "QUICKSEND=1"
                
               ' If Len(RstArchivos!NOMBRE_CMD) = 0 Then
                
               '    Print #1, "CMD=NOTHING"
                   
               ' Else
                
               '     Print #1, "CMD=" & RstArchivos!NOMBRE_CMD
                   
               ' End If
                
                Print #1, "FECHA_CREACION=" & Now
                Print #1, "COMENTARIO=" & Coment & vbCrLf
    
    For C = 1 To RstArchivos.RecordCount
            
                          
            Print #1, "[INFORMACION_ARCHIVO_" & C & "]"
            Print #1, "NOMBRE_ARCHIVO_" & C & "=" & RstArchivos!NOMBRE_ARCHIVO
            Print #1, "TAMAÑO=" & FileLen(RstArchivos!RUTA_ORIGEN)
            Print #1, "RUTA_ORIGEN=" & RstArchivos!RUTA_ORIGEN
                          
                          
            If RstArchivos!RUTA_DESTINO = "Ruta Default" Then
               Print #1, "EXTRAER=999" & vbCrLf
            Else
               Print #1, "EXTRAER=" & RstArchivos!RUTA_DESTINO & vbCrLf
            End If
    
    RstArchivos.MoveNext
    Next C


Close (1)



End Sub



Sub ActualizaHeaderEnvio()

With RstHeader
    'Actualiza el registro del Headerenvio para los datos del archivo comprimido que ya se genero
    
    If .RecordCount > 1 Then
    MsgBox "Error en la actualizacion del HeaderEnvio", vbCritical, "Error en Datos"
    Exit Sub
    End If
   
    RstHeader!NOMBRE_ARCHIVO_COMPRIMIDO = NombreArchivoGenerado & ".dpk"
    RstHeader!Tamaño = FileLen(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk")
    RstHeader!UBICACION = Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk"
    RstHeader!FECHA_REGISTRO = Now
    RstHeader!FECHA_REVISADO = Now
    .Update
    .Requery
    
End With




End Sub




Public Sub EnvioControlado()


   Dim sCmdLine As String
    Dim idProg As Long, iExit As Long
    
    
  'If InxPosEnv <= er Then
    
           sCmdLine = App.path & "\Envios.exe"
           
           idProg = Shell(sCmdLine, vbNormalFocus)
           
           iExit = fWait(idProg)
           
            Call ConsultarRecepcion
            Call ConsultarEnvio
            Call ConsultarPendientes
        
        
        
           If iExit Then
               
               MsgBox "Modulo de Envio no se ha podido iniciar, intentarlo de nuevo mas tarde", vbCritical, "Error de la aplicacion"
               
           Else
           
              ' MsgBox "Proceso Finalizado", vbInformation, "Operacion Exitosa"
            ' Debug.Print "Proceso Finalizado", vbInformation, "Operacion Exitosa"
             
             ' se manda la primera vez, y despues se hace la validacion de volverlo mandar
              ' InxPosEnv = InxPosEnv + 1
           
           'Debug.Print "Finished processing."
           End If

 ' End If
  
  

End Sub


Function fWait(ByVal lProgID As Long) As Long
    ' Wait until proggie exit code <>
    '     STILL_ACTIVE&
    Dim lExitCode As Long, hdlProg As Long
    ' Get proggie handle
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    ' Get current proggie exit code
    GetExitCodeProcess hdlProg, lExitCode


    Do While lExitCode = STILL_ACTIVE&


        DoEvents
            GetExitCodeProcess hdlProg, lExitCode
        Loop
        CloseHandle hdlProg
        fWait = lExitCode

End Function

