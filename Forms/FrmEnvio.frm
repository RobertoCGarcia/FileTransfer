VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEnvio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generador de Archivo para Envio"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cb_CmdAsociado 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5640
      Width           =   4335
   End
   Begin VB.CommandButton cmd_Ruta 
      Caption         =   "(R)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   18
      ToolTipText     =   "Agrega una Ruta al Archivo seleccionado"
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmd_EnvioRapido 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "FrmEnvio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Guarda los Datos para Realizar Envios Rapidos"
      Top             =   6120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CD_Arch 
      Left            =   7080
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   7140
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   5292
            MinWidth        =   5292
            TextSave        =   "02:36 p.m."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Usuario:"
            TextSave        =   "Usuario:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Folio Salida:"
            TextSave        =   "Folio Salida:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd_AddArch 
      Caption         =   " (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      ToolTipText     =   "Agrega un archivo a la lista"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmd_Del 
      Caption         =   "(-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      ToolTipText     =   "Elimina un archivo de la lista"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmd_salir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Picture         =   "FrmEnvio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancela en envio de informacion"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox txtComent 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4680
      Width           =   6852
   End
   Begin VB.CommandButton cmd_quitar 
      Caption         =   "(-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.ComboBox cb_Usuario 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmd_Agregar 
      Caption         =   " (+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin MSComctlLib.ListView Lv_User 
      Height          =   1452
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6852
      _ExtentX        =   12091
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.ListView LV_ArchivosElegidos 
      Height          =   2052
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6852
      _ExtentX        =   12091
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre Archivo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tamaño (Kb)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ruta Completa"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ruta a Extraer"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.CommandButton cmd_EnvioInfo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "FrmEnvio.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Inicia el envio de Informacion"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comando Asociado"
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
      Height          =   300
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   2340
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Favoritos"
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
      Left            =   720
      TabIndex        =   17
      ToolTipText     =   "Guarda los Archivos necesarios para poder realiazar un Envio Rapido"
      Top             =   6720
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   5760
      TabIndex        =   15
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Generar Envio"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   6720
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   1428
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Archivos a Enviar:"
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
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2124
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
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
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   276
   End
End
Attribute VB_Name = "FrmEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private sysTray As Object

Public InxRutArch As Integer ' Para la ruta elegida


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



Private Sub Cb_CmdAsociado_Click()


Envio.bd_Cmd = Cb_CmdAsociado.Text

End Sub

Private Sub cb_Usuario_Click()

'MsgBox IPUsuarios(cb_Usuario.ListIndex + 1, 3)


Lv_User.Enabled = True

End Sub

Private Sub cmd_AddArch_Click()


Dim X As ListItem

    Dim sFile As String
    
    
    
    With CD_Arch
        .DialogTitle = "Abrir"
        .CancelError = False
        
        'Pendiente: establecer los indicadores y atributos del control common dialog
        .Filter = "Todos los archivos (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    'Procedimiento para la ruta a la cual sera extraido cuando sea recibido remotamente

          Set X = LV_ArchivosElegidos.ListItems.Add(, , Dir(sFile))
           X.Tag = sFile
           X.SubItems(1) = Round(FileLen(sFile) / 1024, 4) 'tamaño del archivo kb
           X.SubItems(2) = sFile ' ruta del archivo
           X.SubItems(3) = ""
           

  
    
   ' lblNo.Caption = LV_ArchivosElegidos.ListItems.Count
    
   cmd_EnvioInfo.Enabled = True
   LV_ArchivosElegidos.Enabled = True
   cmd_EnvioRapido.Enabled = True
   
   
End Sub

Private Sub cmd_Agregar_Click()

Dim Inx As Integer
Dim X As ListItem




If Len(cb_Usuario.Text) = 0 Then
MsgBox "Elija Adecuadamente a un usuario de destino!!!", vbCritical, "Dato Necesario"
cb_Usuario.SetFocus
Exit Sub
End If

'FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(f).SubItems(6)

    For Inx = 1 To Lv_User.ListItems.Count
    
    
        If Lv_User.ListItems.Item(Inx).SubItems(1) = cb_Usuario.Text Then
            
            MsgBox "El usuario de destino ya fue agregado, elija otro porfavor", vbInformation, "No Permitido"
            cb_Usuario.Clear
            Call ObtenUsuarios
            cb_Usuario.SetFocus
            Exit Sub
        
        End If
    
    
    
    
    Next Inx



            If IPUsuarios((cb_Usuario.ListIndex + 1), 2) = Usuario.OidUsuario Then
            
                MsgBox "Imposible Agregar al Usuario default en la Lista", vbInformation, "No Permitido"
                cb_Usuario.Clear
                Call ObtenUsuarios
                cb_Usuario.SetFocus
                Exit Sub
            
            End If





       Set X = Lv_User.ListItems.Add(, , cb_Usuario.ListIndex + 1)
           X.Tag = "e"
           
           X.SubItems(1) = cb_Usuario.Text 'IPUsuarios(cb_Usuario.ListIndex + 1)   'IP DESTINO
        
           cmd_AddArch.Enabled = True
           
           'x.SubItems(2) = IPUsuarios(cb_Usuario.ListIndex + 1) ' ip equipo
           ' If TEnv = False Then ' si se eligio el tipo de generar archivo
           ' x.SubItems(3) = FileLen(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk ") / 1024
           ' x.SubItems(5) = Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk " ' ruta completa del archivo
           ' x.SubItems(6) = NombreArchivoGenerado & ".dpk "
           ' Else
           ' x.SubItems(3) = FileLen(RutArchElegido)
           ' x.SubItems(5) = RutArchElegido ' ruta completa del archivo
           ' x.SubItems(6) = Right(RutArchElegido, 12)
           ' End If
           ' x.SubItems(4) = Now ' fecha de envio
           'x.SubItems(7) = "Archivo No Enviado"



End Sub


Private Sub cmd_Del_Click()



If LV_ArchivosElegidos.ListItems.Count > 0 Then


    With LV_ArchivosElegidos
    
    
        If .ListItems.Count > 0 Then
            .ListItems(.SelectedItem.Index).Selected = True
            .ListItems.Remove .SelectedItem.Index
        End If
        
        
        If .ListItems.Count = 0 Then
        
            cmd_EnvioInfo.Enabled = False
            LV_ArchivosElegidos.Enabled = False
            cmd_EnvioRapido.Enabled = False
            
        End If
      
        
    End With


  '  lblNo.Caption = LV_ArchivosElegidos.ListItems.Count

Else


MsgBox "Elija un elemento Adecuadamente", vbCritical, "Error de Seleccion"


End If










End Sub

Private Sub cmd_EnvioInfo_Click()

' existen 2 tipos de envio el programado y el que se hace en el momento

Dim IdEnvio As String
Dim X As Integer

Favoritos.Bd_IDFAVORITO = "F_000000000"
Favoritos.Bd_NOMBREFAVORITO = "NO DEFINIDO"

NombreArchivoGenerado = Usuario.Nick & Mid(CStr(Rnd), 3, 5)
IdEnvio = NombreArchivoGenerado & "_" & Mid(CStr(Rnd), 3, 8)
Envio.OidEnv = IdEnvio

Screen.MousePointer = vbHourglass

If Len(txtComent.Text) = 0 Then
    MsgBox "Comentario necesario para realizar envio.", vbExclamation, "Dato Necesario"
    txtComent.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
End If



Call Archivo_Enviar



'realiza el archivo generado
Call DoBackup(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk", 0)

Screen.MousePointer = vbDefault
RutArchElegido = Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk"
Envio.bd_Ruta_Arch = RutArchElegido

' SE GENERA LA MATRIZ LISTA PARA ENVIAR EL ARCHIVO GENERADO



ReDim Mat_Envio(1 To Lv_User.ListItems.Count, 1 To 6)

For X = 1 To Lv_User.ListItems.Count
  Mat_Envio(X, 1) = IPUsuarios(Lv_User.ListItems.Item(X), 1)
  'Debug.Print Mat_Envio(x, 1) '& vbCrLf
  Mat_Envio(X, 2) = IIf(Ping(IPUsuarios(Lv_User.ListItems.Item(X), 1)), "1", "0")
  '.Print Mat_Envio(x, 2) '& vbCrLf
  If Mat_Envio(X, 2) = "0" Then
  
    Mat_Envio(X, 3) = "Equipo remoto no encontrado con Ping"
    Mat_Envio(X, 4) = "0"
    
  Else
  
    Mat_Envio(X, 3) = "Equipo remoto encontrado"
    Mat_Envio(X, 4) = "9"
    
  End If
  'Debug.Print Mat_Envio(x, 3) & vbCrLf
 
  Mat_Envio(X, 5) = Opciones.PuertoSalida ' puerto

  Mat_Envio(X, 6) = IPUsuarios(Lv_User.ListItems.Item(X), 2) ' UID
  
  '1 col: ip del equipo remoto
  '2 col resultado del status 1 conectado y listo 0 no esta listo
  '3 col comentario del status del posible error o causas de porque fue el error
  '4 cola de envio pendiente 9 se inicio apenas, 1 enviado con exito, 0 error de envio
  '5 puerto remoto
  '6 UID USUARIO ID
  'nombre del archivo con todo y ruta
Next X



'Una vez la matriz generada se procede con el envio
' invocando al programa adecuado
' primero se genera el archivo envio.lst
' que contiene los datos de la matriz



' se generO el archivo CON LA LISTA DE ENVIO AHORA hay que invocar al programa de envio y
' esperar que finalice para enviar el siguiente y asi hasta terminar la lista

''InxPosEnv = 1

  Open App.path & "\Generados\Lista\Envio.lst" For Output As #1
  
                            Print #1, "[INFO_ARCHIVO]"
                            Print #1, "Nombre=" & NombreArchivoGenerado & ".dpk"
                            Envio.bd_Nombre = NombreArchivoGenerado & ".dpk"
                            
                            Print #1, "Ruta=" & RutArchElegido
                            Print #1, "Fecha_Creacion= " & Now
                            Envio.bd_FECHA_CREACION = Now
                     
                            Print #1, "NumeroEnvios= " & Lv_User.ListItems.Count
                            Envio.bd_NoArchivos = LV_ArchivosElegidos.ListItems.Count
                            Envio.bd_NoDestinos = Lv_User.ListItems.Count
                     
                            Print #1, "UsuarioOrigen= " & Usuario.Nombre & " " & Usuario.Apellido
                            
                            
                            Print #1, "UID=" & Opciones.OidUserDefault
                            Envio.bd_UIDorigen = Opciones.OidUserDefault
                     
                            Print #1, "OIDMOV= " & IdEnvio
                            
                  
                            If Len(txtComent.Text) = 0 Then
                               Print #1, "ComentarioMain= Sin comentario" & vbCrLf
                               Envio.bd_ComentarioMain = "Sin comentario"
                            Else
                               Print #1, "ComentarioMain=" & txtComent.Text & vbCrLf
                               Envio.bd_ComentarioMain = txtComent.Text
                            End If
                            
                            Call RegistrarEnvio(1) ' Registro el header de la tabla envio
                  
                        
                        
                        
                         For X = 1 To Lv_User.ListItems.Count
                         
                           ' If Mat_Envio(x, 2) = "1" Then
                             Print #1, "[USUARIO_" & X & "]"
                             
                                    Print #1, "IP = " & Mat_Envio(X, 1)
                                    Envio.bd_IP = Mat_Envio(X, 1)
                                    
                                    Print #1, "PING = " & Mat_Envio(X, 2)
                                    Envio.bd_PING = Mat_Envio(X, 2)
                             
                                    Print #1, "COMENTARIO = " & Mat_Envio(X, 3)
                                    Envio.Bd_COMENTARIO = Mat_Envio(X, 3)
                                    
                                    Print #1, "ENVIO = " & Mat_Envio(X, 4)
                                    Envio.bd_Envio = Mat_Envio(X, 4)
                                    
                                    Print #1, "PUERTO = " & Mat_Envio(X, 5)
                                    Envio.bd_PUERTO = Mat_Envio(X, 5)
                                    
                                    
                                    Print #1, "FECHA_ENVIO = " & Now
                                    
                                    'Nombre del usuario destino
                                     Envio.bd_UsuarioDestino = Mid(Lv_User.ListItems.Item(X).SubItems(1), 5, Len(Lv_User.ListItems.Item(X).SubItems(1)))
                                     'MsgBox Lv_User.ListItems.Item(X).SubItems(1)
                                    Print #1, "UID =" & Mat_Envio(X, 6) & vbCrLf
                                    Envio.bd_UID = Mat_Envio(X, 6) ' el oid del usuario que es destino al que se le va enviar
                            
                            ' Print #1, "NOMBRE_USUARIO= "
                             
                             Call RegistrarEnvio(2) ' Registro el body de la tabla envio
 
                             
                             'Call EnvioControlado
                          '   End If
                         Next
     
     
     
     Close (1)
     
' se registran todos los archivos que se van a enviar
Call RegistrarEnvio(3)
     
' se actualiza el contador de las salidas en 1 para ir avanzando
'con cada salida

     
Call ArchivoOid(2)


Call RegistraEvento("PROGRAMA", 1, "Envio Generado", "Folio Salida: " & Usuario.Salidas)

Call InformacionUsuario(Usuario.OidUsuario, "UPDATE_SALIDAS")
Call InformacionUsuario(Usuario.OidUsuario, "CONSULTA")
     
     
     
     
'el envio se va a hacer en base a lo que se registro en la base de datos

     
     
     
     
'     For X = 1 To Lv_User.ListItems.Count
          Unload Me
'         If Mat_Envio(X, 2) = "1" Then
         
           Call EnvioControlado
           
'         End If
         
'     Next X
     
      
      

'   idProg = Shell(App.path & "\Envios.exe", vbNormalFocus)
      
      
      
  '  MsgBox "=================TODO FINALIZADO!!!!!==============="
    
    
     'AQUI SE INICIA EL REGISTRO DE TODO LO QUE SE ENVIO
     
'LEO EL ARCHIVO COMPLETO YA GENERADO CON TODOS LOS RESULTADOS Y
'LO AGREGO A LA BD DE MOVIMIENTOS

'Call LAEP

Call FileCopy(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk", Opciones.Rutacarpeta_Enviados & "\" & NombreArchivoGenerado & ".dpk")

End Sub

Private Sub cmd_EnvioRapido_Click()

'Unload Me
frmAddFavorito.Show

End Sub

Private Sub cmd_quitar_Click()


If Lv_User.ListItems.Count > 0 Then


    With Lv_User
    
    
        If .ListItems.Count > 0 Then
            .ListItems(.SelectedItem.Index).Selected = True
            .ListItems.Remove .SelectedItem.Index
        End If
        
        
        If .ListItems.Count = 0 Then
           ' cmd_Enviar.Enabled = False
             LV_ArchivosElegidos.Enabled = False
             cmd_AddArch.Enabled = False
            
            
            
       End If
        
        
        
        
    End With

End If



End Sub

Private Sub cmd_Ruta_Click()


If InxRutArch = 0 Then
MsgBox "Elija adecuadamente el elemento de la lista", vbCritical, "Error de Seleccion"
Exit Sub
End If


FrmDireccion.Show


End Sub

Private Sub cmd_Salir_Click()
        Unload Me
       
End Sub







Private Sub Form_Load()


Call ObtenUsuarios
Call Comandos(5)
SBar.Panels(2).Text = Usuario.Nombre & " " & Usuario.Apellido
SBar.Panels(3).Text = "Folio Salida: " & Usuario.Salidas
               


'Lv_User.HideSelection



End Sub






Public Sub ObtenUsuarios()

Dim C As Integer
Dim Str_dato As String
Dim Rst_LstUsuarios As ADODB.Recordset

Str_dato = "SELECT * FROM USUARIOS WHERE STATUS='OK'"
'Rst_LstUsuarios
Set Rst_LstUsuarios = New ADODB.Recordset
    Rst_LstUsuarios.CursorLocation = adUseClient
    Rst_LstUsuarios.CursorType = adOpenDynamic
    Rst_LstUsuarios.LockType = adLockOptimistic
    Rst_LstUsuarios.Open Str_dato, CadenaCnx



ReDim IPUsuarios(1 To Rst_LstUsuarios.RecordCount, 3)


Rst_LstUsuarios.MoveFirst
C = 0

Do While Not Rst_LstUsuarios.EOF

C = C + 1
    cb_Usuario.AddItem C & ".- " & Rst_LstUsuarios.Fields("NOMBRE_USUARIO") & Space(1) & Rst_LstUsuarios.Fields("APELLIDO_USUARIO")
    IPUsuarios(C, 1) = Rst_LstUsuarios.Fields("IP_EQUIPO")
    IPUsuarios(C, 2) = Rst_LstUsuarios.Fields("OID_USUARIO")
    IPUsuarios(C, 3) = Rst_LstUsuarios.Fields("NOMBRE_USUARIO") & Space(1) & Rst_LstUsuarios.Fields("APELLIDO_USUARIO")
  
    
    
  '  Debug.Print IPUsuarios(C, 1) & "     " & IPUsuarios(C, 2) & "  " & vbCrLf
    
    
    Rst_LstUsuarios.MoveNext
    
    
    
    
Loop

Rst_LstUsuarios.Close

Set Rst_LstUsuarios = Nothing

End Sub



Public Function Ruta_Carpeta_Remota() As String

FrmDireccion.Show


End Function

Private Sub LV_ArchivosElegidos_ItemClick(ByVal Item As MSComctlLib.ListItem)

'MsgBox Item.Text


'FrmDireccion.Show
 InxRutArch = Item.Index
 
'Set X = LV_ArchivosElegidos.ListItems.Add(, , Dir(sFile))
'           X.Tag = sFile
'           X.SubItems(1) = Round(FileLen(sFile) / 1024, 4) 'tamaño del archivo kb
'          X.SubItems(2) = sFile ' ruta del archivo
'           X.SubItems(3) = ""



End Sub





Private Sub txtcoment_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Sub Archivo_Enviar()

'Archivo de envio



Dim C As Integer
Dim RutExt As String
Dim NombreArchivoGenerado As String

' Genero el archivo que tiene informacion de todos los archivos elegidos


'Ruta carpeta ensamble



Open Opciones.Rutacarpeta_Ensamble & "\0x00.Raw" For Output As #1

    
            Print #1, "[ENCABEZADO_ARCHIVO]"
            'Print #1, "CLIENTE_FUENTE=" & ArchivoMaestro.CLIENTE_FUENTE
                Print #1, "USUARIO_ORIGEN=" & Usuario.Nombre & " " & Usuario.Apellido
                Print #1, "OID_USUARIO_ORIGEN=" & Opciones.OidUserDefault
                Print #1, "FOLIO_SALIDA=" & Usuario.Salidas
                Print #1, "OID_MOVIMIENTO=" & Envio.OidEnv
                Print #1, "No_ARCHIVOS=" & LV_ArchivosElegidos.ListItems.Count
                Print #1, "QUICKSEND=0"
                
                If Len(Cb_CmdAsociado.Text) = 0 Then
                   Print #1, "CMD=NO"
                Else
                  ' MsgBox Envio.bd_Cmd
                   Print #1, "CMD=" & Envio.bd_Cmd
                End If
                
            'Print #1, "USUARIO_DESTINO=" & ArchivoMaestro.USUARIO_DESTINO
            'Print #1, "OID_USUARIO_DESTINO=" & ArchivoMaestro.OID_USUARIO_DESTINO
                Print #1, "FECHA_CREACION=" & Now
            'Print #1, "FECHA_ENVIO="
            'Print #1, "FECHA_RECEPCION="
            
                If Len(Trim(txtComent.Text)) <> 0 Then
                        
                   Print #1, "COMENTARIO=" & txtComent.Text & vbCrLf
                
                Else
                
                   Print #1, "COMENTARIO=S/C" & vbCrLf
                
                End If
    
    
    For C = 1 To LV_ArchivosElegidos.ListItems.Count
            
                          
            Print #1, "[INFORMACION_ARCHIVO_" & C & "]"
            Print #1, "NOMBRE_ARCHIVO_" & C & "=" & LV_ArchivosElegidos.ListItems.Item(C).Text
            Print #1, "TAMAÑO=" & FileLen(LV_ArchivosElegidos.ListItems.Item(C).SubItems(2))
            Print #1, "RUTA_ORIGEN=" & LV_ArchivosElegidos.ListItems.Item(C).SubItems(2)
                          
                          
            If Len(LV_ArchivosElegidos.ListItems.Item(C).SubItems(3)) = 0 Then
            Print #1, "EXTRAER=999" & vbCrLf
            Else
            Print #1, "EXTRAER=" & LV_ArchivosElegidos.ListItems.Item(C).SubItems(3) & vbCrLf
            End If
    
               
    Next C


Close (1)



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
               
               MsgBox "Modulo de envio no se ha podido iniciar, intentarlo de nuevo mas tarde", vbCritical, "Error de la aplicacion"
               
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


