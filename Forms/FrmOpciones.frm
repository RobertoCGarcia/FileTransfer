VERSION 5.00
Begin VB.Form FrmOpciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones del Sistema"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptHt 
      Caption         =   "Host Normal"
      Height          =   375
      Left            =   7200
      TabIndex        =   42
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton OptServ 
      Caption         =   "Servidor"
      Height          =   375
      Left            =   7200
      TabIndex        =   41
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   5400
      MaxLength       =   15
      TabIndex        =   38
      Top             =   5760
      Width           =   3855
   End
   Begin VB.CommandButton cmd_Reportes 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   8400
      Width           =   375
   End
   Begin VB.CommandButton cmdAdduser 
      Caption         =   "&Administrador Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   13
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmd_ListaGen 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton cmd_RutArch 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton cmd_RutExt 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdRut_Bd 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox txtPuertoEnv 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtPuertoRecep 
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   600
      Width           =   3135
   End
   Begin VB.ComboBox cb_User 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdRut_Ens 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdRut_Gen 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdRuta_Rec 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdRuta_Env 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmd_Salir 
      Caption         =   "&Salir"
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
      Left            =   7560
      TabIndex        =   15
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmd_Guardar 
      Caption         =   "&Guardar"
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
      Left            =   5520
      TabIndex        =   14
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Label Label16 
      Caption         =   $"FrmOpciones.frx":0000
      Height          =   1335
      Left            =   5400
      TabIndex        =   43
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Nodo:"
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
      Left            =   5400
      TabIndex        =   40
      Top             =   4440
      Width           =   1305
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre en la Red:"
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
      Left            =   5400
      TabIndex        =   39
      Top             =   5400
      Width           =   2250
   End
   Begin VB.Label lblReportes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   37
      Top             =   8760
      Width           =   660
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Reportes:"
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
      TabIndex        =   36
      Top             =   8400
      Width           =   1860
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   9360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Generados:"
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
      TabIndex        =   35
      Top             =   7440
      Width           =   2460
   End
   Begin VB.Label lblLista 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   34
      Top             =   7800
      Width           =   660
   End
   Begin VB.Label lblArch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   33
      Top             =   6840
      Width           =   660
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta ""Archivo"":"
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
      TabIndex        =   32
      Top             =   6480
      Width           =   2880
   End
   Begin VB.Label lblExt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   31
      Top             =   5880
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta ""Extraidos"":"
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
      TabIndex        =   30
      Top             =   5520
      Width           =   3120
   End
   Begin VB.Label lblBd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   29
      Top             =   4920
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta Base de Datos:"
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
      TabIndex        =   28
      Top             =   4560
      Width           =   3570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto a Conectarse (Remoto):"
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
      Left            =   5400
      TabIndex        =   27
      Top             =   1080
      Width           =   3765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto para Recepcion (Local):"
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
      Left            =   5400
      TabIndex        =   26
      Top             =   240
      Width           =   3750
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   5280
      Y1              =   0
      Y2              =   9360
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
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
      Left            =   5400
      TabIndex        =   25
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario Default:"
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
      Left            =   5400
      TabIndex        =   24
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Label lblEns 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   23
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label lblgen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   22
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta ""Ensamble"":"
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
      TabIndex        =   21
      Top             =   3600
      Width           =   3180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta ""Generados"":"
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
      TabIndex        =   20
      Top             =   2640
      Width           =   3330
   End
   Begin VB.Label lblRec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\bd"
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
      TabIndex        =   19
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label lblenv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c:\xxx"
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
      TabIndex        =   18
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta ""Recibidos"":"
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
      TabIndex        =   17
      Top             =   1680
      Width           =   3180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta Carpeta ""Enviados"":"
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
      TabIndex        =   16
      Top             =   840
      Width           =   3090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cb_User_Click()

        ' MsgBox "oID. USER : " & Array_OidUsuario(cb_User.ListIndex + 1, 1)

Opciones.OidUserDefault = Array_OidUsuario(cb_User.ListIndex + 1, 1)
lblUser.Caption = cb_User.Text

End Sub

Private Sub cmd_Guardar_Click()


If Len(lblEns.Caption) = 0 Then
    MsgBox "Elija la Ruta de Ensamble adecuadamente", vbCritical, "Dato Necesario"
    Exit Sub
End If



If Len(lblenv.Caption) = 0 Then
    MsgBox "Elija la Ruta de Enviados adecuadamente", vbCritical, "Dato Necesario"
    Exit Sub
End If

    
If Len(lblgen.Caption) = 0 Then
    MsgBox "Elija la Ruta de Generados adecuadamente", vbCritical, "Dato Necesario"
    Exit Sub
End If
    
    

If Len(lblRec.Caption) = 0 Then
    MsgBox "Elija la Ruta de Recibidos adecuadamente", vbCritical, "Dato Necesario"
    Exit Sub
End If
    
    

If Len(lblBd.Caption) = 0 Then
    MsgBox "Elija la Ruta de Base de datos", vbCritical, "Dato Necesario"
    Exit Sub
End If
        
    
  
If Len(txtPuertoRecep.Text) = 0 Then
    MsgBox "Escriba un numero de Puerto valido...", vbCritical, "Dato Necesario"
    Exit Sub
End If
    
    
If Len(txtPuertoEnv.Text) = 0 Then
    MsgBox "Escriba un numero de Puerto valido...", vbCritical, "Dato Necesario"
    Exit Sub
End If
    

    
If Len(lblUser.Caption) = 0 Then
    MsgBox "Elija un Usuario Default adecuadamente", vbCritical, "Dato Necesario"
    Exit Sub
End If
    
    
    
If Len(lblLista.Caption) = 0 Then
    MsgBox "Elija Ruta de la lista de Envio adecuadamente", vbCritical, "Dato Necesario"
    Exit Sub
End If
    
    
If Len(txtNombre.Text) = 0 Then
    MsgBox "Escribe un nombre adecuado del equipo", vbCritical, "Dato Necesario"
    txtNombre.Text = ""
    txtNombre.SetFocus
    Exit Sub
End If
    
    
    
'If OptServ.Value = 0 Then
'    MsgBox "Elija el tipo de Server adecuado", vbCritical, "Dato Necesario"
    'txtNombre.Text = ""
    'txtNombre.SetFocus'
'    Exit Sub
'End If
        
    
If Len(Opciones.TipoHost) = 0 Then
  MsgBox "Elija el tipo adecuado Server/Host", vbCritical, "Dato Necesario"
  Exit Sub
End If

    
'If OptHt.Value = 0 Then
'    MsgBox "Elija el tipo de host adecuado", vbCritical, "Dato Necesario"
    'txtNombre.Text = ""
    'txtNombre.SetFocus
'    Exit Sub
'End If
    
'       Print #1, Opciones.TipoHost ' tipo de host
'        Print #1, Opciones.NombreHost
    
    
    
    
    
    
'If Len(txtPuertoInf.Text) = 0 Then
'    MsgBox "Escriba un numero de Puerto valido...", vbCritical, "Dato Necesario"
'    txtPuertoInf.Text = ""
'    txtPuertoInf.SetFocus
    
    
'    Exit Sub
'End If
    
    
  





   Open App.path & "\dll\Opc.RLJ" For Output As #1
    
        Print #1, lblenv.Caption
        Print #1, lblRec.Caption
        Print #1, lblgen.Caption
        Print #1, lblEns.Caption
        Print #1, Opciones.OidUserDefault
        Print #1, txtPuertoRecep.Text
        Print #1, txtPuertoEnv.Text
        Print #1, lblBd.Caption
        Print #1, Opciones.Rutacarpeta_Extraidos
        Print #1, Opciones.Rutacarpeta_depositoRecibidos
        Print #1, Opciones.RutaListaGenerada
        Print #1, Opciones.RutaReportes
        Print #1, Opciones.TipoHost ' tipo de host
        Print #1, txtNombre.Text ' nombre del host en la red
        
       ' Print #1, txtPuertoInf.Text
        
       
   Close (1)


Call InformacionUsuario(Opciones.OidUserDefault, "CONSULTA")
' Actualiza datos usuarios CUANDO SE CAMBIA DE USUARIO DEFAULT

'FrmMain!lblUser = Usuario.Nombre & " " & Usuario.Apellido
FrmEnvio!SBar.Panels(2).Text = Usuario.Nombre & " " & Usuario.Apellido
FrmMain!SBar.Panels(1).Text = Usuario.Nombre & " " & Usuario.Apellido

'actualizar la barra de estado



'Debug.Print Usuario.Apellido

    

Call LeerInfoArch



FrmRecepcionFile.Winsock1.Close
FrmRecepcionFile.Winsock1.LocalPort = Val(Opciones.PuertoRecepcion)
FrmRecepcionFile.Winsock1.Listen
 
FrmMain.lblport = "Esperando: " & FrmRecepcionFile.Winsock1.LocalPort
FrmMain.lblTipo = ""
FrmMain.lblTipo = FrmMain.lblTipo & " Tipo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost

MsgBox "Cambios Guardados con Exito!!!", vbInformation, "Operacion Exitosa"



'FrmMain!Sck_Info.Close
'FrmMain!Sck_Info.LocalPort = Val(Opciones.PuertoInfo)
'FrmMain!Sck_Info.Listen


Unload Me

'Para actualzar la informacion guardada



'Unload Me






End Sub

Private Sub cmd_RutaBD_Click()




  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblBd.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de la Base de Datos de Movimientos"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     lblBd.Caption = Left(path, pos - 1) & "\srd.mdb"
     Opciones.RutaBd = lblBd.Caption
  End If

  Call CoTaskMemFree(pidl)




End Sub

Private Sub cmd_ListaGen_Click()



Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblLista.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de la Carpeta Extraidos"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblLista.Caption = Left(path, pos - 1) '& "\Registro.mdb"

     Opciones.RutaListaGenerada = lblLista.Caption
     
  End If

  Call CoTaskMemFree(pidl)













End Sub

Private Sub cmd_Reportes_Click()

Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblReportes.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta donde se generan los Reportes del Sistema"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblReportes.Caption = Left(path, pos - 1) '& "\Registro.mdb"

     Opciones.RutaReportes = lblReportes.Caption
     
  End If

  Call CoTaskMemFree(pidl)


End Sub

Private Sub cmd_RutArch_Click()




Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblArch.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de la Carpeta Extraidos"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblArch.Caption = Left(path, pos - 1) '& "\Registro.mdb"

     Opciones.Rutacarpeta_depositoRecibidos = lblArch.Caption
     
  End If

  Call CoTaskMemFree(pidl)

End Sub

Private Sub cmd_RutExt_Click()



Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblExt.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de la Carpeta Extraidos"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblExt.Caption = Left(path, pos - 1) '& "\Registro.mdb"

     Opciones.Rutacarpeta_Extraidos = lblExt.Caption
     
  End If

  Call CoTaskMemFree(pidl)










End Sub

Private Sub cmd_Salir_Click()

Unload Me


End Sub


Private Sub cmdAdduser_Click()

Derecho = 1
FrmPassword.Show

End Sub

Private Sub cmdRut_Bd_Click()

Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblBd.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de la base de Datos"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblBd.Caption = Left(path, pos - 1) & "\Registro.mdb"

     Opciones.RutaBd = lblBd.Caption
     
  End If

  Call CoTaskMemFree(pidl)



End Sub

Private Sub cmdRut_Ens_Click()
Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblEns.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta del Ensamble"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     lblEns.Caption = Left(path, pos - 1)
     Opciones.Rutacarpeta_Ensamble = lblEns.Caption
     
  End If

  Call CoTaskMemFree(pidl)




End Sub

Private Sub cmdRut_Gen_Click()


  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblgen.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de los Generados"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblgen.Caption = Left(path, pos - 1)
     
     Opciones.Rutacarpeta_Generados = lblgen.Caption
     
  End If

  Call CoTaskMemFree(pidl)






End Sub

Private Sub cmdRuta_Env_Click()


  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblenv.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de los Envios"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblenv.Caption = Left(path, pos - 1)
     Opciones.Rutacarpeta_Enviados = lblenv.Caption
     
  End If

  Call CoTaskMemFree(pidl)








End Sub

Private Sub cmdRuta_Rec_Click()


  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblRec.Caption = ""

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = Me.hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = "Selecciona la Ruta de Recibidos"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     
     lblRec.Caption = Left(path, pos - 1)
     
     Opciones.Rutacarpeta_Recibidos = lblRec.Caption
     
  End If

  Call CoTaskMemFree(pidl)


















End Sub


Private Sub Form_Load()

'Call GetUserName



If Len(Dir(App.path & "\dll\Opc.Rlj")) = 0 Then

        MsgBox "Archivo de Opc.Rlj no existe, copielo o creelo", vbCritical, "Archivo Necesario"
        Unload Me
        Exit Sub
Else



        Open App.path & "\dll\Opc.Rlj" For Input As #1
        
        
        Line Input #1, Opciones.Rutacarpeta_Enviados
        Line Input #1, Opciones.Rutacarpeta_Recibidos
        Line Input #1, Opciones.Rutacarpeta_Generados
        Line Input #1, Opciones.Rutacarpeta_Ensamble
        Line Input #1, Opciones.OidUserDefault
        Line Input #1, Opciones.PuertoRecepcion ' puerto local
        Line Input #1, Opciones.PuertoSalida ' puerto remoto
        Line Input #1, Opciones.RutaBd
        Line Input #1, Opciones.Rutacarpeta_Extraidos
        Line Input #1, Opciones.Rutacarpeta_depositoRecibidos
        Line Input #1, Opciones.RutaListaGenerada
        Line Input #1, Opciones.RutaReportes
        Line Input #1, Opciones.TipoHost ' tipo de host
        Line Input #1, Opciones.NombreHost ' nombre del host en la red
        
        
       '' Line Input #1, Opciones.PuertoInfo
        
        
        'env
        'rec
        'gen
        
        Close (1)
        
        
     'cb_User
     
        lblenv = Opciones.Rutacarpeta_Enviados
        lblRec = Opciones.Rutacarpeta_Recibidos
        lblgen = Opciones.Rutacarpeta_Generados
        lblEns = Opciones.Rutacarpeta_Ensamble
        lblBd = Opciones.RutaBd
        lblExt = Opciones.Rutacarpeta_Extraidos
        lblArch = Opciones.Rutacarpeta_depositoRecibidos
        lblLista = Opciones.RutaListaGenerada
        lblReportes = Opciones.RutaReportes
        
 
        lblUser = NombreUsuario(Opciones.OidUserDefault)
        txtPuertoRecep = Opciones.PuertoRecepcion
        txtPuertoEnv = Opciones.PuertoSalida
        
        
        If (Opciones.TipoHost) = "SERVER" Then
        'servidor
        OptServ.Value = True
        OptHt.Value = False
        'SERVER
        End If
        
        If (Opciones.TipoHost) = "HOST" Then
        ' host normal
        'HOST
        OptServ.Value = False
        OptHt.Value = True
        
        End If
        
        
        txtNombre.Text = Opciones.NombreHost
        
        
       ' Line Input #1, Opciones.TipoHost ' tipo de host
       ' Line Input #1, Opciones.NombreHost ' nombre del host en la red
        
        
      '  txtPuertoInf = Opciones.PuertoInfo
          
End If



End Sub









Private Sub OptHt_Click()
Opciones.TipoHost = "HOST"
End Sub

Private Sub OptServ_Click()
Opciones.TipoHost = "SERVER"
End Sub


Private Sub txtPuertoEnv_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
KeyAscii = 0
Beep
End If
End Sub



Private Sub txtPuertoInf_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtPuertoRecep_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
KeyAscii = 0
Beep
End If
End Sub
