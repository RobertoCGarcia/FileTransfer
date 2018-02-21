VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrminfoRecibida1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informacion Recibida"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtComentario 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   7935
   End
   Begin MSDataGridLib.DataGrid DG_ListaArchivos 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Usuario Origen:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Label lblComent 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4680
      TabIndex        =   13
      Top             =   5640
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Comentario Recepcion:"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   5280
      Width           =   1995
   End
   Begin VB.Label lblfin 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   6120
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Hora Fin Envio:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lblIni 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   5640
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hora Inicio Envio:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1545
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   6600
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "IP Remota:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Informacion de la Recepcion"
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
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2985
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Archivos Recibidos:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comentario Enviado:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "FrminfoRecibida1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()

Unload Me


'borro el archivo del oid si existe para evitar cualquier cosa

If Len(Dir(App.path & "\dll\RecepcionOid.oid")) <> 0 Then
    'PORQUE EXISTE SI NO NO SE BORRA NADA.
    Kill (App.path & "\dll\RecepcionOid.oid")
End If

End Sub

Private Sub Form_Load()
'se carga con la info del comentario que se envia y la lista de archivos que se descomprimieraon en las
'rutas que se definen

Dim resp As Variant
Dim strCadena As String
Dim OidExtraido As String



If Len(Dir(App.path & "\dll\RecepcionOid.oid")) <> 0 Then
'EL ARCHIVO EXISTE
                
                
                
                
                strCadena = String(255, 0)
                resp = GetPrivateProfileString("Oid_Recepcion", "OidReciente", "Default", strCadena, 255, App.path & "\dll\RecepcionOid.oid")
                If resp <> 0 Then strCadena = Left$(strCadena, resp)
                OidExtraido = strCadena
                strCadena = ""
                
                
                'MsgBox OidExtraido
                
                'Generar 2 rst el de los archivos que llegaron
                'y el del comentario guardado en el header
                
                Dim RstHeader As ADODB.Recordset
                Dim RstBody As ADODB.Recordset
                Dim RstRecibidos As ADODB.Recordset
                
                
                Dim strSql As String
                
                
                strSql = "SELECT * From HEADER_RECEPCION WHERE HEADER_RECEPCION.OIDRECEP = '" & OidExtraido & "';"
                
                Set RstHeader = New ADODB.Recordset
                    RstHeader.CursorLocation = adUseClient
                    RstHeader.CursorType = adOpenDynamic
                    RstHeader.LockType = adLockPessimistic
                    RstHeader.Open strSql, CadenaCnx
                
                With RstHeader
                ' DEL HEADER SOLO EL COMENTARO QUE SE LE ENVIO
                txtComentario.Text = !Comentario
                .Close
                End With
                
                
                
                strSql = "SELECT RECIBIDOS.OIDRECEP, RECIBIDOS.NOMBRE_ARCHIVO, RECIBIDOS.TAMAÑO, RECIBIDOS.RUTA_DESTINO from RECIBIDOS WHERE RECIBIDOS.OIDRECEP = '" & OidExtraido & "';"
                
                Set RstRecibidos = New ADODB.Recordset
                    RstRecibidos.CursorLocation = adUseClient
                    RstRecibidos.CursorType = adOpenDynamic
                    RstRecibidos.LockType = adLockPessimistic
                    RstRecibidos.Open strSql, CadenaCnx
                
                With RstRecibidos
                'de recibidos solo los archivos recibidos dentro del empaquetado
                Set Me.DG_ListaArchivos.DataSource = RstRecibidos
                
                
                DG_ListaArchivos.Columns(0).Visible = False
                DG_ListaArchivos.Columns(0).Locked = True
                
                
                DG_ListaArchivos.Columns(1).Width = "2500"
                DG_ListaArchivos.Columns(1).Locked = True
                DG_ListaArchivos.Columns(1).Caption = "Nombre Archivo"
                
                
                DG_ListaArchivos.Columns(2).Width = "2500"
                DG_ListaArchivos.Columns(2).Locked = True
                DG_ListaArchivos.Columns(2).Caption = "Tamaño bytes"
                
                
                DG_ListaArchivos.Columns(3).Width = "2500"
                DG_ListaArchivos.Columns(3).Locked = True
                DG_ListaArchivos.Columns(3).Caption = "Ubicado en: "
                
                
                
                End With
                
                
                
                
                
                strSql = "SELECT * From BODYRECEPCION WHERE BODYRECEPCION.OIDRECEP = '" & OidExtraido & "';"
                
                Set RstBody = New ADODB.Recordset
                    RstBody.CursorLocation = adUseClient
                    RstBody.CursorType = adOpenDynamic
                    RstBody.LockType = adLockPessimistic
                    RstBody.Open strSql, CadenaCnx
                
                With RstBody
                'de recibidos solo los archivos recibidos dentro del empaquetado
                
                lblIP = !IP_REMOTA
                lblIni = !PETICION_INICIO
                lblfin = !PETICION_FINAL
                lblComent = !COMENTARIO_RECEPCION
                lblNombre = !USUARIO_ORIGEN
                
                .Close
                End With



                Set RstHeader = Nothing
                Set RstBody = Nothing
                Set RstRecibidos = Nothing
Else

Unload Me


End If






End Sub


Private Sub Form_Unload(Cancel As Integer)


If Len(Dir(App.path & "\dll\RecepcionOid.oid")) <> 0 Then
    'PORQUE EXISTE SI NO NO SE BORRA NADA.
    Kill (App.path & "\dll\RecepcionOid.oid")
End If



End Sub
