VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer IP 2 IP by ((( Nordwind Systemz )))"
   ClientHeight    =   6810
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   10305
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10560
      TabIndex        =   25
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   10035
      TabIndex        =   2
      Top             =   840
      Width           =   10095
      Begin VB.Timer TimStatus 
         Interval        =   1000
         Left            =   2520
         Top             =   3960
      End
      Begin MSComctlLib.ListView LV_Recibido 
         Height          =   975
         Left            =   3120
         TabIndex        =   24
         Top             =   1440
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1720
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
            Text            =   "Usuario Origen"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Recepcion"
            Object.Width           =   2752
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Folio Entrada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Folio Salida"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmd_borrar 
         Caption         =   "Actualizar"
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
         Index           =   2
         Left            =   8640
         TabIndex        =   22
         ToolTipText     =   "Borra los Nombres de los Usuarios que estan pendientes de envio y vuelve a cargar los Nombres Correctos"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmd_borrar 
         Caption         =   "Actualizar"
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
         Index           =   1
         Left            =   8640
         TabIndex        =   21
         ToolTipText     =   "Borra los Nombres de los Usuarios a los que se ha enviado informacion y vuelve a cargar los Nombres Correctos"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmd_borrar 
         Caption         =   "Actualizar"
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
         Index           =   0
         Left            =   8640
         TabIndex        =   20
         ToolTipText     =   "Borra los Nombres de los Usuarios que han enviado y vuelve a cargar los Nombres Correctos"
         Top             =   1200
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV_Enviados 
         Height          =   975
         Left            =   3120
         TabIndex        =   26
         Top             =   2880
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1720
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario Destino"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Envio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Folio Salida"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV_Pendientes 
         Height          =   975
         Left            =   3120
         TabIndex        =   27
         Top             =   4320
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1720
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario Destino"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Creacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Folio Salida"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Shape statusInet 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   1560
         Shape           =   3  'Circle
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informacion:"
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
         TabIndex        =   23
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de Envio:"
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
         Index           =   6
         Left            =   3120
         TabIndex        =   15
         Top             =   4080
         Width           =   2220
      End
      Begin VB.Label lblPen 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "XXXXX"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   4680
         Width           =   555
      End
      Begin VB.Image Img_Pen 
         Height          =   720
         Left            =   240
         MouseIcon       =   "FrmMain.frx":0802
         MousePointer    =   99  'Custom
         Top             =   4440
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivos Pendientes"
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
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   4080
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviados a:"
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
         Index           =   4
         Left            =   3120
         TabIndex        =   12
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Recibidos de:"
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
         Index           =   0
         Left            =   3120
         TabIndex        =   11
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status General"
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
         Left            =   7320
         TabIndex        =   10
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label lblport 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   8070
         TabIndex        =   9
         Top             =   480
         Width           =   75
      End
      Begin VB.Label lblfecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dddddddd"
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
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hoy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivos Enviados"
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
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1950
      End
      Begin VB.Label lblEnv 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "XXXXX"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   3360
         Width           =   555
      End
      Begin VB.Image img_Env 
         Height          =   720
         Left            =   240
         MouseIcon       =   "FrmMain.frx":0954
         MousePointer    =   99  'Custom
         Top             =   3120
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivos Recibidos"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2040
      End
      Begin VB.Label lblBRecep 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "XXXXX"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   1800
         Width           =   555
      End
      Begin VB.Image img_BR 
         Height          =   720
         Left            =   240
         MouseIcon       =   "FrmMain.frx":0AA6
         MousePointer    =   99  'Custom
         Top             =   1560
         Width           =   600
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   9975
      End
   End
   Begin MSComctlLib.ImageList ImgLMain 
      Left            =   360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0BF8
            Key             =   "envio"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0F12
            Key             =   "historial"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":122C
            Key             =   "recepcion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1546
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1860
            Key             =   "opciones"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":273A
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2B8C
            Key             =   "reenvio"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2FDE
            Key             =   "ModoBarra"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3430
            Key             =   "Ocultar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3882
            Key             =   "quicksend"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3CD4
            Key             =   "bitacora"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tb_Main 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1535
      ButtonWidth     =   1640
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImgLMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Envio"
            Key             =   "envio"
            Object.ToolTipText     =   "Genera la informacion a enviar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Recepcion"
            Object.ToolTipText     =   "Muestra informacion acerca del proceso de recepcion de informacion"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Opciones"
            Object.ToolTipText     =   "Opciones del Programa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Historial"
            Object.ToolTipText     =   "Historial de todo lo que se ha recibido y se ha enviado"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Object.ToolTipText     =   "Reinicia todo en el programa"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reenvio"
            Key             =   "Reenvio de las Salidas Pendientes"
            Object.ToolTipText     =   "Da la Posibilidad de Hacer envios que Quedaron Pendientes"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ocultar"
            Key             =   "Oculta el Programa"
            Object.ToolTipText     =   "Oculta la Vista Actual"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modo Barra"
            Key             =   "Hace pequeño el programa, para ocupar menor espacio en la pantalla"
            Object.ToolTipText     =   "Modo rapido accede a las funciones principales"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "QuickSend"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bitácora"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Sale del programa"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "FrmMain.frx":4126
   End
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6510
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/04/2007"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "02:40 a.m."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "No Conectado"
            TextSave        =   "No Conectado"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblent 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   195
      Left            =   1560
      TabIndex        =   19
      Top             =   7920
      Width           =   270
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
      Left            =   0
      TabIndex        =   18
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Label lbsal 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   7560
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Folio Salida:"
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
      Left            =   0
      TabIndex        =   16
      Top             =   7560
      Width           =   1065
   End
   Begin VB.Menu MenuEmengente 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Envio 
         Caption         =   "Envio"
      End
      Begin VB.Menu Recepcion 
         Caption         =   "Recepcion"
      End
      Begin VB.Menu OpcionesMen 
         Caption         =   "Opciones"
      End
      Begin VB.Menu Status 
         Caption         =   "Status"
      End
      Begin VB.Menu Ocultar 
         Caption         =   "Ocultar"
         Visible         =   0   'False
      End
      Begin VB.Menu Reset 
         Caption         =   "Reset"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmd_borrar_Click(Index As Integer)

Select Case Index
    Case Is = 0
            Call ConsultarRecepcion
    Case Is = 1
            Call ConsultarEnvio
    Case Is = 2
            Call ConsultarPendientes
End Select


End Sub



Private Sub Command1_Click()

Dim elementoX As ListItem

Set elementoX = FrmMain.LV_Recibido.ListItems.Add(, , "dfdf")
                            elementoX.Tag = "er"
                            elementoX.SubItems(1) = "fecha movimiento"
                            elementoX.SubItems(2) = "!FOLIO_RECEPCION" ' folio recepcion
                            elementoX.SubItems(3) = "!FOLIO_origen" ' folio origen
                           


End Sub

Private Sub Envio_Click()
FrmEnvio.Show
End Sub

Private Sub Form_Load()
FrmMain.lblfecha = Format$(Date, "dddd, dd MMMM  YYYY")
Call CALCULO_FECHAS(Format$(Date, "dd/mm/yyyy"))
FrmMain.lblTipo = ""
FrmMain.lblTipo = FrmMain.lblTipo & " Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost
SBar.Panels(1).Text = Usuario.Nombre & " " & Usuario.Apellido
ArchivoMaestro.USUARIO_ORIGEN = Usuario.Nombre & " " & Usuario.Apellido
Call ConsultarRecepcion
Call ConsultarEnvio
Call ConsultarPendientes
End Sub
'""""""""""Eventos de la imagen en el task bar




Private Sub Form_Unload(Cancel As Integer)
    
    Call RegistraEvento("PROGRAMA", 1, "Fin Programa", "Salida a peticion del Usuario: " & Usuario.Nick)
    
    
    Dim i As Integer


    For i = Forms.Count - 1 To 0 Step -1
   
        Unload Forms(i)
    
    Next

'  DoEvents
   
Dim OSF As Object 'Objeto para manipular los archivos
Set OSF = CreateObject("scripting.filesystemobject")

Dim Carpeta As Folder
Set Carpeta = OSF.GetFolder(Opciones.Rutacarpeta_Recibidos)

' Carpeta.Files.Count

    If Carpeta.Files.Count > 0 Then
    ' se borra todo lo que hay hay
         Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"
    
    End If
    
    End
    

    
End Sub




Private Sub img_BR_Click()
'ase Is = 3 ' se invoca desde Vista: Hoy Recibidos

If img_BR.Tag = "1" Then
    
    OpcHistorial = 3
    FrmHistorial.OpcBqd = 0
    'Set FrmHistorial.DG_Mov.DataSource = Nothing
    FrmHistorial.cb_mov.Enabled = False
    FrmHistorial.Show
    
Else

MsgBox "No hay nada recibido!!", vbInformation, "Informacion"

    
End If


End Sub

Private Sub img_Env_Click()
'ase Is = 2 ' se invoca desde Vista: Hoy Enviados
If img_Env.Tag = "1" Then

     OpcHistorial = 2
     FrmHistorial.OpcBqd = 1
     'Set FrmHistorial.DG_Mov.DataSource = Nothing
     FrmHistorial.cb_mov.Enabled = False
     FrmHistorial.Show
Else
    MsgBox "No hay nada por Enviar!!", vbInformation, "Informacion"
End If



End Sub

Private Sub Img_Pen_Click()


If Img_Pen.Tag = "1" Then
    'Case Is = 1 ' se invoca desde Vista: Hoy Pendientes
     OpcHistorial = 1
     FrmHistorial.OpcBqd = 2
     'Set FrmHistorial.DG_Mov.DataSource = Nothing
     FrmHistorial.cb_mov.Enabled = False
     FrmHistorial.Show

Else

MsgBox "No hay nada pendiente!!", vbInformation, "Informacion"

    
End If

End Sub









Private Sub Tb_Main_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Index


    Case Is = 1
    ' envio
    FrmEnvio.Show
    
    Case Is = 2
    'recepcion
    FrmRecepcionFile.Show
    Case Is = 3
    ' opciones
    Derecho = 0
    FrmPassword.Show
    
    
    
      
    
    Case Is = 4
     OpcHistorial = 0
     FrmHistorial.Show

    
    Case Is = 5
   
     Call Resetall
     
    Case Is = 6
       'Opcion de Reenvio
       
       
'       FrmReenvio.Show
        FrmReenvioOK.Show

    Case Is = 7
    'ocultar
    FrmMain.Hide
    
    Case Is = 8
    'minimizar modo barra
    
    Call ModoBarra
    

    Case Is = 9
    'favoritos del sistema, envio rapido
    FrmQuickSend.Show
    
    
    
    Case Is = 10
    
    FrmBitacora.Show
    
    
    
    Case Is = 11
    
    Call SalirAll



End Select


End Sub


Public Sub Resetall()

    ' reset
        FrmMain.lblport = "Reset..."
        FrmRecepcionFile.Winsock1.Close
        sendsize = 1024
        FrmRecepcionFile.Winsock1.Close
        FrmRecepcionFile.Winsock1.LocalPort = Opciones.PuertoRecepcion
        FrmRecepcionFile.Winsock1.Listen
        FrmConfiguracion.lblIPactual = FrmRecepcionFile.Winsock1.LocalIP
        
        FrmMain.lblport = "Esperando: " & FrmRecepcionFile.Winsock1.LocalPort
        
        Dim OSF As Object 'Objeto para manipular los archivos
        Set OSF = CreateObject("scripting.filesystemobject")

        Dim Carpeta As Folder
        Set Carpeta = OSF.GetFolder(Opciones.Rutacarpeta_Recibidos)
        
        ' Carpeta.Files.Count
        
            If Carpeta.Files.Count > 0 Then
            ' se borra todo lo que hay hay
                 Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"
            
            End If

    
    
    
        
        Call InformacionUsuario(Opciones.OidUserDefault, "CONSULTA")
        
        Call LeerInfoArch
        
        
        Call CALCULO_FECHAS(Format$(Date, "dd/mm/yyyy"))

        Call ConsultarRecepcion
        Call ConsultarEnvio
        Call ConsultarPendientes
        


End Sub


Public Sub ModoBarra()

    FrmMain.Hide
    Min = True
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost
    FrmMin.Show

End Sub


Public Sub SalirAll()
        ' Salir
        'Unload FrmSystemTray
       ' Call FrmSystemTray.sysTray.DeleteIcon
        Unload FrmMain
       'FrmInfoEnvio.Sck_Envio.Close
        'FrmRecepcionFile.Winsock1.Close
        ' Call sysTray.DeleteIcon
        ' Set sysTray = Nothing
        
        
        
        
        'FrmRecepcionFile.Winsock1.Close
        ' Call sysTray.DeleteIcon
        ' Set sysTray = Nothing
        'Unload FrmSystemTray
        'Unload FrmMin
        'Unload Me
        'End
        
End Sub




Public Function IsNetConnectOnline() As Boolean
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
   ' MsgBox IsNetConnectOnline
End Function

Private Sub TimStatus_Timer()

If IsNetConnectOnline Then
statusInet.BackColor = &HC000& 'verdadero online
SBar.Panels(4).Text = "Conectado a la Red"
Else
statusInet.BackColor = &HFF& ' falso no online
SBar.Panels(4).Text = "Desconectado de la Red"
End If

End Sub
