VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBitacora 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Eventos"
   ClientHeight    =   3435
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV_Eventos 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Evento (Tipo)"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Titulo"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripcion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Usuario"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fecha"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSComCtl2.DTPicker dt_Fi 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   53477377
      UpDown          =   -1  'True
      CurrentDate     =   38701
   End
   Begin MSComCtl2.DTPicker dt_FF 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   53477377
      UpDown          =   -1  'True
      CurrentDate     =   38701
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir del Registro de Eventos"
      Top             =   2760
      Width           =   1092
   End
   Begin VB.CommandButton cmd_Info 
      Caption         =   "Informacion"
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
      Height          =   612
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Muestra Informacion mas detallada del evento"
      Top             =   2760
      Width           =   1092
   End
   Begin VB.CommandButton cmd_Bqd 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Busca los eventos ocurridos"
      Top             =   2760
      Width           =   1092
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
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
      TabIndex        =   6
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Width           =   990
   End
End
Attribute VB_Name = "FrmBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FI As String
Public FF As String

Private Sub cmd_Bqd_Click()

If dt_Fi.Year = dt_FF.Year Then

   
   
   If dt_Fi.Month = dt_FF.Month Then
   
      
      If dt_Fi.Day = dt_FF.Day Then
     

      dt_FF.Value = dt_FF.Value + 1 ' para la busqueda de hoy

      End If

   End If
   


End If

'Format$(Date, "dd/mm/yyyy")
FI = Format(dt_Fi.Value, "mm/dd/yyyy")
FF = Format(dt_FF.Value, "mm/dd/yyyy")


ConsultarEventos (1)


End Sub

Private Sub cmd_OK_Click()

Unload Me

End Sub

Private Sub Form_Load()
dt_Fi.Value = Date
dt_FF.Value = Date
ConsultarEventos (0)
End Sub

Private Sub LV_Eventos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

LV_Eventos.Sorted = True
LV_Eventos.SortOrder = lvwAscending
LV_Eventos.SortKey = ColumnHeader.Index - 1





End Sub
