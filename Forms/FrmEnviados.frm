VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmEnviados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Enviados"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_salir 
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
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DG_Lista 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Enviados"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "FrmEnviados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LclRst As ADODB.Recordset

Private Sub Cmd_Enterado_Click()

' Enterado, se entiende que el usuario se da por enterado de todo lo que se recibio
' y se somete a todo lo recibido





End Sub

Private Sub cmd_Salir_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim strSql As String
Dim Dia As String
Dim mes As String
Dim año As String
Dim Fecha As String

lblfecha.Caption = Format(Date, "dd/mm/YYYY")


'Call CALCULO_FECHAS(Format$(Date, "dd/mm/yyyy"))
'strSql = "SELECT * From HEADERENVIO WHERE (((FECHA_RECEPCION) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((RECIBIDOS.REVISADO)=False));"


Call CALCULO_FECHAS(Format$(Date, "dd/mm/yyyy"))
strSql = "SELECT * From HEADERENVIO WHERE (((HEADERENVIO.FECHA_ENVIO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((HEADERENVIO.REVISADO)=False));"


'Debug.Print strSql


Set LclRst = New ADODB.Recordset
    LclRst.CursorLocation = adUseClient
    LclRst.CursorType = adOpenDynamic
    LclRst.LockType = adLockPessimistic
    LclRst.Open strSql, CadenaCnx

   ' With LclRst


Set Me.DG_Lista.DataSource = LclRst

'DG_Lista.Columns(0).Visible = False
'DG_Lista.Columns(0).Locked = True

'DG_Lista.Columns(1).Visible = False
'DG_Lista.Columns(1).Locked = True

'DG_Lista.Columns(2).Visible = True
'DG_Lista.Columns(2).Width = "2500"
'DG_Lista.Columns(2).Locked = True

'DG_Lista.Columns(3).Visible = False
'DG_Lista.Columns(3).Locked = True

'DG_Lista.Columns(4).Visible = False
'DG_Lista.Columns(4).Locked = True

'DG_Lista.Columns(5).Visible = False
'DG_Lista.Columns(5).Locked = True

'DG_Lista.Columns(6).Visible = True
'DG_Lista.Columns(6).Locked = True
'DG_Lista.Columns(6).Width = "2500"


End Sub


Private Sub Form_Unload(Cancel As Integer)

LclRst.Close
Call CALCULO_FECHAS(Format$(Date, "dd/mm/yyyy"))
Call ConsultarRecepcion

End Sub
