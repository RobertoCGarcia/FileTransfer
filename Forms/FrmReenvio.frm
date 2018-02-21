VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmReenvio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reenvio de Archivos"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2760
      TabIndex        =   12
      Top             =   6840
      Width           =   3975
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "&Cerrar"
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
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Buscar 
      Caption         =   "&Buscar"
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
      Left            =   6840
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DG_Mov 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2355
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3201
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSComCtl2.DTPicker dt_Fi 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38701
   End
   Begin MSComCtl2.DTPicker dt_FF 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38701
   End
   Begin MSComctlLib.ListView Lv_Lista 
      Height          =   1125
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   1984
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Precio"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Salidas Pendientes de Envio"
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
      Top             =   2160
      Width           =   3180
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1110
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label lblreg 
      AutoSize        =   -1  'True
      Caption         =   "Encontrados (0)"
      Height          =   195
      Left            =   6960
      TabIndex        =   4
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Usuarios Destino"
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
      TabIndex        =   3
      Top             =   4080
      Width           =   2190
   End
End
Attribute VB_Name = "FrmReenvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstMov As ADODB.Recordset
Dim RstBody As ADODB.Recordset
Dim RstArch As ADODB.Recordset

Dim OrdenHeaderSql As String
Dim OrdenBodySql As String

Private Sub cmd_Buscar_Click()



If dt_Fi.Year = dt_FF.Year Then

   
   
   If dt_Fi.Month = dt_FF.Month Then
   
      
      If dt_Fi.Day = dt_FF.Day Then
     

      dt_FF.Value = dt_FF.Value + 1 ' para la busqueda de hoy

      End If
      
    
   


   End If
   


End If




    OrdenHeaderSql = "SELECT OIDENV,FOLIO_SALIDA,USUARIO_CREO,COMENTARIO,FECHA_ENVIO, FECHA_REGISTRO,NOMBRE_ARCHIVO_COMPRIMIDO,NO_ARCHIVOS,NO_ENVIOS From HEADERENVIO WHERE (((FECHA_REGISTRO) Between  #" & Format(dt_Fi.Value, "mm/dd/yyyy") & "# And  #" & Format(dt_FF.Value, "mm/dd/yyyy") & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=0)) ORDER BY FOLIO_SALIDA;"

    Set RstMov = New ADODB.Recordset
    RstMov.CursorLocation = adUseClient
    RstMov.CursorType = adOpenDynamic
    RstMov.LockType = adLockPessimistic
    RstMov.Open OrdenHeaderSql, CadenaCnx




    If RstMov.RecordCount > 0 Then
    ' hay registros sobre los cuales trabajar
    lblreg.Caption = "Encontrados: ( " & RstMov.RecordCount & " )"
   '   FrmReenvio.Height = 6915
    
    
    DG_Mov.Enabled = True
    Set DG_Mov.DataSource = RstMov
    
                DG_Mov.Columns(0).Visible = False
                DG_Mov.Columns(0).Locked = True
                
                DG_Mov.Columns(1).Width = "1000"
                DG_Mov.Columns(1).Locked = True
                DG_Mov.Columns(1).Caption = "Folio Salida"
                
                DG_Mov.Columns(2).Width = "2500"
                DG_Mov.Columns(2).Locked = True
                DG_Mov.Columns(2).Caption = "Usuario Creo"
                
                DG_Mov.Columns(3).Width = "2500"
                DG_Mov.Columns(3).Locked = True
                DG_Mov.Columns(3).Caption = "Comentario"
                
                DG_Mov.Columns(4).Width = "2500"
                DG_Mov.Columns(4).Locked = True
                DG_Mov.Columns(4).Caption = "Fecha Envio"
                
                DG_Mov.Columns(5).Width = "2500"
                DG_Mov.Columns(5).Locked = True
                DG_Mov.Columns(5).Caption = "Fecha Registro"
                
                
                DG_Mov.Columns(6).Width = "2500"
                DG_Mov.Columns(6).Locked = True
                DG_Mov.Columns(6).Caption = "Nombre Archivo"
                
                DG_Mov.Columns(7).Width = "2500"
                DG_Mov.Columns(7).Locked = True
                DG_Mov.Columns(7).Caption = "No. Archivos"
                
                
                
                
                DG_Mov.Columns(8).Width = "2500"
                DG_Mov.Columns(8).Locked = True
                DG_Mov.Columns(8).Caption = "No. Usuarios Destino"
    
    Else
    ' no hay nada sobre que trabajar
    lblreg.Caption = "Encontrados (0)"
  '  FrmReenvio.Height = 1770
    
    
    End If





End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub



Private Sub DG_Mov_DblClick()



RstMov.Bookmark = DG_Mov.Bookmark

MsgBox RstMov!OidEnv









End Sub



Private Sub Form_Load()


dt_Fi.Value = Date
dt_FF.Value = Date




End Sub



Sub FormatearGrids()

                DG_Mov.Columns(0).Width = "1000"
                DG_Mov.Columns(0).Locked = True
                DG_Mov.Columns(0).Caption = "Folio"
                
                DG_Mov.Columns(1).Width = "2500"
                DG_Mov.Columns(1).Locked = True
                DG_Mov.Columns(1).Caption = "Usuario Origen"
                
                DG_Mov.Columns.Add (2)
                DG_Mov.Columns(2).Width = "2500"
                DG_Mov.Columns(2).Locked = True
                DG_Mov.Columns(2).Caption = "Comentario"
                
                DG_Mov.Columns.Add (3)
                DG_Mov.Columns(3).Width = "2500"
                DG_Mov.Columns(3).Locked = True
                DG_Mov.Columns(3).Caption = "Fecha Recepcion"
                
                DG_Mov.Columns.Add (4)
                DG_Mov.Columns(4).Width = "2500"
                DG_Mov.Columns(4).Locked = True
                DG_Mov.Columns(4).Caption = "Fecha Registro"
                
                DG_Mov.Columns.Add (5)
                DG_Mov.Columns(5).Width = "2500"
                DG_Mov.Columns(5).Locked = True
                DG_Mov.Columns(5).Caption = "Nombre Archivo"
                
                DG_Mov.Columns.Add (6)
                DG_Mov.Columns(6).Width = "2500"
                DG_Mov.Columns(6).Locked = True
                DG_Mov.Columns(6).Caption = "No. Archivos"
End Sub






Sub BuscarDetalle(Oid As String)

'Dim OrdenBodySql As String
' salida (pendiente)

    OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "' AND ENVIO=9;"


    Set RstBody = New ADODB.Recordset
    RstBody.CursorLocation = adUseClient
    RstBody.CursorType = adOpenDynamic
    RstBody.LockType = adLockPessimistic
    RstBody.Open OrdenBodySql, CadenaCnx


End Sub


