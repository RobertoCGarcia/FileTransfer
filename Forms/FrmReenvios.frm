VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmReenvioOK 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reenvio de Archivos"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FM_Gral 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   9135
      Begin MSDataGridLib.DataGrid DG_Mov 
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2566
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
      Begin MSDataGridLib.DataGrid Dg_Lista 
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
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
      Begin MSDataGridLib.DataGrid DG_Arch 
         Height          =   2175
         Left            =   5640
         TabIndex        =   11
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3836
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
      Begin VB.Label lblMov 
         AutoSize        =   -1  'True
         Caption         =   "Ordenes de Envio Pendientes:"
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
         TabIndex        =   14
         Top             =   240
         Width           =   3165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Usuarios Destino"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Archivos a Enviar"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   2280
         Width           =   1245
      End
   End
   Begin MSComCtl2.DTPicker dt_Fi 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38719
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "&OK"
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
      Left            =   7440
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Busqueda 
      Caption         =   "Buscar"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dt_FF 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38719
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      Caption         =   "(0)"
      Height          =   195
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados:"
      Height          =   195
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmReenvioOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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





Dim RstMov As ADODB.Recordset
Dim RstBody As ADODB.Recordset
Dim RstArch As ADODB.Recordset
Dim OrdenHeaderSql As String
Dim OrdenBodySql As String




Private Sub cb_mov_Click()

Set DG_Mov.DataSource = Nothing


lblReg = ""

lblMov.Caption = "Movimiento"
FormatearGrids

DG_Mov.Enabled = False
'DG_Users.Enabled = False
'DG_Arch.Enabled = False

End Sub

Private Sub cmd_Busqueda_Click()


Dim FI As String
Dim FF As String



'If Len(cb_mov.Text) = 0 Then
'MsgBox "Seleccione un tipo de movimiento", vbInformation, "Criterio Necesario"
'cb_mov.SetFocus
'Exit Sub
'End If



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

'MsgBox dt_Fi.Value & " " & dt_FF.Value


'If dt_Fi.Value = dt_Fi.Value Then
'MsgBox "son iguales"
'End If


    OrdenHeaderSql = "SELECT OIDENV,FOLIO_SALIDA,USUARIO_CREO,COMENTARIO,FECHA_ENVIO, FECHA_REGISTRO,NOMBRE_ARCHIVO_COMPRIMIDO,NO_ARCHIVOS,NO_ENVIOS From HEADERENVIO WHERE (((FECHA_REGISTRO) Between  #" & FI & "# And  #" & FF & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=0)) ORDER BY FOLIO_SALIDA;"

    Set RstMov = New ADODB.Recordset
    RstMov.CursorLocation = adUseClient
    RstMov.CursorType = adOpenDynamic
    RstMov.LockType = adLockPessimistic
    RstMov.Open OrdenHeaderSql, CadenaCnx




    If RstMov.RecordCount > 0 Then
    ' hay registros sobre los cuales trabajar
    lblReg.Caption = "( " & RstMov.RecordCount & " )"
    'lblMov.Caption = "Movimiento: " & cb_mov.Text
    
    
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
                
                
                 FM_Gral.Enabled = True
    
    Else
    ' no hay nada sobre que trabajar
    lblReg.Caption = "No se encontro nada con ese criterio"
    FM_Gral.Enabled = False
    
    
    
    
    End If



' Debug.Print OrdenHeaderSql



End Sub

Private Sub cmd_OK_Click()



Unload Me


End Sub





Private Sub Dg_Lista_DblClick()
' se va a reenviar el archivo elegido
    
    
   RstBody.Bookmark = Dg_Lista.Bookmark
 
    
    
    Open App.path & "\dll\EnvioOid.Oid" For Output As #2
                Print #2, "[Oid_Envio]"
                Print #2, "OidReciente=" & RstBody!OidEnv
    Close (2)


Call EnvioControlado


End Sub









Private Sub DG_Mov_DblClick()


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






RstMov.Bookmark = DG_Mov.Bookmark

'MsgBox RstMov.Fields(0).Value



    Open App.path & "\dll\EnvioOid.Oid" For Output As #2
                Print #2, "[Oid_Envio]"
                Print #2, "OidReciente=" & RstMov.Fields(0).Value
    
    Close (2)
    
    
    Call RegistraEvento("REENVIO", 1, "Reenvio Generado", "Reenvio con folio de salida: " & RstMov!FOLIO_SALIDA & " Hecho por el usuario:  " & RstMov!USUARIO_CREO)
    
           Unload Me
    
            Call EnvioControlado
    

'Call BuscarDetalle(RstMov.Fields(0).Value)
'Call BuscarArchivos(RstMov.Fields(0).Value)

'Call GenerarReporte


'WB_Info.Navigate (Opciones.RutaReportes & "\Reporte.html")

End Sub





Private Sub Form_Load()





If Len(Dir(App.path & "\dll\EnvioOid.Oid")) > 0 Then
'existe y lo borro
Kill App.path & "\dll\EnvioOid.Oid"
End If



dt_Fi.Value = Date
dt_FF.Value = Date


'cb_mov.AddItem "Entrada"
'cb_mov.AddItem "Salida (Enviado)"
'cb_mov.AddItem "Salida (Pendiente)"
FormatearGrids
'txtFechaIni.Text = Format("DD-MM-AA", txtFechaIni.Text)
'WB_Info.Navigate (App.path & "\dll\Inicio.html")



End Sub

Sub FormatearGrids()

'==Entrada
'Folio
'Usuario Origen
'Fecha Recepcion
'Comentario
'Nombre Archivo
'No. Archivos

'== salida
'Folio
'Usuario Origen
'Fecha envio
'Comentario
'Nombre Archivo
'No. Archivos


'DG_Mov
'DG_Users
'DG_Arch

'Movimientos
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



    OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "' AND ENVIO=9;"


    Set RstBody = New ADODB.Recordset
    RstBody.CursorLocation = adUseClient
    RstBody.CursorType = adOpenDynamic
    RstBody.LockType = adLockPessimistic
    RstBody.Open OrdenBodySql, CadenaCnx
    Set Dg_Lista.DataSource = RstBody
    
    
         
    
'0 OIDENV
'1 FOLIO_SALIDA
'2 IP
'3 RESULTADOPING
'4 Comentario
'5 Envio
'6 PUERTO
'7 UID_DESTINO
'8        USUARIO_DESTINO
'9 FECHA_ENVIO
'10        FECHA_REGISTRO
'11 USUARIO_CREO
'12 INICIO_TRANSMICION
'13 FIN_TRANSMICION
         
         
    
    Dg_Lista.Columns(0).Visible = False
    Dg_Lista.Columns(1).Visible = False
    Dg_Lista.Columns(2).Visible = False
    Dg_Lista.Columns(3).Visible = False
    Dg_Lista.Columns(4).Visible = False
    Dg_Lista.Columns(5).Visible = False
    Dg_Lista.Columns(6).Visible = False
    Dg_Lista.Columns(7).Visible = False
        Dg_Lista.Columns(8).Caption = "Usuario Destino"
        Dg_Lista.Columns(8).Visible = True
        Dg_Lista.Columns(8).Locked = True
        Dg_Lista.Columns(8).Width = 2300
    Dg_Lista.Columns(9).Visible = False
        Dg_Lista.Columns(10).Caption = "Fecha Registro"
        Dg_Lista.Columns(10).Visible = True
        Dg_Lista.Columns(10).Locked = True
        Dg_Lista.Columns(10).Width = 1500
    Dg_Lista.Columns(11).Visible = False
    Dg_Lista.Columns(12).Visible = False
    Dg_Lista.Columns(13).Visible = False
           
    
    ' Dg_Lista.Columns(0).Locked = True
   ' Dg_Lista.Columns(0).Caption = "No. Archivos"
    
'    Set DG_Users.DataSource = RstBody


End Sub



Sub BuscarArchivos(Oid As String)

Dim OrdenBodySql As String
'Dim RstArch As ADODB.Recordset



    
    
    'Salida
            OrdenBodySql = "SELECT * from enviados WHERE OIDENV =  '" & Oid & "';"
    
                    Set RstArch = New ADODB.Recordset
                    RstArch.CursorLocation = adUseClient
                    RstArch.CursorType = adOpenDynamic
                    RstArch.LockType = adLockPessimistic
                    RstArch.Open OrdenBodySql, CadenaCnx
                
                    Set DG_Arch.DataSource = RstArch
                
                
                
'0 OidEnv
'1 FOLIO_SALIDA
'2 NOMBRE_ARCHIVO
'3 Tamaño
'4 RUTA_DESTINO
'5 RUTA_ORIGEN
'6 FECHA_CREACION_ARCHIVO
'7 FECHA_REGISTRO
'8 USUARIO_CREO
                
    


    DG_Arch.Columns(0).Visible = False
    DG_Arch.Columns(1).Visible = False
    'DG_Arch.Columns(2).Visible = False
    'DG_Arch.Columns(3).Visible = False
    DG_Arch.Columns(4).Visible = False
    'Dg_Lista.Columns(5).Visible = False
    DG_Arch.Columns(6).Visible = False
    DG_Arch.Columns(7).Visible = False
    DG_Arch.Columns(8).Visible = False
        
        DG_Arch.Columns(2).Caption = "Archivo"
        DG_Arch.Columns(2).Visible = True
        DG_Arch.Columns(2).Locked = True
        DG_Arch.Columns(2).Width = 1500
   
        DG_Arch.Columns(3).Caption = "Tamaño Bytes"
        DG_Arch.Columns(3).Visible = True
        DG_Arch.Columns(3).Locked = True
        DG_Arch.Columns(3).Width = 1000
    

        DG_Arch.Columns(5).Caption = "Ubicacion Local"
        DG_Arch.Columns(5).Visible = True
        DG_Arch.Columns(5).Locked = True
        DG_Arch.Columns(5).Width = 1500
    

End Sub





Private Sub Form_Unload(Cancel As Integer)

Set RstMov = Nothing
Set RstBody = Nothing
Set RstArch = Nothing


End Sub








Public Sub EnvioControlado()


    Dim sCmdLine As String
    Dim idProg As Long, iExit As Long
    
    
  'If InxPosEnv <= er Then
    
           sCmdLine = App.path & "\envios.exe"
           
           idProg = Shell(sCmdLine, vbNormalFocus)
           
           iExit = fWait(idProg)
        
           
            Call ConsultarRecepcion
            Call ConsultarEnvio
            Call ConsultarPendientes
          
           If iExit Then
               
               MsgBox "Modulo de envio no se ha podido iniciar, intentarlo de nuevo mas tarde", vbCritical, "Error de la aplicacion"
               
                 
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


