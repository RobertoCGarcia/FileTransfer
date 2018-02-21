VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmHistorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historial de Movimientos"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dt_Fi 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52625409
      CurrentDate     =   38701
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      Top             =   7680
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DG_Mov 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2566
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
      Left            =   7560
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox cb_mov 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   9135
      Begin SHDocVwCtl.WebBrowser WB_Info 
         Height          =   3855
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   8895
         ExtentX         =   15690
         ExtentY         =   6800
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label lblMov 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   810
      End
   End
   Begin MSComCtl2.DTPicker dt_FF 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52625409
      CurrentDate     =   38701
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      Caption         =   "(0)"
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Movimiento"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "FrmHistorial"
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
Public OpcBqd As String
Public TitleMov As String




Private Sub cb_mov_Click()

Set DG_Mov.DataSource = Nothing

'

  '1  cb_mov.AddItem "Salida (Enviado)"
  '2  cb_mov.AddItem "Salida (Pendiente)"
  '  OpcBqd
OpcBqd = cb_mov.ListIndex



Select Case OpcBqd

Case Is = 0 'cb_mov.AddItem "Entrada"
 TitleMov = "Entrada"
Case Is = 1
 TitleMov = "Salida (Enviado)"
Case Is = 2
 TitleMov = "Salida (Pendiente)"
End Select




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



If Len(cb_mov.Text) = 0 Then
MsgBox "Seleccione un tipo de movimiento", vbInformation, "Criterio Necesario"
cb_mov.SetFocus
Exit Sub
End If


Call Bqd




End Sub

Private Sub cmd_OK_Click()

Unload Me


End Sub


Private Sub DG_Mov_DblClick()






Select Case OpcBqd


Case Is = 0

                DG_Mov.Columns(0).Visible = False
                DG_Mov.Columns(0).Locked = True
                
                DG_Mov.Columns(1).Width = "1000"
                DG_Mov.Columns(1).Locked = True
                DG_Mov.Columns(1).Caption = "Folio Recepcion"
                
                DG_Mov.Columns(2).Width = "2500"
                DG_Mov.Columns(2).Locked = True
                DG_Mov.Columns(2).Caption = "Usuario Origen"
                
                DG_Mov.Columns(3).Width = "2500"
                DG_Mov.Columns(3).Locked = True
                DG_Mov.Columns(3).Caption = "Comentario"
                
                DG_Mov.Columns(4).Width = "2500"
                DG_Mov.Columns(4).Locked = True
                DG_Mov.Columns(4).Caption = "Fecha Recepcion"
                
                
                DG_Mov.Columns(5).Width = "2500"
                DG_Mov.Columns(5).Locked = True
                DG_Mov.Columns(5).Caption = "Fecha Registro"
                
                DG_Mov.Columns(6).Width = "2500"
                DG_Mov.Columns(6).Locked = True
                DG_Mov.Columns(6).Caption = "Nombre Archivo"
                
                DG_Mov.Columns(7).Width = "2500"
                DG_Mov.Columns(7).Locked = True
                DG_Mov.Columns(7).Caption = "No. Archivos"
             


Case Is = 1

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



Case Is = 2

  
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



End Select




RstMov.Bookmark = DG_Mov.Bookmark

'MsgBox RstMov.Fields(0).Value


Call BuscarDetalle(RstMov.Fields(0).Value)
Call BuscarArchivos(RstMov.Fields(0).Value)

Call GenerarReporte


WB_Info.Navigate (Opciones.RutaReportes & "\Reporte.html")

End Sub





Private Sub Form_Load()

dt_Fi.Value = Date
dt_FF.Value = Date

Select Case OpcBqd
    Case Is = 0 'cb_mov.AddItem "Entrada"
      TitleMov = "Entrada"
    Case Is = 1
      TitleMov = "Salida (Enviado)"
    Case Is = 2
      TitleMov = "Salida (Pendiente)"
End Select

  '0  cb_mov.AddItem "Entrada"
  '1  cb_mov.AddItem "Salida (Enviado)"
  '2  cb_mov.AddItem "Salida (Pendiente)"
  '  OpcBqd

Select Case OpcHistorial

Case Is = 0 'normal se invoca desde el menu del main

    cb_mov.AddItem "Entrada"
    cb_mov.AddItem "Salida (Enviado)"
    cb_mov.AddItem "Salida (Pendiente)"
    FormatearGrids
    WB_Info.Navigate (App.path & "\dll\Inicio.html")

Case Is = 1 ' se invoca desde Vista: Hoy Pendientes

    cb_mov.AddItem "Entrada"
    cb_mov.AddItem "Salida (Enviado)"
    cb_mov.AddItem "Salida (Pendiente)"
    FormatearGrids
    WB_Info.Navigate (App.path & "\dll\Inicio.html")
    
    
    
    Call Bqd

Case Is = 2 ' se invoca desde Vista: Hoy Enviados

    cb_mov.AddItem "Entrada"
    cb_mov.AddItem "Salida (Enviado)"
    cb_mov.AddItem "Salida (Pendiente)"
    FormatearGrids
    WB_Info.Navigate (App.path & "\dll\Inicio.html")
    Call Bqd
    
Case Is = 3 ' se invoca desde Vista: Hoy Recibidos

    cb_mov.AddItem "Entrada"
    cb_mov.AddItem "Salida (Enviado)"
    cb_mov.AddItem "Salida (Pendiente)"
    FormatearGrids
    WB_Info.Navigate (App.path & "\dll\Inicio.html")
    Call Bqd
    
End Select


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




Select Case OpcBqd


Case Is = 0
'entradas detalle


    OrdenBodySql = "SELECT * from bodyrecepcion WHERE OIDRECEP =   '" & Oid & "';"

    Set RstBody = New ADODB.Recordset
    RstBody.CursorLocation = adUseClient
    RstBody.CursorType = adOpenDynamic
    RstBody.LockType = adLockPessimistic
    RstBody.Open OrdenBodySql, CadenaCnx



   ' Set DG_Users.DataSource = RstBody
    
'formatear la seccion de detalle de la orden de envio



     

Case Is = 1
' salida (enviado)


'    OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "' AND ENVIO=1;"
'    OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "';"
 '   OrdenBodySql = "SELECT * from BODYENVIO WHERE (((BODYENVIO.ENVIO)=0 Or (BODYENVIO.ENVIO)=9)) AND OIDENV =  '" & Oid & "';"
    OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "';"

    Set RstBody = New ADODB.Recordset
    RstBody.CursorLocation = adUseClient
    RstBody.CursorType = adOpenDynamic
    RstBody.LockType = adLockPessimistic
    RstBody.Open OrdenBodySql, CadenaCnx
    
    
    
    
    'Set DG_Users.DataSource = RstBody


Case Is = 2
' salida (pendiente)

    OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "';"
   ' OrdenBodySql = "SELECT * from BODYENVIO WHERE OIDENV =  '" & Oid & "';"

    Set RstBody = New ADODB.Recordset
    RstBody.CursorLocation = adUseClient
    RstBody.CursorType = adOpenDynamic
    RstBody.LockType = adLockPessimistic
    RstBody.Open OrdenBodySql, CadenaCnx
    
'    Set DG_Users.DataSource = RstBody



End Select


End Sub



Sub BuscarArchivos(Oid As String)

Dim OrdenBodySql As String
'Dim RstArch As ADODB.Recordset






Select Case OpcBqd

    Case Is = 0
    'Entrada
            OrdenBodySql = "SELECT * from Recibidos WHERE OIDRECEP =  '" & Oid & "';"

    
    
                    Set RstArch = New ADODB.Recordset
                    RstArch.CursorLocation = adUseClient
                    RstArch.CursorType = adOpenDynamic
                    RstArch.LockType = adLockPessimistic
                    RstArch.Open OrdenBodySql, CadenaCnx
                
                
                
                   ' Set DG_Arch.DataSource = RstArch

                 
   
    
    
    
    Case Is = 1, 2
    'Salida
            OrdenBodySql = "SELECT * from enviados WHERE OIDENV =  '" & Oid & "';"
    
                    Set RstArch = New ADODB.Recordset
                    RstArch.CursorLocation = adUseClient
                    RstArch.CursorType = adOpenDynamic
                    RstArch.LockType = adLockPessimistic
                    RstArch.Open OrdenBodySql, CadenaCnx
                
                
                
                   ' Set DG_Arch.DataSource = RstArch




     

End Select




End Sub




Sub GenerarReporte()


Open Opciones.RutaReportes & "\Reporte.html" For Output As #1




If RstBody.RecordCount <> 0 Then
                    
                    
                    Print #1, "<html>"
                    Print #1, "<head>"
                    Print #1, "<meta http-equiv='content-type' content=''>"
                    Print #1, "<title></title></head>"
                    'Print #1, "<BODY bgColor=#004a80 leftMargin=0 topMargin=0 marginheight='0' marginwidth='0'>"
                    Print #1, "<BODY bgColor=white leftMargin=0 topMargin=0 marginheight='0' marginwidth='0'>"
                    Print #1, "<h1>" & TitleMov & "</h1>"
                    Print #1, "<hr><center>"
                    
                    
                    Select Case OpcBqd
                    
                    
                    Case Is = 0 ' es una ENTRADA
                                
                                
                                Print #1, "<tt><b><font size=5>Lista de Usuarios</font></b></tt>"
                                Print #1, "<br>"
                                Print #1, "<br>"
                                
                                Print #1, "<table border=1>"
                                Print #1, "<tbody>"
                                Print #1, "<tr>"
                                
                                 'Nota: La informacion relativa a la direccion IP se quita par cuestiones de seguridad
                                
                                        Print #1, "<td style='font-size: 10pt;'><b> No. </b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Usuario que Envio </b></td>"
                                        'Print #1, "<td style='font-size: 10pt;'><b> IP Origen </b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Comentario </b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Inicio Transmicion </b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Fin Transmicion </b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Fecha Recepcion </b> </td>"
                                
                                Print #1, "</tr>"
                                
                                With RstBody
                                        
                                        
                                        .MoveFirst
                                        Do While Not RstBody.EOF
                                        
                                
                                        Print #1, "<tr>"
                                        Print #1, "<td>" & RstBody.Bookmark & "</td>"
                                        Print #1, "<td>" & RstBody!USUARIO_ORIGEN & "</td>"
                                        'Print #1, "<td>" & RstBody!IP_REMOTA & "</td>"
                                        Print #1, "<td>" & RstBody!COMENTARIO_RECEPCION & "</td>"
                                        Print #1, "<td>" & RstBody!PETICION_INICIO & "</td>"
                                      '  Print #1, "<td>" & RstBody!INICIO_TRANSMICION & "</td>"
                                        Print #1, "<td>" & RstBody!PETICION_FINAL & "</td>"
                                        Print #1, "<td>" & RstBody!FECHA_RECEPCION & "</td>"
                                        Print #1, "</tr>"
                                        
                                        .MoveNext
                                                
                                        Loop
                                .Close
                                End With
                                
                                
                                Print #1, "</tbody>"
                                Print #1, "</table>"
                    
                    
                                'Lista de los Archivos comprimidos
                                
                                Print #1, "<br>"
                                Print #1, "<br>"
                                
                                Print #1, "<tt><b><font size=5>Lista de Archivos</font></b></tt>"
                                Print #1, "<br>"
                                Print #1, "<br>"
                                
                    
                    
                                Print #1, "<table border=1>"
                                Print #1, "<tbody>"
                                Print #1, "<tr>"
                                
                                        Print #1, "<td style='font-size: 10pt;'><b>No.</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Nombre</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Tamaño bytes</b></td>"
                                       ' Print #1, "<td style='font-size: 10pt;'><b>Ruta Origen</b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Ruta Destino</b> </td>"
                                
                                Print #1, "</tr>"
                                
                                With RstArch
                                        
                                        .MoveFirst
                                        Do While Not RstArch.EOF
                                        
                                
                                        Print #1, "<tr>"
                                        Print #1, "<td>" & RstArch.Bookmark & "</td>"
                                        Print #1, "<td>" & RstArch!NOMBRE_ARCHIVO & "</td>"
                                        Print #1, "<td>" & RstArch!Tamaño & "</td>"
                                       ' Print #1, "<td>" & RstArch!RUTA_ORIGEN & "</td>"
                                        Print #1, "<td>" & RstArch!RUTA_DESTINO & "</td>"
                                        Print #1, "</tr>"
                                        
                                        .MoveNext
                                                
                                        Loop
                                .Close
                                End With
                                
                                
                                Print #1, "</tbody>"
                                Print #1, "</table>"
                                
                                
                                
                                
                                
                                
                    Case Is = 1, 2 ' Es una salida de informacion ya sea pendiente o ya se envio
                                
                                Print #1, "<tt><b><font size=5>Lista de Usuarios</font></b></tt>"
                                Print #1, "<br>"
                                Print #1, "<br>"
                                
                                Print #1, "<table border=1>"
                                Print #1, "<tbody>"
                                Print #1, "<tr>"
                                
                                        Print #1, "<td style='font-size: 10pt;'><b>No.</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Usuario Destino</b></td>"
                                        'Print #1, "<td style='font-size: 10pt;'><b>IP Usuario</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Fecha Envio</b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Comentario</b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Inicio Transmicion</b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b> Fin Transmicion </b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Fecha Registro</b> </td>"
                                
                                Print #1, "</tr>"
                                
                                With RstBody
                                        
                                  ' If .RecordCount = 0 Then
                                  '      Print #1, "<tr>"
                                  '      Print #1, "<td> Registros no encontrados </td>"
                                  '      Print #1, "</tR>"
                                        
                                  '
                                  ' End If
                                        
                                        .MoveFirst
                                        Do While Not RstBody.EOF
                                        
                                
                                        Print #1, "<tr>"
                                        Print #1, "<td>" & RstBody.Bookmark & "</td>"
                                        Print #1, "<td>" & RstBody!USUARIO_DESTINO & "</td>"
                                        'Print #1, "<td>" & RstBody!ip & "</td>"
                                        Print #1, "<td>" & RstBody!FECHA_ENVIO & "</td>"
                                        Print #1, "<td>" & RstBody!Comentario & "</td>"
                                        Print #1, "<td>" & RstBody!INICIO_TRANSMICION & "</td>"
                                        Print #1, "<td>" & RstBody!FIN_TRANSMICION & "</td>"
                                        Print #1, "<td>" & RstBody!FECHA_REGISTRO & "</td>"
                                        Print #1, "</tr>"
                                        
                                        .MoveNext
                                                
                                        Loop
                                .Close
                                End With
                                
                                
                                Print #1, "</tbody>"
                                Print #1, "</table>"
                    
                    
                                'Lista de los Archivos comprimidos
                                
                                Print #1, "<br>"
                                Print #1, "<br>"
                                
                                Print #1, "<tt><b><font size=5>Lista de Archivos</font></b></tt>"
                                Print #1, "<br>"
                                Print #1, "<br>"
                                
                    
                    
                                Print #1, "<table border=1>"
                                Print #1, "<tbody>"
                                Print #1, "<tr>"
                                
                                        Print #1, "<td style='font-size: 10pt;'><b>No.</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Nombre</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Tamaño bytes</b></td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Ruta Origen</b> </td>"
                                        Print #1, "<td style='font-size: 10pt;'><b>Ruta Destino</b> </td>"
                                
                                Print #1, "</tr>"
                                
                                With RstArch
                                        
                                        .MoveFirst
                                        Do While Not RstArch.EOF
                                        
                                
                                        Print #1, "<tr>"
                                        Print #1, "<td>" & RstArch.Bookmark & "</td>"
                                        Print #1, "<td>" & RstArch!NOMBRE_ARCHIVO & "</td>"
                                        Print #1, "<td>" & RstArch!Tamaño & "</td>"
                                        Print #1, "<td>" & RstArch!RUTA_ORIGEN & "</td>"
                                        Print #1, "<td>" & RstArch!RUTA_DESTINO & "</td>"
                                        Print #1, "</tr>"
                                        
                                        .MoveNext
                                                
                                        Loop
                                
                                End With
                                
                                
                                Print #1, "</tbody>"
                                Print #1, "</table>"
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    End Select
                    
                    Close (1)

Else

                    Print #1, "<html>"
                    Print #1, "<head>"
                    Print #1, "<meta http-equiv='content-type' content=''>"
                    Print #1, "<title></title></head>"
                    Print #1, "<BODY bgColor=#004a80 leftMargin=0 topMargin=0 marginheight='0' marginwidth='0'>"
                    
                    
                    Print #1, "<h1><center> <font color=RED>Error al Registrar el Movimiento (x)</FONT> </h1>"
                    Print #1, "<BR>"
                    Print #1, "<hr>"
                    Print #1, "<BR>"
                    Print #1, "<BR>"
                    Print #1, "<h2>Se sugiere que vuelva a hacer el archivo a enviar Falta parte del Registro</h2>"

                    
                    Print #1, "</center>"
                    Print #1, "</body>"
                    Print #1, "</html>"



                    Close (1)

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set RstMov = Nothing
Set RstBody = Nothing
Set RstArch = Nothing
Set DG_Mov.DataSource = Nothing

End Sub


Sub Bqd()

Dim FI As String
Dim FF As String

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

Select Case OpcBqd

Case Is = 0
'entradas
    OrdenHeaderSql = "SELECT OIDRECEP,FOLIO_RECEPCION,USUARIO_ORIGEN,COMENTARIO,FECHA_RECEPCION, FECHA_CREACION, NOMBRE_ARCHIVO_COMPRIMIDO,NO_ARCHIVOS From HEADER_RECEPCION WHERE (((FECHA_RECEPCION) Between  #" & FI & "# And  #" & FF & "# )) ORDER BY HEADER_RECEPCION.FOLIO_RECEPCION;"
            
           
            
            
                Set RstMov = New ADODB.Recordset
                RstMov.CursorLocation = adUseClient
                RstMov.CursorType = adOpenDynamic
                RstMov.LockType = adLockPessimistic
                RstMov.Open OrdenHeaderSql, CadenaCnx

'0 OIDRECEP
'1 FOLIO_RECEPCION
'2 USUARIO_ORIGEN
'3 Comentario
'4 FECHA_RECEPCION
'5 NOMBRE_ARCHIVO_COMPRIMIDO
'6 No_Archivos


                If RstMov.RecordCount > 0 Then
                ' hay registros sobre los cuales trabajar
                lblReg.Caption = "( " & RstMov.RecordCount & " )"
                lblMov.Caption = "Movimiento: " & cb_mov.Text
                
                
                DG_Mov.Enabled = True
                Set DG_Mov.DataSource = RstMov
                
                
                
                DG_Mov.Columns(0).Visible = False
                DG_Mov.Columns(0).Locked = True
                
                DG_Mov.Columns(1).Width = "1000"
                DG_Mov.Columns(1).Locked = True
                DG_Mov.Columns(1).Caption = "Folio Recepcion"
                
                DG_Mov.Columns(2).Width = "2500"
                DG_Mov.Columns(2).Locked = True
                DG_Mov.Columns(2).Caption = "Usuario Origen"
                
                DG_Mov.Columns(3).Width = "2500"
                DG_Mov.Columns(3).Locked = True
                DG_Mov.Columns(3).Caption = "Comentario"
                
                DG_Mov.Columns(4).Width = "2500"
                DG_Mov.Columns(4).Locked = True
                DG_Mov.Columns(4).Caption = "Fecha Recepcion"
                
                
                DG_Mov.Columns(5).Width = "2500"
                DG_Mov.Columns(5).Locked = True
                DG_Mov.Columns(5).Caption = "Fecha Registro"
                
                DG_Mov.Columns(6).Width = "2500"
                DG_Mov.Columns(6).Locked = True
                DG_Mov.Columns(6).Caption = "Nombre Archivo"
                
                DG_Mov.Columns(7).Width = "2500"
                DG_Mov.Columns(7).Locked = True
                DG_Mov.Columns(7).Caption = "No. Archivos"
                
                  
                Else
                ' no hay nada sobre que trabajar
                lblReg.Caption = "No se encontro nada con ese criterio"
                
                
                
                End If






Case Is = 1
'salida (enviado)

 OrdenHeaderSql = "SELECT OIDENV,FOLIO_SALIDA,USUARIO_CREO,COMENTARIO,FECHA_ENVIO,FECHA_REGISTRO, NOMBRE_ARCHIVO_COMPRIMIDO,NO_ARCHIVOS,NO_ENVIOS From HEADERENVIO WHERE (((FECHA_ENVIO) Between  #" & FI & "# And  #" & FF & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=1)) ORDER BY FOLIO_SALIDA;"

  ' OrdenHeaderSql = "SELECT OIDENV,FOLIO_SALIDA,USUARIO_CREO,COMENTARIO,FECHA_ENVIO,FECHA_REGISTRO, NOMBRE_ARCHIVO_COMPRIMIDO,NO_ARCHIVOS,NO_ENVIOS From HEADERENVIO WHERE FECHA_ENVIO Between  #" & FI & "# And  #" & FF & "#   ORDER BY FOLIO_SALIDA;"


    Set RstMov = New ADODB.Recordset
    RstMov.CursorLocation = adUseClient
    RstMov.CursorType = adOpenDynamic
    RstMov.LockType = adLockPessimistic
    RstMov.Open OrdenHeaderSql, CadenaCnx






    If RstMov.RecordCount > 0 Then
    ' hay registros sobre los cuales trabajar
    lblReg.Caption = "( " & RstMov.RecordCount & " )"
    lblMov.Caption = "Movimiento: " & cb_mov.Text
    
    
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
    lblReg.Caption = "No se encontro nada con ese criterio"
    
    
    
    End If




Case Is = 2
'salida (pendiente)

    OrdenHeaderSql = "SELECT OIDENV,FOLIO_SALIDA,USUARIO_CREO,COMENTARIO,FECHA_ENVIO, FECHA_REGISTRO,NOMBRE_ARCHIVO_COMPRIMIDO,NO_ARCHIVOS,NO_ENVIOS From HEADERENVIO WHERE (((FECHA_REGISTRO) Between  #" & FI & "# And  #" & FF & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=0)) ORDER BY FOLIO_SALIDA;"

    Set RstMov = New ADODB.Recordset
    RstMov.CursorLocation = adUseClient
    RstMov.CursorType = adOpenDynamic
    RstMov.LockType = adLockPessimistic
    RstMov.Open OrdenHeaderSql, CadenaCnx




    If RstMov.RecordCount > 0 Then
    ' hay registros sobre los cuales trabajar
    lblReg.Caption = "( " & RstMov.RecordCount & " )"
    lblMov.Caption = "Movimiento: " & cb_mov.Text
    
    
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
    lblReg.Caption = "No se encontro nada con ese criterio"
    
    
    
    End If





End Select


 'Debug.Print OrdenHeaderSql




End Sub
