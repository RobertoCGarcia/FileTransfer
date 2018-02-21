VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmAdminUsers 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Administracion de Usuarios"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer T_Hora 
      Interval        =   1000
      Left            =   9720
      Top             =   7680
   End
   Begin VB.PictureBox pbx_Usuarios 
      Height          =   9135
      Left            =   120
      ScaleHeight     =   9075
      ScaleWidth      =   10035
      TabIndex        =   19
      Top             =   0
      Width           =   10095
      Begin VB.TextBox txtcoment 
         Enabled         =   0   'False
         Height          =   855
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   6000
         Width           =   3615
      End
      Begin VB.TextBox txtGrupo 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   5160
         Width           =   3615
      End
      Begin VB.ComboBox cb_Grupo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CheckBox chkPermiso 
         Caption         =   "Modificar IP"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   48
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CheckBox chkPermiso 
         Caption         =   "Consultar Info Usuarios"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   47
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CheckBox chkPermiso 
         Caption         =   "Eliminar Usuarios"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   46
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CheckBox chkPermiso 
         Caption         =   "Modificar Usuarios"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   45
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CheckBox chkPermiso 
         Caption         =   "Agregar Usuarios"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   44
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CheckBox chkPermiso 
         Caption         =   "Modificar Configuracion"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   43
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CheckBox chk_Modip 
         Caption         =   "Modificar IP"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5880
         TabIndex        =   9
         Top             =   6600
         Width           =   1935
      End
      Begin VB.TextBox txtipactual 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6120
         Width           =   3615
      End
      Begin VB.TextBox txtip 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   8040
         MaxLength       =   3
         TabIndex        =   13
         Top             =   6840
         Width           =   495
      End
      Begin VB.TextBox txtip 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   7320
         MaxLength       =   3
         TabIndex        =   12
         Top             =   6840
         Width           =   495
      End
      Begin VB.TextBox txtip 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   11
         Top             =   6840
         Width           =   495
      End
      Begin VB.TextBox txtip 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   10
         Top             =   6840
         Width           =   495
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8160
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DG_Usuarios 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8160
         Width           =   1815
      End
      Begin VB.TextBox txtArea 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4800
         Width           =   3615
      End
      Begin VB.CommandButton cmd_Nvo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   8160
         Width           =   1815
      End
      Begin VB.ComboBox Cb_Area 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtPassword2 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   4440
         Width           =   3615
      End
      Begin VB.TextBox txtPassword1 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   4080
         Width           =   3615
      End
      Begin VB.TextBox txtNick 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   3
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox txtApPat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   2
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   1
         Top             =   3000
         Width           =   3615
      End
      Begin VB.CommandButton cmd_cancelar 
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
         Height          =   735
         Left            =   8040
         TabIndex        =   18
         Top             =   8160
         Width           =   1815
      End
      Begin VB.CommandButton cmd_Guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6000
         TabIndex        =   17
         Top             =   8160
         Width           =   1815
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   5880
         TabIndex        =   52
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   5880
         TabIndex        =   51
         Top             =   4920
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Left            =   7920
         TabIndex        =   42
         Top             =   6960
         Width           =   885
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Left            =   7200
         TabIndex        =   41
         Top             =   6960
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Left            =   6480
         TabIndex        =   40
         Top             =   6960
         Width           =   45
      End
      Begin VB.Label Label15 
         Caption         =   $"FrmAddUsers.frx":0000
         Height          =   735
         Left            =   5880
         TabIndex        =   39
         Top             =   7200
         Width           =   3735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "IP Asignada"
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
         Left            =   5880
         TabIndex        =   38
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
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
         TabIndex        =   37
         Top             =   6000
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Entradas"
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
         TabIndex        =   36
         Top             =   5640
         Width           =   1125
      End
      Begin VB.Label lblEntradas 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1920
         TabIndex        =   35
         Top             =   5640
         Width           =   45
      End
      Begin VB.Label lblsalidas 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1920
         TabIndex        =   34
         Top             =   5280
         Width           =   45
      End
      Begin VB.Label lblAccion 
         AutoSize        =   -1  'True
         Caption         =   "Ninguna Accion"
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
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Area"
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
         TabIndex        =   32
         Top             =   4800
         Width           =   405
      End
      Begin VB.Label lblfechaModif 
         AutoSize        =   -1  'True
         Caption         =   "No Seleccionada"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2400
         TabIndex        =   31
         Top             =   7560
         Width           =   1230
      End
      Begin VB.Label lblFechacreacion 
         AutoSize        =   -1  'True
         Caption         =   "No seleccionada"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2400
         TabIndex        =   30
         Top             =   7200
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Modificado"
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
         TabIndex        =   29
         Top             =   7560
         Width           =   1530
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Creacion:"
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
         TabIndex        =   28
         Top             =   7200
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Area"
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
         Left            =   5880
         TabIndex        =   27
         Top             =   3000
         Width           =   405
      End
      Begin VB.Line Line1 
         X1              =   5760
         X2              =   5760
         Y1              =   2880
         Y2              =   7680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar Password"
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
         TabIndex        =   26
         Top             =   4440
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         TabIndex        =   25
         Top             =   4080
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nick"
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
         TabIndex        =   24
         Top             =   3720
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. Salidas"
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
         TabIndex        =   23
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellido"
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
         TabIndex        =   22
         Top             =   3360
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         TabIndex        =   21
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "USUARIOS ACTIVOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   2985
      End
   End
End
Attribute VB_Name = "FrmAdminUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public opcion As String
Dim Rst_Gral As ADODB.Recordset
Dim Rst_Areas As ADODB.Recordset
Dim Str_Gral As String
Dim TipoAccion As String
Dim IdArea As String
Dim NumOidUser As Double
Dim BanderaSel As Boolean
Dim ClaveIni As String
Dim OidUser As String



'////////////////////////////////////////////////////////////////////////////+
'///
'///
'///
            'formulario para la administracion de los usuarios internos del sistema, con derechos
            'sobre cada operacio permitida dentro de este sistema.
            'By Pv.Karlos 2005 (Kwalick)
'///
'///
'///
'///
'///
'////////////////////////////////////////////////////////////////////////////+



Private Sub Cb_Area_Click()

Dim StrFilter As String
Dim Nombrecampo As String

Nombrecampo = Cb_Area.Text

               StrFilter = "NOMBRE_AREA = '" & Nombrecampo & "'"
               Rst_Areas.Filter = StrFilter
               IdArea = Rst_Areas.Fields("oidarea")
               Rst_Areas.Filter = ""
               
             ''  MsgBox IdArea
               
End Sub

Private Sub chk_Modip_Click()

If chk_Modip.Value = 1 Then
chk_Modip.Enabled = False
End If

txtip(0).BackColor = &HFFFFFF
txtip(1).BackColor = &HFFFFFF
txtip(2).BackColor = &HFFFFFF
txtip(3).BackColor = &HFFFFFF




End Sub

Private Sub cmd_Cancelar_Click()
Unload Me
End Sub

Private Sub cmd_guardar_Click()
lblAccion.Caption = "Guardando Cambios..."
Dim Nombrecampo As String
Dim StrFilter As String
Dim i As Integer

Select Case TipoAccion

Case Is = "NUEVO"

            If Len(txtNombre) = 0 Then
            MsgBox "Nombre requerido", vbCritical, "Dato Necesario"
            txtNombre.Text = ""
            txtNombre.SetFocus
            Exit Sub
            End If
            
            If Len(txtApPat) = 0 Then
            MsgBox "Apellido requerido", vbCritical, "Dato Necesario"
            txtApPat.Text = ""
            txtApPat.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtNick) = 0 Then
            MsgBox "Nick requerido", vbCritical, "Dato Necesario"
            txtNick.Text = ""
            txtNick.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtPassword1) = 0 Then
            MsgBox "Password requerido", vbCritical, "Dato Necesario"
            txtPassword1.Text = ""
            txtPassword1.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtPassword2) = 0 Then
            MsgBox "Password requerido", vbCritical, "Dato Necesario"
            txtPassword2.Text = ""
            txtPassword2.SetFocus
            Exit Sub
            End If
            
            
            If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "Password debe de ser Igual, escribalo de nuevo", vbCritical, "Dato Necesario"
            txtPassword1.Text = ""
            txtPassword2.Text = ""
            txtPassword1.SetFocus
            Exit Sub
            End If
            
            
            If (ValidarPassword(txtPassword1)) Then
            ' SI YA EXISTE
            MsgBox "El Password que intenta utilizar ya existe, elija otro diferente", vbCritical, "Password No Valido"
            txtPassword1.Text = ""
            txtPassword2.Text = ""
            txtPassword1.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtComent.Text) = 0 Then
            MsgBox "Comentario requerido!", vbCritical, "Dato Necesario"
            txtComent.SetFocus
            Exit Sub
            End If
            
    
    
            If Len(Cb_Area.Text) = 0 Then
            MsgBox "Elija un area adecuadamente!!", vbCritical, "Dato Necesario"
            Cb_Area.SetFocus
            Exit Sub
            End If
        
        
                
                For i = 0 To 3
                    
                    If Val(txtip(i).Text) < 0 Then
                       MsgBox "Numero no valido, escribirlo de nuevo."
                       txtip(i).Text = ""
                       txtip(i).SetFocus
                       Exit Sub
                    End If
                    
                    If Val(txtip(i).Text) > 255 Then
                       MsgBox "Numero no valido, escribirlo de nuevo."
                       txtip(i).Text = ""
                       txtip(i).SetFocus
                       Exit Sub
                    End If
                    
                    If Len(txtip(i).Text) = 0 Then
                       MsgBox "Llenar el espacio en blanco."
                       txtip(i).Text = ""
                       txtip(i).SetFocus
                       Exit Sub
                    End If
                    
                Next i
                
                
                   With Rst_Gral
                   ' NumOidUser = (CDbl(Dato("OID_USUARIO", , 1, , "DATO", "NOMBRE_DATO")) + 1)
                 'MsgBox NumOidUser
                    
                        .AddNew
                        
0                        Rst_Gral!OID_USUARIO = "USER" & Mid(CStr(Rnd), 3, 5)
1                        Rst_Gral!NOMBRE_USUARIO = txtNombre.Text
2                        Rst_Gral!APELLIDO_USUARIO = txtApPat.Text
3                        Rst_Gral!OID_AREA = IdArea
5                        Rst_Gral!Status = "OK"
6                        Rst_Gral!Nick = txtNick.Text
7                        Rst_Gral!Password = txtPassword1.Text
8                        Rst_Gral!USUARIO_CREO = UsuarioActivo.Nick
9                        Rst_Gral!FECHA_CREACION = Now
10                       Rst_Gral!FECHA_ACTUALIZACION = Now
11                       Rst_Gral!USUARIO_ACTUALIZO = UsuarioActivo.Nick
12                       Rst_Gral!Salidas = 0
13                       Rst_Gral!Entradas = 0
14                       Rst_Gral!IP_equipo = txtip(0).Text & "." & txtip(1).Text & "." & txtip(2).Text & "." & txtip(3).Text
15                       Rst_Gral!Comentario = txtComent.Text
                
                '16       Rst_Gral!Default = txtComent.Text
17                       Rst_Gral!PERMISO_MODCONFIG = chkPermiso(0).Value
18                       Rst_Gral!PERMISO_CONINFOUSER = chkPermiso(1).Value
19                       Rst_Gral!PERMISO_ADDUSERS = chkPermiso(2).Value
20                       Rst_Gral!PERMISO_MODUSERS = chkPermiso(3).Value
21                       Rst_Gral!PERMISO_DELUSERS = chkPermiso(4).Value
22                       Rst_Gral!PERMISO_MODIP = chkPermiso(5).Value
                
                         If Len(cb_Grupo.Text) > 0 Then
                         
                            Rst_Gral!Grupo = cb_Grupo.Text
                         
                         End If
                         
                
                
                .Update
                   
                   End With
                
                  
                
                MsgBox "Usuario Agregado con Exito", vbInformation, "Operacion Exitosa"
                
                Rst_Gral.Requery
                lblTitulo.Caption = lblTitulo.Caption & "   ( " & Rst_Gral.RecordCount & " )"
                Unload Me


Case Is = "MODIFICACION"

        
            If Len(txtNombre) = 0 Then
            MsgBox "Nombre requerido", vbCritical, "Dato Necesario"
            txtNombre.Text = ""
            txtNombre.SetFocus
            Exit Sub
            End If
            
            If Len(txtApPat) = 0 Then
            MsgBox "Apellido requerido", vbCritical, "Dato Necesario"
            txtApPat.Text = ""
            txtApPat.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtNick) = 0 Then
            MsgBox "Nick requerido", vbCritical, "Dato Necesario"
            txtNick.Text = ""
            txtNick.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtPassword1) = 0 Then
            MsgBox "Password requerido", vbCritical, "Dato Necesario"
            txtPassword1.Text = ""
            txtPassword1.SetFocus
            Exit Sub
            End If
            
            
            If Len(txtPassword2) = 0 Then
            MsgBox "Password requerido", vbCritical, "Dato Necesario"
            txtPassword2.Text = ""
            txtPassword2.SetFocus
            Exit Sub
            End If
            
            
            If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "Password debe de ser Igual, escribalo de nuevo", vbCritical, "Dato Necesario"
            txtPassword1.Text = ""
            txtPassword2.Text = ""
            txtPassword1.SetFocus
            Exit Sub
            End If
            
            
            
            If ClaveIni <> txtPassword1.Text Then
                If (ValidarPassword(txtPassword1.Text)) Then
                ' SI YA EXISTE
                MsgBox "El Password que intenta utilizar ya existe, elija otro diferente", vbCritical, "Password No Valido"
                txtPassword1.Text = ""
                txtPassword2.Text = ""
                txtPassword1.SetFocus
                Exit Sub
                End If
            End If
        
                
        
          Rst_Gral.MoveFirst
          
          With Rst_Gral
            

               StrFilter = "OID_USUARIO = '" & OidUser & "'"
               Rst_Gral.Filter = StrFilter
            
               'MsgBox "Reg Filtrados: " & Rst_Gral.RecordCount
               '.Bookmark = DG_Usuarios.Bookmark
               ' .AbsolutePosition = DG_Usuarios.Bookmark
                Rst_Gral!NOMBRE_USUARIO = txtNombre.Text
                Rst_Gral!APELLIDO_USUARIO = txtApPat.Text
                
                If Len(Cb_Area.Text) = 0 Then ' no se modifico el area
                Rst_Gral!OID_AREA = IdArea
                Else
                Rst_Gral!OID_AREA = IdArea
                End If
                
                Rst_Gral!Nick = txtNick.Text
                Rst_Gral!Password = txtPassword1.Text
                'Rst_Gral!USUARIO_CREO = PerfilUser.NickUser
                Rst_Gral!USUARIO_ACTUALIZO = UsuarioActivo.Nick
                
                'Rst_Gral!FECHA_CREACION_USER = lblFechacreacion
                Rst_Gral!FECHA_ACTUALIZACION = Now
         
                Rst_Gral!Comentario = txtComent.Text
                
                If chk_Modip.Value = 1 Then
                 
                   Rst_Gral!IP_equipo = txtip(0).Text & "." & txtip(1).Text & "." & txtip(2).Text & "." & txtip(3).Text
                 
                Else
                
                   Rst_Gral!IP_equipo = txtipactual.Text
                
                End If

                
                
                        '16       Rst_Gral!Default = txtComent.Text
               Rst_Gral!PERMISO_MODCONFIG = chkPermiso(0).Value
               Rst_Gral!PERMISO_CONINFOUSER = chkPermiso(1).Value
               Rst_Gral!PERMISO_ADDUSERS = chkPermiso(2).Value
               Rst_Gral!PERMISO_MODUSERS = chkPermiso(3).Value
               Rst_Gral!PERMISO_DELUSERS = chkPermiso(4).Value
               Rst_Gral!PERMISO_MODIP = chkPermiso(5).Value
        
                 If Len(cb_Grupo.Text) > 0 Then
                 
                    Rst_Gral!Grupo = cb_Grupo.Text
                 
                 End If
               .Update
            End With
            
            Rst_Gral.Filter = ""
            
            
        MsgBox "Cambios guardados con Exito", vbInformation, "Operacion Exitosa"
        Unload Me


End Select

End Sub


Private Sub cmd_Nvo_Click()

Derecho = 2
FrmPassword.Show

End Sub

Private Sub cmdEliminar_Click()

Derecho = 4
FrmPassword.Show

End Sub

Private Sub cmdModificar_Click()

Derecho = 3
FrmPassword.Show

End Sub

Private Sub DG_Usuarios_Click()


Dim StrFilter As String
Dim Nombrecampo As String

BanderaSel = True

txtNombre = Rst_Gral!NOMBRE_USUARIO
txtApPat = Rst_Gral!APELLIDO_USUARIO
txtNick = Rst_Gral!Nick

txtPassword1 = Rst_Gral!Password
txtPassword2 = Rst_Gral!Password

ClaveIni = Rst_Gral!Password

OidUser = Rst_Gral!OID_USUARIO

lblFechacreacion = Rst_Gral!FECHA_CREACION
lblfechaModif = Rst_Gral!FECHA_ACTUALIZACION
'lblUser = Rst_Gral!USUARIO_CREO

 'txtApMat = Rst_Gral!APELLIDO_MATERNO


lblsalidas = Rst_Gral!Salidas
lblEntradas = Rst_Gral!Entradas
txtComent = Rst_Gral!Comentario

txtipactual = Rst_Gral!IP_equipo

'obtengo el nombre del area
'txtArea = Rst_Gral!OID_Area

Nombrecampo = Rst_Gral!OID_AREA

IdArea = Rst_Gral!OID_AREA
StrFilter = "OIDAREA = '" & Nombrecampo & "'"
Rst_Areas.Filter = StrFilter

txtArea = Rst_Areas.Fields("NOMBRE_AREA")

Rst_Areas.Filter = ""


'15 Default
'16 PERMISO_MODCONFIG
'17 PERMISO_CONINFOUSER
'18 PERMISO_ADDUSERS
'19 PERMISO_MODUSERS
'20 PERMISO_DELUSERS
'21 PERMISO_MODIP
'22 Grupo '

txtGrupo.Text = Rst_Gral!Grupo

chkPermiso(0).Value = IIf(Rst_Gral!PERMISO_MODCONFIG = True, 1, 0)  'modificar configuracion
chkPermiso(1).Value = IIf(Rst_Gral!PERMISO_CONINFOUSER = True, 1, 0) 'consultar info usuarios
chkPermiso(2).Value = IIf(Rst_Gral!PERMISO_ADDUSERS = True, 1, 0) 'agregar usuarios
chkPermiso(3).Value = IIf(Rst_Gral!PERMISO_MODUSERS = True, 1, 0) 'modificar usuarios
chkPermiso(4).Value = IIf(Rst_Gral!PERMISO_DELUSERS = True, 1, 0) 'eliminar usuarios
chkPermiso(5).Value = IIf(Rst_Gral!PERMISO_MODIP = True, 1, 0) ' modificar ip



cmdModificar.Enabled = True
cmdEliminar.Enabled = True







  
 

End Sub

Private Sub Form_Load()

cb_Grupo.AddItem "USUARIOS"
cb_Grupo.AddItem "ADMINISTRADOR"
BanderaSel = False
            Dim i As Integer
              Str_Gral = "select * from Areas"
              Set Rst_Areas = New ADODB.Recordset
              Rst_Areas.CursorLocation = adUseClient
              Rst_Areas.CursorType = adOpenDynamic
              Rst_Areas.LockType = adLockOptimistic
              Rst_Areas.Open Str_Gral, CadenaCnx
                Str_Gral = "SELECT * FROM USUARIOS WHERE STATUS='OK'"
                Me.Caption = "Acciones sobre Usuarios"
                Set Rst_Gral = New ADODB.Recordset
                Rst_Gral.CursorLocation = adUseClient
                Rst_Gral.CursorType = adOpenDynamic
                Rst_Gral.LockType = adLockOptimistic
                Rst_Gral.Open Str_Gral, CadenaCnx
                 Rst_Areas.MoveFirst
                 For i = 1 To Rst_Areas.RecordCount
                  Cb_Area.AddItem Rst_Areas.Fields("NOMBRE_AREA")
                  Rst_Areas.MoveNext
                 Next i
            
            lblTitulo.Caption = lblTitulo.Caption & "   ( " & Rst_Gral.RecordCount & " )"
                
              Set DG_Usuarios.DataSource = Rst_Gral
            
            If Rst_Gral.RecordCount = 0 Then
            DG_Usuarios.Enabled = False
            End If
            
            
                
                DG_Usuarios.Columns(1).Width = "3000"
                DG_Usuarios.Columns(1).Caption = "NOMBRE"
                DG_Usuarios.Columns(1).Locked = True
                DG_Usuarios.Columns(2).Width = "3000"
                DG_Usuarios.Columns(2).Caption = "APELLIDO"
                DG_Usuarios.Columns(2).Locked = True
                DG_Usuarios.Columns(0).Visible = False
                DG_Usuarios.Columns(3).Visible = False
                DG_Usuarios.Columns(4).Visible = False
                DG_Usuarios.Columns(5).Visible = False
                DG_Usuarios.Columns(6).Visible = False
                DG_Usuarios.Columns(7).Visible = False
                DG_Usuarios.Columns(8).Visible = False
                DG_Usuarios.Columns(9).Visible = False
                DG_Usuarios.Columns(10).Visible = False
                DG_Usuarios.Columns(11).Visible = False
                DG_Usuarios.Columns(12).Visible = False
                DG_Usuarios.Columns(13).Visible = False
                DG_Usuarios.Columns(14).Visible = False
                DG_Usuarios.Columns(15).Visible = False
                DG_Usuarios.Columns(16).Visible = False
                DG_Usuarios.Columns(17).Visible = False
                DG_Usuarios.Columns(18).Visible = False
                DG_Usuarios.Columns(19).Visible = False
                DG_Usuarios.Columns(20).Visible = False
                DG_Usuarios.Columns(21).Visible = False
               
                         
          
'15 Default
'16 PERMISO_MODCONFIG
'17 PERMISO_CONINFOUSER
'18 PERMISO_ADDUSERS
'19 PERMISO_MODUSERS
'20 PERMISO_DELUSERS
'21 PERMISO_MODIP
'22 Grupo '


End Sub
Private Sub Form_Unload(Cancel As Integer)
Rst_Gral.Close
Set Rst_Gral = Nothing
Set Rst_Areas = Nothing
Unload FrmOpciones
End Sub

Private Sub txtApMat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtApPat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtcoment_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtip_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then
    Exit Sub
    End If
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
    KeyAscii = 0
    Beep
    End If
End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Function ValidarPassword(pass As String) As Boolean
'Valido que no exista un password igual dentro de la tabla
Dim RstLocal As ADODB.Recordset
Rst_Gral.MoveFirst
Set RstLocal = Rst_Gral
    Do While Not Rst_Gral.EOF
      
      If pass = Rst_Gral!Password Then
         ValidarPassword = True
         Exit Function
      End If
        Rst_Gral.MoveNext
    
    Loop
Rst_Gral.MoveFirst
End Function

Public Sub AddUser()

Dim s As Integer
TipoAccion = "NUEVO"
lblAccion.Caption = "Nuevo usuario"
txtNombre.Enabled = True
txtApPat.Enabled = True
txtNick.Enabled = True
txtPassword1.Enabled = True
txtPassword2.Enabled = True
lblFechacreacion.Enabled = True
lblfechaModif.Enabled = True
txtArea.Enabled = True
txtComent.Enabled = True
chk_Modip.Value = 1
Cb_Area.Enabled = True
txtNombre.Text = ""
txtApPat.Text = ""
txtNick.Text = ""
txtPassword1.Text = ""
txtPassword2.Text = ""
lblFechacreacion.Caption = ""
lblfechaModif.Caption = ""
txtArea.Text = ""
txtComent.Text = ""
txtipactual.Text = ""
lblfechaModif = Now
lblFechacreacion = Now
txtNombre.SetFocus
cmd_guardar.Enabled = True
cmd_Nvo.Enabled = False
cmdModificar.Enabled = False
cmdEliminar.Enabled = False
chkPermiso(0).Enabled = True
chkPermiso(1).Enabled = True
chkPermiso(2).Enabled = True
chkPermiso(3).Enabled = True
chkPermiso(4).Enabled = True
chkPermiso(5).Enabled = True
cb_Grupo.Enabled = True
chkPermiso(0).Value = 0 'modificar configuracion
chkPermiso(1).Value = 0  'consultar info usuarios
chkPermiso(2).Value = 0  'agregar usuarios
chkPermiso(3).Value = 0  'modificar usuarios
chkPermiso(4).Value = 0  'eliminar usuarios
chkPermiso(5).Value = 0
txtip(0).Enabled = True
txtip(1).Enabled = True
txtip(2).Enabled = True
txtip(3).Enabled = True
End Sub


Public Sub ModiUsers()


Dim s As Integer
If BanderaSel = False Then
MsgBox "Elija un usuario adecuadamente para modificar", vbExclamation, "Registro requerido"
Exit Sub
End If

TipoAccion = "MODIFICACION"
lblAccion.Caption = "Modificar Usuario Existente"
txtNombre.Enabled = True
txtApPat.Enabled = True
txtNick.Enabled = True
txtPassword1.Enabled = True
txtPassword2.Enabled = True
lblFechacreacion.Enabled = True
lblfechaModif.Enabled = True
txtArea.Enabled = True
txtComent.Enabled = True
Cb_Area.Enabled = True
cb_Grupo.Enabled = True
Me.DG_Usuarios.Enabled = True
txtip(0).Enabled = True
txtip(1).Enabled = True
txtip(2).Enabled = True
txtip(3).Enabled = True
chkPermiso(0).Enabled = True
chkPermiso(1).Enabled = True
chkPermiso(2).Enabled = True
chkPermiso(3).Enabled = True
chkPermiso(4).Enabled = True
chkPermiso(5).Enabled = True
chk_Modip.Enabled = True
txtNombre.SetFocus
cmd_guardar.Enabled = True
End Sub


Public Sub Delusers()

If BanderaSel = False Then
MsgBox "Elija un usuario adecuadamente para Eliminar", vbExclamation, "Registro requerido"
Exit Sub
End If

If MsgBox("Desea Eliminar a el Usuario: " & txtNombre & " " & txtApPat, vbYesNo + vbQuestion) = vbYes Then
 
 
If Rst_Gral.RecordCount = 1 Then
MsgBox "No se puede Eliminar, ya que tiene que quedarse al menos un registro", vbInformation, "No valido"
Exit Sub
End If
    
    With Rst_Gral
           .Bookmark = DG_Usuarios.Bookmark
           If Usuario.OidUsuario = Rst_Gral!OID_USUARIO Then
           MsgBox "No se puede Eliminar, ya que es el usuario activo, elija otro o cambie de usuario por Default", vbCritical, "No valido"
           Exit Sub
           End If
           Rst_Gral!FECHA_ACTUALIZACION = Now
           Rst_Gral!USUARIO_ACTUALIZO = Usuario.Nick
           Rst_Gral!Status = "DEL"
           .Update
    End With
Rst_Gral.Requery
lblTitulo.Caption = lblTitulo.Caption & "   ( " & Rst_Gral.RecordCount & " )"
If Rst_Gral.RecordCount = 0 Then
 DG_Usuarios.Enabled = False
End If
MsgBox "Usuario eliminado con exito", vbInformation, "Operacion Exitosa"
                DG_Usuarios.Columns(1).Width = "3000"
                DG_Usuarios.Columns(1).Caption = "NOMBRE"
                DG_Usuarios.Columns(1).Locked = True
                DG_Usuarios.Columns(2).Width = "3000"
                DG_Usuarios.Columns(2).Caption = "APELLIDO"
                DG_Usuarios.Columns(2).Locked = True
                DG_Usuarios.Columns(3).Width = "3000"
                DG_Usuarios.Columns(3).Caption = "APELLIDO MATERNO"
                DG_Usuarios.Columns(3).Locked = True
                DG_Usuarios.Columns(0).Visible = False
                DG_Usuarios.Columns(4).Visible = False
                DG_Usuarios.Columns(5).Visible = False
                DG_Usuarios.Columns(6).Visible = False
                DG_Usuarios.Columns(7).Visible = False
                DG_Usuarios.Columns(8).Visible = False
                DG_Usuarios.Columns(9).Visible = False
                DG_Usuarios.Columns(10).Visible = False
                DG_Usuarios.Columns(11).Visible = False
                DG_Usuarios.Columns(12).Visible = False
                DG_Usuarios.Columns(13).Visible = False
                DG_Usuarios.Columns(14).Visible = False
                DG_Usuarios.Columns(15).Visible = False
                DG_Usuarios.Columns(16).Visible = False
                DG_Usuarios.Columns(17).Visible = False
                DG_Usuarios.Columns(18).Visible = False
                DG_Usuarios.Columns(19).Visible = False
                DG_Usuarios.Columns(20).Visible = False
                DG_Usuarios.Columns(21).Visible = False
End If
Unload Me
End Sub
