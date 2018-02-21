VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConfiguracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferencias del Sistema"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_guardar 
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
      Left            =   5760
      TabIndex        =   47
      Top             =   4680
      Width           =   1575
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
      TabIndex        =   46
      Top             =   4680
      Width           =   1575
   End
   Begin TabDlg.SSTab TabConfiguracion 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Carpetas del Sistema"
      TabPicture(0)   =   "FrmConfiguracion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblReportes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblArch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblExt"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblBd"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblEns"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblgen"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblRec"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblenv"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label18"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblLog"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmd_Reportes"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmd_ListaGen"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd_RutArch"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_RutExt"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdRut_Bd"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdRut_Ens"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdRut_Gen"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdRuta_Rec"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdRuta_Env"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmd_RutaLog"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Informacion de Red"
      TabPicture(1)   =   "FrmConfiguracion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPuertoRecep"
      Tab(1).Control(1)=   "txtPuertoEnv"
      Tab(1).Control(2)=   "txtNombre"
      Tab(1).Control(3)=   "OptServ"
      Tab(1).Control(4)=   "OptHt"
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(6)=   "lblIPactual"
      Tab(1).Control(7)=   "Label9"
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(9)=   "Label14"
      Tab(1).Control(10)=   "Label15"
      Tab(1).Control(11)=   "Label16"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Usuarios"
      TabPicture(2)   =   "FrmConfiguracion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "lblUpdate"
      Tab(2).Control(2)=   "lblUser"
      Tab(2).Control(3)=   "Label8(1)"
      Tab(2).Control(4)=   "cmdAdduser"
      Tab(2).Control(5)=   "chkLista"
      Tab(2).Control(6)=   "cb_User"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Acciones"
      TabPicture(3)   =   "FrmConfiguracion.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label19"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label20"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label22"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label23"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblRut1"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblRut2"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lblRut3"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblRut4"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblRut5"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Line2"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "cmd_Add"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "CD_Cmd"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Lv_Cmd"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "cmdRuta_1"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "cmdRuta_2"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "cmdRuta_3"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "cmdRuta_4"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "cmdRuta_5"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "cmd_Edit"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "cmd_del"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).ControlCount=   21
      Begin VB.CommandButton cmd_del 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68400
         TabIndex        =   70
         ToolTipText     =   "Borra el comando elegido"
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmd_Edit 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69360
         TabIndex        =   68
         ToolTipText     =   "Edita el comando seleccionado"
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdRuta_5 
         Caption         =   "!"
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
         Left            =   -72840
         TabIndex        =   66
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdRuta_4 
         Caption         =   "!"
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
         Left            =   -72840
         TabIndex        =   65
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdRuta_3 
         Caption         =   "!"
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
         Left            =   -72840
         TabIndex        =   64
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdRuta_2 
         Caption         =   "!"
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
         Left            =   -72840
         TabIndex        =   63
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdRuta_1 
         Caption         =   "!"
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
         Left            =   -72840
         TabIndex        =   62
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmd_RutaLog 
         Caption         =   "..."
         Height          =   255
         Left            =   8280
         TabIndex        =   49
         Top             =   3240
         Width           =   375
      End
      Begin VB.ComboBox cb_User 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chkLista 
         Caption         =   "Enviar Lista de Usuarios"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   1080
         Width           =   2055
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
         Left            =   -74760
         TabIndex        =   38
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtPuertoRecep 
         Height          =   285
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtPuertoEnv 
         Height          =   285
         Left            =   -74760
         TabIndex        =   31
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   -70200
         MaxLength       =   15
         TabIndex        =   30
         Top             =   1800
         Width           =   3855
      End
      Begin VB.OptionButton OptServ 
         Caption         =   "Servidor"
         Height          =   255
         Left            =   -70200
         TabIndex        =   29
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton OptHt 
         Caption         =   "Host Normal"
         Height          =   255
         Left            =   -70200
         TabIndex        =   28
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdRuta_Env 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdRuta_Rec 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdRut_Gen 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdRut_Ens 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdRut_Bd 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmd_RutExt 
         Caption         =   "..."
         Height          =   255
         Left            =   8280
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmd_RutArch 
         Caption         =   "..."
         Height          =   255
         Left            =   8280
         TabIndex        =   3
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmd_ListaGen 
         Caption         =   "..."
         Height          =   255
         Left            =   8280
         TabIndex        =   2
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmd_Reportes 
         Caption         =   "..."
         Height          =   255
         Left            =   8280
         TabIndex        =   1
         Top             =   2640
         Width           =   375
      End
      Begin MSComctlLib.ListView Lv_Cmd 
         Height          =   1335
         Left            =   -70320
         TabIndex        =   67
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Nombre"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comando Asociado"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComDlg.CommonDialog CD_Cmd 
         Left            =   -70200
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_Add 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70320
         TabIndex        =   69
         ToolTipText     =   "Agrega un comando nuevo a la lista"
         Top             =   2040
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   -70440
         X2              =   -70440
         Y1              =   480
         Y2              =   4320
      End
      Begin VB.Label lblRut5 
         AutoSize        =   -1  'True
         Caption         =   "c:\"
         Height          =   195
         Left            =   -74640
         TabIndex        =   61
         Top             =   3960
         Width           =   210
      End
      Begin VB.Label lblRut4 
         AutoSize        =   -1  'True
         Caption         =   "c:\"
         Height          =   195
         Left            =   -74640
         TabIndex        =   60
         Top             =   3240
         Width           =   210
      End
      Begin VB.Label lblRut3 
         AutoSize        =   -1  'True
         Caption         =   "c:\"
         Height          =   195
         Left            =   -74640
         TabIndex        =   59
         Top             =   2520
         Width           =   210
      End
      Begin VB.Label lblRut2 
         AutoSize        =   -1  'True
         Caption         =   "c:\"
         Height          =   195
         Left            =   -74640
         TabIndex        =   58
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label lblRut1 
         AutoSize        =   -1  'True
         Caption         =   "c:\"
         Height          =   195
         Left            =   -74640
         TabIndex        =   57
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Ruta definida 5"
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
         Left            =   -74640
         TabIndex        =   56
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Ruta definida 4"
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
         Left            =   -74640
         TabIndex        =   55
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Ruta definida 3"
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
         Left            =   -74640
         TabIndex        =   54
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Ruta definida 2"
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
         Left            =   -74640
         TabIndex        =   53
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Ruta definida 1"
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
         Left            =   -74640
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLog 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         Left            =   4680
         TabIndex        =   51
         Top             =   3360
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Log:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4680
         TabIndex        =   50
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Actual"
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
         Left            =   -74760
         TabIndex        =   48
         Top             =   2520
         Width           =   1110
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   45
         Top             =   1680
         Width           =   1980
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
         Left            =   -74760
         TabIndex        =   44
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblIPactual 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0.0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74760
         TabIndex        =   42
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label lblUpdate 
         AutoSize        =   -1  'True
         Caption         =   "31/12/2099"
         Height          =   195
         Left            =   -72840
         TabIndex        =   40
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultima Actualizacion"
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   720
         Width           =   1740
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
         Left            =   -74760
         TabIndex        =   37
         Top             =   480
         Width           =   3750
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   1320
         Width           =   3765
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
         Left            =   -70200
         TabIndex        =   35
         Top             =   1440
         Width           =   2250
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
         Left            =   -70200
         TabIndex        =   34
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label16 
         Caption         =   $"FrmConfiguracion.frx":0070
         Height          =   1335
         Left            =   -70200
         TabIndex        =   33
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Line Line1 
         X1              =   4560
         X2              =   4560
         Y1              =   480
         Y2              =   3720
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   3090
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   3180
      End
      Begin VB.Label lblenv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         TabIndex        =   25
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\bd"
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
         TabIndex        =   24
         Top             =   1440
         Width           =   540
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   3330
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   3180
      End
      Begin VB.Label lblgen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         TabIndex        =   21
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label lblEns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         TabIndex        =   20
         Top             =   2760
         Width           =   585
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   3570
      End
      Begin VB.Label lblBd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         TabIndex        =   18
         Top             =   3360
         Width           =   585
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4680
         TabIndex        =   17
         Top             =   600
         Width           =   3120
      End
      Begin VB.Label lblExt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         Left            =   4680
         TabIndex        =   16
         Top             =   840
         Width           =   585
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4680
         TabIndex        =   15
         Top             =   1200
         Width           =   2880
      End
      Begin VB.Label lblArch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label lblLista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         Left            =   4680
         TabIndex        =   13
         Top             =   2160
         Width           =   585
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4680
         TabIndex        =   12
         Top             =   1920
         Width           =   2460
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4680
         TabIndex        =   11
         Top             =   2520
         Width           =   1860
      End
      Begin VB.Label lblReportes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c:\xxx"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   2760
         Width           =   585
      End
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cmd_Oid As String 'Oid del comando seleccionado



Private Sub cb_User_Click()
Opciones.OidUserDefault = Array_OidUsuario(cb_User.ListIndex + 1, 1)
lblUser.Caption = cb_User.Text
End Sub

Private Sub cmd_add_Click()

    Dim sFile As String
    
    
    
    With CD_Cmd
        .DialogTitle = "Elegir Archivo para ejecutar"
        .CancelError = False
        
        'Pendiente: establecer los indicadores y atributos del control common dialog
        .Filter = "Archivos Ejecutables (*.exe, *.bat, *.com)|*.exe; *.bat; *.com"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        Cmd.bd_RutaCmd = .FileName
    End With
    
    

    Cmd.bd_Nombre = InputBox("Nombre del Comando que se va a agregar: ", "Nombre del Comando")
    
        If Len(Cmd.bd_Nombre) = 0 Then
            MsgBox "Nombre del Comando necesario", vbCritical, "Dato Necesario"
            sFile = ""
            Exit Sub
        End If
        
    
 
    
    Call Comandos(0)
    Call Comandos(3)

End Sub

Private Sub cmd_Del_Click()



If MsgBox("Desea Eliminar el comando elegido?", vbYesNo + vbQuestion) = vbYes Then
 
'Borrar el comando
 Call Comandos(2)
 Call Comandos(3)
 
End If


End Sub

Private Sub cmd_Edit_Click()

If Len(Cmd_Oid) = 0 Then
    MsgBox "Elija adecuadamente el comando a editar", vbCritical, "Error en seleccion"
    Exit Sub
End If


Call Comandos(4)
FrmEDitcmd.Show

End Sub

Private Sub cmd_guardar_Click()
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
    
    
    
If Len(lblLog.Caption) = 0 Then
    MsgBox "Elija Ruta del Archivo de log adecuadamente", vbCritical, "Dato Necesario"
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

    
    
If Len(lblRut1.Caption) = 0 Then
  MsgBox "Escriba la Ruta default 1 adecuadamente", vbCritical, "Dato Necesario"
  Exit Sub
End If



If Len(lblRut2.Caption) = 0 Then
  MsgBox "Escriba la Ruta default 2 adecuadamente", vbCritical, "Dato Necesario"
  Exit Sub
End If



If Len(lblRut3.Caption) = 0 Then
  MsgBox "Escriba la Ruta default 3 adecuadamente", vbCritical, "Dato Necesario"
  Exit Sub
End If



If Len(lblRut4.Caption) = 0 Then
  MsgBox "Escriba la Ruta default 4 adecuadamente", vbCritical, "Dato Necesario"
  Exit Sub
End If



If Len(lblRut5.Caption) = 0 Then
  MsgBox "Escriba la Ruta default 5 adecuadamente", vbCritical, "Dato Necesario"
  Exit Sub
End If

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
        Print #1, lblLog.Caption ' log del sistema ruta
        Print #1, lblRut1.Caption ' ruta definida 1
        Print #1, lblRut2.Caption ' ruta definida 2
        Print #1, lblRut3.Caption ' ruta definida 3
        Print #1, lblRut4.Caption ' ruta definida 4
        Print #1, lblRut5.Caption ' ruta definida 5

        
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

FrmConfiguracion.lblIPactual = FrmRecepcionFile.Winsock1.LocalIP

'FrmMain!Sck_Info.Close
'FrmMain!Sck_Info.LocalPort = Val(Opciones.PuertoInfo)
'FrmMain!Sck_Info.Listen


Unload Me

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
  bi.lpszTitle = "Selecciona la Ruta de la Carpeta de Archivos Generados"

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

Private Sub cmd_RutaLog_Click()

Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblLog.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta donde se registran los eventos del sistema"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblLog.Caption = Left(path, pos - 1) '& "\Registro.mdb"

     Opciones.Rutacarpeta_Log = lblLog.Caption
     
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
  bi.lpszTitle = "Selecciona la Ruta de la Carpeta de Archivo"

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

Private Sub cmdRuta_1_Click()


  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblRut1.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta Default 1"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblRut1.Caption = Left(path, pos - 1)
     Opciones.RutaDefinida1 = lblRut1.Caption
     
  End If

  Call CoTaskMemFree(pidl)


End Sub

Private Sub cmdRuta_2_Click()

  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblRut2.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta Default 2"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblRut2.Caption = Left(path, pos - 1)
     Opciones.RutaDefinida2 = lblRut2.Caption
     
  End If

  Call CoTaskMemFree(pidl)



End Sub

Private Sub cmdRuta_3_Click()

  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblRut3.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta Default 3"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblRut3.Caption = Left(path, pos - 1)
     Opciones.RutaDefinida3 = lblRut3.Caption
     
  End If

  Call CoTaskMemFree(pidl)




End Sub

Private Sub cmdRuta_4_Click()

  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblRut4.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta Default 4"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblRut4.Caption = Left(path, pos - 1)
     Opciones.RutaDefinida4 = lblRut4.Caption
     
  End If

  Call CoTaskMemFree(pidl)


End Sub

Private Sub cmdRuta_5_Click()


  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
  lblRut5.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta Default 5"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     lblRut5.Caption = Left(path, pos - 1)
     Opciones.RutaDefinida5 = lblRut5.Caption
     
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


FrmConfiguracion.cb_User.Clear
FrmConfiguracion.lblIPactual = FrmRecepcionFile.Winsock1.LocalIP


Call GetUserName
Call Comandos(3) 'Para recuperar la informacion de loscomandos definidos


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
        Line Input #1, Opciones.Rutacarpeta_Log ' ruta del la carpeta donde se almacenan todos los eventos
        Line Input #1, Opciones.RutaDefinida1 ' ruta definida 1
        Line Input #1, Opciones.RutaDefinida2 ' ruta definida 2
        Line Input #1, Opciones.RutaDefinida3 ' ruta definida 3
        Line Input #1, Opciones.RutaDefinida4 ' ruta definida 4
        Line Input #1, Opciones.RutaDefinida5 ' ruta definida 5
        
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
        lblLog = Opciones.Rutacarpeta_Log
 
        lblUser = NombreUsuario(Opciones.OidUserDefault)
        txtPuertoRecep = Opciones.PuertoRecepcion
        txtPuertoEnv = Opciones.PuertoSalida
        
        lblRut1 = Opciones.RutaDefinida1 ' ruta definida 1
        lblRut2 = Opciones.RutaDefinida2 ' ruta definida 2
        lblRut3 = Opciones.RutaDefinida3 ' ruta definida 3
        lblRut4 = Opciones.RutaDefinida4 ' ruta definida 4
        lblRut5 = Opciones.RutaDefinida5 ' ruta definida 5

        
        
        
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




Public Sub GetUserName()

Dim C As Integer
Dim Str_dato As String
Dim Rst_LstUsuarios As ADODB.Recordset

C = 1

Str_dato = "SELECT * FROM USUARIOS WHERE STATUS='OK'"
'Rst_LstUsuarios
Set Rst_LstUsuarios = New ADODB.Recordset
    Rst_LstUsuarios.CursorLocation = adUseClient
    Rst_LstUsuarios.CursorType = adOpenDynamic
    Rst_LstUsuarios.LockType = adLockOptimistic
    Rst_LstUsuarios.Open Str_dato, CadenaCnx



ReDim Array_OidUsuario(1 To Rst_LstUsuarios.RecordCount, 3)


Rst_LstUsuarios.MoveFirst

Do While Not Rst_LstUsuarios.EOF

FrmConfiguracion.cb_User.AddItem Rst_LstUsuarios.Fields("NOMBRE_USUARIO") & Space(1) & Rst_LstUsuarios.Fields("APELLIDO_USUARIO")


Array_OidUsuario(C, 1) = Rst_LstUsuarios!OID_USUARIO
Array_OidUsuario(C, 2) = Rst_LstUsuarios!Nick
Array_OidUsuario(C, 3) = Rst_LstUsuarios.Fields("NOMBRE_USUARIO") & Space(1) & Rst_LstUsuarios.Fields("APELLIDO_USUARIO")



'Debug.Print c & "  : " & Array_OidCliente(c, 1)
'Debug.Print c & "  : " & Array_OidCliente(c, 2)



C = C + 1
Rst_LstUsuarios.MoveNext
Loop

Rst_LstUsuarios.Close

End Sub


Sub GuardarInfo()


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







Private Sub Lv_Cmd_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Cmd_Oid = Item.Tag

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





Private Sub txtPuertoRecep_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
KeyAscii = 0
Beep
End If
End Sub

