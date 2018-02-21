VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProgressBar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Recibiendo Informacion"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Salir 
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
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar recieved1 
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1296
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   -240
      Top             =   2280
      Width           =   8055
   End
   Begin VB.Label lblPorcent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   540
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje de Recepcion:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   5820
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   -240
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "FrmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_Salir_Click()
Unload Me
End Sub

