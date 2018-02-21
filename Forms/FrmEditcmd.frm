VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmEditcmd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editar Comando"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Rut 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Cnd_Exit 
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
      Left            =   5520
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmd_OK 
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
      Left            =   4080
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtNom 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   6735
   End
   Begin VB.OptionButton OptExe 
      Caption         =   "Normal sin Foco"
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton OptExe 
      Caption         =   "Normal con Foco"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.OptionButton OptExe 
      Caption         =   "Minimizado sin Foco"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.OptionButton OptExe 
      Caption         =   "Minimizado con Foco"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.OptionButton OptExe 
      Caption         =   "Maximizado con Foco"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.OptionButton OptExe 
      Caption         =   "Oculto"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CD_Cmd 
      Left            =   3360
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRut 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Como Ejecutarlo"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
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
      TabIndex        =   1
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "FrmEDitcmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OpcExe As Integer



Private Sub cmd_OK_Click()

Cmd.bd_Nombre = txtNom.Text

Cmd.bd_RutaCmd = lblRut.Caption

Call Comandos(1)
Call Comandos(4)
Unload FrmConfiguracion
Unload Me

End Sub

Private Sub cmd_Rut_Click()
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
       lblRut = .FileName
    End With
End Sub

Private Sub Cnd_Exit_Click()

Unload Me

End Sub

Private Sub OptExe_Click(Index As Integer)
OpcExe = Index
'MsgBox OpcExe
End Sub
