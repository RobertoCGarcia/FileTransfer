VERSION 5.00
Begin VB.Form FrmDireccion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ruta a Extraer"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7200
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5520
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Dir 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      ToolTipText     =   "Permite elegir la Ruta donde se extraera"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmd_RD5 
      Caption         =   "Ruta definida 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      MouseIcon       =   "FrmDireccion.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmDireccion.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtRuta 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   8655
   End
   Begin VB.CommandButton cmd_RD4 
      Caption         =   "Ruta definida 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      MouseIcon       =   "FrmDireccion.frx":0594
      MousePointer    =   99  'Custom
      Picture         =   "FrmDireccion.frx":06E6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmd_RD3 
      Caption         =   "Ruta definida 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      MouseIcon       =   "FrmDireccion.frx":0B28
      MousePointer    =   99  'Custom
      Picture         =   "FrmDireccion.frx":0C7A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmd_RD2 
      Caption         =   "Ruta definida 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      MouseIcon       =   "FrmDireccion.frx":10BC
      MousePointer    =   99  'Custom
      Picture         =   "FrmDireccion.frx":120E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmd_RD1 
      Caption         =   "Ruta definida 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MouseIcon       =   "FrmDireccion.frx":1650
      MousePointer    =   99  'Custom
      Picture         =   "FrmDireccion.frx":17A2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Escriba la direccion donde se extraera el archivo o elije una Ruta Definida:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Cancel_Click()
'FrmEnvio.LV_ArchivosElegidos.ListItems.Item(Item.Index).SubItems(3) = ""
Unload Me

End Sub

Private Sub cmd_Dir_Click()



  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer
    
 
 'lblenv.Caption = ""

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
  bi.lpszTitle = "Selecciona la Ruta de la Carpeta donde se Extraera el Archivo"

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS
  
 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)
 
 'the dialog has closed, so parse & display the
 'user's returned folder selection contained in pidl
  path = Space$(MAX_PATH)
    
  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     
     txtRuta.Text = Left(path, pos - 1)
     
    ' Opciones.Rutacarpeta_Enviados = lblenv.Caption
     
  End If

  Call CoTaskMemFree(pidl)


End Sub

Private Sub cmd_OK_Click()

If FrmEnvio.LV_ArchivosElegidos.ListItems.Count <= 0 Then
    MsgBox "No existe Archivo valido para asignar Ruta", vbCritical, "Archivo necesario"
    FrmEnvio.cmd_Ruta.SetFocus
    Unload Me
    Exit Sub
End If


FrmEnvio.LV_ArchivosElegidos.ListItems.Item(FrmEnvio.InxRutArch).SubItems(3) = txtRuta.Text
FrmEnvio.cmd_Ruta.SetFocus
Unload Me
End Sub

Private Sub cmd_RD1_Click()

txtRuta.Text = Opciones.RutaDefinida1



End Sub

Private Sub cmd_RD2_Click()

txtRuta.Text = Opciones.RutaDefinida2

End Sub

Private Sub cmd_RD3_Click()

txtRuta.Text = Opciones.RutaDefinida3

End Sub

Private Sub cmd_RD4_Click()

txtRuta.Text = Opciones.RutaDefinida4

End Sub

Private Sub cmd_RD5_Click()

txtRuta.Text = Opciones.RutaDefinida5

End Sub

