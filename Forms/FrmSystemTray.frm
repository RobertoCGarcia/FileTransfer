VERSION 5.00
Begin VB.Form FrmSystemTray 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   6165
   Icon            =   "FrmSystemTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Envio 
         Caption         =   "Envio"
      End
      Begin VB.Menu QuickSend 
         Caption         =   "QuickSend"
      End
      Begin VB.Menu Recepcion 
         Caption         =   "Recepcion"
      End
      Begin VB.Menu Reenvio 
         Caption         =   "Reenvio"
      End
      Begin VB.Menu Historial 
         Caption         =   "Historial"
      End
      Begin VB.Menu ui 
         Caption         =   "-"
      End
      Begin VB.Menu Modo_Barra 
         Caption         =   "Modo Barra"
      End
      Begin VB.Menu Opciones 
         Caption         =   "Opciones"
      End
      Begin VB.Menu Reset 
         Caption         =   "Reset"
      End
      Begin VB.Menu Bitácora 
         Caption         =   "Bitácora"
      End
   End
End
Attribute VB_Name = "FrmSystemTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents sysTray As SystemTray.Application
Attribute sysTray.VB_VarHelpID = -1

Private Sub Bitácora_Click()

FrmBitacora.Show

End Sub

Private Sub Envio_Click()

 FrmEnvio.Show

End Sub

Private Sub Form_Load()

    Set sysTray = New SystemTray.Application
    Set sysTray = CreateObject("SystemTray.Application")
    Call sysTray.CreateIcon(App.path & "\img\connect.ico", "Parabolick")

End Sub

Private Sub Form_Unload(Cancel As Integer)


Call sysTray.DeleteIcon
Set sysTray = Nothing
End Sub

Private Sub Historial_Click()
     OpcHistorial = 0
     FrmHistorial.Show
End Sub

Private Sub Modo_Barra_Click()

FrmMain.ModoBarra

    
End Sub

Private Sub Opciones_Click()
    ' opciones
    Derecho = 0
    FrmPassword.Show
End Sub

Private Sub QuickSend_Click()

FrmQuickSend.Show

End Sub

Private Sub Recepcion_Click()

FrmRecepcionFile.Show

End Sub

Private Sub Reenvio_Click()
        FrmReenvioOK.Show
End Sub

Private Sub Reset_Click()

   FrmMain.Resetall

End Sub



Private Sub sysTray_ButtonDown(ByVal Button As Integer)
'
    Debug.Print "sysTray_ButtonDown  " & Button '2
    If Button = 2 Then
    PopupMenu Menu
    Else
         
    FrmMin.Hide
    FrmMain.Show
     
         
    End If
    
    
End Sub


