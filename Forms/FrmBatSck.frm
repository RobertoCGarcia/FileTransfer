VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmBatSck 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sck1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck2 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck3 
      Left            =   1080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck4 
      Left            =   1560
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck5 
      Left            =   2040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck6 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck7 
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck8 
      Left            =   1080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck9 
      Left            =   1560
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck10 
      Left            =   2040
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck11 
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck12 
      Left            =   600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck13 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck14 
      Left            =   1560
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck15 
      Left            =   2040
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck16 
      Left            =   120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck17 
      Left            =   600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck18 
      Left            =   1080
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck19 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck20 
      Left            =   2040
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck21 
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck22 
      Left            =   600
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck23 
      Left            =   1080
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck24 
      Left            =   1560
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck25 
      Left            =   2040
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmBatSck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub sck1_Connect()


              
              sck1.SendData "FILESIZE " & FileLen(RutArchElegido)
              DoEvents
              Sleep 2000
              
               'Debug.Print "Evento connect : " & "SENDNAME " & NombreArchivoGenerado & ".dpk"

              sck1.SendData "SENDNAME " & NombreArchivoGenerado & ".dpk"

              DoEvents
              Sleep 2000
    
             ' lblstatus1 = "Enviando Archivo"
            '  FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(Index).SubItems(7) = "Enviando Archivo..."
              
'             Call UpdateFile(Opciones.Rutacarpeta_Generados & "\" & Mat_Envio(C, 1), "ENV")
              Call SendFile(Opciones.Rutacarpeta_Generados & "\" & NombreArchivoGenerado & ".dpk", 1)
     
           '  MsgBox "Terminado el Envio!!!", vbInformation, "Operacion Exitosa"
             
           '  Unload FrmInfoEnvio


             
             ' Sck_Envio(Index).SendData "ARCHIVE " & "1010101010101010101010101010101010101010000101010101010101010101001010"
             ' DoEvents
             ' Sleep 1000



   ' recieved.Value = 0
  '  byteslabel = ""
 '   thespeed = ""
'    lbltime = ""



End Sub

Private Sub sck1_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next
Dim temprecieve As String
DoEvents



sck1.GetData temprecieve
'Debug.Print "Recibido: " & temprecieve



If InStr(1, temprecieve, "FILESIZE ") <> 0 Then
    filesize(1) = Mid(temprecieve, 10, sendsize)
    
    
ElseIf InStr(1, temprecieve, "SENDNAME ") <> 0 Then
    
    thename = Mid(temprecieve, 10, sendsize)
    currentint(1) = 0
    currentint(1) = FileLen(Opciones.Rutacarpeta_Recibidos & thename)
    
    If currentint(1) >= filesize(1) Then
        sck1.SendData "ALLDONE"
        DoEvents
        Close #2
        sck1.Close
        FrmInfoEnvio.AddStat "DONE!"
        Beep
        Exit Sub
    Else
        
        currentint(1) = currentint(1) + 1
        sck1.SendData "RESUME " & currentint(1)
    End If
    Close #2
    
    Open Opciones.Rutacarpeta_Recibidos & thename For Binary Access Write As #2
    
ElseIf InStr(1, temprecieve, "THEDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    currentint(1) = currentint(1) + sendsize
    FrmInfoEnvio.recieved.Value = Int(currentint(1) / filesize(1) * 100)
    FrmInfoEnvio.AddStat "Recieved " & Int((currentint(1) / filesize(1)) * 100) & "%"
    FrmInfoEnvio.Caption = "File Transfer - " & Int((currentint(1) / filesize(1)) * 100) & "%"
    
    Put #2, , temprecieve
    sck1.SendData "SENDMORE"
    
    rate(1) = rate(1) + 1
    FrmInfoEnvio.byteslabel = "Recieved " & currentint(1) & " of " & filesize(1) & " bytes"
ElseIf InStr(1, temprecieve, "SRTDATA ") <> 0 Then
    
    temprecieve = Mid(temprecieve, 9, sendsize + 1)
    FrmInfoEnvio.recieved.Value = Int(currentint(1) / filesize(1) * 100)
    FrmInfoEnvio.AddStat "Recieved " & Int((currentint(1) / filesize(1)) * 100) & "%"
    FrmInfoEnvio.Caption = "File Transfer - " & Int((currentint(1) / filesize(1)) * 100) & "%"
    
    Put #2, currentint(1), temprecieve
    sck1.SendData "SENDMORE"
    
    rate(1) = rate(1) + 1
    FrmInfoEnvio.byteslabel = "Recieved " & currentint(1) & " of " & filesize(1) & " bytes"
    currentint(1) = currentint(1) + sendsize
ElseIf InStr(1, temprecieve, "RESUME ") <> 0 Then
    
    currentint(1) = Mid(temprecieve, 8, 20)
    filestart(1) = currentint(1)
    sendmore(1) = 1
ElseIf temprecieve = "SENDMORE" Then
    
    sendmore(1) = 1
    
ElseIf temprecieve = "ALLDONE" Then
    
    Close #2
    sck1.Close
    FrmInfoEnvio.AddStat "HECHO !"
    Beep
    
    
    
    
'ElseIf InStr(1, temprecieve, "OK ") <> 0 Then

    
End If




End Sub

Private Sub sck1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

'Debug.Print "eRROR SCK1"

MsgBox "Error en Socket: 1 : " & Description & "  " & Number, vbCritical, "Error de Transferencia"
FrmInfoEnvio.lblstatus = "Error de transferencia :  " & Description & " " & Number

EnvC = False


Mat_Envio(InxCe, 3) = Description & "  " & Number
Mat_Envio(InxCe, 4) = "0" ' error
InxCe = InxCe + 1
'EnvC = False
sck1.Close
sck1.Close

If InxCe <= UBound(Mat_Envio) Then
    ' si es menor o igual se llama ya que faltan algunos por
    'enviarse, de lo contrario sale del procedimiento
    Call EnvioProgramado

Else
    
    Exit Sub

End If





'FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(Inx).SubItems(7) = "No Enviado Error..."
'FrmMainEnvio.LV_ArchivosElegidos.ListItems.Item(Inx).SubItems(7) = "Enviando Archivo..."



End Sub



