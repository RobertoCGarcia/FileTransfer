VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   Icon            =   "FrmMin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImgLMain 
      Left            =   360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":030A
            Key             =   "envio"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":0624
            Key             =   "historial"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":093E
            Key             =   "recepcion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":0C58
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":0F72
            Key             =   "opciones"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":1E4C
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":229E
            Key             =   "reenvio"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":26F0
            Key             =   "ModoBarra"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":2B42
            Key             =   "Ocultar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":2F94
            Key             =   "quicksend"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMin.frx":33E6
            Key             =   "bitacora"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tb_Main 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   1535
      ButtonWidth     =   1826
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImgLMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Envio"
            Key             =   "envio"
            Object.ToolTipText     =   "Genera la informacion a enviar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Recepcion"
            Object.ToolTipText     =   "Muestra informacion acerca del proceso de recepcion de informacion"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Opciones"
            Object.ToolTipText     =   "Opciones del Programa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Historial"
            Object.ToolTipText     =   "Historial de todo lo que se ha recibido y se ha enviado"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Object.ToolTipText     =   "Reinicia todo en el programa"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reenvio"
            Key             =   "Reenvio de las Salidas Pendientes"
            Object.ToolTipText     =   "Da la Posibilidad de Hacer envios que Quedaron Pendientes"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ocultar"
            Key             =   "Oculta el Programa"
            Object.ToolTipText     =   "Oculta la Vista Actual"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modo Normal"
            Key             =   "Hace pequeño el programa, para ocupar menor espacio en la pantalla"
            Object.ToolTipText     =   "Modo rapido accede a las funciones principales"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "QuickSend"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bitácora"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Sale del programa"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "FrmMin.frx":3838
   End
End
Attribute VB_Name = "FrmMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Tb_Main_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index


    Case Is = 1
    ' envio
    FrmEnvio.Show
    
    Case Is = 2
    'recepcion
    FrmRecepcionFile.Show
    Case Is = 3
    ' opciones
    Derecho = 0
    FrmPassword.Show
    
    
    Case Is = 4
    
     FrmHistorial.Show

    
    Case Is = 5
    ' reset
        FrmMain.lblport = "Reset..."
        FrmRecepcionFile.Winsock1.Close
        sendsize = 1024
        FrmRecepcionFile.Winsock1.Close
        FrmRecepcionFile.Winsock1.LocalPort = Opciones.PuertoRecepcion
        FrmRecepcionFile.Winsock1.Listen
        
        FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost

        FrmMain.lblport = "Esperando: " & FrmRecepcionFile.Winsock1.LocalPort
        Call InformacionUsuario(Opciones.OidUserDefault, "CONSULTA")
        
        Call LeerInfoArch
        
        
        Call CALCULO_FECHAS(Format$(Date, "dd/mm/yyyy"))

        Call ConsultarRecepcion
        Call ConsultarEnvio
        Call ConsultarPendientes
        


    Case Is = 6
       'Opcion de Reenvio
       
       
'       FrmReenvio.Show
        FrmReenvioOK.Show

    Case Is = 7
    'ocultar
        FrmMin.Hide
    
    Case Is = 8
    'minimizar
    
        If Min = True Then
        
            FrmMain.Show
            Min = False
            FrmMin.Hide
        End If
    
    Case Is = 9
    'favoritos del sistema, envio rapido
    FrmQuickSend.Show
    
    
    
    Case Is = 10
    
    FrmBitacora.Show
    
    
    
    Case Is = 11
    

    FrmMain.SalirAll


End Select


End Sub
