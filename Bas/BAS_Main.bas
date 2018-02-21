Attribute VB_Name = "BAS_Main"
Option Explicit
Public CadenaCnx As String
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Sub Main()

Load FrmSystemTray

Randomize
Call LeerInfoArch
Derecho = 11
CadenaCnx = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & Opciones.RutaBd & "; Persist Security Info=False"

Call InformacionUsuario(Opciones.OidUserDefault, "CONSULTA")


If Len(Dir(App.path & "\dll\RecepcionOid.oid")) <> 0 Then
    'PORQUE EXISTE SI NO NO SE BORRA NADA.
    Kill (App.path & "\dll\RecepcionOid.oid")
End If


'Borro el archivo que tiene la informacion que se recibe cuando se descomprime
'al momento de recibirlo

'If Len(Dir(Opciones.Rutacarpeta_Ensamble & "\0x00.raw")) <> 0 Then
    'PORQUE EXISTE SI NO NO SE BORRA NADA.
'    Kill (Opciones.Rutacarpeta_Ensamble & "\0x00.raw")
'End If


sendsize = 1024


Dim OSF As Object 'Objeto para manipular los archivos
Set OSF = CreateObject("scripting.filesystemobject")

Dim Carpeta As Folder
Set Carpeta = OSF.GetFolder(Opciones.Rutacarpeta_Recibidos)

' Carpeta.Files.Count

    If Carpeta.Files.Count > 0 Then
    ' se borra todo lo que hay hay
         Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"
    
    End If


FrmRecepcionFile.Winsock1.Close
FrmRecepcionFile.Winsock1.LocalPort = Val(Opciones.PuertoRecepcion)
FrmRecepcionFile.Winsock1.Listen

FrmConfiguracion.lblIPactual = FrmRecepcionFile.Winsock1.LocalIP

Call RegistraEvento("PROGRAMA", 1, "Inicio Programa", "Esperendo conexion remota, Puerto: " & Opciones.PuertoRecepcion & " en IP " & FrmRecepcionFile.Winsock1.LocalIP)


Load FrmMain
'FrmMain.lblfolio = Usuario.Salidas
 
FrmMain.lblport = "Esperando: " & FrmRecepcionFile.Winsock1.LocalPort
FrmMain.Show



End Sub








