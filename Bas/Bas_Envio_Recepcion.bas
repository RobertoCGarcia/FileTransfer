Attribute VB_Name = "Bas_Envio"
Option Explicit

Public C As Long ' Es el indice dentro de la matriz para saber dopnde se encuentra
'Public TerminoProceso As Boolean

Dim NumRep As Long

Public Port As Integer ' temporal
Public InxCe As Long 'indice de la cola de envio
Public EnvC As Boolean 'Envio correcto?





'Public EnvC As Boolean
Public mainbuffer As String '''del
Public sendsize  As Integer ''tamaño de envio
Public sendmore(1 To 5) As Integer
Public thename As String ' no se ocupa en el envio
Public filesize(1 To 5) As Long ' Termino del archivo a enviar
Public currentint(1 To 5) As Long ' porcentaje de envio del archivo dentro de cada socket
Public rate(1 To 5) As Integer
Public filestart(1 To 5) As Long ' donde inicia el archivo




Public Sub UpdateFile(RutaArch As String, Opc As String)

Dim strTemp As String

strTemp = Now

Select Case Opc

Case Is = "ENV"

WritePrivateProfileString "ENCABEZADO_ARCHIVO", "FECHA_ENVIO", strTemp, RutaArch

Case Is = "RCP"

WritePrivateProfileString "ENCABEZADO_ARCHIVO", "FECHA_RECEPCION", strTemp, RutaArch


End Select

End Sub





