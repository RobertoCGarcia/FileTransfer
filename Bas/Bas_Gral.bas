Attribute VB_Name = "Bas_Gral"
Option Explicit
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)



    Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
    Public Const INTERNET_RAS_INSTALLED As Long = &H10
    Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
    Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40



Type Favorito
'Estrucctura para manejar la lista de favoritos

    Bd_IDFAVORITO As String
    Bd_OIDENV As String
    Bd_NOMBREFAVORITO As String
    Bd_USUARIO_CREO As String
    Bd_COMENTARIO As String
    Bd_ID_LISTA_USUARIOS As String


End Type


Type Opciones_Grales

 Rutacarpeta_Enviados As String
 Rutacarpeta_Recibidos As String
 Rutacarpeta_Generados As String
 Rutacarpeta_Ensamble As String
 
 Rutacarpeta_Extraidos As String
 Rutacarpeta_depositoRecibidos As String
 RutaListaGenerada As String
 RutaReportes As String
 Rutacarpeta_Log As String
 
 OidUserDefault As String
 PuertoRecepcion As String
 PuertoSalida As String
 PuertoInfo As String
 RutaBd As String
 
 TipoHost As String
 NombreHost As String
 
 RutaDefinida1 As String
 RutaDefinida2 As String
 RutaDefinida3 As String
 RutaDefinida4 As String
 RutaDefinida5 As String
 
 

End Type



Type ArchivoMaestro_Datos

   Nombre As String
   CLIENTE_FUENTE As String
   USUARIO_ORIGEN As String
   OID_USUARIO_ORIGEN As String
   FOLIO_SALIDA As String
   OID_MOVIMIENTO As String
   No_Archivos As String
   USUARIO_DESTINO As String
   OID_USUARIO_DESTINO As String
   FECHA_CREACION As String
   FECHA_ENVIO As String

End Type



Type Usuario_Datos
' Son los Datos del Usuario default del sistema, entonces
' pueden existir 2 usuarios el default y el de la clave
' para solicitar permiso para realizar alguna accion sobre el
' sistema

        OidUsuario As String
        Nombre As String
        Apellido As String
        Password As String
        Nick As String
        Area As String
        Comentario As String
        fechaCreacion As String
        Entradas As Double
        Salidas As Double
        IP_equipo As String
        
        
        'Permisos del usuario
        P_ModConfig As Boolean
        P_ConInfoUsers As Boolean
        P_AddUsers As Boolean
        P_ModUsers As Boolean
        P_DelUsers As Boolean
        p_ModIP As Boolean
        Grupo As String
                
End Type



Type Usuario_Activo
' usuario que se obtiene al poner el password requerido para realizar
' ciertas acciones

        OidUsuario As String
        Nombre As String
        Apellido As String
        Password As String
        Nick As String
        Area As String
        Comentario As String
        fechaCreacion As String
        Entradas As Double
        Salidas As Double
        IP_equipo As String
        
        
        'Permisos del usuario
        P_ModConfig As Boolean
        P_ConInfoUsers As Boolean
        P_AddUsers As Boolean
        P_ModUsers As Boolean
        P_DelUsers As Boolean
        p_ModIP As Boolean
        Grupo As String


End Type


Type CmdDef

    bd_Oid_Cmd As String
    bd_RutaCmd As String
    bd_Nombre As String

End Type


Public Cmd As CmdDef 'Comando asociado a un envio
Public Min As Boolean
Public OpcHistorial As Integer
Public ArchExiste As Boolean
Public Opciones As Opciones_Grales
Public ArchivoMaestro As ArchivoMaestro_Datos
Public Usuario As Usuario_Datos
Public Favoritos As Favorito
Public UsuarioActivo As Usuario_Activo 'Usuario que se obtiene del password solicitado
Public IPUsuarios() As String
Public Mat_Envio() As String


Public Sub Extraer_Archivos(NombreArch As String)


FrmRecepcionFile.Extraer = False



Sleep 1500

Dim resp As Variant
Dim strCadena As String
Dim No_Archivos As Long
Dim C As Integer
Dim NomArchEmp As String
Dim RutaExtraer As String


Dim UsuarioOrigen As String
Dim Comentario As String
Dim Folio As String


'//Primero se recibe el archivo .dpk en la carpeta recibidos

'Despues se extrae todo su contenido en la carpeta extraidos
'despues se lee el archivo Info.rtl
'que indica a donde se tiene que poner cada archivo que contiene la carpeta extraidos
'Se copia cada archivo de la carpeta extraidos a donde indica la info
'Se borra la carpeta extraidos
'Se copia el archivo Recibido de la carpeta recibidos a la carpeta Archivo
'Se borra la carpeta Recibidos

     FrmMain.lblport = "Extrayendo Archivo..."
     FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Extrayendo Archivo..."
   
        'Ruta carpeta Extraidos sacar todos los archivos en esta carpeta
        'Ruta carpeta Recibidos
        Call ListFileContents(Opciones.Rutacarpeta_Recibidos & "\" & NombreArch)
        DoRestore Opciones.Rutacarpeta_Ensamble & "\"
        
        Sleep 4000

     FrmMain.lblport = "Contenido Extraido!!!"
     FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Contenido Extraido!!!"

    FrmMain.lblport = "Leyendo Informacion 0x00.Raw"
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Leyendo Informacion 0x00.Raw"
    
    
    Call ProcesarArchivo(NombreArch)
    
    
    FrmMain.lblport = "Lectura Terminada"
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Lectura Terminada"
    
    
    FrmMain.lblport = "Copiando Archivo(s)..."
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Copiando Archivo(s)..."
    

    Call MoverArchivos
    FrmMain.lblport = "Terminado..."
    FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Terminado..."


'genero un archivo con el Oid de la ultima recepcion
'el archivo se llama CurrentOid.Key que contiene el
'solamente esa cadena de caracteres.

FrmMain.lblport = "Registrando Entrada..."
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Registrando Entrada..."

    Call RegistrarRecepcion(1) 'Header Recepcion
    Call RegistrarRecepcion(2) 'Body de la Recepcion
    Call RegistrarRecepcion(3) 'Registros de los Archivos Recibidos
FrmMain.lblport = "Fin Registro Entrada"
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Fin Registro Entrada"


Call ArchivoOid(1)

FrmMain.lblport = "Moviendo Archivo Principal..."
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Moviendo Archivo Principal..."



Kill Opciones.Rutacarpeta_Ensamble & "\" & "*.*"
'Copiar el archivo recibido el que viene comprimido
Call FileCopy(Opciones.Rutacarpeta_Recibidos & "\" & NombreArch, Opciones.Rutacarpeta_depositoRecibidos & "\" & NombreArch)

Kill Opciones.Rutacarpeta_Recibidos & "\" & "*.*"

FrmMain.lblport = "Moviendo Terminado"
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & "  Moviendo Terminado"

'Se actualiza al sig folio para la sig entrada, despues de que ya fueron
'guardados todos los movimientos

Call InformacionUsuario(Usuario.OidUsuario, "UPDATE_ENTRADAS")
Call InformacionUsuario(Usuario.OidUsuario, "CONSULTA") ' se actualiza la entrada activa del usuario

FrmMain.lblport = "Terminado con exito"
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & " Terminado con exito"

FrmRecepcionFile.Winsock1.Close
FrmRecepcionFile.Winsock1.LocalPort = Val(Opciones.PuertoRecepcion)
FrmRecepcionFile.Winsock1.Listen
FrmRecepcionFile.AddStat "Esperando Conexion Remota en: " & Val(Opciones.PuertoRecepcion)

FrmMain.lblport = "Esperando: " & FrmRecepcionFile.Winsock1.LocalPort
FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost
FrmRecepcionFile.lblstatus = ""

FrminfoRecibida1.Show


Call ConsultarRecepcion




End Sub


Sub ProcesarArchivo(Nombre As String)


'En este sub se lee la informacion que llega adentro del archivo 0x00.raw
'1ero se descomprime todo en ensamble y despues se lee todo en este sub
'despues se copia todo a la ruta que se indique

Dim resp As Variant
Dim strCadena As String
Dim No_Archivos As Long
Dim C As Integer
Dim NomArchEmp As String
Dim RutaExtraer As String
Dim RutaOrigen As String
Dim Tamaño As String



Dim UsuarioOrigen As String
Dim Comentario As String
Dim Folio As String
 'leer el archivo de 0x00.Raw para sacar el numero de archivos

    strCadena = String(255, 0)
'1    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "USUARIO_ORIGEN", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    UsuarioOrigen = strCadena
    Recepcion.bd_UsuarioOrigen = UsuarioOrigen
    strCadena = ""
            



'2
    
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "No_ARCHIVOS", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    No_Archivos = Val(strCadena)
    Recepcion.bd_NoArchivos = No_Archivos
    strCadena = ""
            
                 
    ' usuario de origen
    ' UsuarioOrigen
    
            
            
'3    ' comentario agregado al archivo
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "COMENTARIO", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    Comentario = strCadena
    Recepcion.Bd_COMENTARIO = Comentario
    strCadena = ""


'4    'folio de salida dentro del archivo
    
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "FOLIO_SALIDA", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    Folio = strCadena
    Recepcion.bd_FolioSalida = Folio
    strCadena = ""


'5    'Oid del usuario que envio este archivo
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "OID_USUARIO_ORIGEN", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    Recepcion.bd_UID_Origen = strCadena
    strCadena = ""


    'Oid del movimiento que se genero en la bd de envio para llevar un  control mejor
'6    'tener el mismo oid en las dos bd la del receptor y la del emisor
    
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "OID_MOVIMIENTO", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    Recepcion.bd_OidRecepcion = strCadena
    strCadena = ""
 
'7    'FECHA_CREACION
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "FECHA_CREACION", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    Recepcion.bd_FechaCreacionArchivoDpk = strCadena
    strCadena = ""
 
'100    Comando Asociado a ejecutar cuando se recibe el archivo
    strCadena = String(255, 0)
    resp = GetPrivateProfileString("ENCABEZADO_ARCHIVO", "CMD", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
    If resp <> 0 Then strCadena = Left$(strCadena, resp)
    Recepcion.bd_CmdEjecutar = strCadena
    
    'MsgBox "Comando asociado: " & Recepcion.bd_CmdEjecutar
    
    
    strCadena = ""
 
 
 '8
    Recepcion.bd_NombreArchivoDpk = Nombre



'1 col: Nombre archivo
'2 col: Ruta a extraer
'3 col: Ruta Origen
'4 col: Tamaño del archivo
'Recepcion.bd_InfoArchivoMatriz(c, 1)

'Debug.Print "Inx 1 Nombre archivo 1 Ruta a extraer 1 Ruta Origen 1 Tamaño del archivo 1"

ReDim Recepcion.bd_InfoArchivoMatriz(Recepcion.bd_NoArchivos, 4)


    For C = 1 To Recepcion.bd_NoArchivos
    
                FrmMain.lblport = "Leyendo Informacion Archivo: " & C & "/" & No_Archivos
'1
                strCadena = String(255, 0)
                resp = GetPrivateProfileString("INFORMACION_ARCHIVO_" & C, "NOMBRE_ARCHIVO_" & C, "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
                If resp <> 0 Then strCadena = Left$(strCadena, resp)
                NomArchEmp = Trim(strCadena)
                Recepcion.bd_NombreArchivo = NomArchEmp
                strCadena = ""
                Recepcion.bd_InfoArchivoMatriz(C, 1) = Recepcion.bd_NombreArchivo
                
                
'2
                strCadena = String(255, 0)
                resp = GetPrivateProfileString("INFORMACION_ARCHIVO_" & C, "EXTRAER", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
                If resp <> 0 Then strCadena = Left$(strCadena, resp)
                RutaExtraer = Trim(strCadena)
                
                'Verificar que la ruta exista
                     
                     
                     If Len(Dir(RutaExtraer, vbDirectory)) > 0 Then
                                  
                              If RutaExtraer = "999" Then ' si se cumple no se agrego ninguna ruta y se deja en la carpeta extraidos
                                       RutaExtraer = Opciones.Rutacarpeta_Extraidos & "\"
                              Else
                                       RutaExtraer = RutaExtraer & "\"
                              End If
                     Else
                               RutaExtraer = Opciones.Rutacarpeta_Extraidos & "\"
                     End If
                     
                     
                     Recepcion.bd_RutaExtraer = RutaExtraer
                     
                     Recepcion.bd_InfoArchivoMatriz(C, 2) = Recepcion.bd_RutaExtraer
                                                        
'3
                strCadena = String(255, 0)
                resp = GetPrivateProfileString("INFORMACION_ARCHIVO_" & C, "RUTA_ORIGEN", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
                If resp <> 0 Then strCadena = Left$(strCadena, resp)
                RutaOrigen = Trim(strCadena)
                Recepcion.bd_InfoArchivoMatriz(C, 3) = RutaOrigen
                                                                                
'4
                strCadena = String(255, 0)
                resp = GetPrivateProfileString("INFORMACION_ARCHIVO_" & C, "TAMAÑO", "Default", strCadena, 255, Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
                If resp <> 0 Then strCadena = Left$(strCadena, resp)
                Tamaño = Trim(strCadena)
                Recepcion.bd_InfoArchivoMatriz(C, 4) = Tamaño
                                         

                                         
                
               ' Debug.Print C & " 1 " & Recepcion.bd_InfoArchivoMatriz(C, 1) & " 1 " & Recepcion.bd_InfoArchivoMatriz(C, 2) & " 1 " & Recepcion.bd_InfoArchivoMatriz(C, 3) & " 1 " & Recepcion.bd_InfoArchivoMatriz(C, 4)
                                         
     'se registran los archivos recibidos de modo diferente
                 
                   
                 
    
                 ' Call RegistrarRecepcion(3) ' Registrar los archivos que se recibieron
   
     Next C




' desde aqui se invoca al comando asociado si existe, si se definio adecuadamente, si se agrego
' esa opcion, si no no se hace nada y solo se extrae a la ruta definida


 If Recepcion.bd_CmdEjecutar <> "NO" Then
 ' se definio algun comando, ahora buscarlo y ejecutarlo
 
 Call Comandos(6)
 
 
 End If
 
 
 
 






End Sub


Sub MoverArchivos()

Dim C As Integer

'1 col: Nombre archivo
'2 col: Ruta a extraer
'3 col: Ruta Origen
'4 col: Tamaño del archivo
'Recepcion.bd_InfoArchivoMatriz(c, 1)

    For C = 1 To Recepcion.bd_NoArchivos
    
    FileCopy Opciones.Rutacarpeta_Ensamble & "\" & Recepcion.bd_InfoArchivoMatriz(C, 1), Recepcion.bd_InfoArchivoMatriz(C, 2) & Recepcion.bd_InfoArchivoMatriz(C, 1)

    Next C
    
End Sub




Sub UpdateCaption(txt As String)

FrmMin.Caption = "Tipo Nodo: " & Opciones.TipoHost & " Nombre: " & Opciones.NombreHost & " " & txt


End Sub
