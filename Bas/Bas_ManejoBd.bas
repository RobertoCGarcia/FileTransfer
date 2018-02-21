Attribute VB_Name = "Bas_manejoBd"
Option Explicit

Public Array_OidUsuario() As String
Public FECHA_INICIAL As Date
Public FECHA_FINAL As Date

Type RegistroEnvio

 
    bd_UIDorigen As String
    bd_Nombre As String
    bd_IP As String
    bd_PING As String
    Bd_COMENTARIO As String
    bd_PUERTO As String
    bd_UID As String
    bd_Ruta_Arch As String
    OidEnv As String
    bd_FECHA_CREACION As String
    bd_Cmd As String
    
    
    bd_ComentarioMain As String
    bd_Envio As String
    bd_UsuarioDestino As String
    
    bd_NoDestinos As Long 'Numero de Usuarios que se les envia la informacion
    bd_NoArchivos As Long 'Numero de archivos que van empaquetados en el archivo *.dpk


End Type



Type RegistroRecepcion

'//////////////////////////////Header
    bd_OidRecepcion As String
    bd_NombreArchivoDpk As String
    bd_UsuarioOrigen As String
    Bd_COMENTARIO As String
    bd_NoArchivos As String
    bd_FolioSalida As String
    bd_FechaCreacionArchivoDpk As String
    bd_CmdEjecutar As String 'Comando a ejecutar cuando se recibe el archivo
'//////////////////////////////Header

   
'//////////////////////////////Nombre del Archivo que se va a descomprimir
   bd_RemoteHost As String
   bd_RemotePort As String
   bd_RemoteIp As String
   bd_StatusRecepcion As String ' 1 exitiso  0 it'was a mistake
   bd_ComentarioRecepcion As String
   bd_Peticion_Inicio As String
   bd_Peticion_Final As String ' fin de la transmicion se cierra el socket
   bd_UID_Origen As String 'oid del archivo que envia la informacion
   
   bd_InfoArchivoMatriz() As String 'Contiene informacion de cada
   
   
   '///ARCHIVOS RECIBIDOS
   
   bd_NombreArchivo As String
   bd_RutaExtraer As String
  




End Type


Public Derecho As Integer
Public Recepcion As RegistroRecepcion
Public Envio As RegistroEnvio




Public Function NombreUsuario(Oid As String) As String


Dim Str_Usuario As String
Dim LclRst_Uservalid As ADODB.Recordset

Str_Usuario = "SELECT * FROM USUARIOS WHERE Oid_usuario='" & Oid & "' AND STATUS='OK';"

Set LclRst_Uservalid = New ADODB.Recordset
    LclRst_Uservalid.CursorLocation = adUseClient
    LclRst_Uservalid.CursorType = adOpenDynamic
    LclRst_Uservalid.LockType = adLockPessimistic
    LclRst_Uservalid.Open Str_Usuario, CadenaCnx

NombreUsuario = LclRst_Uservalid.Fields("NOMBRE_USUARIO") & Space(1) & LclRst_Uservalid.Fields("APELLIDO_USUARIO")
LclRst_Uservalid.Close



End Function




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

FrmOpciones!cb_User.AddItem Rst_LstUsuarios.Fields("NOMBRE_USUARIO") & Space(1) & Rst_LstUsuarios.Fields("APELLIDO_USUARIO")


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



Public Sub InformacionUsuario(Oid As String, Accion As String)


Dim Str_Usuario As String
Dim LclRst_Uservalid As ADODB.Recordset

Str_Usuario = "SELECT * FROM USUARIOS WHERE Oid_usuario='" & Oid & "' AND STATUS='OK';"

Set LclRst_Uservalid = New ADODB.Recordset
    LclRst_Uservalid.CursorLocation = adUseClient
    LclRst_Uservalid.CursorType = adOpenDynamic
    LclRst_Uservalid.LockType = adLockPessimistic
    LclRst_Uservalid.Open Str_Usuario, CadenaCnx


        Select Case Accion
        
        
            Case Is = "CONSULTA"
            
                        Usuario.OidUsuario = LclRst_Uservalid.Fields("OID_USUARIO")
                        Usuario.Nombre = LclRst_Uservalid.Fields("NOMBRE_USUARIO")
                        Usuario.Apellido = LclRst_Uservalid.Fields("APELLIDO_USUARIO")
                        Usuario.Area = LclRst_Uservalid.Fields("OID_AREA")
                        Usuario.fechaCreacion = LclRst_Uservalid.Fields("FECHA_CREACION")
                        Usuario.Nick = LclRst_Uservalid.Fields("NICK")
                        Usuario.Password = LclRst_Uservalid.Fields("PASSWORD")
                        Usuario.Entradas = Val(LclRst_Uservalid.Fields("ENTRADAS"))
                        Usuario.Salidas = Val(LclRst_Uservalid.Fields("SALIDAS"))
                        Usuario.Comentario = LclRst_Uservalid.Fields("COMENTARIO")
                        Usuario.IP_equipo = LclRst_Uservalid.Fields("ip_equipo")
                        
                        'Permisos del usuario actual o seleccionado
                        
                        Usuario.P_ModConfig = LclRst_Uservalid.Fields("PERMISO_MODCONFIG")
                        Usuario.P_ConInfoUsers = LclRst_Uservalid.Fields("PERMISO_CONINFOUSER")
                        Usuario.P_AddUsers = LclRst_Uservalid.Fields("PERMISO_ADDUSERS")
                        Usuario.P_ModUsers = LclRst_Uservalid.Fields("PERMISO_MODUSERS")
                        Usuario.P_DelUsers = LclRst_Uservalid.Fields("PERMISO_DELUSERS")
                        Usuario.p_ModIP = LclRst_Uservalid.Fields("PERMISO_MODIP")
                        Usuario.Grupo = LclRst_Uservalid.Fields("GRUPO")
                                               
                        
                        LclRst_Uservalid.Close
                        FrmMain.lbsal = Usuario.Salidas
                        FrmMain.lblEnt = Usuario.Entradas
                        
                        
                        
                        
                        
                        
                        Set LclRst_Uservalid = Nothing
                        
                        Exit Sub
                        
            Case Is = "UPDATE_SALIDAS"
                        ' SE GUARDA EL REGISTRO SIGUIENTE
                        
                        With LclRst_Uservalid
                            
                              LclRst_Uservalid.Fields("SALIDAS") = (Usuario.Salidas + 1)
                              LclRst_Uservalid.Fields("FECHA_ACTUALIZACION") = Now
                             .Update
                             .Close
                             FrmMain.lbsal = Usuario.Salidas
                        End With
                        
                        Set LclRst_Uservalid = Nothing
                        
                        Exit Sub
                         
            
            Case Is = "UPDATE_ENTRADAS"
            
                
                        With LclRst_Uservalid
                            
                           '  LclRst_Uservalid.Fields ("OID_USUARIO")
                           '  LclRst_Uservalid.Fields ("NOMBRE_USUARIO")
                           '  LclRst_Uservalid.Fields ("APELLIDO_USUARIO")
                           '  LclRst_Uservalid.Fields ("OID_AREA")
                            ' LclRst_Uservalid.Fields ("FECHA_CREACION")
                            ' LclRst_Uservalid.Fields ("NICK")
                            ' LclRst_Uservalid.Fields ("PASSWORD")
                              'Usuario.Salidas = (Usuario.Entradas + 1)
                              LclRst_Uservalid.Fields("ENTRADAS") = (Usuario.Entradas + 1)
                              ' LclRst_Uservalid.Fields("SALIDAS") = Usuario.Salidas
                              LclRst_Uservalid.Fields("FECHA_ACTUALIZACION") = Now
                                FrmMain.lblEnt = Usuario.Entradas
                             .Update
                             '.Requery
                             .Close
                             
                             Set LclRst_Uservalid = Nothing
                              Exit Sub
                        End With
            
            
            
         End Select
         
         
     
End Sub

Public Sub LeerInfoArch()


Open App.path & "\dll\Opc.Rlj" For Input As #1


'Call ArchivoOid
'Recepcion.bd_OidRecepcion

        Line Input #1, Opciones.Rutacarpeta_Enviados
        Line Input #1, Opciones.Rutacarpeta_Recibidos
        Line Input #1, Opciones.Rutacarpeta_Generados
        Line Input #1, Opciones.Rutacarpeta_Ensamble
        Line Input #1, Opciones.OidUserDefault
        Line Input #1, Opciones.PuertoRecepcion
        Line Input #1, Opciones.PuertoSalida
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
     '   Line Input #1, Opciones.PuertoInfo



Close (1)


End Sub


Public Sub RegistrarRecepcion(Opc As Integer)



Dim strSql As String
Dim LclRst As ADODB.Recordset
Dim C As Integer



Select Case Opc

    Case Is = 1 'agregar registro al header
    
        strSql = "SELECT * from HEADER_RECEPCION;"
        Set LclRst = New ADODB.Recordset
            LclRst.CursorLocation = adUseClient
            LclRst.CursorType = adOpenDynamic
            LclRst.LockType = adLockPessimistic
            LclRst.Open strSql, CadenaCnx
        
        
            With LclRst
                
                .AddNew
                
                    !OIDRECEP = Recepcion.bd_OidRecepcion
                    !USUARIO_ORIGEN = Recepcion.bd_UsuarioOrigen
                    !FECHA_RECEPCION = Now
                    !FECHA_RECEPCION_ORIGEN = Recepcion.bd_FechaCreacionArchivoDpk
                    !Comentario = Recepcion.Bd_COMENTARIO
                    !UID_USUARIO_REMOTO = Recepcion.bd_UID_Origen
                    !NOMBRE_ARCHIVO_COMPRIMIDO = Recepcion.bd_NombreArchivoDpk
                    !No_Archivos = Recepcion.bd_NoArchivos
                    !Tamaño = FileLen(Opciones.Rutacarpeta_Recibidos & "\" & Recepcion.bd_NombreArchivoDpk)
                    !UBICACION = Opciones.Rutacarpeta_Recibidos
                    !FOLIO_RECEPCION = Usuario.Entradas
                    !FOLIO_ORIGEN = Recepcion.bd_FolioSalida
                    !FECHA_CREACION = Now
                    !USUARIO_CREACION = Usuario.Nick
                    !REVISADO = False
                    !FECHA_REVISADO = "18/04/84"
                
                .Update
                .Close
            End With
    
             Exit Sub
    
    Case Is = 2 'agregar registro al body
    
        strSql = "SELECT * FROM BODYRECEPCION;"
        Set LclRst = New ADODB.Recordset
            LclRst.CursorLocation = adUseClient
            LclRst.CursorType = adOpenDynamic
            LclRst.LockType = adLockPessimistic
            LclRst.Open strSql, CadenaCnx
        
          With LclRst
            
            .AddNew
                !OIDRECEP = Recepcion.bd_OidRecepcion
                !IP_REMOTA = Recepcion.bd_RemoteIp
                !HOSTNAME_REMOTO = IIf(Len(Recepcion.bd_RemoteHost) = 0, "S/N", Recepcion.bd_RemoteHost)
                !COMENTARIO_RECEPCION = Recepcion.bd_ComentarioRecepcion  '''
                !PETICION_INICIO = Recepcion.bd_Peticion_Inicio ' INICIO DE LA CNX REMOTA PARA RECIBIR EL ARCHIVO
                !PETICION_FINAL = Recepcion.bd_Peticion_Final ' CUANDO SE TERMINO DE ENVIAR EL ARCHIVO Y SE CIERRA LA CNX
                !USUARIO_ORIGEN = Recepcion.bd_UsuarioOrigen
                !RECEPCION_RESULTADO = Recepcion.bd_StatusRecepcion
                !puerto = Recepcion.bd_RemotePort
                !UID_ORIGEN = Recepcion.bd_UID_Origen ' NOMBRE ARCHIVO
                !FECHA_RECEPCION = Now ''''
                !FECHA_CREACION = Now
                !USUARIO_CREO = Usuario.Nick

            .Update
            .Close
        End With
    
            Exit Sub
    
    Case Is = 3 'agregar registro de los archivos empaquetados que se han enviado



        strSql = "SELECT * FROM RECIBIDOS;"
        Set LclRst = New ADODB.Recordset
            LclRst.CursorLocation = adUseClient
            LclRst.CursorType = adOpenDynamic
            LclRst.LockType = adLockPessimistic
            LclRst.Open strSql, CadenaCnx
        
      
           
        With LclRst
           
            
            For C = 1 To Recepcion.bd_NoArchivos
            
                        
            '1 col: Nombre archivo
            '2 col: Ruta a extraer
            '3 col: Ruta Origen
            '4 col: Tamaño del archivo
            'Recepcion.bd_InfoArchivoMatriz(c, 1)
             .AddNew
            !OIDRECEP = Recepcion.bd_OidRecepcion
            !NOMBRE_ARCHIVO = Recepcion.bd_InfoArchivoMatriz(C, 1)
            !Tamaño = FileLen(Opciones.Rutacarpeta_Extraidos & "/" & Recepcion.bd_InfoArchivoMatriz(C, 1))
            !RUTA_DESTINO = Recepcion.bd_InfoArchivoMatriz(C, 2)
            !RUTA_ORIGEN = Recepcion.bd_InfoArchivoMatriz(C, 3)
            !FECHA_CREACION = Now
            !USUARIO_RECIBIO = Usuario.Nick
            
            
            .Update
            
            Next C
            
            .Close
            
            

        End With

          Exit Sub

End Select







End Sub




Public Sub ConsultarRecepcion()

FrmMain.LV_Recibido.ListItems.Clear

Dim strSql As String
Dim LclRst As ADODB.Recordset
'Dim LclRstRecep As ADODB.Recordset ' Rst de la recepcion
Dim Dia As String
Dim mes As String
Dim año As String
Dim Fecha As String
Dim elementoX As ListItem


strSql = "SELECT * From HEADER_RECEPCION WHERE (((FECHA_RECEPCION) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((REVISADO)=False));"


Set LclRst = New ADODB.Recordset
    LclRst.CursorLocation = adUseClient
    LclRst.CursorType = adOpenDynamic
    LclRst.LockType = adLockPessimistic
    LclRst.Open strSql, CadenaCnx

    With LclRst
        

     If .RecordCount > 0 Then
     ' hay registros
     FrmMain.img_BR.Picture = LoadPicture(App.path & "\img\1.ico")
     FrmMain.lblBRecep = "Hay recibidos: " & .RecordCount
     FrmMain.img_BR.Tag = "1"

             
                    With LclRst
                       FrmMain.LV_Recibido.ListItems.Clear
                             
                        Do While Not .EOF
                            
                            Set elementoX = FrmMain.LV_Recibido.ListItems.Add(, , !USUARIO_ORIGEN)
                            elementoX.Tag = !OIDRECEP
                            elementoX.SubItems(1) = !FECHA_RECEPCION ' fecha movimiento
                            elementoX.SubItems(2) = !FOLIO_RECEPCION ' folio recepcion
                            elementoX.SubItems(3) = !FOLIO_ORIGEN ' folio origen
                            
                            
                            .MoveNext
                        Loop
                            
                       .Close
                    Exit Sub
                    End With

     Else
     'No hay registros
     
     FrmMain.img_BR.Picture = LoadPicture(App.path & "\img\0.ico")
     FrmMain.lblBRecep = "No hay recibidos"
     FrmMain.img_BR.Tag = "0"
     FrmMain.LV_Recibido.ListItems.Clear
     
     End If
     
    .Close
    
    
    
'    Exit Sub
    
    End With

    Set LclRst = Nothing


End Sub



Public Sub CALCULO_FECHAS(FECHA_ELEJIDA As Date)


Dim Fecha As String
Dim FechaTmp As String

Dim suma As Double

Dim Dia As String
Dim mes As String
Dim año As String


'MsgBox "fecha normal    " & FECHA_ELEJIDA
'MsgBox "fecha + 1 en dia:     " &

Dia = Mid$(CStr(FECHA_ELEJIDA), 1, 2)
Dia = CDbl(Dia) + 1
'MsgBox "dia + 1:     " & Dia

'mes = Mid$(CStr(FECHA_ELEJIDA), 4, 2)
mes = Month(FECHA_ELEJIDA)
'dia = CDbl(dia) + 1
'MsgBox mes

'año = Mid$(CStr(FECHA_ELEJIDA), 7, 2)
año = Year(FECHA_ELEJIDA)
'MsgBox año

If Dia >= 31 Then
    mes = (CDbl(mes) + 1)
    Dia = "01"



End If



If mes > 12 Then
    año = (CDbl(año) + 1)
    mes = "01"
    Dia = "01"
End If


Fecha = mes & "/" & Dia & "/" & año


'MsgBox Format(suma, "00")
'MsgBox CDate(fecha)

'dia mes año

'asigno fechas
FECHA_INICIAL = CDate(FECHA_ELEJIDA)
FECHA_FINAL = CDate(Fecha)
'FechaTmp = Format("dd/mm/yyyy", Fecha)

'MsgBox "Inicial: " & FECHA_INICIAL & "   Final:  " & FECHA_FINAL

End Sub





Public Sub RegistrarEnvio(Opc As Integer, Optional Favorito As Boolean)


Dim strSql As String
Dim LclRst As ADODB.Recordset
Dim C As Integer



Select Case Opc

    Case Is = 1 'agregar registro al header
    
        strSql = "SELECT * FROM HEADERENVIO;"
        Set LclRst = New ADODB.Recordset
            LclRst.CursorLocation = adUseClient
            LclRst.CursorType = adOpenDynamic
            LclRst.LockType = adLockPessimistic
            LclRst.Open strSql, CadenaCnx
        
        With LclRst
            
            .AddNew
            
                !OidEnv = Envio.OidEnv
                '!FECHA_ENVIO = "18/04/84"
                !STATUS_ENVIO = 0
                !Comentario = Envio.bd_ComentarioMain
                !NOMBRE_ARCHIVO_COMPRIMIDO = Envio.bd_Nombre
                !NOMBRE_CMD = IIf(Len(Envio.bd_Cmd) = 0, "NOTHING", Envio.bd_Cmd)
                
                !No_Archivos = Envio.bd_NoArchivos
                !NO_ENVIOS = Envio.bd_NoDestinos
                
                If Favorito = True Then
                !Tamaño = 0
                !Favorito = True ' indica la opcion de que si es un favorito
                Else
                !Tamaño = FileLen(Envio.bd_Ruta_Arch)
                !Favorito = False
                End If
                
                
                
                !UBICACION = Envio.bd_Ruta_Arch
                !FOLIO_SALIDA = Usuario.Salidas
                !USUARIO_CREACION = Envio.bd_UIDorigen
                !REVISADO = False
                !FECHA_REVISADO = "18/04/84"
                !FECHA_REGISTRO = Envio.bd_FECHA_CREACION 'fecha en que se creo el archivo a enviar
                !USUARIO_CREO = Usuario.Nombre & " " & Usuario.Apellido
'                !INICIO_TRANSMICION = ""
'                !FIN_TRANSMICION = " "

                 'seccion del favorito
                  !IDFAVORITO = Favoritos.Bd_IDFAVORITO
                  !NOMBRE_FAVORITO = Favoritos.Bd_NOMBREFAVORITO
                 
                 
                 
            .Update
            .Close
        End With
    
               Set LclRst = Nothing
    
    Case Is = 2 'agregar registro al body
    
        strSql = "SELECT * FROM BODYENVIO;"
        Set LclRst = New ADODB.Recordset
            LclRst.CursorLocation = adUseClient
            LclRst.CursorType = adOpenDynamic
            LclRst.LockType = adLockPessimistic
            LclRst.Open strSql, CadenaCnx
        
        With LclRst
            
            .AddNew
            
                !OidEnv = Envio.OidEnv
                !RESULTADOPING = Envio.bd_PING
                !FOLIO_SALIDA = Usuario.Salidas
                !ip = Envio.bd_IP
                !Comentario = Envio.Bd_COMENTARIO 'comentario de cada envio
                !Envio = Envio.bd_Envio
                !puerto = Envio.bd_PUERTO
                !uid_destino = Envio.bd_UID
                '!FECHA_ENVIO = "18/04/84" 'fecha en que se envia,en que se hace pasar por el socket hacia el otro usuario
                !USUARIO_DESTINO = Envio.bd_UsuarioDestino
                !FECHA_REGISTRO = Now
                !USUARIO_CREO = Usuario.OidUsuario
         
            
            .Update
            
            .Close
            
        End With
    
             Set LclRst = Nothing
    
    Case Is = 3 'agregar registro de los archivos empaquetados que se han enviado



        strSql = "SELECT * FROM ENVIADOS;"
        Set LclRst = New ADODB.Recordset
            LclRst.CursorLocation = adUseClient
            LclRst.CursorType = adOpenDynamic
            LclRst.LockType = adLockPessimistic
            LclRst.Open strSql, CadenaCnx
        
       
        With LclRst
            
                
                         
                For C = 1 To FrmEnvio.LV_ArchivosElegidos.ListItems.Count
                         
                         .AddNew
                             'error en el paso de registro de ruta origen
                             !OidEnv = Envio.OidEnv
                             !FOLIO_SALIDA = Usuario.Salidas
                             !NOMBRE_ARCHIVO = FrmEnvio.LV_ArchivosElegidos.ListItems.Item(C).Text
                             !Tamaño = FileLen(FrmEnvio.LV_ArchivosElegidos.ListItems.Item(C).SubItems(2))
                             !RUTA_ORIGEN = FrmEnvio.LV_ArchivosElegidos.ListItems.Item(C).SubItems(2)
                             !RUTA_DESTINO = IIf(Len(FrmEnvio.LV_ArchivosElegidos.ListItems.Item(C).SubItems(3)) <> 0, FrmEnvio.LV_ArchivosElegidos.ListItems.Item(C).SubItems(3), "Ruta Default")
                             !FECHA_CREACION_ARCHIVO = Now
                             !USUARIO_CREO = Usuario.OidUsuario
                             !FECHA_REGISTRO = Now
                         
                         .Update
                Next C
                        .Close


        End With

        Set LclRst = Nothing

End Select





End Sub





Public Sub ConsultarEnvio()


Dim strSql As String
Dim LclRst As ADODB.Recordset
Dim LclRstEnv As ADODB.Recordset
Dim Dia As String
Dim mes As String
Dim año As String
Dim Fecha As String
Dim elementoX As ListItem


FrmMain.LV_Enviados.ListItems.Clear

strSql = "SELECT * From HEADERENVIO WHERE (((FECHA_REGISTRO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=1));"


Set LclRst = New ADODB.Recordset
    LclRst.CursorLocation = adUseClient
    LclRst.CursorType = adOpenDynamic
    LclRst.LockType = adLockPessimistic
    LclRst.Open strSql, CadenaCnx


    With LclRst
        

     If .RecordCount > 0 Then
             ' hay registros
             FrmMain.img_Env.Picture = LoadPicture(App.path & "\img\1.ico")
             FrmMain.lblenv = "Hay enviados: " & .RecordCount
             FrmMain.img_Env.Tag = "1"
             'FrmMain.MouseIcon = LoadPicture(App.path & "\img\SELECT.cur")
             ' SE CARGA LA LISTA DE USUARIOS A LOS QUE SE LES HA ENVIADO INFORMACION EN EL DIA ACTUAL
             
                    strSql = "SELECT * From BODYENVIO WHERE (((FECHA_REGISTRO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((ENVIO)=1));"
            
                    Set LclRstEnv = New ADODB.Recordset
                    LclRstEnv.CursorLocation = adUseClient
                    LclRstEnv.CursorType = adOpenDynamic
                    LclRstEnv.LockType = adLockPessimistic
                    LclRstEnv.Open strSql, CadenaCnx
                    
                    
'Los archivos pendientes de envio se pueden clasificar en 3 deacuerdo al valor del campo
'Envio
                    
                    With LclRstEnv
                        
                        FrmMain.LV_Enviados.ListItems.Clear
                        
                        Do While Not .EOF
                            
                            

                            Set elementoX = FrmMain.LV_Enviados.ListItems.Add(, , !USUARIO_DESTINO)
                            elementoX.Tag = "er"
                            elementoX.SubItems(1) = !FECHA_ENVIO ' fecha registro
                            elementoX.SubItems(2) = !FOLIO_SALIDA ' folio salida
                           ' elementoX.SubItems(3) = "!FOLIO_origen" '
                        .MoveNext
                        Loop
                    
                    .Close
                    End With
            
     
     Else
     'no hay registros
            
            FrmMain.img_Env.Picture = LoadPicture(App.path & "\img\0.ico")
            FrmMain.lblenv = "No hay enviados"
            FrmMain.img_Env.Tag = "0"
            FrmMain.LV_Enviados.ListItems.Clear
     
     End If
           
    .Close
    
    
    
    Set LclRst = Nothing
    Set LclRstEnv = Nothing
    
    End With




End Sub








Public Sub ArchivoOid(Opc As Integer)
'el archivo solo contiene el oid del archivo que se recibe
'sirve para identificar el archivo mas reciente

Select Case Opc

Case Is = 1 ' recepcion cuando llega el archivo

    Open App.path & "\dll\RecepcionOid.Oid" For Output As #1
                Print #1, "[Oid_Recepcion]"
                Print #1, "OidReciente=" & Recepcion.bd_OidRecepcion
    Close (1)

Case Is = 2 ' cuando se envia un archivo se guarda el oid en otro archivo

    Open App.path & "\dll\EnvioOid.Oid" For Output As #2
                Print #2, "[Oid_Envio]"
                Print #2, "OidReciente=" & Envio.OidEnv
    
    Close (2)

End Select


End Sub






Public Sub ConsultarPendientes()


Dim strSql As String
Dim LclRst As ADODB.Recordset
Dim LclRstEnv As ADODB.Recordset
Dim Dia As String
Dim mes As String
Dim año As String
Dim Fecha As String
Dim b As Long
Dim elementoX As ListItem

FrmMain.LV_Pendientes.ListItems.Clear

strSql = "SELECT * From HEADERENVIO WHERE (((FECHA_REGISTRO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=0));"
''Debug.Print "Cadena sql de envio pendientes consulta: " & strSql
'Debug.Print strSql


Set LclRst = New ADODB.Recordset
    LclRst.CursorLocation = adUseClient
    LclRst.CursorType = adOpenDynamic
    LclRst.LockType = adLockPessimistic
    LclRst.Open strSql, CadenaCnx


    With LclRst

    
     If .RecordCount > 0 Then
             ' hay registros
             FrmMain.Img_Pen.Picture = LoadPicture(App.path & "\img\1.ico")
             FrmMain.lblPen = "Pendientes de Envio: " & .RecordCount
             FrmMain.Img_Pen.Tag = "1"
             'FrmMain.MouseIcon = LoadPicture(App.path & "\img\SELECT.cur")
             ' SE CARGA LA LISTA DE USUARIOS A LOS QUE SE LES HA ENVIADO INFORMACION EN EL DIA ACTUAL
             
                    strSql = "SELECT * From BODYENVIO WHERE (((BODYENVIO.ENVIO)=0 Or (BODYENVIO.ENVIO)=9)) and (((FECHA_REGISTRO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ));"
                    
                    Set LclRstEnv = New ADODB.Recordset
                    LclRstEnv.CursorLocation = adUseClient
                    LclRstEnv.CursorType = adOpenDynamic
                    LclRstEnv.LockType = adLockPessimistic
                    LclRstEnv.Open strSql, CadenaCnx
                    
        With LclRstEnv
                                 
                        Do While Not .EOF
                               Set elementoX = FrmMain.LV_Pendientes.ListItems.Add(, , !USUARIO_DESTINO)
                                   elementoX.Tag = "er"
                                   elementoX.SubItems(1) = !FECHA_REGISTRO ' fecha registro
                                   elementoX.SubItems(2) = !FOLIO_SALIDA ' folio salida
                           
                        .MoveNext
                        Loop
                              
                  
                    .Close
                    End With
            
     
     Else
     'no hay registros
            
            FrmMain.Img_Pen.Picture = LoadPicture(App.path & "\img\0.ico")
            FrmMain.lblPen = "No hay pendientes"
            FrmMain.Img_Pen.Tag = "0"
     
     
     End If
           
    .Close
    End With



Set LclRst = Nothing
Set LclRstEnv = Nothing


End Sub

Public Function DerechosUser(Clave As String, Derecho As Integer) As Boolean

'Por medio de la clave busco los derechos que tiene
'y se manda la respuesta de si lo tiene permitodo o no
'PERMISO_MODCONFIG 0
'PERMISO_CONINFOUSER 1
'PERMISO_ADDUSERS 2
'PERMISO_MODUSERS 3
'PERMISO_DELUSERS 4
'PERMISO_MODIP 5
Dim Str_Gral As String
Dim Rst_Gral As ADODB.Recordset

                
                        Str_Gral = "SELECT * FROM USUARIOS WHERE STATUS='OK' AND PASSWORD='" & Clave & "'"
                                      
                        Set Rst_Gral = New ADODB.Recordset
                        Rst_Gral.CursorLocation = adUseClient
                        Rst_Gral.CursorType = adOpenDynamic
                        Rst_Gral.LockType = adLockOptimistic
                        Rst_Gral.Open Str_Gral, CadenaCnx
                        
                
                If Rst_Gral.RecordCount = 0 Then
                
                        MsgBox "Password No Valido, Teclearlo de Nuevo", vbInformation, "No Valido"
                        FrmPassword.txtpass = ""
                        FrmPassword.txtpass.SetFocus
                        Rst_Gral.Close
                        Set Rst_Gral = Nothing
                        Exit Function

                End If
                

                UsuarioActivo.Nombre = Rst_Gral!NOMBRE_USUARIO
                UsuarioActivo.Apellido = Rst_Gral!APELLIDO_USUARIO
                UsuarioActivo.Grupo = Rst_Gral!Grupo
                UsuarioActivo.Comentario = Rst_Gral!Comentario
                UsuarioActivo.Nick = Rst_Gral!Nick
                UsuarioActivo.OidUsuario = Rst_Gral!OID_USUARIO
                UsuarioActivo.P_AddUsers = Rst_Gral!PERMISO_ADDUSERS
                UsuarioActivo.P_ConInfoUsers = Rst_Gral!PERMISO_CONINFOUSER
                UsuarioActivo.P_DelUsers = Rst_Gral!PERMISO_DELUSERS
                UsuarioActivo.P_ModConfig = Rst_Gral!PERMISO_MODCONFIG
                UsuarioActivo.p_ModIP = Rst_Gral!PERMISO_MODIP
                UsuarioActivo.P_ModUsers = Rst_Gral!PERMISO_MODUSERS

                
                Select Case Derecho
                
                Case Is = 0
                    
                    'Modificar configuiracion del programa
                    
                    If Rst_Gral!PERMISO_MODCONFIG Then
                    
                    'MsgBox "VALIDO"
                    Unload FrmPassword
                    
                    FrmConfiguracion.Show
                    DerechosUser = True
                    
                    Else
                    Unload FrmPassword
                    MsgBox "No Tiene Derecho para poder Modificar la Configuracion del Programa: " & UsuarioActivo.Nombre & "  " & UsuarioActivo.Apellido, vbInformation, "Sin Derecho"
                    DerechosUser = False
                    
                    End If
                    
                Case Is = 1
                    'consultar informacion de los usuarios
                    
                    If Rst_Gral!PERMISO_CONINFOUSER Then
                    
                    Unload FrmPassword
                    Unload FrmConfiguracion
                    
                    FrmAdminUsers.Show
                    DerechosUser = True
                                       
                    Else
                    Unload FrmPassword
                    Unload FrmOpciones
                    MsgBox "No Tiene Derecho para poder Consultar la Informacion de los Usuarios Registrados: " & UsuarioActivo.Nombre & "  " & UsuarioActivo.Apellido, vbInformation, "Sin Derecho"
                    DerechosUser = False
                    End If
                    
                    
                Case Is = 2
                
                    'Agregar usuarios
                    
                    If Rst_Gral!PERMISO_ADDUSERS Then
                    'MsgBox "VALIDO"
                    Unload FrmPassword
                    FrmAdminUsers.AddUser
                    DerechosUser = True
                    Else
                    Unload FrmPassword
                    'MsgBox "NO VALIDO"
                    MsgBox "No Tiene Derecho para poder Agregar un Usuario Nuevo: " & UsuarioActivo.Nombre & "  " & UsuarioActivo.Apellido, vbInformation, "Sin Derecho"
                    DerechosUser = False
                    End If
                
                Case Is = 3
                    'Modificar usuarios
                    
                    If Rst_Gral!PERMISO_MODUSERS Then
                    Unload FrmPassword
                    FrmAdminUsers.ModiUsers
                    DerechosUser = True
                    Else
                    Unload FrmPassword
                    MsgBox "No Tiene Derecho para poder Modificar un Usuario: " & UsuarioActivo.Nombre & "  " & UsuarioActivo.Apellido, vbInformation, "Sin Derecho"
                    DerechosUser = False
                    End If
                
                
                Case Is = 4
                
                    'Borrar Usuarios
                    
                    If Rst_Gral!PERMISO_DELUSERS Then
                    Unload FrmPassword
                    FrmAdminUsers.Delusers
                    DerechosUser = True
                    Else
                    Unload FrmPassword
                    MsgBox "No Tiene Derecho para Borrar a un Usuario: " & UsuarioActivo.Nombre & "  " & UsuarioActivo.Apellido, vbInformation, "Sin Derecho"
                    DerechosUser = False
                    End If
                
                
                Case Is = 5
                    'Modificar ip
                    'Nota: Puede requerirse pero realmente no tiene mucho sentido el no
                    'permitir modificar la ip del usuario destino
                    
                    If Rst_Gral!PERMISO_MODIP Then
                    MsgBox "VALIDO"
                    DerechosUser = True
                    Else
                    MsgBox "NO VALIDO"
                    DerechosUser = False
                    End If
                
                Case Else
                
                MsgBox "Password No Valido, Teclearlo de Nuevo", vbInformation, "No Valido"
                FrmPassword.txtpass = ""
                FrmPassword.txtpass.SetFocus
                
                End Select
                

Rst_Gral.Close
Set Rst_Gral = Nothing

End Function





Public Sub RegistraFavoritos()

Dim strSql As String
Dim RstFavoritos As ADODB.Recordset
strSql = "select * from Favoritos"


            Set RstFavoritos = New ADODB.Recordset
            RstFavoritos.CursorLocation = adUseClient
            RstFavoritos.CursorType = adOpenDynamic
            RstFavoritos.LockType = adLockPessimistic
            RstFavoritos.Open strSql, CadenaCnx
        
        With RstFavoritos
            
            .AddNew
            
                !IDFAVORITO = Favoritos.Bd_IDFAVORITO
                !OidEnv = Favoritos.Bd_OIDENV
                !NOMBRE_FAVORITO = Favoritos.Bd_NOMBREFAVORITO
                !Comentario = Favoritos.Bd_COMENTARIO
                !USUARIO_CREO = Usuario.Nick
                !Status = "OK"
                !FECHA_CREACION = Now
                !USUARIO_ACTUALIZO = Usuario.Nick
                !FECHA_ACTUALIZACION = Now
                !ID_LISTA_USUARIOS = Favoritos.Bd_ID_LISTA_USUARIOS
                
                                
            .Update
            .Close
            
        End With
         
       Set RstFavoritos = Nothing
       
End Sub


'Sirve para manejar los comandos que se agregan, modifican o se borran

Public Sub Comandos(Opc As Integer)

Dim Rst As ADODB.Recordset
Dim strSql  As String
Dim X As ListItem
Dim Result As Double


Select Case Opc

Case Is = 0 'Agregar

    'Nombre
    'RutaCmd
    'fechaCreacion
    'UsuarioCreo
    'FechaActualizacion
    'UsuarioActualizo
    'Status
    'ExecuteLike
    
        strSql = "INSERT INTO Tb_Cmd(Nombre, RutaCmd, fechaCreacion,UsuarioCreo,FechaActualizacion,UsuarioActualizo,Status,ExecuteLike) Values('" & Cmd.bd_Nombre & "', '" & Cmd.bd_RutaCmd & "', '" & Now & "', '" & Usuario.Nick & "', '" & Now & "','" & Usuario.Nick & "', 'OK ',0)"
    
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
        Set Rst = Nothing
        MsgBox "Agregado con Exito", vbInformation, "Operacion Exitosa"
        

Case Is = 1 'Modificar



    'strSql = " Tb_Cmd(Nombre, RutaCmd, fechaCreacion,UsuarioCreo,FechaActualizacion,UsuarioActualizo,Status,ExecuteLike) Values('" & Cmd.bd_Nombre & "', '" & Cmd.bd_RutaCmd & "', '" & Now & "', '" & Usuario.Nick & "', '" & Now & "','" & Usuario.Nick & "', 'OK ',0)"
    
        strSql = "UPDATE Tb_Cmd SET Tb_Cmd.Nombre = '" & Cmd.bd_Nombre & "', Tb_Cmd.RutaCmd = '" & Cmd.bd_RutaCmd & "', Tb_Cmd.FechaActualizacion = '" & Now & "', Tb_Cmd.UsuarioActualizo = '" & Usuario.Nick & "', Tb_Cmd.ExecuteLike = '" & FrmEDitcmd.OpcExe & "' WHERE (((Tb_Cmd.OidCmd)='" & FrmConfiguracion.Cmd_Oid & "'));"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
        Set Rst = Nothing
        MsgBox "Actualizado con Exito", vbInformation, "Operacion Exitosa"
        





Case Is = 2 'Eliminar

        strSql = "UPDATE Tb_Cmd SET Tb_Cmd.status = 'Del',Tb_Cmd.FechaActualizacion = '" & Now & "', Tb_Cmd.UsuarioActualizo = '" & Usuario.Nick & "' WHERE (((Tb_Cmd.OidCmd)='" & FrmConfiguracion.Cmd_Oid & "'));"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
        Set Rst = Nothing
        MsgBox "Borrado Adecuadamente", vbInformation, "Operacion Exitosa"




Case Is = 3 'Recuperar el comando agregado y agregarlo al listview

        strSql = "select * from Tb_Cmd where status='OK '"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
        
        If Rst.RecordCount = 0 Then
           FrmConfiguracion.Lv_Cmd.Enabled = False
           Set Rst = Nothing
           Exit Sub
        Else
           FrmConfiguracion.Lv_Cmd.Enabled = True
           FrmConfiguracion.Lv_Cmd.ListItems.Clear
                   
           
        End If
                
        

     
   
     With Rst
     
        .MoveFirst
        Do While Not .EOF
        
           Set X = FrmConfiguracion.Lv_Cmd.ListItems.Add(, , Rst!Nombre)
           X.Tag = Rst!OidCmd
           X.SubItems(1) = Rst!RutaCmd 'Ruta del comando a ejecutarse
           .MoveNext
           
        Loop
        
     End With
           
        Set Rst = Nothing

Case Is = 4 'Recuperar info para el form editar


        strSql = "select * from Tb_Cmd where oidCmd='" & FrmConfiguracion.Cmd_Oid & "'"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
        
        If Rst.RecordCount = 0 Then
            MsgBox "No encontrado"
            Set Rst = Nothing
            Exit Sub
        End If

        FrmEDitcmd.lblRut = Rst!RutaCmd
        FrmEDitcmd.txtNom = Rst!Nombre
        FrmEDitcmd.OptExe(Rst!ExecuteLike).Value = True
            

      
Case Is = 5 'Recuperar los comandos definidos

      
        strSql = "select * from Tb_Cmd where status='OK '"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
      
           
       With Rst
     
        .MoveFirst
        
            Do While Not .EOF
            
               FrmEnvio.Cb_CmdAsociado.AddItem Rst!Nombre
              .MoveNext
                            
            Loop
            
       End With
           
        Set Rst = Nothing
      
      
      
      
      
      
      
Case Is = 6 'busca el comando que es recibido y si existe regresa su informacion y datos


        strSql = "select * from Tb_Cmd where status='OK '"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.CursorType = adOpenDynamic
        Rst.LockType = adLockPessimistic
        Rst.Open strSql, CadenaCnx
        
        If Rst.RecordCount = 0 Then
        '   FrmConfiguracion.Lv_Cmd.Enabled = False
           Set Rst = Nothing
           Exit Sub
        Else
        '   FrmConfiguracion.Lv_Cmd.Enabled = True
        '   FrmConfiguracion.Lv_Cmd.ListItems.Clear
                        
        
         With Rst
        
              .MoveFirst
            Do While Not .EOF
          
               If Trim(Recepcion.bd_CmdEjecutar) = Trim(Rst!Nombre) Then
                
                Call Shell(Rst!RutaCmd, Rst!ExecuteLike)
                Exit Sub
                
                 
               End If
               
              .MoveNext
            Loop
          
         End With
           
        Set Rst = Nothing
                 
                   
                   
           
        End If
                
        

     
   
      
End Select


End Sub
