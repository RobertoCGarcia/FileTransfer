Attribute VB_Name = "Bas_Bitacora"
Option Explicit
'Desarrollado por: Garcia Cortes Roberto Carlos
''Modulo para manejo del registro de los eventos
'que se van produciendo en el programa al correr
'***

Public Sub RegistraEvento(IdEvento As String, Img As Integer, Titulo As String, Descrip As String)

Dim Rst As ADODB.Recordset
Dim strSql  As String
'OIDEVENTO
'TIPO
'IMAGEN
'TITULO
'DESCRIPCION
'USUARIO
'FECHA_EVENTO
'FECHA_CREACION
'USUARIO_CREO

    strSql = "INSERT INTO Tab_Log(TIPO, IMAGEN, TITULO,DESCRIPCION,USUARIO,FECHA_EVENTO,FECHA_CREACION,USUARIO_CREO) Values('" & IdEvento & "', '" & Img & "', '" & Titulo & "', '" & Descrip & "', '" & Usuario.Nick & "','" & Now & "', '" & Now & "','" & Usuario.Nick & "')"

    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.CursorType = adOpenDynamic
    Rst.LockType = adLockPessimistic
    Rst.Open strSql, CadenaCnx
    Set Rst = Nothing

End Sub



Public Sub ConsultarEventos(Opc As Integer)

Dim strSql As String
Dim Rst As ADODB.Recordset
Dim elementoX As ListItem


FrmBitacora.LV_Eventos.ListItems.Clear

Select Case Opc

Case Is = 0 'la opcion es cuando inicia el form
    
    strSql = "SELECT * From Tab_Log WHERE (((FECHA_EVENTO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ));" ' FECHA ACTUAL

Case Is = 1 ' la opcion es cuando se define un periodo de fechas inicial y final para realizar la busqueda solicitada
    
    strSql = "SELECT * From Tab_Log WHERE (((FECHA_EVENTO) Between  #" & FrmBitacora.FI & "# And  #" & FrmBitacora.FF & "# ));"  ' FECHA ACTUAL

End Select


    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.CursorType = adOpenDynamic
    Rst.LockType = adLockPessimistic
    Rst.Open strSql, CadenaCnx

    
    
    With Rst
          
     If .RecordCount = 0 Then
     FrmBitacora.LV_Eventos.Enabled = False
     Exit Sub
     Else
     FrmBitacora.LV_Eventos.Enabled = True
     End If
     
       
    .MoveFirst
     Do While Not .EOF
       Set elementoX = FrmBitacora.LV_Eventos.ListItems.Add(, , !TIPO)
       elementoX.Tag = "er"
       elementoX.SubItems(1) = !Titulo ' titulo evento
       elementoX.SubItems(2) = !Descripcion ' descripcion evento
       elementoX.SubItems(3) = !Usuario ' usuario que genero el evento
       elementoX.SubItems(4) = !fecha_evento ' comentario evento
       
       .MoveNext
     Loop
       .Close
    End With
        
    
Set Rst = Nothing


End Sub
