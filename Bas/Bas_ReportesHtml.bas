Attribute VB_Name = "Bas_Reportes"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Sub ReporteHtml(Opc As Integer)
Dim ArchHtml As String
Dim OrdenSql As String
Dim OrdenSqlRecep As String




Dim RstHeaderEnvio As ADODB.Recordset
Dim RstBodyEnvio As ADODB.Recordset


Dim RstHeaderRecepcion As ADODB.Recordset
Dim RstBodyRecepcion As ADODB.Recordset


        OrdenSql = "SELECT * From HEADERENVIO WHERE (((FECHA_REGISTRO) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((HEADERENVIO.STATUS_ENVIO)=1)) ORDER BY HEADERENVIO.FECHA_ENVIO;"
        
        Set RstHeaderEnvio = New ADODB.Recordset
            RstHeaderEnvio.CursorLocation = adUseClient
            RstHeaderEnvio.CursorType = adOpenDynamic
            RstHeaderEnvio.LockType = adLockPessimistic
            RstHeaderEnvio.Open OrdenSql, CadenaCnx




        OrdenSqlRecep = "SELECT * From HEADER_RECEPCION WHERE (((FECHA_CREACION) Between  #" & Format(Date, "mm/dd/yyyy") & "# And  #" & FECHA_FINAL & "# ) AND ((REVISADO)=False));"

        
        Set RstHeaderEnvio = New ADODB.Recordset
            RstHeaderEnvio.CursorLocation = adUseClient
            RstHeaderEnvio.CursorType = adOpenDynamic
            RstHeaderEnvio.LockType = adLockPessimistic
            RstHeaderEnvio.Open OrdenSqlRecep, CadenaCnx


Select Case Opc


Case Is = 1
'Reporte de envio de hoy
'1ero se genera el rst de envio

Case Is = 2
'Reporte de recepcion hoy

Case Is = 3
'Reporte de ambos envio y recepcion de hoy

ArchHtml = "Reporte" & Day(Date) & Month(Date) & Year(Date) & ".html"



Open Opciones.RutaReportes & "\" & ArchHtml For Output As #1

Print #1, "<html>"
Print #1, "<head>"
Print #1, "<title>Reporte de Movimientos General:  " & FECHA_INICIAL & "</title>"
Print #1, "</head>"
Print #1, "<body>"
Print #1, "<center><tt><b><font size=7 COLOR=BLACK>"
Print #1, "Reporte de Movimientos" & " (" & Date & ")"
Print #1, "</font><tt></b></center>"
Print #1, "<hr>"
Print #1, "<TABLE BORDER=1 WIDTH=750 align=center>"

Print #1, "<TR>"
Print #1, "<TD ALIGN=CENTER><b>Usuario Activo</b></TD>"
Print #1, "<TD ALIGN=CENTER><b>Nick</b></TD>"
Print #1, "<TD ALIGN=CENTER><b>IP Equipo</b></TD>"
Print #1, "<TD ALIGN=CENTER><b>Area Usuario</b></TD>"
Print #1, "</TR>"

Print #1, "<TR>"
Print #1, "<TD ALIGN=CENTER>" & Usuario.Nombre & " " & Usuario.Apellido & "</TD>"
Print #1, "<TD ALIGN=CENTER>" & Usuario.Nick & "</TD>"
Print #1, "<TD ALIGN=CENTER>" & Usuario.IP_equipo & "</TD>"
Print #1, "<TD ALIGN=CENTER>" & Usuario.Area & "</TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"

'///////////////////////////////// Seccion de Envio /////////////////////////////////////


Print #1, "<hr><br><br><br><br>"
Print #1, "<font size=5 COLOR=black><tt><b><u>Lista de Enviados" & "( " & RstHeaderEnvio.RecordCount & " )" & "</u></b></tt></FONT>"
Print #1, "<br><br><br><hr>"

If RstHeaderEnvio.RecordCount > 0 Then

                                    Print #1, "<TABLE BORDER=1 align=center>"
                                    Print #1, "<TR>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=10%><b>Folio Salida</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=20%><b>Fecha Creacion</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=20%><b>Comentario</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=10%><b>Nombre Archivo Comprimido</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=10%><b>No. Archivos</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=15%><b>Tamaño bytes</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER HEIGHT=20%><b>Fecha Envio</b></TD>"
                                    Print #1, "</TR>"
                                   ' Print #1, "</table>"
                                    
                        
                                    RstHeaderEnvio.MoveFirst
                            
                                    Do While Not RstHeaderEnvio.EOF
                                   ' Print #1, "<TABLE BORDER=1 WIDTH=100% align=center>"
                                    Print #1, "<TR>"
                                    Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstHeaderEnvio!FOLIO_SALIDA & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstHeaderEnvio!FECHA_REGISTRO & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstHeaderEnvio!Comentario & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstHeaderEnvio!NOMBRE_ARCHIVO_COMPRIMIDO & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstHeaderEnvio!No_Archivos & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=15%><font size=2>" & RstHeaderEnvio!TAMAÑO & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstHeaderEnvio!FECHA_ENVIO & "</font></TD>"
                                    Print #1, "</TR>"
                  '    MsgBox RstHeaderEnvio!OidEnv



                                    OrdenSql = "SELECT * From bodyENVIO WHERE  bodyENVIO.oidenv = '" & RstHeaderEnvio!OidEnv & "';"
                                    
                                    Set RstBodyEnvio = New ADODB.Recordset
                                        RstBodyEnvio.CursorLocation = adUseClient
                                        RstBodyEnvio.CursorType = adOpenDynamic
                                        RstBodyEnvio.LockType = adLockPessimistic
                                        RstBodyEnvio.Open OrdenSql, CadenaCnx
                                        
            
                              
                                    Print #1, "<TR>"
                                    Print #1, "<TD COLSPAN=7 ALIGN=left HEIGHT=100%>"
                                    
                                    Print #1, "<table border=0 width=100%>"
                                    
                                            Print #1, "<TR>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%>Usuario Destino</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%>IP Destino</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%>Fecha Envio</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%>Fecha Registro</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%>Puerto</TD>"
                                            Print #1, "</TR>"
                        
                                        RstBodyEnvio.MoveFirst
                                        Do While Not RstBodyEnvio.EOF
                                            Print #1, "<TR>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstBodyEnvio!USUARIO_DESTINO & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstBodyEnvio!ip & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstBodyEnvio!FECHA_ENVIO & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstBodyEnvio!FECHA_REGISTRO & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstBodyEnvio!puerto & "</font></TD>"
                                            Print #1, "</TR>"
                                            RstBodyEnvio.MoveNext
                                        Loop
 
                        
                                       Print #1, "</table>"
                                       
                                       Print #1, "</TD></TR>"
            
                                        RstBodyEnvio.Close
                                       RstHeaderEnvio.MoveNext
                                       
                                       Loop
                                      
                                      Print #1, "</TABLE>"
End If
                            
Print #1, "<br><br><br>"
'///////////////////////////////// Seccion de Envio (Fin) /////////////////////////////////////

'///////////////////////////////// Seccion de Recepcion /////////////////////////////////////


Print #1, "<hr><br><br><br><br>"
Print #1, "<font size=5 COLOR=black><tt><b><u>Lista de Recibidos" & "( " & RstHeaderRecepcion.RecordCount & " )" & "</u></b></tt></FONT>"
Print #1, "<br><br><br><hr>"
If RstHeaderRecepcion.RecordCount > 0 Then

                                    Print #1, "<TABLE BORDER=1 align=center>"
                                    Print #1, "<TR>"
                                    Print #1, "<TD ALIGN=CENTER><b>Folio Entrada</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Fecha Creacion</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Comentario</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Nombre Archivo Comprimido</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>No. Archivos</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Tamaño bytes</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Fecha Recepcion</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Hora Inicio</b></TD>"
                                    Print #1, "<TD ALIGN=CENTER><b>Hora Fin</b></TD>"
                             
                                    
                                    Print #1, "</TR>"
                                   ' Print #1, "</table>"
                                    
                        
                                    RstHeaderRecepcion.MoveFirst
                            
                                    Do While Not RstHeaderEnvio.EOF
                                   ' Print #1, "<TABLE BORDER=1 WIDTH=100% align=center>"
                                    Print #1, "<TR>"
                                    Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstHeaderRecepcion!FOLIO_SALIDA & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstHeaderRecepcion!FECHA_REGISTRO & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstHeaderRecepcion!Comentario & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstHeaderRecepcion!NOMBRE_ARCHIVO_COMPRIMIDO & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstHeaderRecepcion!No_Archivos & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=15%><font size=2>" & RstHeaderRecepcion!TAMAÑO & "</font></TD>"
                                    Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstHeaderRecepcion!FECHA_ENVIO & "</font></TD>"
                                    Print #1, "</TR>"
                  '    MsgBox RstHeaderEnvio!OidEnv



                                    OrdenSql = "SELECT * From bodyENVIO WHERE  bodyENVIO.oidenv = '" & RstHeaderEnvio!OidEnv & "';"
                                    
                                    Set RstBodyEnvio = New ADODB.Recordset
                                        RstBodyEnvio.CursorLocation = adUseClient
                                        RstBodyEnvio.CursorType = adOpenDynamic
                                        RstBodyEnvio.LockType = adLockPessimistic
                                        RstBodyEnvio.Open OrdenSql, CadenaCnx
                                        
            
                              
                                    Print #1, "<TR>"
                                    Print #1, "<TD COLSPAN=7 ALIGN=left HEIGHT=100%>"
                                    
                                    Print #1, "<table border=0 width=100%>"
                                    
                                            Print #1, "<TR>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%>Usuario Destino</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%>IP Destino</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%>Fecha Envio</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%>Fecha Registro</TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%>Puerto</TD>"
                                            Print #1, "</TR>"
                        
                                        RstBodyEnvio.MoveFirst
                                        Do While Not RstBodyEnvio.EOF
                                            Print #1, "<TR>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstBodyEnvio!USUARIO_DESTINO & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstBodyEnvio!ip & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=20%><font size=2>" & RstBodyEnvio!FECHA_ENVIO & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstBodyEnvio!FECHA_REGISTRO & "</font></TD>"
                                            Print #1, "<TD ALIGN=left HEIGHT=10%><font size=2>" & RstBodyEnvio!puerto & "</font></TD>"
                                            Print #1, "</TR>"
                                            RstBodyEnvio.MoveNext
                                        Loop
 
                        
                                       Print #1, "</table>"
                                       
                                       Print #1, "</TD></TR>"
            
                                        RstBodyEnvio.Close
                                       RstHeaderEnvio.MoveNext
                                       
                                       Loop
                                      
                                      Print #1, "</TABLE>"
                            
 End If
                            
                            
Print #1, "<br><br><br>"






Print #1, "<br><br><br>"
Print #1, "<font size=5 COLOR=black><tt><b><u>Lista de Pendientes</u></b></tt></font>"
Print #1, "<br><br><br>"
Print #1, "</body>"
Print #1, "</html>"
Close (1)

RstHeaderEnvio.Close

Call ShellExecute(FrmMain.hwnd, vbNullString, Opciones.RutaReportes & "\" & ArchHtml, vbNullString, _
             vbNullString, 1)
             



End Select

End Sub
