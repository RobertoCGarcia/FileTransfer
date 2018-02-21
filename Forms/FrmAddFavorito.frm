VERSION 5.00
Begin VB.Form frmAddFavorito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Favorito"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcoment 
      Height          =   1725
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   6615
   End
   Begin VB.TextBox txtNomFav 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6615
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion o Comentario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2190
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del Favorito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "frmAddFavorito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_add_Click()

Dim IdEnvio As String
Dim X As Integer

If Len(txtNomFav.Text) = 0 Then
MsgBox "Escriba un nombre Valido", vbCritical, "Nombre Necesario"
txtNomFav.Text = ""
txtNomFav.SetFocus
Exit Sub
End If

If Len(txtcoment.Text) = 0 Then
MsgBox "Escriba una Descripcion Valida", vbCritical, "Descripcion Necesaria"
txtcoment.Text = ""
txtcoment.SetFocus
Exit Sub
End If




Favoritos.Bd_NOMBREFAVORITO = txtNomFav.Text
Favoritos.Bd_COMENTARIO = txtcoment.Text


'Ya que algunos datos necesarios ya estan definidos se procede a
'guardar la informacion en las tablas siguientes:
Favoritos.Bd_ID_LISTA_USUARIOS = "L_" & Usuario.Nick & Mid(CStr(Rnd), 5, 6)
Favoritos.Bd_IDFAVORITO = "F_" & Usuario.Nick & Mid(CStr(Rnd), 3, 6)
NombreArchivoGenerado = Usuario.Nick & Mid(CStr(Rnd), 3, 5)
IdEnvio = NombreArchivoGenerado & "_" & Mid(CStr(Rnd), 3, 8)
Envio.OidEnv = IdEnvio
Favoritos.Bd_OIDENV = Envio.OidEnv



ReDim Mat_Envio(1 To FrmEnvio.Lv_User.ListItems.Count, 1 To 6)

For X = 1 To FrmEnvio.Lv_User.ListItems.Count
    Mat_Envio(X, 1) = IPUsuarios(FrmEnvio.Lv_User.ListItems.Item(X), 1)
    'Debug.Print Mat_Envio(x, 1) '& vbCrLf
    Mat_Envio(X, 2) = IIf(Ping(IPUsuarios(FrmEnvio.Lv_User.ListItems.Item(X), 1)), "1", "0")
    '.Print Mat_Envio(x, 2) '& vbCrLf
    If Mat_Envio(X, 2) = "0" Then
       Mat_Envio(X, 3) = "Equipo remoto no encontrado con Ping"
       Mat_Envio(X, 4) = "0"
    Else
       Mat_Envio(X, 3) = "Equipo remoto encontrado"
       Mat_Envio(X, 4) = "9"
    End If
    'Debug.Print Mat_Envio(x, 3) & vbCrLf
    Mat_Envio(X, 5) = Opciones.PuertoSalida ' puerto
    Mat_Envio(X, 6) = IPUsuarios(FrmEnvio.Lv_User.ListItems.Item(X), 2) ' UID
  
    '1 col: ip del equipo remoto
    '2 col resultado del status 1 conectado y listo 0 no esta listo
    '3 col comentario del status del posible error o causas de porque fue el error
    '4 cola de envio pendiente 9 se inicio apenas, 1 enviado con exito, 0 error de envio
    '5 puerto remoto
    '6 UID USUARIO ID
    'nombre del archivo con todo y ruta
Next X



'Una vez la matriz generada se procede con el envio
' invocando al programa adecuado
' primero se genera el archivo envio.lst
' que contiene los datos de la matriz



' se generO el archivo CON LA LISTA DE ENVIO AHORA hay que invocar al programa de envio y
' esperar que finalice para enviar el siguiente y asi hasta terminar la lista

''InxPosEnv = 1

  Open App.path & "\Generados\Lista\Envio.lst" For Output As #1
  
                            Print #1, "[INFO_ARCHIVO]"
                            Print #1, "Nombre=" & NombreArchivoGenerado & ".dpk"
                            Envio.bd_Nombre = "FAVORITO"
                            
                            Print #1, "Ruta=" & "FAVORITO"
                            Print #1, "Fecha_Creacion= " & Now
                            Envio.bd_FECHA_CREACION = Now
                     
                            Print #1, "NumeroEnvios= " & FrmEnvio.Lv_User.ListItems.Count
                            Envio.bd_NoArchivos = FrmEnvio.LV_ArchivosElegidos.ListItems.Count
                            Envio.bd_NoDestinos = FrmEnvio.Lv_User.ListItems.Count
                     
                            Print #1, "UsuarioOrigen= " & Usuario.Nombre & " " & Usuario.Apellido
                            
                            
                            Print #1, "UID=" & Usuario.OidUsuario
                            Envio.bd_UIDorigen = Usuario.OidUsuario
                     
                            Print #1, "OIDMOV= " & IdEnvio
                            
                  
                            If Len(txtcoment.Text) = 0 Then
                               Print #1, "ComentarioMain= Sin comentario" & vbCrLf
                               Envio.bd_ComentarioMain = "Sin comentario"
                            Else
                               Print #1, "ComentarioMain=" & txtcoment.Text & vbCrLf
                               Envio.bd_ComentarioMain = txtcoment.Text
                            End If
                            
                            Call RegistrarEnvio(1, True) ' Registro el header de la tabla envio
                  
                        
                        
                        
                         For X = 1 To FrmEnvio.Lv_User.ListItems.Count
                         
                           ' If Mat_Envio(x, 2) = "1" Then
                             Print #1, "[USUARIO_" & X & "]"
                             
                                    Print #1, "IP = " & Mat_Envio(X, 1)
                                    Envio.bd_IP = Mat_Envio(X, 1)
                                    
                                    Print #1, "PING = " & Mat_Envio(X, 2)
                                    Envio.bd_PING = Mat_Envio(X, 2)
                             
                                    Print #1, "COMENTARIO = " & Mat_Envio(X, 3)
                                    Envio.Bd_COMENTARIO = Mat_Envio(X, 3)
                                    
                                    Print #1, "ENVIO = " & Mat_Envio(X, 4)
                                    Envio.bd_Envio = Mat_Envio(X, 4)
                                    
                                    Print #1, "PUERTO = " & Mat_Envio(X, 5)
                                    Envio.bd_PUERTO = Mat_Envio(X, 5)
                                    
                                    
                                    Print #1, "FECHA_ENVIO = " & Now
                                    
                                    'Nombre del usuario destino
                                     Envio.bd_UsuarioDestino = Mid(FrmEnvio.Lv_User.ListItems.Item(X).SubItems(1), 5, Len(FrmEnvio.Lv_User.ListItems.Item(X).SubItems(1)))
                                                                
                                    Print #1, "UID =" & Mat_Envio(X, 6) & vbCrLf
                                    Envio.bd_UID = Mat_Envio(X, 6) ' el oid del usuario que es destino al que se le va enviar
                            
                            ' Print #1, "NOMBRE_USUARIO= "
                             
                             Call RegistrarEnvio(2, True) ' Registro el body de la tabla envio
 
                             
                             'Call EnvioControlado
                          '   End If
                         Next
     
     
     
     Close (1)
     
' se registran todos los archivos que se van a enviar
Call RegistrarEnvio(3, True)
Call RegistraFavoritos
' se actualiza el contador de las salidas en 1 para ir avanzando
'con cada salida

     
Call ArchivoOid(2)

Call InformacionUsuario(Usuario.OidUsuario, "UPDATE_SALIDAS")
Call InformacionUsuario(Usuario.OidUsuario, "CONSULTA")


 Call ConsultarRecepcion
 Call ConsultarEnvio
 Call ConsultarPendientes

Unload FrmEnvio

MsgBox "El Favorito:  " & Favoritos.Bd_NOMBREFAVORITO & "   Fue Guardado Correctamente. Puede Ocuparse en la Opcion QuickSend", vbInformation, "Operacion Exitosa"
Unload Me




End Sub

Private Sub cmd_Cancel_Click()
Unload Me
End Sub




Private Sub txtcoment_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomFav_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
