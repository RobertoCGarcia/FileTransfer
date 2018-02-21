Attribute VB_Name = "Bas_Compresion"
Private Declare Function compress2 Lib "zlib.dll" (Dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long, ByVal Level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (Dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long) As Long



Private Type DPack
    SigHead As String * 3
    NoOfFiles As Long
End Type

Private Type TFileStruct
    m_Filename() As String
    m_Filesize() As Long
    m_FileData() As String
End Type



Public s_Head As DPack
Public s_File As TFileStruct
Public lzFilename As String
Public NombreArchivoGenerado As String
Public RutArchElegido As String


Public Function CompressStr(StrB As String, CompLevel As Long) As Long

Dim StrDstLen As Long, StrSrcLen As Long, TResult As Long
Dim StrDst As String
    StrSrcLen = Len(StrB) ' Assign StrSrcLen with StrB string length
    StrDst = String(StrSrcLen + (StrSrcLen * 0.01) + 12, 0) ' Resize buffer
    StrDstLen = Len(StrDst) ' Update StrDstLen with StrDst string length
    TResult = compress2(ByVal StrDst, StrDstLen, ByVal StrB, StrSrcLen, CompLevel)   ' Call the zlib dll to compress the data
    
    StrB = Left(StrDst, StrDstLen) ' Strip away any junk
    StrDst = "" ' Clear the buffer
    CompressStr = TResult ' Return zlib code
    
End Function

Public Function StrDecompress(StrB As String, OldSize As Long) As Long
Dim StrDst As String
Dim StrDstLen As Long
Dim StrSrcLen As Long
Dim TResult As Long

    StrSrcLen = Len(StrB) ' Assign StrSrcLen with StrB string length
    StrDst = String(OldSize + (OldSize * 0.01) + 12, 0)
    StrDstLen = Len(StrDst) ' Update StrDstLen with StrDst string length

    TResult = uncompress(ByVal StrDst, StrDstLen, ByVal StrB, StrSrcLen)   ' Call the zlib dll to compress the data
    StrB = Left(StrDst, StrDstLen) ' Strip away any junk
    StrDst = "" ' Clear the buffer
    StrDecompress = TResult ' Return zlib code
End Function

Function Fixpath(lzpath As String) As String
    If Right(lzpath, 1) = "\" Then Fixpath = lzpath Else Fixpath = lzpath & "\"
End Function

Function FindFile(lzFile As String) As Boolean
    If Dir(lzFile) = "" Then FindFile = False Else FindFile = True
End Function


Public Function IsVialdHeader(lzFile As String) As Boolean
' This function checks to see if we have a vaild file header
Dim nFile As Long
    nFile = FreeFile
    s_Head.NoOfFiles = 0
    s_Head.SigHead = ""
    Open lzFile For Binary As #nFile
        Get #nFile, , s_Head
    Close #nFile
    If (Not s_Head.SigHead = "DPK") Then IsVialdHeader = False Else IsVialdHeader = True

End Function




Public Sub DoBackup(lzbkFile As String, Tipo As Integer, Optional RstFiles As ADODB.Recordset)

Dim i_File As Long, i As Long
Dim m_Filename As String
Dim Str_Buffer As String
Dim mResult As Long
Dim C As Integer

Select Case Tipo


Case Is = 0 ' envio normal, no es favorito

        C = 1
        
            i_File = FreeFile ' Free file pointer
            'numero de archivo a enviar
            
            For i = 0 To FrmEnvio!LV_ArchivosElegidos.ListItems.Count - 1 ' Look to we get the end of the listbox listcount
            
                m_Filename = FrmEnvio!LV_ArchivosElegidos.ListItems.Item(C).SubItems(2) ' Assign m_Filename Ruta Completa del archivo a enzipar
                Str_Buffer = OpenFile(m_Filename)  ' Assign Str_Buffer with the file contents
                s_Head.NoOfFiles = (i + 1) ' Up date the filecount see s_head type
                ReDim Preserve s_File.m_Filename(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_FileData(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_Filesize(s_Head.NoOfFiles) ' Resize array
                s_File.m_Filename(s_Head.NoOfFiles) = GetFilename(m_Filename) ' Assign array with filename
                s_File.m_Filesize(s_Head.NoOfFiles) = Len(Str_Buffer)   ' Assign array with the size of the file
                mResult = CompressStr(Str_Buffer, 9) ' Call Zlib see bas for function
                s_File.m_FileData(s_Head.NoOfFiles) = Str_Buffer ' Assign the compressed data to the array
            
            C = C + 1
                
            Next
            
            
        'C:\Compresion\Ensamble\0x00.Raw
            'Agregar el archivo que tiene la informacion acerca de donde se extraera la informacion de estos
            
                m_Filename = Opciones.Rutacarpeta_Ensamble & "\0x00.Raw" ' Assign m_Filename Ruta Completa del archivo a enzipar
                Str_Buffer = OpenFile(m_Filename)  ' Assign Str_Buffer with the file contents
                s_Head.NoOfFiles = (i + 1) ' Up date the filecount see s_head type
                ReDim Preserve s_File.m_Filename(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_FileData(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_Filesize(s_Head.NoOfFiles) ' Resize array
                s_File.m_Filename(s_Head.NoOfFiles) = GetFilename(m_Filename) ' Assign array with filename
                s_File.m_Filesize(s_Head.NoOfFiles) = Len(Str_Buffer)   ' Assign array with the size of the file
                mResult = CompressStr(Str_Buffer, 9) ' Call Zlib see bas for function
                s_File.m_FileData(s_Head.NoOfFiles) = Str_Buffer ' Assign the compressed data to the array
            
            
            s_Head.SigHead = "DPK" ' Assign header information
            Open lzbkFile For Binary As #i_File ' Open the file in binary mode
                Put #i_File, , s_Head ' Save header information
                Put #i_File, , s_File ' Save file data information
            Close #i_File ' Close the file
            
          '  MsgBox "Todos los Archivos fueron comprimirdos con exito en:" _
          '  & vbNewLine & vbNewLine & lzbkFile, vbInformation, "Proceso Finalizado"
            
            
            'borro el archivo que se envio como apoyo de informacion, ya que no se necesita
            
            If Len(Dir(Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")) <> 0 Then
            'El archivo existe
              Kill (Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
            End If
            
            
            
            
            ' show message to user that files have been backed up
            m_Filename = "" ' Clear fie buffer
            CleanVars ' Call Clean vars sub
            

Case Is = 1 ' El respaldo se realiza por medio de un favorito definido previamente



        C = 1
        
            i_File = FreeFile ' Free file pointer
            'numero de archivo a enviar
            RstFiles.MoveFirst
            
            For i = 0 To RstFiles.RecordCount - 1 ' Look to we get the end of the listbox listcount
            
                m_Filename = RstFiles!RUTA_ORIGEN ' Assign m_Filename Ruta Completa del archivo a enzipar
                Str_Buffer = OpenFile(m_Filename)  ' Assign Str_Buffer with the file contents
                s_Head.NoOfFiles = (i + 1) ' Up date the filecount see s_head type
                ReDim Preserve s_File.m_Filename(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_FileData(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_Filesize(s_Head.NoOfFiles) ' Resize array
                s_File.m_Filename(s_Head.NoOfFiles) = GetFilename(m_Filename) ' Assign array with filename
                s_File.m_Filesize(s_Head.NoOfFiles) = Len(Str_Buffer)   ' Assign array with the size of the file
                mResult = CompressStr(Str_Buffer, 9) ' Call Zlib see bas for function
                s_File.m_FileData(s_Head.NoOfFiles) = Str_Buffer ' Assign the compressed data to the array
            
            C = C + 1
            RstFiles.MoveNext
            
            Next
            
            
        'C:\Compresion\Ensamble\0x00.Raw
            'Agregar el archivo que tiene la informacion acerca de donde se extraera la informacion de estos
            
                m_Filename = Opciones.Rutacarpeta_Ensamble & "\0x00.Raw" ' Assign m_Filename Ruta Completa del archivo a enzipar
                Str_Buffer = OpenFile(m_Filename)  ' Assign Str_Buffer with the file contents
                s_Head.NoOfFiles = (i + 1) ' Up date the filecount see s_head type
                ReDim Preserve s_File.m_Filename(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_FileData(s_Head.NoOfFiles) ' Resize array
                ReDim Preserve s_File.m_Filesize(s_Head.NoOfFiles) ' Resize array
                s_File.m_Filename(s_Head.NoOfFiles) = GetFilename(m_Filename) ' Assign array with filename
                s_File.m_Filesize(s_Head.NoOfFiles) = Len(Str_Buffer)   ' Assign array with the size of the file
                mResult = CompressStr(Str_Buffer, 9) ' Call Zlib see bas for function
                s_File.m_FileData(s_Head.NoOfFiles) = Str_Buffer ' Assign the compressed data to the array
            
            
            s_Head.SigHead = "DPK" ' Assign header information
            Open lzbkFile For Binary As #i_File ' Open the file in binary mode
                Put #i_File, , s_Head ' Save header information
                Put #i_File, , s_File ' Save file data information
            Close #i_File ' Close the file
            
          '  MsgBox "Todos los Archivos fueron comprimirdos con exito en:" _
          '  & vbNewLine & vbNewLine & lzbkFile, vbInformation, "Proceso Finalizado"
            
            
            'borro el archivo que se envio como apoyo de informacion, ya que no se necesita
            
            If Len(Dir(Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")) <> 0 Then
            'El archivo existe
              Kill (Opciones.Rutacarpeta_Ensamble & "\0x00.Raw")
            End If
            
            
            
            
            ' show message to user that files have been backed up
            m_Filename = "" ' Clear fie buffer
            CleanVars ' Call Clean vars sub
            
         '''   lstAdd.Clear ' Clear listbox contents
        'Debug.Print "Terminado el archivo   " & lzbkFile



        'MsgBox "Archivo generado con exito!!!!"


End Select


End Sub





Sub CleanVars()
  
    'Limpiar Todo lo que no se necesita mas
    
    
    s_Head.NoOfFiles = 0
    s_Head.SigHead = ""
    Erase s_File.m_FileData
    Erase s_File.m_Filename
    Erase s_File.m_Filesize
End Sub








Public Function OpenFile(lzFile As String) As String
Dim nFile As Long, fData As String
    nFile = FreeFile ' Pointer to the file
    Open lzFile For Binary As #nFile ' open the file in binary mode
        fData = Space(LOF(nFile))
        Get #nFile, , fData      ' get file contents
    Close #nFile    ' close the file

    OpenFile = fData
    fData = ""
    
End Function




Function GetFilename(lzPathFile As String) As String
Dim i As Long
    For i = Len(lzPathFile) To 1 Step -1
       If InStr(i, lzPathFile, "\", vbBinaryCompare) Then
            GetFilename = Mid(lzPathFile, i + 1, Len(lzPathFile))
            Exit For
            Exit Function
            i = 0
        Else
            GetFilename = lzPathFile
       End If
    Next
    i = 0
End Function






Public Function ListFileContents(lzFile As String)
' The Function add all the files in the compressed file

Dim nFile As Long, i As Long
    CleanVars
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        Get #nFile, , s_Head
        Get #nFile, , s_File
    Close #nFile
    
    For i = 1 To s_Head.NoOfFiles
     '   lstfiles1.AddItem s_File.m_Filename(I)
     
   ''  Debug.Print s_File.m_Filename(I)
     
    
    Next
    i = 0
    
    
End Function







Sub DoRestore(lzExtractPath As String)
Dim i As Long, StrData As String
    nFile = FreeFile ' Pointer to a free file
    For i = 1 To s_Head.NoOfFiles ' Look though the listbox
        
        'lblfilestat.Caption = s_File.m_Filename(I)
        
     ''   Debug.Print "Extrayendo contenido:  " & s_File.m_Filename(I)
        
        
        ' Update the fiel status caption this is not realy needed
        ' If you have time or know you choud add a processbar
        DoEvents
        StrData = s_File.m_FileData(i) ' Get file data form array
        StrDecompress StrData, s_File.m_Filesize(i) ' Call Zib see bas file for more info
        Open lzExtractPath & s_File.m_Filename(i) For Binary As #nFile ' Open the file in binary mode
            Put #nFile, , StrData ' Save data to the file
        Close #nFile ' Close the file
    Next
  '  MsgBox "Todos " & s_Head.NoOfFiles & " Fueron Extraidos a: " & lzExtractPath, vbInformation, "Proceso Finalizado"



End Sub









