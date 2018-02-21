Attribute VB_Name = "Bas_Minimizar"
Option Explicit

'Modulo para minimizar en el taskbar
'Module:  SysTray.BAS
'               To put an icon in system tray
'Author:    Pheeraphat Sawangphian
'E-Mail:     tooh@asianet.co.th
'URL:       http://www.geocities.com/Hollywood/Lot/6166
'Note:       Put the following lines in Form Load event.
'                       SystemTray.cbSize = Len(SystemTray)
'                       SystemTray.hWnd = Me.hWnd
'                       SystemTray.uId = vbNull
'                       SystemTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'                       SystemTray.ucallbackMessage = WM_MOUSEMOVE
'                       SystemTray.hIcon = <Icon>
'                       SystemTray.szTip = <Tip> & vbNullChar
'                       Call Shell_NotifyIcon(NIM_ADD, SystemTray)
'                Where <Icon>   is an icon you want to show in system tray.
'                            <Tip>     is a tip that shown when mouse moved over the icon in system tray
'                You can detect mouse event in Form MouseMove event.
'                To remove an icon from system tray:
'                        Call Shell_NotifyIcon(NIM_DELETE, SystemTray)


Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const NIM_ADD = &H0          'add an icon into system tray
Public Const NIM_MODIFY = &H1    'update an icon in system tray
Public Const NIM_DELETE = &H2    'remove an icon from system tray

Public Const NIF_MESSAGE = &H1  'I want return messages
Public Const NIF_ICON = &H2          'adding an icon
Public Const NIF_TIP = &H4             'adding a tip

'rodent constant need for the callback
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200

'system tray notification definitions
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public SystemTray As NOTIFYICONDATA

Public SystemTrayIcon As Integer ' las 2 variables para la barra de tareas
Public SystemTrayTip As String  ' el texto y la imagen



Public Sub Minimizar()


FrmMenu.Icon = FrmMenu!ImgMain.Picture


    SystemTrayIcon = 0
    SystemTrayTip = "Administracion de Usuarios"
    
    SystemTray.cbSize = Len(SystemTray)                                         'size of system tray notification
    SystemTray.hwnd = FrmMenu.hwnd                                                     'form handle
    SystemTray.uId = vbNull
    SystemTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE  'show icon and tip in system tray and return messages
    SystemTray.ucallbackMessage = WM_MOUSEMOVE                'return messages in mousemove event when user do something with that icon
    SystemTray.hIcon = FrmMenu.Icon                                                    'specify an icon to show in system tray
    SystemTray.szTip = SystemTrayTip & vbNullChar                          'specify tip text

    Call Shell_NotifyIcon(NIM_ADD, SystemTray)                               'add an icon into system tray


End Sub
