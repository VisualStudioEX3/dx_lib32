Attribute VB_Name = "Global_Mod"
'===============================================================================
' Proyecto dx_lib32                                        
'-------------------------------------------------------------------------------
'                                                          
' Copyright (C) 2001 - 2010, José Miguel Sánchez Fernández 
'                                                          
' This file is part of dx_lib32 project.
'
' dx_lib32 project is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, version 2 of the License.
'
' dx_lib32 is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public License
' along with dx_lib32 project. If not, see <http://www.gnu.org/licenses/>.
'===============================================================================

'===============================================================================
' Name: Global_Mod
' Purpose: Declaraciones, Tipos y Constantes del API de Windows
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 07/08/2001
' Modified: 10/03/2010
'===============================================================================

'Constantes:
'______________________________________________________________________________________________________________________________________________________________________________________________________

'Constante que define un identificador de error de la libreria:
Public Const SubError = vbObjectError + 513

'Constante que define que una funcion ha realizado la operacion satisfactoriamente:
Public Const ERROR_SUCCESS As Long = 0&

'Numero PI: (Constante externa del API)
Public Const PI As Single = 3.14159265358979

'Constantes para el control de ventanas:
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'Constantes para leer el registro de Windows (32 Bits):
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

'Constantes para el control del joystick:
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1

Public Const JOY_POVCENTERED = -1
Public Const JOY_POVFORWARD = 0
Public Const JOY_POVRIGHT = 9000
Public Const JOY_POVLEFT = 27000
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNU = &H10
Public Const JOY_RETURNV = &H20
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNRAWDATA = &H100&
Public Const JOY_RETURNPOVCTS = &H200&
Public Const JOY_RETURNCENTERED = &H400&
Public Const JOY_USEDEADZONE = &H800&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Public Const JOY_CAL_READALWAYS = &H10000
Public Const JOY_CAL_READRONLY = &H2000000
Public Const JOY_CAL_READ3 = &H40000
Public Const JOY_CAL_READ4 = &H80000
Public Const JOY_CAL_READXONLY = &H100000
Public Const JOY_CAL_READYONLY = &H200000
Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000

'Constantes del cuadro de dialogo Abrir/Guardar como...
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_EXPLORER = &H80000
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

'Constante para abrir un fichero:
Public Const OF_EXIST = &H4000 'Especifica si existe.

'Constantes para ventana de directorios:
Public Const BIF_BROWSEFORCOMPUTER = 1000
Public Const BIF_BROWSEFORPRINTER = 2000
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_RETURNFSANCESTORS = 8
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_STATUSTEXT = 4

Public Const MAX_SIZE = 255

'Constante para SystemParametersInfo():
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_ENABLEWINKEYS = 97 'Desactiva las teclas y combinaciones de teclas de sistema en plataformas 9x.

'Constantes para busqueda de archivos:
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20

Public Const MAX_PATH = 260

'Bits o parametros de estilo de la ventana:
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SYSMENU = &H80000
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER As Long = &H800000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const GWL_STYLE = (-16)
Public Const WGAME_STYLE As Long = Not WS_CAPTION

'Constante para ShowWindow():
Public Const SW_NORMAL As Long = 1
Public Const SW_MAXIMIZE As Long = 3
Public Const SW_MINIMIZE As Long = 6

'Sistema metrico de windows:
Public Const SM_CYCAPTION = 4
Public Const SM_CYBORDER = 6
Public Const SM_CXBORDER = 5
Public Const SM_CXEDGE As Long = 45
Public Const SM_CYEDGE As Long = 46
Public Const SM_CYMENU = 15
Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen

'Constantes para mostrar u ocultar una ventana:
Public Const SW_HIDE As Long = 0
Public Const SW_RESTORE As Long = 9

'Constate para repintar la ventana al instante:
Public Const RDW_UPDATENOW As Long = &H100
Public Const RDW_NOERASE As Long = &H20

'Constantes para ventana de consola de texto:
Public Const STD_OUTPUT_HANDLE = -11&

Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_FONTCHANGE As Long = &H1D

'______________________________________________________________________________________________________________________________________________________________________________________________________

'Tipos de Datos:
'______________________________________________________________________________________________________________________________________________________________________________________________________

Public Type POINTAPI 'Tipo coordenadas de la API de Windows.
    X As Long
    Y As Long
End Type

Public Type Size 'Tipo Tamaño:
    cX As Long
    cY As Long
End Type

'Informacion de Memoria:
Public Type MEMORYSTATUS                ' size of 'Type' = 8 x 4 bytes = 32 (a Long is 4 Bytes)
    dwLength As Long                    ' This need to be set at the size of this 'Type'  = 32
    dwMemoryLoad As Long                ' Cantidad de memoria RAM utilizada.
    dwTotalPhys As Long                 ' Cantidad de memoria RAM total.
    dwAvailPhys As Long                 ' Cantidad de memoria RAM disponbile.
    dwTotalPageFile As Long             ' Cantidad de memoria total de pagina.
    dwAvailPageFile As Long             ' Cantidad de memoria disponible de pagina.
    dwTotalVirtual As Long              ' Cantidad de memoria virtual total.
    dwAvailVirtual As Long              ' Cantidad de memoria virtual disponible.
End Type

'Informacion extendida del joystick:
Public Type JOYINFOEX
    dwSize As Long                '  size of structure
    dwFlags As Long               '  flags to indicate what to return
    dwXpos As Long                '  x position
    dwYpos As Long                '  y position
    dwZpos As Long                '  z position
    dwRpos As Long                '  rudder/4th axis position
    dwUpos As Long                '  5th axis position
    dwVpos As Long                '  6th axis position
    dwButtons As Long             '  button states
    dwButtonNumber As Long        '  current button number pressed
    dwPOV As Long                 '  point of view state
    dwReserved1 As Long           '  reserved for communication between winmm driver
    dwReserved2 As Long           '  reserved for future expansion
End Type

'Estructura para la ventana Abrir/Guardar como...
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Hora y fecha del sistema:
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'Informacion de la version del Sistema Operativo:
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Estructura para abrir archivos:
Public Type OFSTRUCT '136 bytes in length
    cBytes As String * 1
    fFixedDisk As String * 1
    nErrCode As Integer
    reserved As String * 4
    szPathName As String * 128
End Type

'Estructura para ventana de directorios:
Public Type BROWSEINFO
         hwndOwner As Long
         pidlRoot As Long
         pszDisplayName As String
         lpszTitle As String
         ulFlags As Long
         lpfn As Long
         lParam As Long
         iImage As Long
End Type

'Estructura de fecha:
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

'Estructura de busqueda de archivos:
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

'Informacion de cabecera de Mapa de Bits para Win32:
Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Cabecera de Mapa de Bits para Win32:
Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Metrica de la fuente de texto:
Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

' Almacena los estados del teclado:
Public Type KeyboardBytes
     kbByte(0 To 255) As Byte
     
End Type


'______________________________________________________________________________________________________________________________________________________________________________________________________

'Funciones de la API de Windows:
'______________________________________________________________________________________________________________________________________________________________________________________________________

'Declaración de la función API para manejar los tiempos.
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Hace sonar el altavoz de la CPU.
Public Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'Declaracion de la función API para mostrar u ocultar el puntero del raton de Windows.
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long

'Devuelve la posicion del Raton.
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Cambia la posicion del Raton.
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

'Declaración de las funciones API's para escribir y leer archivos INI.
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Declaraciones para el envio de teclas.
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Declaracion para obtener la cantidad de memoria RAM.
Public Declare Sub apiMemStatus Lib "kernel32" Alias "GlobalMemoryStatus" (lpBuffer As MEMORYSTATUS)

'Definición de la API bloquear para el manejo de los parametros del sistema:
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'Funciones para el control del joystick.
Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Public Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal Id As Long) As Long
Public Declare Function joySetCapture Lib "winmm.dll" (ByVal hWnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long

'Funciones de dispositivos MCI. (Soporte para CD Audio).
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Funciones para obtener la ruta de los directorio del sistema (Windows, System, Temporal):
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Posiciona una ventana sobre todas las demas:
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'Funciones para acceder al registro de Windows (32 Bits):
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

'Definición de las API's que llaman los cuadros de dialogo Abrir y Guardar:
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'Devuelve hora y fecha del sistema:
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'Devuelven la version de DirectX instalada en el sistema:
Public Declare Function DirectXSetupGetVersion Lib "dsetup.dll" (dwVersion As Long, dwRevision As Long) As Long

'Devuelve informacion sobre el Sistema Operativo:
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'Devuelve el hWnd de la ventana que tiene el foco:
Public Declare Function GetFocus Lib "user32" () As Long

'Devuelve el hWnd de la ventana sobre la que se esta trabajando:
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'Devuelve el las dimensiones de la ventana:
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'Devuelve el las dimensiones del area de cliente de la ventana:
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'Abre la ventana de seleccion de directorio:
Public Declare Function BrowseFolderDlg Lib "shell32.dll" Alias "SHBrowseForFolder" (lpBrowseInfo As BROWSEINFO) As Long

'Devuelve la ruta del directorio seleccionado en la ventana de seleccion de directorio:
Public Declare Function GetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal PointerToIDList As Long, ByVal pszPath As String) As Long

'Funciones para realizar listados de archivos y directorios en una ruta especifica:
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'Funciones para modificar el estilo de la ventana:
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Funcion para copiar porciones de memoria:
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Funcion para congelar la ejecucion del programa:
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Funcion para averiguar mediante el hWnd si una ventana esta minimizada:
Public Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long

'Funcion para mostrar la ventana:
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'Funciones para crear Cronometros, lo mismo que el objeto Timer pero sin control grafico que lo represente:
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'Situa el foco a la ventana indicada:
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'Repinta la ventana:
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

'Funciones para modificar el tamaño del area de cliente de la ventana:
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Devuelve el hWnd del escritorio:
Public Declare Function GetDesktopWindow Lib "user32" () As Long

'Funciones para crear y manipular una consola de texto MS-DOS:
Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Public Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
Public Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal Title As String) As Long
Public Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttr As Long) As Long

'Devuelve el hWnd de la ventana buscada:
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Comprueba si un archivo existe o no:
Public Declare Function FileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Verifica si la ruta es un directorio (se utiliza para comprobar si un directorio existe o no)
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'Verifica si un directorio esta vacio:
Public Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long

'Seleccion de funciones para crear una fuente en un DC propio:
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long

'Funciones para obtener informacion acerca de una fuente de texto:
Public Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hdc As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetCharWidth32 Lib "gdi32.dll" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, ByRef lpBuffer As Long) As Long

'Contadores de alta precision:
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' Funciones para gestion de ventanas:
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long

'Indica si un array esta vacio:
Public Declare Sub GetSafeArrayPointer Lib "msvbvm60.dll" Alias "GetMem4" (pArray() As Any, Ret As Long)

'Crea una region rectangular:
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Aplica y obtiene una region de una ventana:
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function GetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long) As Long

'Importa o descarga una fuente externa al directorio Font de Windows:
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

'Escanea los estados del teclado:
Public Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long

'Devuelve el valor ASCII de una tecla:
Public Declare Function ToAscii Lib "user32.dll" (ByVal uVirtKey As Long, ByVal uScanCode As Long, ByRef lpbKeyState As Byte, ByRef lpwTransKey As Long, ByVal fuState As Long) As Long

' Genera un nombre de archivo temporal:
Private Declare Function APIGetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
     ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" ( _
     ByRef Destination As Any, _
     ByVal Length As Long, _
     ByVal Fill As Byte)

'______________________________________________________________________________________________________________________________________________________________________________________________________

'Variables globales de la dll:
'______________________________________________________________________________________________________________________________________________________________________________________________________

Public D3D_FullScreen As Boolean 'Inidica si el render de D3D ha sido inicializado a pantalla completa.
Public hConsole As Long 'Handle of console window

'______________________________________________________________________________________________________________________________________________________________________________________________________

'Funciones de ayuda:
'______________________________________________________________________________________________________________________________________________________________________________________________________

Public Sub LoadFont(CharBuffer() As Long, CharHeight As Long, Name As String, Size As Long, Bold As Boolean, Italic As Boolean, UnderLine As Boolean, Strikethrough As Boolean)
    Dim CharMetric As TEXTMETRIC
    
    'Create a device context, compatible with the screen
    mDC = CreateCompatibleDC(GetDC(0))
    
    'Select the new font into the form's device context and delete the old font
    DeleteObject SelectObject(mDC, CreateFont(-MulDiv(Size, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72), 0, 0, 0, Bold, Italic, UnderLine, Strikethrough, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, Name))
    
    GetTextMetrics mDC, CharMetric
    GetCharWidth32 mDC, 32, 255, CharBuffer(0)
    
    CharHeight = CharMetric.tmHeight
    
    'clean up
    DeleteDC mDC

End Sub

'Devuelve la longitud en pixeles de una cadena de texto:
Function GetStringWidth(CharBuffer() As Long, Text As String) As Long
Dim i As Long, j As Long
    Dim LenStr As Long
    
    LenStr = Len(Text)
    
    For i = 1 To LenStr
        j = j + CharBuffer(Asc(Mid(Text, i, 1)))
        
        DoEvents
        
    Next i
    
    GetStringWidth = j

End Function

'Esta funcion devuelve una cadena de texto con todos los caracteres desde la izquierda hasta el
'primer caracter de espacio que encuentre sin incluirlo.
Public Function Parser_1(Text As String) As String
    Dim l As Long
    
    For l = 1 To Len(Text)
        If Mid(Text, l, 1) = Chr(32) Then
            Parser_1 = VBA.Left$(Text, l - 1)
            Exit For
        End If
    Next l

End Function

'Formatea a 2 digitos. El primer valor es 0 si la cadena es de 1 digito.
Public Function sFormat(ByVal sInput As String) As String
    If Len(sInput) = 1 Then sInput = "0" & sInput
    sFormat = sInput

End Function

'Aplica un tamaño a la ventana:
Public Sub Set_SizeWindow(hWnd As Long, Height As Long, Width As Long, Windowed As Boolean)
    'Variables de valores metricos de la ventana:
    Dim BorderX As Long
    Dim BorderY As Long
    Dim EdgeX As Long
    Dim EdgeY As Long
    Dim CaptionY As Long
    Dim MenuY As Long
    
    'Dimensiones del area de la ventana:
    Dim WinRect As RECT
    
    Dim Top As Long
    
    BorderX = GetSystemMetrics(SM_CXBORDER)
    BorderY = GetSystemMetrics(SM_CYBORDER)
    EdgeX = GetSystemMetrics(SM_CXEDGE)
    EdgeY = GetSystemMetrics(SM_CYEDGE)
    CaptionY = GetSystemMetrics(SM_CYCAPTION)
    
    'Obtenemos el nuevo tamaño para la ventana y la centramos en pantalla:
    With WinRect
        .Right = Width + (BorderX * 2) + (EdgeX * 2)
        .bottom = Height + (BorderY * 2) + (EdgeY * 2) + CaptionY + MenuY
        .Left = GetSystemMetrics(SM_CXSCREEN) / 2 - (.Right / 2)
        .Top = GetSystemMetrics(SM_CYSCREEN) / 2 - (.bottom / 2)
        
        If Windowed Then
            Call SetWindowRgn(hWnd, GetWindowRgn(hWnd, 0), False)
            Top = HWND_NOTOPMOST
            
            ' Recuperar barra de tareas:
            '...
        
        Else
            Call SetWindowRgn(hWnd, CreateRectRgn(0, 0, Width, Height), True)
            Top = HWND_TOPMOST
            
            ' Ocultar barra de tareas:
            '...
        
        End If
    
        Call MoveWindow(hWnd, .Left, .Top, .Right, .bottom, True)
        'Call RedrawWindow(hwnd, ByVal 0&, ByVal 0&, RDW_UPDATENOW Or RDW_NOERASE)
        
        Call SetWindowPos(hWnd, Top, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
        
    End With

End Sub

Public Function ConsoleHandler(ByVal ctrlType As Long) As Long
    
    ConsoleHandler = 0  'Si se cierra la consola se cierra el programa.
    
    'ConsoleHandler = 1  'This tells the console window to ignore all console
                        'signals. If you don't do this, closing the console window
                        'or typing Ctrl-Break would cause your program to end.
                
End Function

'Obtiene la extexion de la lista de filtros en el cuadro de dialogo Guardar Como para añdirla al nombre de archivo de salida:
Public Function GetExtension(sfilter As String, pos As Long) As String
    Dim Ext() As String
    
    Ext = Split(sfilter, vbNullChar)
    
    If pos = 1 And Ext(pos) <> "*.*" Then
        GetExtension = LCase("." & Replace(Ext(pos), "*.", ""))
        
    ElseIf (pos = 1 And Ext(pos) = "*.*") Or InStr(Ext(pos + 1), "*.*") Then
        GetExtension = vbNullString
        
'    ElseIf InStr(Ext(pos + 1), "*.*") Then
'       GetExtension = vbNullString
'
    Else
       GetExtension = LCase("." & Replace(Ext(pos + 1), "*.", ""))
       
    End If
    
    'MsgBox GetExtension & vbNewLine & sfilter
    
End Function

' Importa una fuente desde archivo y devuelve su nombre:
Public Function LoadFontFile(Filename As String) As String
    'Dim listA() As String, listB() As String
    Dim lngRet As Long
    Dim cfont As New CFontPreview
    
    'listA = GetListFonts
    
    lngRet = AddFontResource(Filename)
    Call SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
    
    
    
'    If lngRet > 0 Then
'        listB = GetListFonts
'
'        Dim i As Integer
'        For i = 0 To UBound(listA)
'            If Not listA(i) = listB(i) Then
'                LoadFontFile = listB(i)
'                Exit Function
'
'            End If
'
'        Next i
'
'        LoadFontFile = listB(UBound(listB))
'
'    End If
  
End Function

'' Devuelve una lista con las fuentes del sistema:
'Public Function GetListFonts() As String()
'    Dim list() As String
'    Dim i As Integer
'
'    ReDim list(VB.Screen.FontCount - 1) As String
'
'    For i = 0 To VB.Screen.FontCount - 1
'        list(i) = VB.Screen.Fonts(i)
'
'    Next
'
'    GetListFonts = list
'
'End Function

' Rota un punto sobre el eje Z:
Public Sub RotatePoint(X As Long, Y As Long, cX As Long, cY As Long, Radio As Long, Angle As Single)
    If (Angle <= -360) Or (Angle >= 360) Then Angle = 0
    Select Case (Radio)
        Case 0 ', 360, -360
            cX = X + Radio
            cY = Y
        Case 90, -270
            cX = X
            cY = Y + Radio
        Case 180, -180
            cX = X - Radio
            cY = Y
        Case 270, -90
            cX = X
            cY = Y - Radio
        Case Else   ' Si el angulo no es perpendicular se calcula su rotacion:
            Dim aRdn As Single
            aRdn = (PI / 180) * CSng(Angle) ' Pasamos el angulo en grados a radianes.
            cX = X + Radio * Cos(aRdn)
            cY = Y + Radio * Sin(aRdn)
    End Select
End Sub

' Calcula la distancia en pixeles de un punto a otro:
Public Function GetPointDist(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Long
    On Local Error Resume Next
    
    Dim Dx As Long, Dy As Long
    
    Dx = Abs(X2 - X1)
    Dy = Abs(Y2 - Y1)
    
    GetPointDist = CLng(Sqr(Dx * Dx + Dy * Dy))

End Function

' Genera un nombre de archivo temporal:
Public Function GetTempFileName() As String
    Dim sTemp As String
    
    'Create a buffer
    sTemp = String(260, 0)
    'Get a temporary filename
    Call APIGetTempFileName(Environ("TEMP"), CStr(Timer), 0, sTemp)
    'Remove all the unnecessary chr$(0)'s
    sTemp = Left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
    
    GetTempFileName = sTemp
End Function
