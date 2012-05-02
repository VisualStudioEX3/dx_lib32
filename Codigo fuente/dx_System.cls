VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dx_System_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 188
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
' Name: dx_System
' Purpose: Clase de funciones complementarias
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 07/08/2001
' Modified: 14/03/2010
'===============================================================================
Option Explicit

'Constantes:
Private Const ID_PAK_Header = "PACK" 'Formato: Quake1, Quake2, Hexen2, Half-Life

'Enumeraciones:
'===============================================================================
' Name: SYS_ErrorCodes
'    SYS_OK - La operación se ha realizado con éxito.
'    SYS_EMPTYLIST - La lista esta vacía.
'    SYS_FILENOTFOUND - No se encontró el archivo.
'    SYS_INVALIDFORMAT = Formato incorrecto.
'    SYS_INVALIDPATH = Ruta no valida.
'    SYS_INVALIDARG = Argumento no valido.
'    SYS_KEYNOFOUND = No se encontró la clave.
'    SYS_NOTCREATED - No se pudo crear el objeto o recurso.
'    SYS_HITIMERNOTSUPPORT - El sistema no soporta cronometros de alta precision.
'    SYS_UNKNOWNERROR - Error desconocido.
'    SYS_UNKNOWNERROR - Error desconocido.
' Purpose: Codigos de error.
' Remarks: Definen errores o resultados de algunas funciones de esta clase.
'===============================================================================
Public Enum SYS_ErrorCodes
    SYS_OK = 0
    SYS_EMPTYLIST = (vbObjectError + 5000)
    SYS_FILENOTFOUND = (vbObjectError + 5001)
    SYS_INVALIDFORMAT = (vbObjectError + 5002)
    SYS_INVALIDPATH = (vbObjectError + 5003)
    SYS_INVALIDARG = (vbObjectError + 5004)
    SYS_KEYNOFOUND = (vbObjectError + 5005)
    SYS_NOTCREATED = (vbObjectError + 5006)
    SYS_HITIMERNOTSUPPORT = (vbObjectError + 5007)
    SYS_UNKNOWNERROR = (vbObjectError + 5999)

End Enum

'===============================================================================
' Name: SYS_Path
'    WINDOWS_DIR - Ruta del directorio de Windows.
'    SYSTEM_DIR - Ruta del directorio de sistema de Windows.
'    TEMP_DIR - Ruta del directorio temporal de Windows. En sistemas NT devuelve la ruta del directorio temporal del usuario.
'    TEMP_DIR - Ruta del directorio temporal de Windows. En sistemas NT devuelve la ruta del directorio temporal del usuario.
' Purpose: Constantes de rutas del sistema.
' Remarks: Estas constantes se usan con la función SYS_GetPath.
'===============================================================================
Public Enum SYS_Path
    WINDOWS_DIR = 0
    SYSTEM_DIR = 1
    TEMP_DIR = 2
    
End Enum

'===============================================================================
' Name: REG_Path
'    HKEY_CLASSES_ROOT -
'    HKEY_CURRENT_USER - Ruta hacia los registros del usuario actual.
'    HKEY_LOCAL_MACHINE - Ruta hacia los registros del sistema.
'    HKEY_USERS - Ruta hacia los registros de los usuarios.
'    HKEY_PERFORMANCE_DATA -
'    HKEY_CURRENT_CONFIG - Ruta hacia los registros de configuración.
'    HKEY_DYN_DATA -
'    HKEY_DYN_DATA -
' Purpose: Constantes de rutas del registro de Windows.
' Remarks: Estas constantes se usan con las funciones de acceso al registro de Windows.
'===============================================================================
Public Enum REG_Path
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    
End Enum

'===============================================================================
' Name: EGA_Color
'    Black - Negro
'    DarkBlue - Azul oscuro.
'    DarkGreen - Verde oscuro.
'    DarkCyan - Cian oscuro.
'    DarkRed - Granate.
'    Purple = Púrpura.
'    Oliva = Verde oliva.
'    Silver = Plata o gris claro.
'    Grey = Gris.
'    Blue = Azul.
'    Green = Verde.
'    Cyan = Cian.
'    Red = Rojo.
'    Violet = Violeta.
'    Yellow = Amarillo.
'    White = Blanco.
'    White = Blanco.
' Purpose: Constantes de colores EGA para la consola de depuración.
' Remarks: Estas constantes se usan para definir los colores de los textos de la consola de depuración.
'===============================================================================
Public Enum EGA_Color
    Black = &H0
    DarkBlue = &H1
    DarkGreen = &H2
    DarkCyan = &H3
    DarkRed = &H4
    Purple = &H5
    Oliva = &H6
    Silver = &H7
    Grey = &H8
    Blue = &H9
    Green = &HA
    Cyan = &HB
    Red = &H1C
    Violet = &HD
    Yellow = &HE
    White = &HF

End Enum

'===============================================================================
' Name: Capture_Mode
'    Screen - Toda la pantalla.
'    WindowsFocus - Solo la ventana que tiene el foco.
'    WindowsFocus - Solo la ventana que tiene el foco.
' Purpose: Constantes del modo de captura de pantalla.
' Remarks: Estas constantes definen el comportamiento de la función SYS_Capture.
'===============================================================================
Public Enum Capture_Mode
    Screen = 0
    WindowsFocus = 1

End Enum

'===============================================================================
' Name: File_Attribute
'    Archive - Archivo.
'    Directory - Directorio.
'    Normal - Atributo de archivo normal.
'    Read_Only - Atributo de solo lectura.
'    Hidden - Atributo de archivo oculto.
'    System - Atributo de archivo de sistema.
'    Temporary - Atributo de archivo temporal.
'    Compressed - Atributo de archivo comprimido.
'    Compressed - Atributo de archivo comprimido.
' Purpose: Constantes de atributos de archivo.
' Remarks: Estas constantes se utilizan para definir la búsqueda en la función FILE_List.
'===============================================================================
Public Enum File_Attribute
    Archive = FILE_ATTRIBUTE_ARCHIVE
    Directory = FILE_ATTRIBUTE_DIRECTORY
    Normal = FILE_ATTRIBUTE_NORMAL
    Read_Only = FILE_ATTRIBUTE_READONLY
    Hidden = FILE_ATTRIBUTE_HIDDEN
    System = FILE_ATTRIBUTE_SYSTEM
    Temporary = FILE_ATTRIBUTE_TEMPORARY
    Compressed = FILE_ATTRIBUTE_COMPRESSED

End Enum

'===============================================================================
' Name: Sort_Mode
'    Increment - Ordena la lista de forma creciente, de menor a mayor.
'    Decrement - Ordena la lista de forma decreciente, de mayor a menor.
'    Decrement - Ordena la lista de forma decreciente, de mayor a menor.
' Purpose: Constantes del modo de ordenación para listas.
' Remarks: Estas constantes definen el comportamiento de ordenación de varias funciones de la clase.
'===============================================================================
Public Enum Sort_Mode
    Increment = 0
    Decrement = 1

End Enum

'===============================================================================
' Name: OS_Plataform
'    Win_9X - Versiones de Windows 95, 98 y Me (Millenium).
'    Win_NT - Versiones de Windows NT, 2000, XP y 2003.
'    Win_NT - Versiones de Windows NT, 2000, XP y 2003.
' Purpose: Constantes que definen la plataforma de Windows.
' Remarks: Estas constantes estan asociadas solo a la estructura de datos OS_Info.
'===============================================================================
Public Enum OS_Plataform
    Win_9x = 1
    Win_NT = 2

End Enum

'===============================================================================
' Name: DEBUG_FileOutput
'    PlainText - Texto plano, sin formato.
'    HTML - Formateado en HTML. Si las entradas en la consola se les aplica color este modo respetara el formato.
'    HTML - Formateado en HTML. Si las entradas en la consola se les aplica color este modo respetara el formato.
' Purpose: Constantes que definen la salida de las entradas de la consola de depuracion en el archivo de registros.
' Remarks: Estas constantes se utilizan para definir el formato de salida en el archivo de registro.
'===============================================================================
'Public Enum DEBUG_FileOutput
'    PlainText = 0
'    HTML = 1
'
'End Enum

'Estructuras de datos:
'===============================================================================
' Name: Memory_Info
'   ByRef TotalPhys As Long - Cantidad total de memoria física.
'   ByRef AvailPhys As Long - Cantidad disponible de memoria física.
'   ByRef TotalPageFile As Long - Cantidad total de memoria de archivo de pagina.
'   ByRef AvailPageFile As Long - Cantidad disponible de memoria de archivo de pagina.
'   ByRef TotalVirtual As Long - Cantidad total de memoria virtual.
'   ByRef AvailVirtual As Long - Cantidad disponible de memoria virtual.
' Purpose: Estructura que almacena las cantidades de memoria del sistema.
' Remarks: Esta estructura se utiliza como valor de retorno de la función SYS_GetMemory.
'===============================================================================
Public Type Memory_Info
    TotalPhys As Long
    AvailPhys As Long
    TotalPageFile As Long
    AvailPageFile As Long
    TotalVirtual As Long
    AvailVirtual As Long
    
End Type

'===============================================================================
' Name: Sys_Time
'   ByRef Year As Long - Valor que define el año.
'   ByRef Month As Long - Valor que representa el mes del año.
'   ByRef Day As Long - Valor que representa el día del mes.
'   ByRef Hour As Long - Valor que representa la hora del día.
'   ByRef Minute As Long - Valor que representa el minuto de la hora.
'   ByRef Second As Long - Valor que representa el segundo del minuto.
'   ByRef Milliseconds As Long - Valor que define el milisegundo del segundo.
'   ByRef DayOfWeek As Long - Valor que define el día de la semana.
' Purpose: Estructura que almacena la fecha y hora del sistema.
' Remarks: Esta estructura se utiliza como valor de retorno de la función SYS_GetTime.
'===============================================================================
Public Type Sys_Time
    Year As Long
    Month As Long
    Day As Long
    Hour As Long
    Minute As Long
    Second As Long
    Milliseconds As Long
    DayOfWeek As Long
    
End Type

'===============================================================================
' Name: OS_Info
'   ByRef MajorVersion As Long - Numero de la versión del sistema operativo.
'   ByRef MinorVersion As Long - Numero de la subversión del sistema operativo.
'   ByRef BuildNumber As Long - Numero de compilación.
'   ByRef PlatformId As OS_Plataform - Identificador de plataforma. Indica si la plataforma esta basada en Windows 9x (95/98/Me) o Windows NT (NT, 2000, XP, 2003).
'   ByRef WindowsId As String - Cadena de texto que devuelve el nombre y edición de la plataforma. Por ejemplo: "Microsoft Windows XP Home Edition"
' Purpose: Estructura que almacena la información del sistema operativo.
' Remarks: Esta estructura se utiliza como valor de retorno de la función SYS_GetOSInfo.
'===============================================================================
Public Type OS_Info
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformID As OS_Plataform
    WindowsId As String
    
End Type

'Estructura archivos PAK:
Private Type PAKFileHeader
  ident As Long
  dirofs As Long
  dirlen As Long
  
End Type

Private Type FileInPAK
   named As String * 56
   filepos As Long
   filelen As Long
   
End Type

'===============================================================================
' Name: PAK_FileInfo
'   ByRef Filename As String - Nombre y ruta del archivo en el paquete.
'   ByRef OffSet As Long - Posición o byte donde comienza el archivo.
'   ByRef Size As String - Tamaño y longitud del archivo en bytes.
' Purpose: Estructura que almacena la información de un archivo contenido en un paquete de formato PAK.
' Remarks: Esta estructura se utiliza como valor de retorno y argumento en las funciones de lectura y escritura de archivos contenidos en un paquete de formato PAK.
'===============================================================================
Public Type PAK_FileInfo  'Similar a FileInPAK pero con cadenas de texto dinamicas.
    Filename As String  'Nombre del archivo.
    Offset As Long      'Byte donde comienza el archivo.
    Size As Long        'Longitud en bytes del archivo.
    
End Type

'===============================================================================
' Name: Processor_Info
'   ByRef Name As String - Nombre del procesador.
'   ByRef Identifier As String - Identificador del modelo y familia a la que pertenece.
'   ByRef VendorIdentifier As String - Nombre del fabricante.
'   ByRef Mhz As Integer - Megahercios del procesador. En plataformas Windows 9x (95/98/Me) esta variable devuelve 0.
' Purpose: Estructura que almacena la información del procesador principal.
' Remarks: Esta estructura se utiliza como valor de retorno en la funcione SYS_GetProcessorInfo.
'===============================================================================
Public Type Processor_Info
    Name As String 'Nombre del procesador.
    Identifier As String 'Identificador del modelo y familia.
    VendorIdentifier As String 'Nombre del fabricante.
    Mhz As Integer 'Megaherzios del procesador.

End Type

'Variables:
Dim m_Timer() As Long       'Cronometros.
Dim m_HITimer() As Currency 'Cronometros de alta precision.
Dim QueryFreq As Currency   'Frecuencia del cronometro de alta precision.

Dim TraceFile As String     'Indica el nombre del archivo donde se guardara el registro de sucesos.
Dim TraceBuffer() As String 'Buffer donde se almacenan los envios a la consola para despues ser guardados en un fichero.

'===============================================================================
' Name: FILE_Exists
' Input:
'   ByVal FileName As String - Nombre y ruta del archivo.
' Output:
'   Boolean - Devuelve verdadero si el archivo existe.
' Purpose: Determina si un archivo existe o no.
' Remarks:
'===============================================================================
Public Function FILE_Exists(Filename As String) As Boolean
Attribute FILE_Exists.VB_HelpID = 244
On Local Error Resume Next

FILE_Exists = CBool(Global_Mod.FileExists(Filename))

End Function

'===============================================================================
' Name: DIR_Exists
' Input:
'   ByVal Path As String - Nombre y ruta del directorio.
' Output:
'   Boolean - Devuelve verdadero si el directorio existe.
' Purpose: Determina si un directorio existe o no.
' Remarks:
'===============================================================================
Public Function DIR_Find(Path As String) As Boolean
Attribute DIR_Find.VB_HelpID = 243
On Local Error Resume Next

DIR_Find = CBool(Global_Mod.PathIsDirectory(Path))

End Function

'===============================================================================
' Name: DIR_Exists
' Input:
'   ByVal Path As String - Nombre y ruta del directorio.
' Output:
'   Boolean - Devuelve verdadero si el directorio esta vació.
' Purpose: Determina si un Directorio esta vació o no.
' Remarks:
'===============================================================================
Public Function DIR_IsEmpty(Path As String) As Boolean
Attribute DIR_IsEmpty.VB_HelpID = 242
On Local Error Resume Next

DIR_IsEmpty = CBool(Global_Mod.PathIsDirectoryEmpty(Path))

End Function

'===============================================================================
' Name: SYS_Beep
' Input:
'   ByVal Frec As Long - Frecuencia en hertzios de la señal.
'   ByVal Duration As Long - Duración en milisegundos de la señal.
' Output:
' Purpose: Reproduce una señal acústica mediante el altavoz interno de la CPU.
' Remarks: La señal acústica se procesa asincrónicamente. Mientras la señal se este procesando el programa detendrá su ejecución hasta que la señal finalice.
'===============================================================================
Public Sub SYS_Beep(Frec As Long, Duration As Long)
Attribute SYS_Beep.VB_HelpID = 241
Call APIBeep(Frec, Duration)

End Sub

'===============================================================================
' Name: SYS_ShowCursor
' Input:
'   ByVal Show As Boolean - Indica si se muestra o no el cursor del ratón de Windows.
' Output:
' Purpose: Muestra u oculta el cursor del ratón que dibuja Windows.
' Remarks: Esta función simplemente desactiva la presentación del cursor del ratón de Windows sobre las ventanas del programa pero no deshabilita la lectura de las coordenadas del cursor.
'===============================================================================
Public Sub SYS_ShowCursor(Show As Boolean)
Attribute SYS_ShowCursor.VB_HelpID = 240
Call ShowCursor(Show)

End Sub

'===============================================================================
' Name: INI_Read
' Input:
'   ByVal Filename As String - Nombre y ruta del archivo de configuración INI.
'   ByVal Key_Value As String - Valor de la sección donde se encuentra la clave a leer.
'   ByVal Key_Name As String - Nombre de la clave.
'   Optional ByVal Default As String - Valor por defecto que devolverá la funcion en caso de no existir la clave o de estar vacía.
' Output:
'   String - Cadena de texto con el valor de clave.
' Purpose: Lee una clave de un archivo de configuración INI.
' Remarks:
'===============================================================================
Public Function INI_Read(Filename As String, Key_Value As String, Key_Name As String, Optional ByVal Default As String) As String
Attribute INI_Read.VB_HelpID = 239
On Error GoTo ErrOut

Dim Size As Integer
Dim Value As String

'Comprobamos que el archivo existe.
If Not FILE_Exists(Filename) Then Err.Raise 53

'Se define el tamaño maximo de caracteres
'que podra tener la variable Value
Value = Space(200)
'Se utiliza la función para obtener
'el valor de la clave
Size = GetPrivateProfileString(Key_Value, Key_Name, vbNullString, Value, Len(Value), Filename)
'Si el tamaño es mayor a -1 entonces
'se ha encontrado el valor de la clave
If Size > 0 Then
    Value = VBA.Left$(Value, Size)
 
Else
    INI_Read = Default

End If

'Devolver el dato...
'Verificar que el dato no sea nulo,
'en caso de ser nulo de se devuelve
'el valor por defecto (Default)
If VBA.Right$(VBA.Trim$(Value), 1) = Chr(0) Then Value = VBA.Left$(VBA.Trim$(Value), Len(VBA.Trim$(Value)) - 1)

If Len(Value) Then
    INI_Read = VBA.Trim$(Value)
 
Else
    INI_Read = Default
 
End If

Exit Function

ErrOut:
    INI_Read = Default
    
End Function

'===============================================================================
' Name: INI_Write
' Input:
'   ByVal Filename As String - Nombre y ruta del archivo de configuración INI. Si el archivo no existe se crea automáticamente.
'   ByVal Key_Value As String - Valor de la sección donde se encuentra la clave. Si la clave no existe se crea automáticamente.
'   ByVal Key_Name As String - Nombre de la clave. Si la clave no existe se crea una automaticamente.
'   ByVal Value As String - Contenido que se escribirá en la clave.
' Output:
'   String - Cadena de texto con el valor de clave.
' Purpose: Escribe una clave de un archivo de configuración INI.
' Remarks: No es preciso crear un archivo INI previamente, si este no existe, al igual que alguna sección o clave, se creara automáticamente con los valores pasados en los argumentos.
'===============================================================================
Public Sub INI_Write(Filename As String, Key_Value As String, Key_Name As String, Value As String)
Attribute INI_Write.VB_HelpID = 238
On Error GoTo ErrOut

Dim Size As Integer

'Escribimos el valor de la clave en el INI
Size = WritePrivateProfileString(Key_Value, Key_Name, Value, Filename)
Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: REG_GetString
' Input:
'   ByVal HKey As REG_Path - Ruta principal del registro al que accederemos.
'   ByVal Path As String - Ruta de la sección.
'   ByVal Key_Name As String - Nombre de la clave.
' Output:
'   String - Cadena de texto almacenada en la clave del registro.
' Purpose: Lee una clave de texto del Registro de Windows.
' Remarks:
'===============================================================================
Public Function REG_GetString(HKey As REG_Path, Path As String, Key_Name As String) As String
Attribute REG_GetString.VB_HelpID = 236
On Local Error Resume Next

Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim lValueType As Long

Call RegOpenKey(HKey, Path, keyhand)
lResult = RegQueryValueEx(keyhand, Key_Name, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, Key_Name, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            REG_GetString = VBA.Left$(strBuf, intZeroPos - 1)
        Else
            REG_GetString = strBuf
        End If
    End If
End If

End Function

'===============================================================================
' Name: REG_SaveString
' Input:
'   ByVal HKey As REG_Path - Ruta principal del registro al que accederemos.
'   ByVal Path As String - Ruta de la sección.
'   ByVal Key_Name As String - Nombre de la clave.
'   ByVal Data As String - Cadena de texto que se almacenara en la clave.
' Output:
'   Long - Devuelve SYS_OK en caso de realizar la operación correctamente. Si la ruta no es valida se devuelve SYS_INVALIDPATH, y si el valor no es valido se devuelve SYS_INVALIDARG.
' Purpose: Escribe una cadena de texto en el registro de Windows.
' Remarks:
'===============================================================================
Public Function REG_SaveString(HKey As REG_Path, Path As String, Key_Name As String, Data As String) As Long
Attribute REG_SaveString.VB_HelpID = 235
On Local Error Resume Next

Dim keyhand As Long

If RegCreateKey(HKey, Path, keyhand) = ERROR_SUCCESS Then
    If RegSetValueEx(keyhand, Key_Name, 0, REG_SZ, ByVal Data, Len(Data)) = ERROR_SUCCESS Then
        Call RegCloseKey(keyhand)
        REG_SaveString = SYS_OK
        Exit Function
    
    End If
    
    REG_SaveString = SYS_INVALIDARG

End If

REG_SaveString = SYS_INVALIDPATH

End Function

'===============================================================================
' Name: REG_GetDWord
' Input:
'   ByVal HKey As REG_Path - Ruta principal del registro al que accederemos.
'   ByVal Path As String - Ruta de la sección.
'   ByVal Key_Name As String - Nombre de la clave.
' Output:
'   Long - Valor almacenado en la clave del registro.
' Purpose: Lee una clave numérica del registro de Windows.
' Remarks:
'===============================================================================
Public Function REG_GetDWord(ByVal HKey As REG_Path, ByVal Path As String, ByVal Key_Name As String) As Long
Attribute REG_GetDWord.VB_HelpID = 234
On Local Error Resume Next

Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

r = RegOpenKey(HKey, Path, keyhand)

lDataBufSize = 4
    
lResult = RegQueryValueEx(keyhand, Key_Name, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        REG_GetDWord = lBuf
    End If

End If

r = RegCloseKey(keyhand)
    
End Function

'===============================================================================
' Name: REG_SaveDWord
' Input:
'   ByVal HKey As REG_Path - Ruta principal del registro al que accederemos.
'   ByVal Path As String - Ruta de la sección.
'   ByVal Key_Name As String - Nombre de la clave.
'   ByVal Data As Long - Valor numérico que se almacenara en la clave.
' Output:
'   Long - Devuelve SYS_OK en caso de realizar la operación correctamente. Si la ruta no es valida se devuelve SYS_INVALIDPATH, y si el valor no es valido se devuelve SYS_INVALIDARG.
' Purpose: Escribe un valor numérico en el registro de Windows.
' Remarks:
'===============================================================================
Public Function REG_SaveDword(ByVal HKey As REG_Path, ByVal Path As String, ByVal Key_Name As String, ByVal Data As Long) As Long
Attribute REG_SaveDword.VB_HelpID = 233
On Local Error Resume Next

Dim keyhand As Long

If RegCreateKey(HKey, Path, keyhand) = ERROR_SUCCESS Then
    If RegSetValueEx(keyhand, Key_Name, 0&, REG_DWORD, Data, 4) = ERROR_SUCCESS Then
        Call RegCloseKey(keyhand)
        REG_SaveDword = SYS_OK
        Exit Function
        
    End If
    
    REG_SaveDword = SYS_INVALIDARG
    
End If

REG_SaveDword = SYS_INVALIDPATH
    
End Function

'===============================================================================
' Name: REG_DeleteKeySection
' Input:
'   ByVal HKey As REG_Path - Ruta principal del registro al que accederemos.
'   ByVal Path As String - Ruta de la sección.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Borra una sección del registro de Windows.
' Remarks:
'===============================================================================
Public Function REG_DeleteKeySection(ByVal HKey As REG_Path, ByVal Path As String) As Boolean
Attribute REG_DeleteKeySection.VB_HelpID = 232
On Local Error Resume Next

If RegDeleteKey(HKey, Path) = ERROR_SUCCESS Then REG_DeleteKeySection = True

End Function

'===============================================================================
' Name: REG_DeleteKeyValue
' Input:
'   ByVal HKey As REG_Path - Ruta principal del registro al que accederemos.
'   ByVal Path As String - Ruta de la sección.
'   ByVal Key_Name As String - Ruta de la clave.
' Output:
'   Long - Devuelve SYS_OK en caso de realizar la operación correctamente. Si la ruta no es valida se devuelve SYS_INVALIDPATH, y si el valor no existe se devuelve SYS_KEYNOFOUND.
' Purpose: Borra una clave del registro de Windows.
' Remarks:
'===============================================================================
Public Function REG_DeleteKeyValue(ByVal HKey As REG_Path, ByVal Path As String, ByVal Key_Name As String) As Long
Attribute REG_DeleteKeyValue.VB_HelpID = 231
On Local Error Resume Next

Dim keyhand As Long

If RegOpenKey(HKey, Path, keyhand) = ERROR_SUCCESS Then
    If RegDeleteValue(keyhand, Key_Name) = ERROR_SUCCESS Then
        Call RegCloseKey(keyhand)
        REG_DeleteKeyValue = SYS_OK
        Exit Function
    
    End If
    
    REG_DeleteKeyValue = SYS_KEYNOFOUND
    
End If

REG_DeleteKeyValue = SYS_INVALIDPATH

End Function

'===============================================================================
' Name: SYS_GetMemory
' Input:
'   ByVal Data As Memory_Info - Parámetro de salida que nos devuelve la informacion de la memoria del sistema.
' Output:
' Purpose: Devuelve la información sobre la memoria del sistema.
' Remarks:
'===============================================================================
Public Sub SYS_GetMemory(Data As Memory_Info)
Attribute SYS_GetMemory.VB_HelpID = 230
On Error GoTo ErrOut

Dim Inf As MEMORYSTATUS

Call apiMemStatus(Inf)

With Data
    .TotalPhys = Inf.dwTotalPhys
    .AvailPhys = Inf.dwAvailPhys
    .TotalVirtual = Inf.dwTotalVirtual
    .AvailVirtual = Inf.dwAvailVirtual
    .TotalPageFile = Inf.dwTotalPageFile
    .AvailPageFile = Inf.dwAvailPageFile
End With

Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: SYS_LockWinKeys
' Input:
'   ByVal Locked As Boolean - Desactiva o no las combinaciones de teclas especiales de Windows.
' Output:
' Purpose: Desactiva el juego de combinaciones de teclas especiales de Windows.
' Remarks: Esta función desactiva el juego de combinaciones de teclas especiales de Windows, tales como por ejemplo Ctrl+Alt+Supr, Ctrl+Esc, Alt+Tab, Logo Windows, tecla de menú contextual, etc... <br><b>Importante:</b> Esta función solo esta soportada en las plataformas Windows 9x (95/98/Me). En plataformas NT no hace efecto alguno.
'===============================================================================
Public Sub SYS_LockWINKeys(Locked As Boolean)
Attribute SYS_LockWINKeys.VB_HelpID = 229
On Local Error Resume Next

Call SystemParametersInfo(SPI_ENABLEWINKEYS, Locked, CStr(1), 0)

End Sub

'===============================================================================
' Name: SYS_ScreenCapture
' Input:
'   ByVal Filename As String - Nombre y ruta del archivo que se creara.
'   ByVal Mode As Capture_Mode - Modo de captura.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Realiza una captura de la pantalla o de la ventana que tiene el foco y la guarda en un archivo de formato Mapa de Bits de Windows (*.BMP).
' Remarks: Esta función no funciona correctamente con la clase grafica. Si desea hacer una captura de pantalla de un programa ejecutando la clase grafica dx_GFX deberá llamar a la función apropiada en la clase, SURF_ScreenCapture.
'===============================================================================
Public Function SYS_ScreenCapture(Filename As String, Mode As Capture_Mode) As Boolean
Attribute SYS_ScreenCapture.VB_HelpID = 228
On Error GoTo ErrOut
Dim SurfPic As Picture
    
    If Mode > 1 Then Mode = 1 Else If Mode < 0 Then Mode = 0
    
    Call keybd_event(44, Mode, 0, 0)  'Simula la pulsacion de la tecla de Imprimir Pantalla.
    
    If Not Clipboard.GetFormat(vbCFBitmap) Then GoTo ErrOut
    
    Set SurfPic = Clipboard.GetData(vbCFBitmap)  'Creamos la captura.
    Call SavePicture(SurfPic, Filename)  'Guardamos la captura.
    
    SYS_ScreenCapture = True
    
Exit Function
    
ErrOut:
    Set SurfPic = Nothing

End Function

'===============================================================================
' Name: SYS_GetPath
' Input:
'   ByVal Path As SYS_Path - Directorio del sistema del que obtendremos su ruta.
' Output:
' Purpose: Devuelve la ruta de un directorio del sistema.
' Remarks: Esta función devuelve la ruta de los directorios Windows, System y temporal de Windows o del usuario.
'===============================================================================
Public Function SYS_GetPath(Path As SYS_Path) As String
Attribute SYS_GetPath.VB_HelpID = 227
On Local Error Resume Next

Dim Temp As String * 256
Dim X As Integer
    
    Select Case Path
        Case 0 'Directorio Windows:
            X = GetWindowsDirectory(Temp, Len(Temp))
            
        Case 1 'Directorio System:
            X = GetSystemDirectory(Temp, Len(Temp))
        
        Case 2 'Directorio Temp (Directorio Temporal):
            X = GetTempPath(Len(Temp), Temp)
    
    End Select
    
    SYS_GetPath = VBA.Left$(Temp, X)
    
    If Path = TEMP_DIR Then SYS_GetPath = VBA.Left$(SYS_GetPath, Len(SYS_GetPath) - 1)
    
End Function

'===============================================================================
' Name: DLG_SaveFile
' Input:
'   ByVal hWnd As Long - Identificador de la ventana que hará de cliente. Si no existe ventana cliente este parámetro se establecerá a 0.
'   ByVal Filter As String - Cadena de texto que define el filtro de las extensiones de archivo que se mostraran. Por ejemplo: "Archivos de texto *.TXT|*.txt|Todos los archivos *.*|*.*". También se puede definir un mismo elemento del filtro para que reconozca varios tipos de extensiones, un ejemplo: "Todos los archivos de imagen|*.bmp;*.gif;*.jpg,*.png".
'   ByVal Title As String - Cadena de texto que define el titulo de la ventana de dialogo.
'   ByVal InitDir As String - Ruta del directorio que mostrara el explorador. Si no se especifica ruta se toma la ruta por defecto.
'   Optional ByVal Filename As String - Nombre de archivo que se mostrara por defecto en el cuadro de texto <b>Nombre</b>.
'   Optional ByVal FilterIndex As Long - Índice de la lista de filtros que indica la extensión de archivo por defecto que se mostrara en explorador.
' Output:
'   String - Nombre y ruta del archivo establecido.
' Purpose: Abre el cuadro de dialogo de guardar archivo de Windows.
' Remarks: Los cuadros de dialogo de Windows son interfaces de ayuda que solo devuelven información, no realizan ninguna acción por el usuario como por ejemplo la escritura en archivos.
'===============================================================================
Public Function DLG_SaveFile(hWnd As Long, Filter As String, Title As String, InitDir As String, Optional Filename As String, Optional FilterIndex As Long) As String
Attribute DLG_SaveFile.VB_HelpID = 226
On Local Error Resume Next

    Dim ofn As OPENFILENAME
    Dim A As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hWnd
    ofn.hInstance = App.hInstance
    
    If VBA.Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For A = 1 To Len(Filter)
        If Mid(Filter, A, 1) = "|" Then Mid(Filter, A, 1) = Chr(0)
    Next
    
        ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        If Not Filename = vbNullString Then ofn.lpstrFile = Filename & Space(254 - Len(Filename))
        ofn.nFilterIndex = FilterIndex
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT Or OFN_EXPLORER
        A = GetSaveFileName(ofn)

        If A Then
            DLG_SaveFile = Trim$(ofn.lpstrFile)
            If VBA.Right$(Trim$(DLG_SaveFile), 1) = Chr(0) Then DLG_SaveFile = VBA.Left$(Trim$(ofn.lpstrFile), Len(Trim$(ofn.lpstrFile)) - 1)
            
            Dim Ext As String
            Ext = GetExtension(ofn.lpstrFilter, ofn.nFilterIndex)
            If Not UCase(Right(DLG_SaveFile, 4)) = UCase(Ext) Then DLG_SaveFile = DLG_SaveFile + Ext

        Else
            DLG_SaveFile = vbNullString
            
        End If
        
End Function

'===============================================================================
' Name: DLG_OpenFile
' Input:
'   ByVal hWnd As Long - Identificador de la ventana que hará de cliente. Si no existe ventana cliente este parámetro se establecerá a 0.
'   ByVal Filter As String - Cadena de texto que define el filtro de las extensiones de archivo que se mostraran. Por ejemplo: "Archivos de texto *.TXT|*.txt|Todos los archivos *.*|*.*". También se puede definir un mismo elemento del filtro para que reconozca varios tipos de extensiones, un ejemplo: "Todos los archivos de imagen|*.bmp;*.gif;*.jpg,*.png".
'   ByVal Title As String - Cadena de texto que define el titulo de la ventana de dialogo.
'   ByVal InitDir As String - Ruta del directorio que mostrara el explorador. Si no se especifica ruta se toma la ruta por defecto.
'   Optional ByVal Filename As String - Nombre de archivo que se mostrara por defecto en el cuadro de texto <b>Nombre</b>.
'   Optional ByVal FilterIndex As Long - Índice de la lista de filtros que indica la extensión de archivo por defecto que se mostrara en explorador.
' Output:
'   String - Nombre y ruta del archivo establecido.
' Purpose: Abre el cuadro de dialogo de abrir archivo de Windows.
' Remarks: Los cuadros de dialogo de Windows son interfaces de ayuda que solo devuelven información, no realizan ninguna acción por el usuario como por ejemplo la escritura en archivos.
'===============================================================================
Public Function DLG_OpenFile(hWnd As Long, Filter As String, Title As String, InitDir As String, Optional Filename As String, Optional FilterIndex As Long) As String
Attribute DLG_OpenFile.VB_HelpID = 225
On Local Error Resume Next

    Dim ofn As OPENFILENAME
    Dim A As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hWnd
    ofn.hInstance = App.hInstance
    
    If VBA.Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid(Filter, A, 1) = Chr(0)
    Next
    
        ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        If Not Filename = vbNullString Then ofn.lpstrFile = Filename & Space$(254 - Len(Filename))
        ofn.nFilterIndex = FilterIndex
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        A = GetOpenFileName(ofn)

        If A Then
            DLG_OpenFile = Trim$(ofn.lpstrFile)
            If VBA.Right$(VBA.Trim$(DLG_OpenFile), 1) = Chr(0) Then DLG_OpenFile = VBA.Left$(VBA.Trim$(ofn.lpstrFile), Len(VBA.Trim$(ofn.lpstrFile)) - 1)
            
        Else
            DLG_OpenFile = vbNullString
            
        End If
        
End Function

'===============================================================================
' Name: MATH_Rand
' Input:
'   ByVal Max As Single - Valor máximo del rango de selección.
'   ByVal Min As Single - Valor mínimo del rango de selección.
' Output:
'   Single - Devuelve un numero aleatorio según el intervalo establecido.
' Purpose: Devuelve un numero aleatorio entre el primer numero y el segundo.
' Remarks: Para evitar secuencias idénticas cada vez que llame a esta función utilice la función MATH_RandSeed para generar una nueva semilla en el generador de números aleatorios.
'===============================================================================
Public Function MATH_Rand(Max As Single, Min As Single) As Single
Attribute MATH_Rand.VB_HelpID = 224
On Local Error Resume Next

MATH_Rand = CSng((Max - Min + 1) * Rnd + Min)
'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)

End Function

'===============================================================================
' Name: MATH_RandSeed
' Input:
'   ByVal Seed As Single - Valor que se utilizara como semilla para el generador de numeros aleatorios.
' Output:
' Purpose: Establece una semilla para crear un numero aleatorio.
' Remarks: El uso de MATH_RandSeed es similar a la instrucción <b>Randomize</b> de Visual Basic. La semilla generada con Randomize es compatible para la función MATH_Rand.
'===============================================================================
Public Sub MATH_RandSeed(Seed As Single)
Attribute MATH_RandSeed.VB_HelpID = 223
On Local Error Resume Next

Call Randomize(Seed)

End Sub

'===============================================================================
' Name: SYS_GetTime
' Input:
'   ByVal Data As Sys_Time - Parámetro de salida que nos devuelve la fecha y hora del sistema.
' Output:
' Purpose: Devuelve la fecha y hora del sistema.
' Remarks: Esta función devuelve por separado el año, mes, día, hora, minuto, segundo, milisegundo y día de la semana del sistema.
'===============================================================================
Public Sub SYS_GetTime(Data As Sys_Time)
Attribute SYS_GetTime.VB_HelpID = 222
On Error GoTo ErrOut

Dim Inf As SYSTEMTIME

Call GetLocalTime(Inf)

With Data
    .Day = CLng(Inf.wDay)
    .Month = CLng(Inf.wMonth)
    .Year = CLng(Inf.wYear)
    .Hour = CLng(Inf.wHour)
    .Minute = CLng(Inf.wMinute)
    .Second = CLng(Inf.wSecond)
    .Milliseconds = CLng(Inf.wMilliseconds)
    .DayOfWeek = CLng(Inf.wDayOfWeek)
    
End With

ErrOut:

End Sub

'===============================================================================
' Name: SYS_GetDirectXVersion
' Input:
'   ByVal Version As Long - Valor que define la versión de las librerías de DirectX.
'   ByVal Revision As Long - Valor que define la revisión de la versión de las librerias de DirectX.
' Output:
' Purpose: Devuelve la versión de DirectX instalada en el sistema.
' Remarks:
'===============================================================================
Public Sub SYS_GetDirectXVersion(Version As Long, Revision As Long)
Attribute SYS_GetDirectXVersion.VB_HelpID = 221
On Error GoTo ErrOut

Dim lMajor As Long, lMinor As Long, Str As String

If DirectXSetupGetVersion(lMajor, lMinor) <> 1 Then Exit Sub

lMajor = lMajor - (lMajor And &H40000)
lMinor = (lMinor - (100 * lMajor))

Version = lMajor
Revision = lMinor

ErrOut: 'Si no se encuentra la dll significa que DirectX no esta instalado o no se ha encontrado la dll.

End Sub

'===============================================================================
' Name: SYS_GetOSInfo
' Input:
'   ByVal Data As OS_Info - Parámetro de salida que nos devuelve la información del sistema operativo.
' Output:
' Purpose: Devuelve información sobre el Sistema Operativo.
' Remarks:
'===============================================================================
Public Sub SYS_GetOSInfo(Data As OS_Info)
Attribute SYS_GetOSInfo.VB_HelpID = 220
On Error GoTo ErrOut

Dim Inf As OSVERSIONINFO

Inf.dwOSVersionInfoSize = 148

Call GetVersionEx(Inf)

With Data
    .MajorVersion = Inf.dwMajorVersion
    .MinorVersion = Inf.dwMinorVersion
    .PlatformID = Inf.dwPlatformId
    
    If .PlatformID = 1 Then
        .BuildNumber = Inf.dwBuildNumber And &HFFFF&
        .WindowsId = REG_GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName")
    ElseIf .PlatformID = 2 Then
        .BuildNumber = Inf.dwBuildNumber
        .WindowsId = REG_GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductName")
    End If
    
End With

Exit Sub

ErrOut:

End Sub

'===============================================================================
' Name: SYS_EnableScreenSaver
' Input:
'   ByVal Enable As Boolean - Indica si el salvapantallas se ejecutara o no.
' Output:
' Purpose: Activa o desactiva el evento que ejecuta el salvapantallas de Windows.
' Remarks:
'===============================================================================
Public Sub SYS_EnableScreenSaver(Enable As Boolean)
Attribute SYS_EnableScreenSaver.VB_HelpID = 219
On Local Error Resume Next

Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, Enable, CStr(1), 0)

End Sub

'===============================================================================
' Name: DLG_BrowseFolder
' Input:
'   ByVal hWnd As Long - Identificador de la ventana que hará de cliente. Si no existe ventana cliente este parámetro se establecerá a 0.
'   ByVal Title As String - Cadena de texto que define el titulo de la ventana de dialogo.
' Output:
'   String - Nombre y ruta del directorio seleccionado.
' Purpose: Abre el cuadro de dialogo de selección de directorios de Windows.
' Remarks: Los cuadros de dialogo de Windows son interfaces de ayuda que solo devuelven información, no realizan ninguna acción por el usuario como por ejemplo la escritura en archivos.
'===============================================================================
Public Function DLG_BrowseFolder(hWnd As Long, Title As String) As String
Attribute DLG_BrowseFolder.VB_HelpID = 218
On Local Error Resume Next

Dim mBrowseInfo As BROWSEINFO
Dim mPointerToIDList As Long
Dim mResult As Long
Dim mPathBuffer As String
Dim sReturn As String

        ' this string is returned as result if you have canceled the action
         'sReturn = "You canceled operation"
         sReturn = ""
         With mBrowseInfo
                     ' set parent window (hWnd of this window)
                      .hwndOwner = hWnd
                      ' set start folder (0=desktor folder)
                      .pidlRoot = 0
                      ' dialog title
                      .lpszTitle = Title
                     ' pointer to a buffer that receives the display name of the folder selected by the user
                      .pszDisplayName = String(MAX_SIZE, Chr(0))
                      ' value specifying the types of folders to be listed in the dialog box as well as other options
                      .ulFlags = BIF_RETURNONLYFSDIRS
         End With
        ' returns a pointer to an item identifier list that specifies the location of the selected folder relative to the root of the name space
         mPointerToIDList = BrowseFolderDlg(mBrowseInfo)

         If mPointerToIDList <> 0& Then
                      ' create a buffer
                      mPathBuffer = String(MAX_SIZE, Chr(0))
                      ' now get the selected path
                      mResult = GetPathFromIDList(ByVal mPointerToIDList, ByVal mPathBuffer)
                      ' returned path
                      sReturn = VBA.Left$(mPathBuffer, InStr(mPathBuffer, Chr(0)) - 1)  'and return just that
         End If

DLG_BrowseFolder = sReturn

End Function

'===============================================================================
' Name: PAK_Load
' Input:
'   ByVal Filename As String - Nombre y ruta del archivo PAK.
'   ByVal FileList() As PAK_FileInfo - Parámetro de salida que genera una lista con la informacion de los archivos contenidos en el paquete.
'   Optional ByVal Files As Long - Parámetro de salida que devuelve el numero de archivos contenidos en el paquete.
'   Optional ByVal Size As Long - Parámetro de salida que devuelve el tamaño en bytes del paquete.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_FILENOTFOUND si el archivo no se encuentra en la ruta especificada, SYS_INVALIDFORMAT si no el formato del archivo no es correcto, SYS_EMPTYLIST si el paquete no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Lee el contenido de un paquete y genera una lista con la información.
' Remarks:
'===============================================================================
Public Function PAK_Load(Filename As String, FileList() As PAK_FileInfo, Optional Files As Long, Optional Size As Long) As Long
Attribute PAK_Load.VB_HelpID = 217
On Error GoTo ErrOut

Dim PakHeader As PAKFileHeader, PakFiles() As FileInPAK, i As Long, hFile As Integer
    
    'Comprobamos que el archivo existe:
    If Not CBool(Global_Mod.FileExists(Filename)) Then
        PAK_Load = SYS_FILENOTFOUND
        Exit Function
    
    End If
    
    'Comprobamos si se trata de un archivo *.PAK:
    If Not PAK_IsPAKFormat(Filename) Then
        PAK_Load = SYS_INVALIDFORMAT
        Exit Function
    
    End If
    
    'Abrimos el archivo para lectura en modo binario:
    hFile = FreeFile
    Open Filename For Binary Lock Read As hFile
    
    'Leemos la cabecera:
    Get #hFile, , PakHeader
    
    'Esta linea solo recuerda el uso de dirlen y dirofs de la cabecera:
    'Debug.Print Int(PakHeader.dirlen / 64) & " archivos" & " en " & CInt(PakHeader.dirofs / 1024 ^ 2) & " Mb (" & LOF(1) & " Bytes)"
    
    'Se comprueba si el paquete contiene archivos:
    If Int(PakHeader.dirlen / 64) = 0 Then
        PAK_Load = SYS_EMPTYLIST
        Exit Function
    
    End If
    
    'Damos tamaño al array para crear una lista de los archivos contenidos en el paquete:
    ReDim PakFiles(Int(PakHeader.dirlen / 64)) As FileInPAK
    ReDim FileList(UBound(PakFiles) - 1) As PAK_FileInfo
    
    Seek #hFile, PakHeader.dirofs + 1

    Get #hFile, , PakFiles()
    
    'Creamos la lista:
    For i = 0 To UBound(PakFiles) - 1
        FileList(i).Filename = VBA.Trim$(VBA.Left$(PakFiles(i).named, InStr(PakFiles(i).named, Chr(0)) - 1))
        FileList(i).Offset = PakFiles(i).filepos
        FileList(i).Size = PakFiles(i).filelen
    Next i
    
    Close #hFile
    
    Files = Int(PakHeader.dirlen / 64) 'Archivos empaquetados.
    Size = PakHeader.dirofs 'Tamaño en bytes.

PAK_Load = SYS_OK
Exit Function

ErrOut:
Close #hFile

If Err.Number = 9 Then PAK_Load = SYS_EMPTYLIST Else PAK_Load = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: PAK_ExtractFile
' Input:
'   ByVal PAKFile As String - Nombre y ruta del archivo PAK.
'   ByVal GetFile As PAK_FileInfo - Información del archivo a extraer.
'   ByVal SaveFile As String - Nombre y ruta donde se extraerá el archivo.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Extrae un archivo de un archivo paquete PAK y lo guarda en disco.
' Remarks: Para obtener la información del parámetro GetFile primero se deberá obtener una lista con la información de los archivos contenidos en el paquete mediante PAK_Load.
'===============================================================================
Public Function PAK_ExtractFile(PAKFile As String, GetFile As PAK_FileInfo, SaveFile As String) As Boolean
Attribute PAK_ExtractFile.VB_HelpID = 216
On Error GoTo ErrOut

Dim Data() As Byte
Dim i As Long, hFile(1) As Integer
    
    'Accedemos al paquete para leer el archivo:
    hFile(0) = FreeFile
    Open PAKFile For Binary Lock Read As hFile(0)
    
    'Creamos el archivo en disco:
    hFile(1) = FreeFile
    Open SaveFile For Binary As hFile(1)
    
        ReDim Data(GetFile.Size - 1) As Byte
        
        Get #hFile(0), GetFile.Offset + 1, Data
        Put #hFile(1), , Data
        
    Close #hFile(1)
    Close #hFile(0)

PAK_ExtractFile = True

ErrOut:
Close #hFile(1)
Close #hFile(0)

End Function

'===============================================================================
' Name: PAK_ExtractFileInMemory
' Input:
'   ByVal PAKFile As String - Nombre y ruta del archivo PAK.
'   ByVal GetFile As PAK_FileInfo - Información del archivo a extraer.
'   ByVal FileData() As Byte - Array de bytes donde se almacenara el archivo.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Extrae un archivo de un archivo paquete PAK y lo almacena en memoria.
' Remarks: Para obtener la información del parámetro GetFile primero se deberá obtener una lista con la información de los archivos contenidos en el paquete mediante PAK_Load.
'===============================================================================
Public Function PAK_ExtractFileInMemory(PAKFile As String, GetFile As PAK_FileInfo, FileData() As Byte) As Boolean
Attribute PAK_ExtractFileInMemory.VB_HelpID = 215
On Error GoTo ErrOut

Dim i As Long, hFile As Integer

    'Accedemos al paquete para leer el archivo:
    hFile = FreeFile
    Open PAKFile For Binary Lock Read As hFile
        
        ReDim FileData(GetFile.Size - 1) As Byte
        
        Get #hFile, GetFile.Offset + 1, FileData
        
    Close #hFile

PAK_ExtractFileInMemory = True

ErrOut:
Close #hFile

End Function

'===============================================================================
' Name: PAK_ExtractFileInMemory
' Input:
'   ByVal Filename As String - Nombre y ruta del archivo PAK.
' Output:
'   Boolean - Devuelve verdadero si formato del archivo coincide con el formato sin compresión estándar PAK de los motores de: Quake, Quake2, Heretic2 y Half-Life.
' Purpose: Comprueba que se trata de un archivo con formato PAK.
' Remarks: No es necesario que compruebe el formato de un archivo que va a ser cargado con PAK_Load ya que dicha comprobación se realiza automáticamente. Esta función solo esta pensada para realizar comprobaciones en otras acciones.
'===============================================================================
Public Function PAK_IsPAKFormat(Filename As String) As Boolean
Attribute PAK_IsPAKFormat.VB_HelpID = 214
On Local Error Resume Next

Dim Header As String

Header = ID_PAK_Header '_DEFAULT

    If Trim(Filename) <> "" And Dir(Filename) <> "" Then
        Dim TmpHandle As Integer
        Dim strHeader As String
        
        TmpHandle = FreeFile
        strHeader = Space(Len(Header))
        
        Open Filename For Binary As #TmpHandle
        Get #TmpHandle, , strHeader
        Close #TmpHandle
        
        If strHeader = Header Then PAK_IsPAKFormat = True
    End If
    
End Function

'===============================================================================
' Name: FILE_List
' Input:
'   ByVal List() As String - Parámetro de salida que devuelve una lista con los archivos y subdirectorios contenidos en un directorio especifico.
'   Optional ByVal Path As String - Ruta del directorio que leeremos. Se pueden usar comodines para filtrar extensiones, por ejemplo "C:\Windows\*.*", "C:\Windows\*.bmp"...
'   Optional ByVal Attrib As File_Attribute - Constante que define el atributo de los archivos que se listaran.
'   Optional ByVal Sort As Boolean - Indica si se ordenan los elementos de la lista.
'   Optional ByVal SortMode As Sort_Mode - Constante que define el modo de ordenación.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Lista los archivos y subdirectorios contenidos en una ruta especifica.
' Remarks:
'===============================================================================
Public Function FILE_List(list() As String, Optional Path As String, Optional Attrib As File_Attribute, Optional Sort As Boolean, Optional SortMode As Sort_Mode) As Long
Attribute FILE_List.VB_HelpID = 213
    On Error GoTo ErrOut
    
    Dim i As Long, j As Long
    Dim Text As String
    Dim Data As WIN32_FIND_DATA
    Dim Handle As Long
    Dim Ret As Long
    
    If Path = vbNullString Then Path = "*.*"
    
    Handle = FindFirstFile(Path, Data)
    
    Do
        If Attrib = 0 Or Data.dwFileAttributes = Attrib Then
            ReDim Preserve list(i) As String
            list(i) = VBA.Trim$(VBA.Left$(Data.cFileName, InStr(Data.cFileName, Chr(0)) - 1))
            i = i + 1
        End If
    
        Ret = FindNextFile(Handle, Data)
        DoEvents
        
    Loop Until Ret = 0
                    
    If Sort Then Call SORT_StringList(list(), SortMode)
        
    Call FindClose(Handle)
    
    FILE_List = SYS_OK
    Exit Function
    
ErrOut:
    If Err.Number = 9 Then FILE_List = SYS_EMPTYLIST Else FILE_List = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_StringList
' Input:
'   ByVal List() As String - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenacion.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de cadenas de texto.
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_StringList(list() As String, SortMode As Sort_Mode) As Long
Attribute SORT_StringList.VB_HelpID = 212
On Error GoTo ErrOut

Dim Temp As String, X As Integer, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 0 And UCase(list(X)) > UCase(list(X + 1)) Or SortMode = 1 And UCase(list(X)) < UCase(list(X + 1)) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_StringList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_StringList = SYS_EMPTYLIST Else SORT_StringList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_ByteList
' Input:
'   ByVal List() As Byte - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenación.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de bytes.
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_ByteList(list() As Byte, Optional SortMode As Sort_Mode) As Long
Attribute SORT_ByteList.VB_HelpID = 211
On Error GoTo ErrOut

Dim Temp As Byte, X As Long, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 1 And list(X) < list(X + 1) Or SortMode = 0 And list(X) > list(X + 1) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_ByteList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_ByteList = SYS_EMPTYLIST Else SORT_ByteList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_IntegerList
' Input:
'   ByVal List() As Integer - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenacion.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de números enteros cortos (2 bytes)
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_IntegerList(list() As Integer, Optional SortMode As Sort_Mode) As Long
Attribute SORT_IntegerList.VB_HelpID = 210
On Error GoTo ErrOut

Dim Temp As Integer, X As Long, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 1 And list(X) < list(X + 1) Or SortMode = 0 And list(X) > list(X + 1) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_IntegerList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_IntegerList = SYS_EMPTYLIST Else SORT_IntegerList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_LongList
' Input:
'   ByVal List() As Long - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenación.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de números enteros largos (4 bytes)
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_LongList(list() As Long, Optional SortMode As Sort_Mode) As Long
Attribute SORT_LongList.VB_HelpID = 209
On Error GoTo ErrOut

Dim Temp As Long, X As Long, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 1 And list(X) < list(X + 1) Or SortMode = 0 And list(X) > list(X + 1) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_LongList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_LongList = SYS_EMPTYLIST Else SORT_LongList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_SingleList
' Input:
'   ByVal List() As Single - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenación.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de números decimales de precisión simple (4 bytes)
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_SingleList(list() As Single, Optional SortMode As Sort_Mode) As Long
Attribute SORT_SingleList.VB_HelpID = 208
On Error GoTo ErrOut

Dim Temp As Single, X As Long, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 1 And list(X) < list(X + 1) Or SortMode = 0 And list(X) > list(X + 1) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_SingleList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_SingleList = SYS_EMPTYLIST Else SORT_SingleList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_DoubleList
' Input:
'   ByVal List() As Double - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenación.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de números decimales de doble precisión (8 bytes)
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_DoubleList(list() As Double, Optional SortMode As Sort_Mode) As Long
Attribute SORT_DoubleList.VB_HelpID = 207
On Error GoTo ErrOut

Dim Temp As Double, X As Long, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 1 And list(X) < list(X + 1) Or SortMode = 0 And list(X) > list(X + 1) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_DoubleList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_DoubleList = SYS_EMPTYLIST Else SORT_DoubleList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: SORT_CurrencyList
' Input:
'   ByVal List() As Currency - Lista a ordenar.
'   ByVal SortMode As Sort_Mode - Constante que define el modo de ordenación.
' Output:
'   Long - Devuelve SYS_OK si la operación se ha realizado con éxito, SYS_EMPTYLIST si el directorio no contiene archivos y SYS_UNKNOWNERROR en caso de ocurrir cualquier otro error.
' Purpose: Ordena una lista de números decimales de punto fijo o decimal de alta precision (8 bytes)
' Remarks: La lista ordenada sobrescribe a lista original introducida en el parámetro List.
'===============================================================================
Public Function SORT_CurrencyList(list() As Currency, Optional SortMode As Sort_Mode) As Long
On Error GoTo ErrOut

Dim Temp As Currency, X As Long, Sorted As Boolean

X = UBound(list)

If SortMode < 0 Then SortMode = 0 Else If SortMode > 1 Then SortMode = 1

Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(list) - 1
    If SortMode = 1 And list(X) < list(X + 1) Or SortMode = 0 And list(X) > list(X + 1) Then
        Temp = list(X + 1)
        list(X + 1) = list(X)
        list(X) = Temp
        Sorted = False
    End If
    DoEvents
Next X
DoEvents
Loop

SORT_CurrencyList = SYS_OK
Exit Function

ErrOut:
If Err.Number = 9 Then SORT_CurrencyList = SYS_EMPTYLIST Else SORT_CurrencyList = SYS_UNKNOWNERROR

End Function

'===============================================================================
' Name: MATH_Atan2
' Input:
'   ByVal X As Double - Valor de X.
'   ByVal Y As Double - Valor de Y.
' Output:
'   Double - Devuelve el arco tangente en radianes.
' Purpose: Devuelve el arco tangente de dos números.
' Remarks:
'===============================================================================
Public Function MATH_Atan2(X As Double, Y As Double) As Double
Attribute MATH_Atan2.VB_HelpID = 206
On Error GoTo ErrOut
  Dim Theta As Double

  If (Abs(X) < 0.0000001) Then
    If (Abs(Y) < 0.0000001) Then
      Theta = 0#
    ElseIf (Y > 0#) Then
      Theta = 1.5707963267949
    Else
      Theta = -1.5707963267949
    End If
  Else
    Theta = Atn(Y / X)
  
    If (X < 0) Then
      If (Y >= 0#) Then
        Theta = 3.14159265358979 + Theta
      Else
        Theta = Theta - 3.14159265358979
      End If
    End If
  End If
    
  MATH_Atan2 = Theta

ErrOut:

End Function

'===============================================================================
' Name: MATH_GetAngle
' Input:
'   ByVal X1 As Long - Coordenada X del primer punto.
'   ByVal Y1 As Long - Coordenada Y del primer punto.
'   ByVal X2 As Long - Coordenada X del segundo punto.
'   ByVal Y2 As Long - Coordenada Y del segundo punto.
' Output:
'   Single - Devuelve el angulo en grados.
' Purpose: Devuelve el angulo de entre dos puntos.
' Remarks: El angulo se calcula en sentido de las agujas del reloj.
'===============================================================================
Public Function MATH_GetAngle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Single
Attribute MATH_GetAngle.VB_HelpID = 205
    On Error Resume Next
    Dim tmp As Single
    tmp = CSng(MATH_Atan2(CDbl(X1) - CDbl(X2), CDbl(Y1) - CDbl(Y2)) * 180 / PI)
    If tmp < 0 Then
        MATH_GetAngle = 360 + tmp
    Else
        MATH_GetAngle = tmp
    End If
End Function

'===============================================================================
' Name: MATH_GetDist
' Input:
'   ByVal X1 As Long - Coordenada X del primer punto.
'   ByVal Y1 As Long - Coordenada Y del primer punto.
'   ByVal X2 As Long - Coordenada X del segundo punto.
'   ByVal Y2 As Long - Coordenada Y del segundo punto.
' Output:
'   Single - Devuelve la distancia en píxeles.
' Purpose: Calcula la distancia entre dos puntos.
' Remarks:
'===============================================================================
Public Function MATH_GetDist(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Long
Attribute MATH_GetDist.VB_HelpID = 204
    MATH_GetDist = Global_Mod.GetPointDist(X1, Y1, X2, Y2)

End Function

'===============================================================================
' Name: MATH_PointInLine
' Input:
'   ByVal X1 As Long - Coordenada X del primer punto de la linea.
'   ByVal Y1 As Long - Coordenada Y del primer punto de la linea.
'   ByVal X2 As Long - Coordenada X del segundo punto de la linea.
'   ByVal Y2 As Long - Coordenada Y del segundo punto de la linea.
'   ByVal X3 As Long - Coordenada X del punto del vector a intersectar.
'   ByVal Y3 As Long - Coordenada Y del punto del vector a intersectar.
' Output:
'   Boolean - Devuelve Verdadero si el punto colisiona con la linea.
' Purpose: Calcula si las coordenadas de un punto forman parte de una linea.
' Remarks:
'===============================================================================
Public Function MATH_PointInLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long) As Boolean
    Dim yprima As Double
    Dim aux As Double
   
    If ((X1 <= X3) And (X3 <= X2)) Or ((X2 <= X3) And (X3 <= X1)) Then
        aux = X2 - X1
       
        If aux > 0 Then
            yprima = ((Y2 - Y1) / (X2 - X1)) * (X3 - X1) + Y1
           
            MATH_PointInLine = (Abs(yprima - Y3) < 1#)
           
        ElseIf aux < 0 Then
            yprima = ((Y1 - Y2) / (X1 - X2)) * (X3 - X2) + Y2
           
            MATH_PointInLine = (Abs(yprima - Y3) < 1#)
           
        Else
            MATH_PointInLine = ((X3 = X1) And (((Y1 <= Y3) And (Y3 <= Y2)) Or _
                                                     ((Y2 <= Y3) And (Y3 <= Y1))))
        End If
   
    End If
    
End Function

'===============================================================================
' Name: MATH_IntersectRect
' Input:
'   ByVal DestRect As GFX_Rect - Rectángulo de destino.
'   ByVal SourceRect As GFX_Rect - Rectángulo de origen.
' Output:
'   Boolean - Devuelve verdadero si los dos rectángulos se interfieren.
' Purpose: Calcula la intersección entre dos rectángulos.
' Remarks: Esta función comprueba si dos áreas rectangulares colisionan entre si. Esta función puede ser útil para implementar colisiones entre objetos o gráficos.
'===============================================================================
Public Function MATH_IntersectRect(DestRect As GFX_Rect, SourceRect As GFX_Rect) As Boolean
Attribute MATH_IntersectRect.VB_HelpID = 203
    Dim A As RECT
    A.Left = DestRect.X
    A.Top = DestRect.Y
    A.Right = DestRect.X + DestRect.Width
    A.bottom = DestRect.Y + DestRect.Height
    
    Dim B As RECT
    B.Left = SourceRect.X
    B.Top = SourceRect.Y
    B.Right = SourceRect.X + SourceRect.Width
    B.bottom = SourceRect.Y + SourceRect.Height
    
    MATH_IntersectRect = CBool(Global_Mod.IntersectRect(A, B, A))
    'MATH_IntersectRect = Not (((DestRect.X + DestRect.Width) < SourceRect.X) Or ((DestRect.Y + DestRect.Height) < SourceRect.Y) Or ((SourceRect.X + SourceRect.Width) < DestRect.X) Or ((SourceRect.Y + SourceRect.Height) < DestRect.Y))

End Function

'===============================================================================
' Name: MATH_PointInRect
' Input:
'   ByVal X As Long - Coordenada X del punto.
'   ByVal Y As Long - Coordenada Y del punto.
'   ByVal DestRect As GFX_Rect - Rectángulo de destino.
' Output:
'   Boolean - Devuelve verdadero si las coordenadas del punto están dentro del area del rectángulo.
' Purpose: Calcula si las coordenadas de un punto se encuentran dentro del área de un rectángulo.
' Remarks: Esta función comprueba si un punto esta dentro del área de un rectangulo. Esta función puede ser útil para implementar colisiones entre objetos o gráficos.
'===============================================================================
Public Function MATH_PointInRect(X As Long, Y As Long, DestRect As GFX_Rect) As Boolean
Attribute MATH_PointInRect.VB_HelpID = 202
    Dim rct As RECT
    rct.Left = DestRect.X
    rct.Top = DestRect.Y
    rct.Right = DestRect.X + DestRect.Width
    rct.bottom = DestRect.Y + DestRect.Height
    
    MATH_PointInRect = CBool(Global_Mod.PtInRect(rct, X, Y))
    
    'MATH_PointInRect = Not (((DestRect.X + DestRect.Width) < X) Or ((DestRect.Y + DestRect.Height) < Y) Or (X < DestRect.X) Or (Y < DestRect.Y))

End Function

'===============================================================================
' Name: CRYPT_EncodeString
' Input:
'   ByVal Text As string - Cadena de texto que será codificada.
'   ByVal Key As String - Clave de codificación.
' Output:
'   String - Devuelve la cadena de texto codificada.
' Purpose: Codifica una cadena de texto.
' Remarks: Esta función codifica una cadena de texto acorde el valor de la clave de codificación. Dicha clave será necesaria para poder decodificar después la cadena de texto.
'===============================================================================
Public Function CRYPT_EncodeString(Text As String, Key As String) As String
Attribute CRYPT_EncodeString.VB_HelpID = 201
On Error GoTo ErrOut

Dim B As String, S As String, i As Long, j As Long
Dim A1 As Long, A2 As Long, A3 As Long, P As String
   
   j = 1
   
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
     
   Next
    
   For i = 1 To Len(Text)
     A1 = Asc(Mid$(P, j, 1))
     
     j = j + 1
     
     If j > Len(P) Then j = 1
     
     A2 = Asc(Mid$(Text, i, 1))
     A3 = A1 Xor A2
     
     B = Hex$(A3)
     
     If Len(B) < 2 Then B = "0" + B
     S = S + B
   
   Next
   
   CRYPT_EncodeString = S

ErrOut:

End Function

'===============================================================================
' Name: CRYPT_DecodeString
' Input:
'   ByVal Text As string - Cadena de texto que será decodificada.
'   ByVal Key As String - Clave de descodificación.
' Output:
'   String - Devuelve la cadena de texto decodificada.
' Purpose: Decodifica una cadena de texto.
' Remarks: Si la clave de descodificación es diferente a la que se uso para codificar la cadena de texto original, el resultado será una cadena de texto con los datos erróneos.
'===============================================================================
Public Function CRYPT_DecodeString(Text As String, Key As String) As String
Attribute CRYPT_DecodeString.VB_HelpID = 200
On Error GoTo ErrOut

Dim B As String, S As String, i As Long, j As Long
Dim A1 As Long, A2 As Long, A3 As Long, P As String
   
   j = 1
   
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
     
   Next
   
   For i = 1 To Len(Text) Step 2
     A1 = Asc(Mid$(P, j, 1))
     
     j = j + 1
     
     If j > Len(P) Then j = 1
     
     B = Mid$(Text, i, 2)
     A3 = Val("&H" + B)
     A2 = A1 Xor A3
     S = S + Chr$(A2)
   
   Next
   
   CRYPT_DecodeString = S
   
ErrOut:

End Function

'===============================================================================
' Name: TIMER_Create
' Input:
' Output:
'   Long - Devuelve el identificador del cronometro. Si no se puede crear el cronometro se devuelve SYS_NOTCREATED.
' Purpose: Crea un cronometro.
' Remarks: Los cronómetros son útiles para medir tiempos y realizar acciones periódicas, como animaciones por ejemplo.
'===============================================================================
Public Function TIMER_Create() As Long
Attribute TIMER_Create.VB_HelpID = 199
On Error GoTo ErrOut
Dim i As Long

i = Get_FreeID(0)
m_Timer(i) = GetTickCount()

TIMER_Create = i
Exit Function

ErrOut:
TIMER_Create = SYS_NOTCREATED
End Function

'===============================================================================
' Name: TIMER_Kill
' Input:
'   ByVal Timer As Long - Identificador del cronometro.
' Output:
' Purpose: Destruye un cronometro.
' Remarks:
'===============================================================================
Public Sub TIMER_Kill(Timer As Long)
Attribute TIMER_Kill.VB_HelpID = 198
On Error GoTo ErrOut

If Timer = UBound(m_Timer) Then
    If Timer = 0 Then Erase m_Timer Else ReDim Preserve m_Timer(UBound(m_Timer) - 1) As Long
    
Else
    m_Timer(Timer) = 0

End If

ErrOut:
End Sub

'===============================================================================
' Name: TIMER_Reset
' Input:
'   ByVal Timer As Long - Identificador del cronometro.
' Output:
' Purpose: Reinicia un cronometro.
' Remarks: Reinicia el valor de la cuenta de un cronometro a 0.
'===============================================================================
Public Sub TIMER_Reset(Timer As Long)
Attribute TIMER_Reset.VB_HelpID = 197
On Local Error Resume Next

m_Timer(Timer) = GetTickCount()

End Sub

'===============================================================================
' Name: TIMER_GetValue
' Input:
'   ByVal Timer As Long - Identificador del cronometro.
' Output:
'   Long - Tiempo en milisegundos desde que se creo o reinicio el cronometro.
' Purpose: Lee el valor de un cronometro.
' Remarks:
'===============================================================================
Public Function TIMER_GetValue(Timer As Long) As Long
Attribute TIMER_GetValue.VB_HelpID = 196
On Local Error Resume Next

TIMER_GetValue = GetTickCount() - m_Timer(Timer)

End Function

'===============================================================================
' Name: HITIMER_Create
' Input:
' Output:
'   Long - Devuelve el identificador del cronometro de alta precision. Si no se puede crear el cronometro se devuelve SYS_NOTCREATED, si el sistema no tiene soporte para cronometros de alta precision se devolvera SYS_HITIMERNOTSUPPORT.
' Purpose: Crea un cronometro de alta precision.
' Remarks: Los cronómetros son útiles para medir tiempos y realizar acciones periódicas, como animaciones por ejemplo.
'===============================================================================
Public Function HITIMER_Create() As Long
On Error GoTo ErrOut
Dim i As Long
i = Get_FreeID(1)

Call Global_Mod.QueryPerformanceCounter(m_HITimer(i))

If m_HITimer(i) = 0@ Then
    HITIMER_Create = SYS_HITIMERNOTSUPPORT
    

Else
    HITIMER_Create = i

End If

Exit Function

ErrOut:
HITIMER_Create = SYS_NOTCREATED

End Function

'===============================================================================
' Name: HITIMER_Kill
' Input:
'   ByVal Timer As Long - Identificador del cronometro de alta precision.
' Output:
' Purpose: Destruye un cronometro de alta precision.
' Remarks:
'===============================================================================
Public Sub HITIMER_Kill(Timer As Long)
On Error GoTo ErrOut

If Timer = UBound(m_HITimer) Then
    If Timer = 0 Then Erase m_HITimer Else ReDim Preserve m_HITimer(UBound(m_HITimer) - 1) As Currency
    
Else
    m_HITimer(Timer) = 0@

End If

ErrOut:

End Sub

'===============================================================================
' Name: HITIMER_Reset
' Input:
'   ByVal Timer As Long - Identificador del cronometro de alta precision.
' Output:
' Purpose: Reinicia un cronometro de alta precision.
' Remarks: Reinicia el valor de la cuenta de un cronometro de alta precision a 0,0.
'===============================================================================
Public Sub HITIMER_Reset(Timer As Long)
On Local Error Resume Next

Call Global_Mod.QueryPerformanceCounter(m_HITimer(Timer))

End Sub

'===============================================================================
' Name: HITIMER_GetValue
' Input:
'   ByVal Timer As Long - Identificador del cronometro de alta precision.
' Output:
'   Long - Tiempo en milisegundos con parte decimal desde que se creo o reinicio el cronometro de alta precision.
' Purpose: Lee el valor de un cronometro de alta precision.
' Remarks:
'===============================================================================
Public Function HITIMER_GetValue(Timer As Long) As Currency
On Local Error Resume Next
Dim tCount As Currency

Call Global_Mod.QueryPerformanceCounter(tCount)
HITIMER_GetValue = ((tCount - m_HITimer(Timer)) / QueryFreq) * 1000

End Function

'===============================================================================
' Name: TIMER_CreateProcess
' Input:
'   ByVal hWnd As Long - Identificador de la ventana que hará de cliente. Si no existe ventana cliente este parámetro se establecerá a 0.
'   ByVal Interval As Long - Intervalo de tiempo en milisegundos que define la frecuencia de ejecución del proceso.
'   ByVal ProcessAddress As Long - Dirección de memoria del procedimiento que lanzaremos con el cronometro. La dirección de memoria de un procedimiento se obtiene mediante la instrucción AddressOf de Visual Basic.
' Output:
'   Long - Identificador de la instancia del cronometro que ejecutara nuestro procedimiento.
' Purpose: Crea un proceso que se ejecutara a intervalos.
' Remarks: Esta función crea un proceso en memoria que se ejecutara a los intervalos de tiempo establecidos y en un hilo independiente evitando asi interrumpir la ejecución del programa principal.<br>
'Solo se puede asociar un procedimiento Sub al proceso, y este debe estar declarado en un modulo de código BAS.
'===============================================================================
Public Function TIMER_CreateProcess(hWnd As Long, Interval As Long, ProcessAddress As Long) As Long
Attribute TIMER_CreateProcess.VB_HelpID = 195
On Error Resume Next

TIMER_CreateProcess = Global_Mod.SetTimer(hWnd, 0&, Interval, ProcessAddress)

End Function

'===============================================================================
' Name: TIMER_KillProcess
' Input:
'   ByVal hWnd As Long - Identificador de la ventana que hará de cliente. Si no existe ventana cliente este parámetro se establecerá a 0.
'   ByVal TimerProcess As Long - Identificador de la instancia del proceso en memoria.
' Output:
' Purpose: Destruye un proceso.
' Remarks: Si no se destruye un proceso se puede correr el riesgo de dejar el procedimiento asociado en ejecución una vez finalizado nuestro programa, dejando nuestro programa en ejecución en 2º plano.
'===============================================================================
Public Sub TIMER_KillProcess(hWnd As Long, TimerProcess As Long)
Attribute TIMER_KillProcess.VB_HelpID = 194
On Error Resume Next

Call Global_Mod.KillTimer(hWnd, TimerProcess)

End Sub

'===============================================================================
' Name: TIMER_Count
' Input:
' Output:
'   Long - Numero de cronómetros.
' Purpose: Devuelve el numero de cronómetros creados.
' Remarks: Solo se devuelve el numero de los cronómetros de tiempo. El numero de procesos no se contabiliza.
'===============================================================================
Public Property Get TIMER_Count() As Long
Attribute TIMER_Count.VB_HelpID = 193
On Local Error Resume Next
TIMER_Count = UBound(m_Timer) + 1

End Property

'===============================================================================
' Name: HITIMER_Count
' Input:
' Output:
'   Long - Numero de cronómetros de alta precision.
' Purpose: Devuelve el numero de cronómetros de alta precision creados.
' Remarks: Solo se devuelve el numero de los cronómetros de alta precision.
'===============================================================================
Public Property Get HITIMER_Count() As Long
On Local Error Resume Next
HITIMER_Count = UBound(m_HITimer) + 1

End Property

'Devuelve un identificador para los cronometros:
Private Function Get_FreeID(Flag As Byte) As Long
On Local Error Resume Next

Dim i As Long, j As Long

Select Case Flag
    Case 0 'Cronometros:
        i = UBound(m_Timer)
        
        If Err.Number = 9 Then
            ReDim m_Timer(0) As Long
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If Not m_Timer(j) = 0 Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve m_Timer(i + 1) As Long
            
            Get_FreeID = UBound(m_Timer)
            
        End If
        
    Case 1 'Cronometros de alta precision:
        i = UBound(m_HITimer)
        
        If Err.Number = 9 Then
            ReDim m_HITimer(0) As Currency
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If Not m_HITimer(j) = 0@ Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve m_HITimer(i + 1) As Currency
            
            Get_FreeID = UBound(m_HITimer)
            
        End If
            
End Select

End Function

'===============================================================================
' Name: DEBUG_OpenLogFile
' Input:
'   ByVal FileName As String - Nombre y ruta del archivo de registros que se creara al finalizar la clase.
' Output:
' Purpose: Abre un buffer en memoria para registrar la entradas de la consola de depuración para despues ser guardadas en un archivo de registro.
' Remarks: Esta consola de depuración no es mas que una ventana de comandos del sistema, una ventana de MS-DOS. La idea de usar esta consola de depuración no es otra que la de obtener una sencilla herramienta para poder volcar mensajes a la hora de depurar nuestro programa o para el uso que se desee.
'===============================================================================
'Public Sub DEBUG_OpenLogFile(FileName As String)
'    TraceExec = True
'
'End Sub

'===============================================================================
' Name: DEBUG_OpenConsole
' Input:
'   ByVal Title As String - Titulo de la consola de depuración.
'   Optional ByVal SetDefaultWndPos As Boolean = True - Indica si la ventana de la consola de depuración se creara en una posición por defecto o no.
' Output:
'   Long - Instancia de la consola de depuración.
' Purpose: Abre una consola de depuración.
' Remarks: Esta consola de depuración no es mas que una ventana de comandos del sistema, una ventana de MS-DOS. La idea de usar esta consola de depuración no es otra que la de obtener una sencilla herramienta para poder volcar mensajes a la hora de depurar nuestro programa o para el uso que se desee.
'===============================================================================
Public Function DEBUG_OpenConsole(Title As String, Optional SetDefaultWndPos As Boolean = True) As Long
Attribute DEBUG_OpenConsole.VB_HelpID = 192
On Error GoTo ErrOut

Dim csl_hWnd As Long, csl_Rect As RECT

    'If we successfully open a console window then
    If AllocConsole() Then
        'Set title caption:
        SetConsoleTitle Title
        
        'we get the handle to the new console window.
        hConsole = GetStdHandle(STD_OUTPUT_HANDLE)

        'If we didn't successfully get the handle
        If hConsole = 0 Then
            'we should disconnect from the console otherwise
            FreeConsole
        Else
            'we need to add our custom console handler to
            'avoid the default console handler and
            SetConsoleCtrlHandler AddressOf ConsoleHandler, True
            
            If Not SetDefaultWndPos Then
                'Colocamos la ventana de la consola en la coordenada 0,0:
                csl_hWnd = Global_Mod.FindWindow(vbNullString, Title)
                Call Global_Mod.GetWindowRect(csl_hWnd, csl_Rect)
                Call Global_Mod.MoveWindow(csl_hWnd, 0, 0, csl_Rect.Right, csl_Rect.bottom, True)
            
            End If
            
            'set the return value to indicate success.
            DEBUG_OpenConsole = csl_hWnd
        End If

    End If

ErrOut:

End Function

'===============================================================================
' Name: DEBUG_CloseConsole
' Input:
' Output:
' Purpose: Cierra la consola de depuración.
' Remarks: Si trata de cerrar la consola de depuración manualmente probablemente no pueda o muestre un mensaje de aviso del sistema preguntándole si desea terminar la ejecución del programa.
'===============================================================================
Public Sub DEBUG_CloseConsole()
Attribute DEBUG_CloseConsole.VB_HelpID = 191
On Error GoTo ErrOut

    'Close the console handle and clear the cached handle
    If CloseHandle(hConsole) <> 0 Then hConsole = 0
    
    'Disconnects this process from the console.  It will close if no other
    'processes are using the console.
    FreeConsole

ErrOut:

End Sub

'===============================================================================
' Name: DEBUG_SendText
' Input:
'   ByVal Text As String - Cadena de texto que se imprimirá en la consola.
'   Optional ByVal Color As EGA_Color - Constante que indica el color que se utilizara para imprimir el texto. Los colores solo pueden ser valores de la lista EGA_Color.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Envía una cadena de texto a la consola de depuración.
' Remarks:
'===============================================================================
Public Function DEBUG_SendText(Text As String, Optional Color As EGA_Color = EGA_Color.Silver) As Boolean
Attribute DEBUG_SendText.VB_HelpID = 190
On Error GoTo ErrOut

    Dim sOut As String, cWritten As Long, Ret As Long

    sOut = Text
    
    'Aplicamos el formato al texto:
    SetConsoleTextAttribute hConsole, Color
    
    'Write the text to the console
    DEBUG_SendText = WriteConsole(hConsole, ByVal sOut, Len(sOut), cWritten, ByVal 0&)
    
    ReDim Preserve TraceBuffer(UBound(TraceBuffer) + 1) As String
    TraceBuffer(UBound(TraceBuffer)) = Text
    
    'Devolvemos el formato por defecto:
    SetConsoleTextAttribute hConsole, &H7

ErrOut:

End Function

Private Sub Class_Initialize()
    Call Global_Mod.QueryPerformanceFrequency(QueryFreq) 'Obtenemos la frecuencia del cronometro de alta precision del sistema.
    
    ReDim TraceBuffer(1) As String
    TraceBuffer(0) = "Reporte iniciado el " & Date & " a las " & Time$
    TraceBuffer(1) = vbNullString
    
End Sub

Private Sub Class_Terminate()
    Dim i As Long, hFile As Integer
    
    'Destruimos los cronometros de tiempo:
    Erase m_Timer
    Erase m_HITimer

    'Cerramos la consola de depuracion en caso de estar inicializada:
    If Global_Mod.hConsole <> 0 Then
        Call Global_Mod.CloseHandle(Global_Mod.hConsole)
        Call Global_Mod.FreeConsole

        'Si se especifico se crea el archivo y se guarda la informacion:
        If Not TraceFile = vbNullString Then
            ReDim Preserve TraceBuffer(UBound(TraceBuffer) + 2) As String
            TraceBuffer(UBound(TraceBuffer) - 1) = vbNullString
            TraceBuffer(UBound(TraceBuffer)) = "Reporte terminado el " & Date & " a las " & Time$
            
            'On Error Resume Next
            
            hFile = FreeFile
            Open TraceFile For Output As #hFile
                For i = 0 To UBound(TraceBuffer)
                    Print #hFile, Replace(Trim(TraceBuffer(i)), vbNewLine, "")
                
                Next i
            
            Close #hFile
                
        End If
        
    End If

End Sub

'===============================================================================
' Name: SYS_GetProcessorInfo
' Input:
'   ByVal Data As Processor_Info - Parámetro de salida que nos devuelve la informacion del procesador principal del sistema.
' Output:
' Purpose: Devuelve la información sobre el procesador de la CPU.
' Remarks: En plataformas Windows 9x (95/98/Me) la propiedad Mhz de la estructura Processor_Info devuelve 0.
'===============================================================================
Public Sub SYS_GetProcessorInfo(Data As Processor_Info)
Attribute SYS_GetProcessorInfo.VB_HelpID = 189
On Error GoTo ErrOut

With Data
    .Name = REG_GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "ProcessorNameString")
    .Identifier = REG_GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "Identifier")
    .VendorIdentifier = REG_GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "VendorIdentifier")
    .Mhz = REG_GetDWord(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "~MHz")
    
End With

ErrOut:

End Sub

'Indica el nombre del archivo de depuracion que se utilizara para guardar el reporte:
Public Property Get DebugFile() As String
    DebugFile = TraceFile
    
End Property

Public Property Let DebugFile(ByVal Filename As String)
    TraceFile = Filename
    
End Property

'===============================================================================
' Name: MATH_RotatePoint
' Input:
'   ByVal X As Long - Coordenada X del punto a rotar. Devuelve el resultado de rotar X.
'   ByVal Y As Long - Coordenada Y del punto a rotar. Devuelve el resultado de rotar Y.
'   ByVal Radius As Long - Radio de rotacion.
'   ByVal Angle As Single - Angulo de rotacion.
' Output:
' Purpose: Calcula la rotacion de un punto a partir de un radio y un angulo.
' Remarks: Esta funcion es idonea para calcular movimientos de coordenadas en angulos y distancias.
'===============================================================================
Public Sub MATH_RotatePoint(X As Long, Y As Long, Radius As Long, Angle As Single)
    Dim cX As Long, cY As Long
    Call Global_Mod.RotatePoint(X, Y, cX, cY, Radius, Angle)
    X = cX
    Y = cY
End Sub
