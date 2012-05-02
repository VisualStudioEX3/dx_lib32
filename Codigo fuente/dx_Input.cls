VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dx_Input_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 151
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
' Name: dx_Input
' Purpose: Clase de lectura de periféricos de entrada
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 07/08/2001
' Modified: 11/02/2010
'===============================================================================

Option Explicit

'Parametros por defecto. Recibe informacion fuera del foco de la ventana de la aplicacion:
Private Const DI_Params = DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

Private Const BufferSize = 10

'===============================================================================
' Name: Key_Const
'    Key_Escape - Escape.
'    Key_F1 - Tecla función 1
'    Key_F2 - Tecla función 2
'    Key_F3 - Tecla función 3
'    Key_F4 - Tecla función 4
'    Key_F5 - Tecla función 5
'    Key_F6 - Tecla función 6
'    Key_F7 - Tecla función 7
'    Key_F8 - Tecla función 8
'    Key_F9 - Tecla función 9
'    Key_F10 - Tecla función 10
'    Key_F11 - Tecla función 11
'    Key_F12 - Tecla función 12
'    Key_F13 - Tecla función 13
'    Key_F14 - Tecla función 14
'    Key_F15 - Tecla función 15
'    Key_BackSlash - Llave de cierre (})
'    Key_1 - 1 en el teclado principal.
'    Key_2 - 2 en el teclado principal.
'    Key_3 - 3 en el teclado principal.
'    Key_4 - 4 en el teclado principal.
'    Key_5 - 5 en el teclado principal.
'    Key_6 - 6 en el teclado principal.
'    Key_7 - 7 en el teclado principal.
'    Key_8 - 8 en el teclado principal.
'    Key_9 - 9 en el teclado principal.
'    Key_0 - 0 en el teclado principal.
'    Key_Equals - Interrogante de inicio (¿)
'    Key_Apostrophe - Llave de inicio ({)
'    Key_Slash - Barra invertida hacia adelante (/) en teclado principal.
'    Key_Back - Retroceso.
'    Key_Tab - Tabulador.
'    Key_Q - Q
'    Key_W - W
'    Key_E - E
'    Key_R - R
'    Key_T - T
'    Key_Y - Y
'    Key_U - U
'    Key_I - I
'    Key_O - O
'    Key_P - P
'    Key_Enter - Retorno de carro.
'    Key_LBracket - Corchete de inicio ([)
'    Key_RBracket - Corchete de cierre (])
'    Key_Caps - Fijador de mayúsculas (Caps Lock)
'    Key_A - A
'    Key_S - S
'    Key_D - D
'    Key_F - F
'    Key_G - G
'    Key_H - H
'    Key_J - J
'    Key_K - K
'    Key_L - L
'    Key_Grave - Numerador (º)
'    Key_SemiColon - Ñ
'    Key_LShift - Shift izquierdo.
'    Key_Minus - Interrogante de cierre (?)
'    Key_Z - Z
'    Key_X - X
'    Key_C - C
'    Key_V - V
'    Key_B - B
'    Key_N - N
'    Key_M - M
'    Key_Comma - Coma (,)
'    Key_Period - Punto (.) en el teclado principal.
'    Key_SysRQ - Imprimir pantalla (Impr Pant PetSis)
'    Key_RShift - Shift derecho.
'    Key_MinorMajor - Signos de mayor que y menor que (< >)
'    Key_LControl - Control izquierdo.
'    Key_LWin - Menu inicio izquierdo.
'    Key_LAlt - Alternativa izquierda.
'    Key_Space - Barra de espacio.
'    Key_RAlt - Alternativo derecho (Alt Gr)
'    Key_RWin - Menu inicio derecho.
'    Key_Apps - Menu contextual.
'    Key_RControl - Control derecho.
'    Key_Scroll - Bloqueo desbloqueo (Scroll Lock)
'    Key_Pause - Pausa interrupción.
'    Key_Insert - Insertar.
'    Key_Delete - Suprimir.
'    Key_Home - Inicio.
'    Key_End - Fin.
'    Key_PagUp - Pagina anterior o retroceso de pagina (RePag)
'    Key_PagDown - Pagina siguiente o avance de pagina (AvPag)
'    Key_Left - Izquierda.
'    Key_Up - Arriba.
'    Key_Down - Abajo.
'    Key_Right - Derecha.
'    Key_NumLock - Bloque numérico.
'    Key_NumpadEnter - Introducir (Enter) en bloque numérico.
'    Key_Decimal - Punto decimal (.) en bloque numérico.
'    Key_Numpad0 - 0 en bloque numérico.
'    Key_Numpad1 - 1 en bloque numérico.
'    Key_Numpad2 - 2 en bloque numérico.
'    Key_Numpad3 - 3 en bloque numérico.
'    Key_Numpad4 - 4 en bloque numérico.
'    Key_Numpad5 - 5 en bloque numérico.
'    Key_Numpad6 - 6 en bloque numérico.
'    Key_Numpad7 - 7 en bloque numérico.
'    Key_Numpad8 - 8 en bloque numérico.
'    Key_Numpad9 - 9 en bloque numérico.
'    Key_Add - Sumar (+) en bloque numérico.
'    Key_SubsTract - Restar (-) en bloque numérico.
'    Key_Multiply - Multiplicar (*) en bloque numérico.
'    Key_Divide- Dividir (/) en bloque numérico.
'    Japan_Key_Kana - "Kana" en teclado japonés.
'    Japan_Key_Convert - "Convertir" en teclado japonés.
'    Japan_Key_NoConvert- "No convertir" en teclado japonés.
'    Japan_Key_Yen - Yen en teclado japonés.
'    Japan_Key_CircumFlex - "CircumFlex" en teclado japonés.
'    Japan_Key_At - "At" en teclado japonés.
'    Japan_Key_Colon - (~) en teclado japonés.
'    Japan_Key_Underline - Subrayado en teclado japonés.
'    Japan_Key_Kanji - "Kanji" en teclado japonés.
'    Japan_Key_Stop - Stop en teclado japonés.
'    Japan_Key_Ax - "Ax" en teclado japonés.
'    Japan_Key_Unlabeled - "Sin etiquetar" en teclado japonés.
'    PC98_Key_NumPadEquals - Interrogante de inicio (¿) en bloque numérico (Solo teclados NEC PC98)
'    PC98_Key_NumpadComma - Coma en bloque numérico (Solo teclados NEC PC98)
' Purpose: Constantes de teclado.
' Remarks:
'===============================================================================
Public Enum Key_Const
    Key_Escape = DIK_ESCAPE
    
    'Teclas de Funciones:
    Key_F1 = DIK_F1
    Key_F2 = DIK_F2
    Key_F3 = DIK_F3
    Key_F4 = DIK_F4
    Key_F5 = DIK_F5
    Key_F6 = DIK_F6
    Key_F7 = DIK_F7
    Key_F8 = DIK_F8
    Key_F9 = DIK_F9
    Key_F10 = DIK_F10
    Key_F11 = DIK_F11
    Key_F12 = DIK_F12
    Key_F13 = DIK_F13
    Key_F14 = DIK_F14
    Key_F15 = DIK_F15
    
    'Teclas numericas del teclado:
    Key_BackSlash = DIK_BACKSLASH
    Key_1 = DIK_1 'On main keyboard
    Key_2 = DIK_2
    Key_3 = DIK_3
    Key_4 = DIK_4
    Key_5 = DIK_5
    Key_6 = DIK_6
    Key_7 = DIK_7
    Key_8 = DIK_8
    Key_9 = DIK_9
    Key_0 = DIK_0
    
    Key_Equals = DIK_EQUALS
    Key_Apostrophe = DIK_APOSTROPHE
    Key_Slash = DIK_SLASH 'Forward slash (/)on main keyboard
    
    Key_Back = DIK_BACK 'BACKSPACE
    
    'Teclas Alfabeticas:
    Key_Tab = DIK_TAB
    
    Key_Q = DIK_Q
    Key_W = DIK_W
    Key_E = DIK_E
    Key_R = DIK_R
    Key_T = DIK_T
    Key_Y = DIK_Y
    Key_U = DIK_U
    Key_I = DIK_I
    Key_O = DIK_O
    Key_P = DIK_P
    
    Key_Enter = DIK_RETURN 'ENTER on main keyboard
    
    Key_LBracket = DIK_LBRACKET '[
    Key_RBracket = DIK_RBRACKET ']
    
    Key_Caps = DIK_CAPITAL 'CAPS LOCK
    
    Key_A = DIK_A
    Key_S = DIK_S
    Key_D = DIK_D
    Key_F = DIK_F
    Key_G = DIK_G
    Key_H = DIK_H
    Key_J = DIK_J
    Key_K = DIK_K
    Key_L = DIK_L
    
    Key_Grave = DIK_GRAVE 'Grave accent (`)
    Key_SemiColon = DIK_SEMICOLON
    
    Key_LShift = DIK_LSHIFT 'Left SHIFT
    Key_Minus = DIK_MINUS
    
    Key_Z = DIK_Z
    Key_X = DIK_X
    Key_C = DIK_C
    Key_V = DIK_V
    Key_B = DIK_B
    Key_N = DIK_N
    Key_M = DIK_M
    
    Key_Comma = DIK_COMMA
    Key_Period = DIK_PERIOD 'On main keyboard
    Key_SysRQ = DIK_SYSRQ
    Key_RShift = DIK_RSHIFT 'Right SHIFT
    
    Key_MinorMajor = 86 '<>
    
    Key_LControl = DIK_LCONTROL 'Left CTRL
    Key_LWin = DIK_LWIN 'Left Windows key
    Key_LAlt = DIK_LMENU 'Left ALT
    Key_Space = DIK_SPACE 'Spacebar
    Key_RAlt = DIK_RMENU 'Right ALT
    Key_RWin = DIK_RWIN 'Right Windows key
    Key_Apps = DIK_APPS 'Application key
    Key_RControl = DIK_RCONTROL 'Right CTRL key
        
    Key_Scroll = DIK_SCROLL  'SCROLL LOCK
    Key_Pause = DIK_PAUSE
    
    Key_Insert = DIK_INSERT
    Key_Delete = DIK_DELETE
    Key_Home = DIK_HOME
    Key_End = DIK_END
    Key_PagUp = DIK_PRIOR 'PAGE UP
    Key_PagDown = DIK_NEXT 'PAGE DOWN
    
    Key_Left = DIK_LEFT 'Left arrow
    Key_Up = DIK_UP 'Up arrow
    Key_Down = DIK_DOWN 'Down arrow
    Key_Right = DIK_RIGHT 'Right arrow
    
    'Teclas del bloque numerico:
    Key_NumLock = DIK_NUMLOCK
    Key_NumpadEnter = DIK_NUMPADENTER
    Key_Decimal = DIK_DECIMAL 'Period (decimal point) on numeric keypad
    
    Key_Numpad0 = DIK_NUMPAD0
    Key_Numpad1 = DIK_NUMPAD1
    Key_Numpad2 = DIK_NUMPAD2
    Key_Numpad3 = DIK_NUMPAD3
    Key_Numpad4 = DIK_NUMPAD4
    Key_Numpad5 = DIK_NUMPAD5
    Key_Numpad6 = DIK_NUMPAD6
    Key_Numpad7 = DIK_NUMPAD7
    Key_Numpad8 = DIK_NUMPAD8
    Key_Numpad9 = DIK_NUMPAD9
    
    Key_Add = DIK_ADD 'Plus sign on numeric keypad
    Key_SubsTract = DIK_SUBTRACT 'Hyphen (minus sign) on numeric keypad
    Key_Multiply = DIK_MULTIPLY 'Asterisk on numeric keypad
    Key_Divide = DIK_DIVIDE 'Forward slash (/)on numeric keypad
    
    'Teclas pertenecientes al teclado Japones:
    Japan_Key_Kana = DIK_KANA 'On Japanese keyboard
    Japan_Key_Convert = DIK_CONVERT 'On Japanese keyboard
    Japan_Key_NoConvert = DIK_NOCONVERT 'On Japanese keyboard
    Japan_Key_Yen = DIK_YEN 'On Japanese keyboard
    Japan_Key_CircumFlex = DIK_CIRCUMFLEX 'On Japanese keyboard
    Japan_Key_At = DIK_AT 'On Japanese keyboard
    Japan_Key_Colon = DIK_COLON 'On Japanese keyboard
    Japan_Key_Underline = DIK_UNDERLINE 'On Japanese keyboard
    Japan_Key_Kanji = DIK_KANJI 'On Japanese keyboard
    Japan_Key_Stop = DIK_STOP 'On Japanese keyboard
    Japan_Key_Ax = DIK_AX 'On Japanese keyboard
    Japan_Key_Unlabeled = DIK_UNLABELED 'On Japanese keyboard
    
    'Teclas pertenecientes al teclado NEC PC98:
    PC98_Key_NumPadEquals = DIK_NUMPADEQUALS 'On numeric keypad (NEC PC98)
    PC98_Key_NumpadComma = DIK_NUMPADCOMMA 'Comma on NEC PC98 numeric keypad
End Enum

'===============================================================================
' Name: Mouse_Button
'   Left_Button - Botón izquierdo.
'   Right_Button - Botón derecho.
'   Middle_Button - Botón central.
' Purpose: Constantes para ratón.
' Remarks: Se utilizan como valor en el parámetro de la función Mouse_Hit.
'===============================================================================
Public Enum Mouse_Button
    Left_Button = 1
    Right_Button = 2
    Middle_Button = 3
End Enum

'===============================================================================
' Name: Joystick_Button
'    PSX_Triangle_Button - Triangulo en gamepad de PlayStation (Botón 1)
'    PSX_Circle_Button - Circulo en gamepad de PlayStation (Botón 2)
'    PSX_Equis_Button - Equis en gamepad de PlayStation (Botón 3)
'    PSX_Square_Button - Cuadrado en gamepad de PlayStation (Botón 4)
'    PSX_L2_Button - L2 en gamepad de PlayStation (Botón 5)
'    PSX_R2_Button - R2 en gamepad de PlayStation (Botón 6)
'    PSX_L1_Button - L1 en gamepad de PlayStation (Botón 7)
'    PSX_R1_Button - R1 en gamepad de PlayStation (Botón 8)
'    SNES_X_Button - X en gamepad de Super Nintendo (Botón 1)
'    SNES_A_Button - A en gamepad de Super Nintendo (Botón 2)
'    SNES_B_Button - B en gamepad de Super Nintendo (Botón 3)
'    SNES_Y_Button - Y en gamepad de Super Nintendo (Botón 4)
'    SNES_L_Button - L en gamepad de Super Nintendo (Botón 7)
'    SNES_R_Button - R en gamepad de Super Nintendo (Botón 8)
'    Start_Button - Empezar (Botón 9)
'    Select_Button - Seleccionar (Botón 10)
'    Extra1_Button - Botón función extra 1 (Botón 11)
'    Extra2_Button - Botón función extra 2 (Botón 12)
'    Left_Arrow - Dirección izquierda según valor del eje primario.
'    Up_Arrow - Dirección arriba según valor del eje primario.
'    Right_Arrow - Dirección derecha según valor del eje primario.
'    Down_Arrow - Dirección abajo según valor del eje primario.
'    Down_Arrow - Dirección abajo según valor del eje primario.
' Purpose: Constantes para los botones del joystick y gamepads.
' Remarks: Se utilizan como valor en el parámetro de la función Joystick_Hit. Se han nombrado las constantes con la nomenclatura de dos de los gamepads mas conocidos, el de Sony PlayStation y el de Súper Nintendo, para facilitar la localización de los botones a la hora de realizar una implementación en código.
'===============================================================================
Public Enum Joystick_Button
    PSX_Triangle_Button = 1
    PSX_Circle_Button = 2
    PSX_Equis_Button = 3
    PSX_Square_Button = 4
    PSX_L2_Button = 5
    PSX_R2_Button = 6
    PSX_L1_Button = 7
    PSX_R1_Button = 8
    
    SNES_X_Button = 1
    SNES_A_Button = 2
    SNES_B_Button = 3
    SNES_Y_Button = 4
    SNES_L_Button = 7
    SNES_R_Button = 8
    
    Start_Button = 9
    Select_Button = 10
    
    Extra1_Button = 11
    Extra2_Button = 12
    
    Left_Arrow = 13
    Up_Arrow = 14
    Right_Arrow = 15
    Down_Arrow = 16

End Enum
    
'===============================================================================
' Name: Mouse_Data
'   ByRef Left_Button As Boolean - Estado botón izquierdo.
'   ByRef Right_Button As Boolean - Estado botón derecho.
'   ByRef Middle_Button As Boolean - Estado botón central.
'   ByRef X As Long - Coordenada X. Define la posición en el eje horizontal del cursor.
'   ByRef Y As Long - Coordenada Y. Define la posición en el eje vertical del cursor.
'   ByRef Z As Long - Avance del eje Z. Define el avance de la rueda del raton en valores positivo y negativo.
' Purpose: Información de retorno del ratón.
' Remarks: Esta estructura es la que se usa como retorno en el objeto Mouse.
'===============================================================================
Public Type Mouse_Data
    Left_Button As Boolean
    Right_Button As Boolean
    Middle_Button As Boolean
    X As Long
    Y As Long
    Z As Long
End Type

'===============================================================================
' Name: Joystick_Data
'   ByRef Button(1 To 12) As Boolean - Estado de los 12 botones.
'   ByRef X As Long - Coordenada X. Define la posición en el eje horizontal.
'   ByRef Y As Long - Coordenada Y. Define la posición en el eje vertical.
' Purpose: Información de retorno del joystick o gamepad.
' Remarks: Esta estructura es la que se usa como retorno en el objeto Joystick.
'===============================================================================
Public Type Joystick_Data
    'Button 1 'X - Triangulo
    'Button 2 'A - Circulo
    'Button 3 'B - X (Equis)
    'Button 4 'Y - Cuadrado
    'Button 5 'L2
    'Button 6 'R2
    'Button 7 'L1
    'Button 8 'R1
    'Button 9 'Start - Pause
    'Button 10 'Select
    'Button 11 'Desconocido
    'Button 12 'Desconocido
    Button(1 To 12) As Boolean 'Botones
    X As Long 'Eje X
    Y As Long 'Eje Y
    
End Type
    
Dim Dx As New DirectX8
Dim Di As DirectInput8

Dim diDevEnum As DirectInputEnumDevices8

'Guarda las posiciones del Raton.
Dim CursorCoord As POINTAPI

Dim JoyCaps As DIDEVCAPS
Dim DiProp_Dead As DIPROPLONG
Dim DiProp_Range As DIPROPRANGE
Dim DiProp_Saturation As DIPROPLONG
Dim AxisPresent(1 To 8) As Boolean

Dim TimerID As Long

Dim DI_Init As Boolean

Private keybHitState(255) As Boolean    ' Estados de las teclas del teclado.
Private mouseHitState As Mouse_Data     ' Estados de los botones del joystick.
Private joyHitState(3, 1 To 16) As Boolean  ' Estados de los botones del joystick.

'===============================================================================
' Name: Init
' Input:
'   ByVal hWnd As Long - Identificador de la ventana.
' Output:
' Purpose: Inicializa el sistema de lectura de la entrada de dispositivos de usuarios.
' Remarks: Inicializa el sistema de lectura de la entrada de dispositivos de usuarios y los subsistemas de la clase habilitando las rutinas y la memoria para leer y escribir la entrada de dispositivos de usuario.
'===============================================================================
Public Function Init(hWnd As Long) As Boolean
Attribute Init.VB_HelpID = 165
On Error GoTo ErrOut
Dim Params As CONST_DISCLFLAGS
Dim diProp As DIPROPLONG

Dim didoEnum As DirectInputEnumDeviceObjects
Dim dido As DirectInputDeviceObjectInstance
Dim i As Integer
Dim sGuid As String

If Not DI_Init Then

    Input_hWnd = hWnd

    Set Di = Dx.DirectInputCreate()

    'Inicializamos el control de Teclado:
    Set Di_Key = Di.CreateDevice("GUID_SysKeyboard")
    Call Di_Key.SetCommonDataFormat(DIFORMAT_KEYBOARD)

    Call Di_Key.SetCooperativeLevel(Input_hWnd, DI_Params)
    Call Di_Key.Acquire

    'Inicializamos el control de Raton:
    Set Di_Mouse = Di.CreateDevice("GUID_SysMouse")
    Call Di_Mouse.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call Di_Mouse.SetCooperativeLevel(Input_hWnd, DI_Params)

    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = BufferSize

    Call Di_Mouse.SetProperty("DIPROP_BUFFERSIZE", diProp)
    Call Di_Mouse.Acquire

    'Inicializamos el control del joistick:
    Set diDevEnum = Di.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)

    m_Joysticks = diDevEnum.GetCount

    If Not m_Joysticks = 0 Then
    
        ReDim Di_Joy(m_Joysticks - 1) As DirectInputDevice8
        ReDim JoyButtonPressId(UBound(Di_Joy)) As Long
        ReDim JoyData(UBound(Di_Joy)) As Joystick_Data_Event
        ReDim JoyState(UBound(Di_Joy), 1 To 16 + 2) ' + 2 -> Para detectar el evento 'hit' de la rueda del raton (giro +, giro -)
    
        '// Este codigo es opcional, en teoria no deberia afectar notoriamente en el funcionamiento de la entrada de
        'los joysticks:

        'Set deadzone for X and Y axis to 10 percent of the range of travel:
        With DiProp_Dead
            .lData = 1000
            .lHow = DIPH_BYOFFSET
                
        End With
    
        'Set saturation zones for X and Y axis to 5 percent of the range:
        With DiProp_Saturation
            .lData = 9500
            .lHow = DIPH_BYOFFSET
                 
        End With
    
        'NOTE Some devices do not let you set the range
        On Local Error Resume Next

        'Set range for all axes:
        With DiProp_Range
            .lHow = DIPH_DEVICE
            .lMin = 0
            .lMax = 10000
            
        End With
    
        '//

        'For i = 0 To m_Joysticks - 1
        For i = 1 To m_Joysticks
            Set Di_Joy(i - 1) = Di.CreateDevice(diDevEnum.GetItem(i).GetGuidInstance)
    
            With Di_Joy(i - 1)
                Call .SetCommonDataFormat(DIFORMAT_JOYSTICK)
    
                DiProp_Dead.lObj = DIJOFS_X
                Call .SetProperty("DIPROP_DEADZONE", DiProp_Dead)
        
                DiProp_Dead.lObj = DIJOFS_Y
                Call .SetProperty("DIPROP_DEADZONE", DiProp_Dead)

                DiProp_Saturation.lObj = DIJOFS_X
                Call .SetProperty("DIPROP_SATURATION", DiProp_Saturation)
        
                DiProp_Saturation.lObj = DIJOFS_Y
                Call .SetProperty("DIPROP_SATURATION", DiProp_Saturation)
    
                Call .SetProperty("DIPROP_RANGE", DiProp_Range)
    
                Call .Acquire
    
            End With

        Next i
    
    End If

    'Activamos la rutina que controla el estado de los dispositivos:
    'TimerID = Global_Mod.SetTimer(hWnd, 0&, 0, AddressOf InputEventControl)
    
    Init = True
    DI_Init = Init

End If

ErrOut:
Set diDevEnum = Nothing
Set didoEnum = Nothing
Set dido = Nothing

End Function

'===============================================================================
' Name: Key_ASCII
' Input:
' Output:
'   Long - Código de la tecla que esta siendo presionada.
' Purpose: Devuelve el valor ASCII de la tecla presionada.
' Remarks: La funcion devuelve el valor ASCII correspondiente a la tecla presionada afectandole el estado de las teclas de mayusculas (Shift) y el bloque mayusculas (Caps). El uso de acentos no afecta al caracter leido. No permite insertar combinaciones ALT+(n ASCII) para obtener caracteres especiales.
'===============================================================================
Public Function Key_ASCII() As Long
    On Error GoTo ErrOut
    
    Dim i As Long, Ret As Long
    Static Pressed As Boolean, Key As Long ' Variables estaticas para controlar la espera de pulsacion.
    
    If Pressed And APIKeybState.kbByte(Key) > 8 Then Exit Function Else Pressed = False
    
    For i = 8 To 127 'Desde 8 (retroceso, retorno de carro...) hasta 127 (codigos comunes)
        If APIKeybState.kbByte(i) > 8 Then
            If ToAscii(i, 0, APIKeybState.kbByte(0), Ret, 0) = 1 Then
                Key = i
                Pressed = True
                Key_ASCII = Ret
                Exit Function
            
            End If
            
        End If

    Next i
    
ErrOut:
    
End Function

'===============================================================================
' Name: Key_Press
' Input:
' Output:
'   Long - Código de la tecla que esta siendo presionada.
' Purpose: Devuelve el código de la tecla que esta siendo presionada en el teclado.
' Remarks:
'===============================================================================
Public Function Key_Press() As Long
Attribute Key_Press.VB_HelpID = 164
Key_Press = KeyPressId

End Function

'===============================================================================
' Name: Mouse_Press
' Input:
' Output:
'   Long - Código del botón que esta siendo presionado.
' Purpose: Devuelve el código del botón del ratón que esta siendo presionado.
' Remarks:
'===============================================================================
Public Function Mouse_Press() As Long
Attribute Mouse_Press.VB_HelpID = 163
Mouse_Press = MouseButtonPressId

End Function

'===============================================================================
' Name: Key
' Input:
'   ByVal KeyValue As Key_Const - Código de la tecla.
' Output:
'   Boolean - Devuelve verdadero si la tecla esta siendo presionada.
' Purpose: Comprueba el estado de una tecla en el teclado.
' Remarks:
'===============================================================================
Public Function Key(KeyValue As Key_Const) As Boolean 'El parametro KeyValue acepta variables de tipo Long.
Attribute Key.VB_HelpID = 162
On Error Resume Next

'Dim Key_State As DIKEYBOARDSTATE

If DI_Init Then
'    Call Di_Key.GetDeviceStateKeyboard(Key_State)
'
'    If (Err.Number = DIERR_NOTACQUIRED) Or (Err.Number = DIERR_INPUTLOST) Then Di_Key.Acquire

    If GetForegroundWindow = hWnd Then If Key_State.Key(KeyValue) Then Key = True

End If

End Function

'===============================================================================
' Name: Key_Hit
' Input:
'   ByVal KeyValue As Key_Const - Código de la tecla.
' Output:
'   Boolean - Devuelve verdadero en cuanto se presiona la tecla. Si la tecla sigue presionada se devuelve falso.
' Purpose: Detecta una pulsación en una tecla.
' Remarks: A diferencia de la función Key, Key_Hit solo detecta la acción de pulsar la tecla pero no devuelve de continuo el estado de dicha tecla.
'===============================================================================
Public Function Key_Hit(KeyValue As Key_Const) As Boolean
Attribute Key_Hit.VB_HelpID = 161
'    If DI_Init Then If KeyState(KeyValue) = 1 Then Key_Hit = True
    If Me.Key(KeyValue) And Not keybHitState(KeyValue) Then
        keybHitState(KeyValue) = True
        Key_Hit = True
    End If
End Function

'===============================================================================
' Name: Mouse_SetPossition
' Input:
'   ByVal X As Long - Coordenada horizontal de la pantalla o del área cliente de la ventana.
'   ByVal Y As Long - Coordenada vertical de la pantalla o del área cliente de la ventana.
' Output:
' Purpose: Establece las coordenadas para el cursor del ratón.
' Remarks: En modo ventana se toma las coordenadas del área de cliente de la ventana donde se ha inicializado la clase. En modo ventana se toman las coordenadas de la pantalla.
'===============================================================================
Public Sub Mouse_SetPossition(X As Long, Y As Long)
Attribute Mouse_SetPossition.VB_HelpID = 160
On Error GoTo ErrOut
Dim MousePos As POINTAPI

If DI_Init Then
    With MousePos
        .X = X
        .Y = Y
        
    End With

    Call ClientToScreen(hWnd, MousePos)
    Call SetCursorPos(MousePos.X, MousePos.Y)

End If

ErrOut:
End Sub

'===============================================================================
' Name: Mouse
' Input:
' Output:
'   Mouse_Data - Estructura que devuelve la información del estado de los botones y las coordenadas del cursor.
' Purpose: Devuelve el estado de los botones del ratón y las coordenadas del cursor.
' Remarks: En modo ventana se toma las coordenadas del área de cliente de la ventana donde se ha inicializado la clase. En modo ventana se toman las coordenadas de la pantalla.
'===============================================================================
Public Property Get Mouse() As Mouse_Data
Attribute Mouse.VB_HelpID = 159
On Error Resume Next

If DI_Init Then
    With Mouse
        .Left_Button = MouseData.Left_Button
        .Middle_Button = MouseData.Middle_Button
        .Right_Button = MouseData.Right_Button
        .X = MouseData.X
        .Y = MouseData.Y
        .Z = MouseData.Z

    End With

End If

End Property

'===============================================================================
' Name: Mouse_Hit
' Input:
'   ByVal Button As Mouse_Button - Código del botón.
' Output:
'   Boolean - Devuelve verdadero en cuanto se presiona el botón. Si el botón sigue presionado se devuelve falso.
' Purpose: Detecta una pulsación en un botón del ratón.
' Remarks: A diferencia del objeto Mouse, Mouse_Hit solo detecta la acción de pulsar el botón pero no devuelve de continuo el estado de dicho botón.
'===============================================================================
Public Function Mouse_Hit(Button As Mouse_Button) As Boolean
Attribute Mouse_Hit.VB_HelpID = 158
'    On Error Resume Next
'    If DI_Init Then Mouse_Hit = (MouseState(Button - 1) = 1)
    If (Button = Left_Button And Me.Mouse.Left_Button) And Not mouseHitState.Left_Button Then
        mouseHitState.Left_Button = True
        Mouse_Hit = True
        
    ElseIf (Button = Right_Button And Me.Mouse.Right_Button) And Not mouseHitState.Right_Button Then
        mouseHitState.Right_Button = True
        Mouse_Hit = True
        
    ElseIf (Button = Middle_Button And Me.Mouse.Middle_Button) And Not mouseHitState.Middle_Button Then
        mouseHitState.Middle_Button = True
        Mouse_Hit = True
        
    End If
End Function

'===============================================================================
' Name: Joystick
' Input:
'   ByVal Joy As Long - Índice del joystick o gamepad que se va a leer.
' Output:
'   Joystick_Data - Estructura que devuelve la información del estado de los botones y las coordenadas del eje primario.
' Purpose: Devuelve el estado de los botones del joystick o gamepad y las coordenadas de eje primario.
' Remarks:
'===============================================================================
Public Property Get Joystick(Joy As Long) As Joystick_Data
Attribute Joystick.VB_HelpID = 157
On Error Resume Next

Dim i As Long

If DI_Init Then
    If Joy < 0 Or Joy > m_Joysticks - 1 Then Exit Function
    With Joystick
        For i = 1 To 12
            .Button(i) = JoyData(Joy).Button(i)
        Next i
        .X = JoyData(Joy).X
        .Y = JoyData(Joy).Y
    End With
End If
End Property

'===============================================================================
' Name: Joystick_Press
' Input:
'   ByVal Joy As Long - Índice del joystick o gamepad que se va a leer.
' Output:
'   Long - Código del botón que esta siendo presionado.
' Purpose: Devuelve el código del botón que esta siendo presionado en el joystick o gamepad.
' Remarks:
'===============================================================================
Public Function Joystick_Press(Joy As Long) As Long
Attribute Joystick_Press.VB_HelpID = 156
    If Joy < 0 Or Joy > m_Joysticks - 1 Then Exit Function
    Joystick_Press = JoyButtonPressId(Joy)
End Function

'===============================================================================
' Name: Joystick_Hit
' Input:
'   ByVal Joy As Long - Índice del joystick o gamepad que se va a leer.
'   ByVal Button As Joystick_Button - Código del botón.
' Output:
'   Boolean - Devuelve verdadero en cuanto se presiona el botón. Si el botón sigue presionado se devuelve falso.
' Purpose: Detecta una pulsación en un botón del joystick o gamepad.
' Remarks: A diferencia del objeto Joystick, Joystick_Hit solo detecta la acción de pulsar el botón pero no devuelve de continuo el estado de dicho botón.
'===============================================================================
Public Function Joystick_Hit(Joy As Long, Button As Joystick_Button) As Boolean
Attribute Joystick_Hit.VB_HelpID = 155
'    On Error Resume Next
'    If DI_Init And m_Joysticks > 0 Then If JoyState(Joy, Button) = 1 Then Joystick_Hit = True

    If Joy < 0 Or Joy > m_Joysticks - 1 Then Exit Function
    
    With Me.Joystick(Joy)
        Select Case Button
            Case 1 To 12
                If .Button(Button) And Not joyHitState(Joy, Button) Then
                    joyHitState(Joy, Button) = True
                    Joystick_Hit = True
                End If
            Case Else
                If Button = Left_Arrow And .X < 4000 And Not joyHitState(Joy, Joystick_Button.Left_Arrow) Then
                    joyHitState(Joy, Joystick_Button.Left_Arrow) = True
                    Joystick_Hit = True
                ElseIf Button = Right_Arrow And .X > 6000 And Not joyHitState(Joy, Joystick_Button.Right_Arrow) Then
                    joyHitState(Joy, Joystick_Button.Right_Arrow) = True
                    Joystick_Hit = True
                ElseIf Button = Up_Arrow And .Y < 4000 And Not joyHitState(Joy, Joystick_Button.Up_Arrow) Then
                    joyHitState(Joy, Joystick_Button.Up_Arrow) = True
                    Joystick_Hit = True
                ElseIf Button = Down_Arrow And .Y > 6000 And Not joyHitState(Joy, Joystick_Button.Down_Arrow) Then
                    joyHitState(Joy, Joystick_Button.Down_Arrow) = True
                    Joystick_Hit = True
                End If
        End Select
    End With
End Function

' Actualiza los estados de la clase:
'===============================================================================
' Name: Update
' Purpose: Actualiza los estados de lectura de entrada.
' Remarks: Este método se encarga de actualizar correctamente los estados de lectura de los diferentes dispositivos de entrada vinculados a la clase.
'===============================================================================
Public Sub Update()
    Call DirectInputEvent_Mod.InputEventControl
    Call UpdateKeyb
    Call UpdateMouse
    Call UpdateJoystick
End Sub

' Codigo encargado de actualizar los estados del teclado:
Private Sub UpdateKeyb()
    Dim Key As Long
    For Key = 0 To 255
        If Not Me.Key(Key) And keybHitState(Key) Then keybHitState(Key) = False
    Next
End Sub

' Codigo encargado de actualizar los estados del raton:
Private Sub UpdateMouse()
    If Not Me.Mouse.Left_Button And mouseHitState.Left_Button Then mouseHitState.Left_Button = False
    If Not Me.Mouse.Right_Button And mouseHitState.Right_Button Then mouseHitState.Right_Button = False
    If Not Me.Mouse.Middle_Button And mouseHitState.Middle_Button Then mouseHitState.Middle_Button = False
End Sub

' Codigo encargado de actualizar los estados de los joysticks disponibles:
Private Sub UpdateJoystick()
    Dim Joystick As Long, Button As Long
    For Joystick = 0 To Me.Joystick_Count - 1
        With Me.Joystick(Joystick)
            For Button = 1 To 16
                Select Case Button
                    Case 1 To 12
                        If Not .Button(Button) And joyHitState(Joystick, Button) Then joyHitState(Joystick, Button) = False
                    Case Else
                        If Button = Left_Arrow And (.X > 4000 And .X < 6000) And joyHitState(Joystick, Joystick_Button.Left_Arrow) Then
                            joyHitState(Joystick, Joystick_Button.Left_Arrow) = False
                        End If
                        
                        If Button = Right_Arrow And (.X < 6000 And .X > 4000) And joyHitState(Joystick, Joystick_Button.Right_Arrow) Then
                            joyHitState(Joystick, Joystick_Button.Right_Arrow) = False
                        End If
                        
                        If Button = Up_Arrow And (.Y > 4000 And .Y < 6000) And joyHitState(Joystick, Joystick_Button.Up_Arrow) Then
                            joyHitState(Joystick, Joystick_Button.Up_Arrow) = False
                        End If
                        
                        If Button = Down_Arrow And (.Y < 6000 And .Y > 4000) And joyHitState(Joystick, Joystick_Button.Down_Arrow) Then
                            joyHitState(Joystick, Joystick_Button.Down_Arrow) = False
                        End If
                End Select
            Next
        End With
    Next
End Sub

'===============================================================================
' Name: Terminate
' Purpose: Cierra el sistema de entrada y los subsistemas de la clase.
' Remarks: Este método se encarga de liberar la memoria ocupada por los recursos de la clase y cerrar las referencias de la misma. Es aconsejable llamar a este método justo antes de terminar el programa para así asegurarnos que cualquier proceso de la clase se termina correctamente sin permanecer en memoria.
'===============================================================================
Public Sub Terminate()
Attribute Terminate.VB_HelpID = 154
On Error Resume Next

Dim i As Byte

If DI_Init Then
    Di_Key.Unacquire
    Di_Mouse.Unacquire

    If Not m_Joysticks = 0 Then
        For i = 0 To m_Joysticks - 1
            Di_Joy(i).Unacquire
        
            Set Di_Joy(i) = Nothing

        Next i
    
        Erase Di_Joy, JoyState, JoyButtonPressId

    End If

    'Call Global_Mod.KillTimer(hWnd, TimerID)

    Set Di_Key = Nothing
    Set Di_Mouse = Nothing

    Set Dx = Nothing
    Set Di = Nothing

End If

End Sub

'===============================================================================
' Name: hWnd
' Input:
' Output:
'   Long - Identificador de la ventana.
' Purpose: Devuelve el identificador de la ventana que esta usando para ejecutar las rutinas de la clase.
' Remarks:
'===============================================================================
Public Property Get hWnd() As Long
Attribute hWnd.VB_HelpID = 153
hWnd = Input_hWnd

End Property

'===============================================================================
' Name: Joystick_Count
' Input:
' Output:
'   Long - Numero de joysticks o gamepads conectados al equipo.
' Purpose: Devuelve el numero de joysticks o gamepads conectados al equipo.
' Remarks:
'===============================================================================
Public Property Get Joystick_Count() As Long
Attribute Joystick_Count.VB_HelpID = 152
Joystick_Count = m_Joysticks

End Property

''===============================================================================
'' Name: Frame
'' Purpose: Actualiza los eventos de entrada.
'' Remarks: Este método se encarga de actualizar los datos recogidos den los envetos de lectura de entrada. Debe ser llamada una sola vez por ciclo (igual que la instruccion Frame de la clase dx_GFX)
''===============================================================================
'Public Sub Frame()
'    Call InputEventControl
'End Sub
