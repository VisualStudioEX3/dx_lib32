Attribute VB_Name = "DirectInputEvent_Mod"
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
' Name: DirectShowEvent
' Purpose: Modulo de control y eventos de entrada
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 01/09/2006
' Modified: 16/02/2010
'===============================================================================

Option Explicit

'Estructura de datos del raton:
Public Type Mouse_Data_Event
    Left_Button As Boolean
    Right_Button As Boolean
    Middle_Button As Boolean
    X As Long
    Y As Long
    Z As Long
End Type

'Estructura de datos del joystick:
Public Type Joystick_Data_Event
    Button(1 To 12) As Boolean 'Botones
    X As Long 'Eje X
    Y As Long 'Eje Y
    
End Type

Public Input_hWnd As Long

Public m_Joysticks As Long 'Numero de joysticks conectados.

'Interfaz de lectura de los dispositivos de entrada:
Public Di_Key As DirectInputDevice8
Public Di_Mouse As DirectInputDevice8
Public Di_Joy() As DirectInputDevice8

'Almacena el identicador de la tecla o boton pulsado:
Public KeyPressId As Long
Public MouseButtonPressId As Long
Public JoyButtonPressId() As Long

Public Key_State As DIKEYBOARDSTATE
Public MouseData As Mouse_Data_Event
Public JoyData() As Joystick_Data_Event

'Registra los estados de cada boton:
Public KeyState(255) As Long, APIKeybState As KeyboardBytes
Public MouseState(2) As Long
Public JoyState() As Long 'JoyState(n, 1 To 16) 12 action buttons + 4 arrows

Public Sub InputEventControl()
    GetEventKeyb
    GetEventMouse
    If Not m_Joysticks = 0 Then GetEventJoy
End Sub

'Devuelve el estado del teclado:
Private Sub GetEventKeyb()
    On Error Resume Next
    
    Dim i As Long
'    Dim Key_State As DIKEYBOARDSTATE
    
    'Escanea el teclado desde la API de Windows para la lectura ASCII:
    Call Global_Mod.GetKeyboardState(APIKeybState)
    
    Call Di_Key.GetDeviceStateKeyboard(Key_State)
    
    If (Err.Number = DIERR_NOTACQUIRED) Or (Err.Number = DIERR_INPUTLOST) Then Di_Key.Acquire
    
    KeyPressId = 0
    
    For i = 0 To 255
        If GetForegroundWindow = Input_hWnd Then
            If Key_State.Key(i) > 0 Then
                KeyPressId = i
                'KeyState(i) = KeyState(i) + 1
'            Else
'
'                'KeyPressId = 0
'                'KeyState(i) = 0
            End If
        End If
    Next i
End Sub

'Devuelve el estado del raton:
Private Sub GetEventMouse()
    On Error GoTo ErrOut
    Dim ClientRect As RECT
    Dim point As POINTAPI
    Dim Mouse_State As DIMOUSESTATE
    
    'Adquirimos las posiciones del raton para el Cursor.
    Call GetCursorPos(point)
    
    If Global_Mod.D3D_FullScreen Then
        Call ScreenToClient(Global_Mod.GetDesktopWindow(), point)
    Else
        Call ScreenToClient(Input_hWnd, point)
    End If
    
    MouseData.X = point.X
    MouseData.Y = point.Y
    
    On Error Resume Next
    Call Di_Mouse.GetDeviceStateMouse(Mouse_State)
    If (Err.Number = DIERR_NOTACQUIRED) Or (Err.Number = DIERR_INPUTLOST) Then Di_Mouse.Acquire
    
    'Indica si el eje Z realiza avance:
    MouseData.Z = Mouse_State.lZ
    
    On Error GoTo ErrOut
    
    MouseButtonPressId = 0
    
    'Solo devolvera verdadero si se ha pulsado dentro de la ventana:
    If GetForegroundWindow = Input_hWnd Then
        Call GetClientRect(Input_hWnd, ClientRect)
        If PtInRect(ClientRect, MouseData.X, MouseData.Y) Then
            If Mouse_State.Buttons(0) <> 0 Then
                MouseData.Left_Button = True
                MouseButtonPressId = 1
                MouseState(0) = MouseState(0) + 1
                Debug.Print MouseState(0)
            Else
                MouseData.Left_Button = False
                MouseState(0) = 0
            End If

            If Mouse_State.Buttons(1) <> 0 Then
                MouseData.Right_Button = True
                MouseButtonPressId = 2
                MouseState(1) = MouseState(1) + 1
            Else
                MouseData.Right_Button = False
                MouseState(1) = 0
            End If

            If Mouse_State.Buttons(2) <> 0 Then
                MouseData.Middle_Button = True
                MouseButtonPressId = 3
                MouseState(2) = MouseState(2) + 1
            Else
                MouseData.Middle_Button = False
                MouseState(2) = 0
            End If

        End If

    End If
    
ErrOut:

End Sub

'Devuelve el estado de los joysticks o pads:
Private Sub GetEventJoy()
    On Error Resume Next
    
    Dim i As Long, j As Byte
    Dim Joy_State As DIJOYSTATE
    
    If (m_Joysticks > 0) Then
        For i = 0 To UBound(Di_Joy)
            JoyButtonPressId(i) = 0
            
            Call Di_Joy(i).GetDeviceStateJoystick(Joy_State)
            If (Err.Number = DIERR_NOTACQUIRED) Or (Err.Number = DIERR_INPUTLOST) Then Di_Joy(i).Acquire
            If GetForegroundWindow = Input_hWnd Then
                With JoyData(i)
                    For j = 1 To 12
                        .Button(j) = Joy_State.Buttons(j - 1)
                        If .Button(j) Then
                            JoyButtonPressId(i) = j
                            JoyState(i, j) = JoyState(i, j) + 1
                        Else
                            JoyState(i, j) = 0
                        End If
                    Next j

                    .X = Joy_State.X
                    If .X > 7500 Then
                        JoyButtonPressId(i) = 15
                        JoyState(i, j) = JoyState(i, 15) + 1
                    ElseIf .X < 2500 Then
                        JoyButtonPressId(i) = 13
                        JoyState(i, j) = JoyState(i, 13) + 1
                    ElseIf .X = 5000 Then
                        JoyState(i, 15) = 0
                        JoyState(i, 13) = 0
                    End If

                    .Y = Joy_State.Y
                    If .Y > 7500 Then
                        JoyButtonPressId(i) = 16
                        JoyState(i, 14) = JoyState(i, 14) + 1
                    ElseIf .Y < 2500 Then
                        JoyButtonPressId(i) = 14
                        JoyState(i, 16) = JoyState(i, 16) + 1
                    ElseIf .Y = 5000 Then
                        JoyState(i, 14) = 0
                        JoyState(i, 16) = 0
                    End If
                End With
            End If
        Next i
    End If
End Sub
