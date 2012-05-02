VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dx_Video_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 166
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
' Name: dx_Video
' Purpose: Clase de funciones de reproducción de video
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 12/04/2004
' Modified: 11/11/2009
'===============================================================================

Option Explicit

Private m_hWnd As Long                      'hWnd del objeto que visualizara video.
Private TimerID As Long

'===============================================================================
' Name: VIDEO_ErrorCodes
'    VIDEO_OK - La operación se ha realizado con éxito.
'    VIDEO_FILENOTLOAD - Error al cargar el archivo.
'    VIDEO_UNKNOWNERROR - Error desconocido.
'    VIDEO_UNKNOWNERROR - Error desconocido.
' Purpose: Códigos de error.
' Remarks: Definen errores o resultados de algunas funciones de esta clase.
'===============================================================================
Public Enum VIDEO_ErrorCodes
    VIDEO_OK = 0
    VIDEO_FILENOTLOAD = (vbObjectError + 4000)
    VIDEO_UNKNOWNERROR = (vbObjectError + 4999)

End Enum

Private Type Video_Pool
    Buffer As IMediaControl
    Video As IBasicVideo
    Cursor As IMediaPosition
    Free As Boolean
    
End Type

Private m_Video() As Video_Pool             'Variable que almacenara los videos.


'===============================================================================
' Name: Init
' Input:
'   ByVal hWnd As Long - Identificador de la ventana.
' Output:
' Purpose: Inicializa el sistema de video.
' Remarks: Inicializa el sistema de video y los subsistemas de la clase habilitando las rutinas y la memoria para trabajar con la reproducción de video.
'===============================================================================
Public Sub Init(hWnd As Long)
Attribute Init.VB_HelpID = 187
On Local Error Resume Next

'Inidicamos el objeto que visualizara los videos (Un Formulario o un PictureBox por ejemplo):
m_hWnd = hWnd

'Activamos la rutina que controla el la reproduccion en bucle:
TimerID = Global_Mod.SetTimer(0, 0, 0, AddressOf VideoEventControl)

End Sub

'===============================================================================
' Name: Terminate
' Purpose: Cierra el sistema de video y los subsistemas de la clase.
' Remarks: Este método se encarga de liberar la memoria ocupada por los recursos de la clase y cerrar las referencias de la misma. Es aconsejable llamar a este metodo justo antes de terminar el programa para así asegurarnos que cualquier proceso de la clase se termina correctamente sin permanecer en memoria.
'===============================================================================
Public Sub Terminate()
On Local Error Resume Next

If Not m_hWnd = 0 Then
    Call Global_Mod.KillTimer(0, TimerID)
    Erase m_Video
    
    With VideoBuffer
        Set .DSAudio = Nothing
        Set .DSControl = Nothing
        Set .DSEvent = Nothing
        Set .DSPosition = Nothing
        Set .DSVideo = Nothing
        Set .DSWindow = Nothing

    End With
        
End If

End Sub
            
'===============================================================================
' Name: VIDEO_Count
' Input:
' Output:
'   Long - Numero de secuencias de video en memoria.
' Purpose: Devuelve el numero de efectos de secuencias de video en memoria.
' Remarks:
'===============================================================================
Public Property Get VIDEO_Count() As Long
Attribute VIDEO_Count.VB_HelpID = 185
On Local Error Resume Next
VIDEO_Count = UBound(m_Video) + 1

End Property
            
'===============================================================================
' Name: VIDEO_Load
' Input:
'   ByVal FileName As String - Nombre y ruta del archivo de video.
' Output:
'   Long - Devuelve el código identificador para la secuencia de video en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar la secuencia de video se devuelve -1.
' Purpose: Carga un archivo de video desde disco como una secuencia de video.
' Remarks: Esta función soporta de forma nativa formatos *.AVI y *.MPG (MPEG1 y 2). Se pueden procesar otros formatos de audio como el *.WMV o DivX si se instalan los codecs necesarios para DirectShow. Se pueden cargar tantas secuencias de video como se deseen teniendo como limite la memoria del sistema.
'===============================================================================
Public Function VIDEO_Load(Filename As String) As Long
Attribute VIDEO_Load.VB_HelpID = 184
On Local Error GoTo ErrLine

Dim i As Long

If Not m_hWnd = 0 And Global_Mod.FileExists(Filename) Then
    i = -1
    i = Get_FreeID()

    With m_Video(i)
            
        Set .Buffer = New FilgraphManager
        Call .Buffer.RenderFile(Filename)
        Set .Video = .Buffer
        Set .Cursor = .Buffer
            
        .Free = False
            
    End With

    VIDEO_Load = i
    Exit Function
    
End If
            
ErrLine:
If Not i = -1 Then Call Me.VIDEO_Unload(i)
VIDEO_Load = VIDEO_FILENOTLOAD

End Function
            
'===============================================================================
' Name: VIDEO_Play
' Input:
'   ByVal Video As Long - Identificador de la secuencia de video.
'   ByVal X As Long - Coordenada horizontal de la pantalla donde se renderizara el video.
'   ByVal Y As Long - Coordenada vertical de la pantalla donde se renderizara el video.
'   ByVal Width As Long - Ancho con el que se renderizara el video. Si el valor es 0 se tomara el ancho original del video.
'   ByVal Height As Long - Alto con el que se renderizara el video. Si el valor es 0 se tomara el alto original del video.
' Output:
'   Boolean - Devuelve verdadero si la operación se realiza con éxito.
' Purpose: Reproduce una secuencia de video.
' Remarks: Esta clase esta diseñada para poder reproducir videos para cinematicas o introducciones en el programa. Dado que la implementación del mismo no permite integrarse a la perfección con la clase grafica se recomienda alternar el uso de una clase con la otra y se limita por ello la posibilidad de reproducir solo 1 video a la vez. Es aconsejable a la hora de reproducir una secuencia de video no llamar a la función Frame de la clase grafica dx_GFX.
'===============================================================================
Public Function VIDEO_Play(Video As Long, X As Long, Y As Long, Width As Long, Height As Long) As Boolean
Attribute VIDEO_Play.VB_HelpID = 183
On Error GoTo ErrOut

If Not m_hWnd = 0 Then
    With VideoBuffer
        Set .DSControl = m_Video(Video).Buffer
    
        On Local Error Resume Next
        Set .DSAudio = .DSControl
        .DSAudio.volume = 0
        .DSAudio.Balance = 0
        
        On Error GoTo ErrOut
        Set .DSWindow = .DSControl
        .DSWindow.WindowStyle = &H6000000
        .DSWindow.Top = Y
        .DSWindow.Left = X
        If (Width > 0 Or Height > 0) Then
            .DSWindow.Width = Width
            .DSWindow.Height = Height
        Else
            Dim realWidth As Long, realHeight As Long
            Call m_Video(Video).Video.GetVideoSize(realWidth, realHeight)
            .DSWindow.Width = realWidth
            .DSWindow.Height = realHeight
        End If
        
        .DSWindow.Owner = m_hWnd
        
        Set .DSEvent = .DSControl
        Set .DSPosition = m_Video(Video).Cursor
        If ObjPtr(.DSPosition) Then .DSPosition.Rate = 1#
        .DSPosition.CurrentPosition = 0
        
        .DSControl.Run
        
        .Playing = True
        
    End With

    VIDEO_Play = True
    
End If

Exit Function

ErrOut:
End Function

'===============================================================================
' Name: VIDEO_IsPlaying
' Input:
' Output:
'   Boolean - Devuelve verdadero si la secuencia de video esta ejecutándose.
' Purpose: Comprueba si la secuencia de video esta en reproducción.
' Remarks: Esta función es útil a la hora de determinar si una secuencia de video ha dejado de reproducirse o no.
'===============================================================================
Public Function VIDEO_IsPlaying() As Boolean
Attribute VIDEO_IsPlaying.VB_HelpID = 182
On Local Error Resume Next

If Not m_hWnd = 0 Then VIDEO_IsPlaying = VideoBuffer.Playing

End Function

'===============================================================================
' Name: VIDEO_Pause
' Input:
' Output:
' Purpose: Detiene la reproducción de una secuencia de video manteniendo la posición de lectura.
' Remarks:
'===============================================================================
Public Sub VIDEO_Pause()
Attribute VIDEO_Pause.VB_HelpID = 181
On Error GoTo ErrOut

If Not m_hWnd = 0 Then
    If VideoBuffer.Playing Then
        Call VideoBuffer.DSControl.Pause
        VideoBuffer.Playing = False
    
    Else
        Call VideoBuffer.DSControl.Run
        VideoBuffer.Playing = True
        
    End If
    
End If

ErrOut:

End Sub

'===============================================================================
' Name: VIDEO_Stop
' Input:
' Output:
' Purpose: Detiene la reproducción de una secuencia de video y resetea la posicion de lectura.
' Remarks: Esta función resetea la posición de lectura al inicio de la secuencia de video. Para realizar una pausa en la secuencia de video utilizar VIDEO_Pause.
'===============================================================================
Public Sub VIDEO_Stop()
Attribute VIDEO_Stop.VB_HelpID = 180
On Error GoTo ErrOut

If Not m_hWnd = 0 Then
    Call VideoBuffer.DSControl.Stop
    VideoBuffer.DSPosition.CurrentPosition = 0
    VideoBuffer.Playing = False

End If

ErrOut:

End Sub

'===============================================================================
' Name: VIDEO_GetRate
' Input:
' Output:
'   Long - Valor comprendido entre 0 y 220 que define la velocidad de reproduccion de la secuencia de video.
' Purpose: Devuelve la velocidad de reproducción de la secuencia de video.
' Remarks: Por defecto la velocidad de reproducción se establece a 100.
'===============================================================================
Public Function VIDEO_GetRate() As Long
Attribute VIDEO_GetRate.VB_HelpID = 179
On Local Error Resume Next

If Not m_hWnd = 0 Then VIDEO_GetRate = CLng(VideoBuffer.DSPosition.Rate * 100)

End Function

'===============================================================================
' Name: VIDEO_GetPan
' Input:
' Output:
'   Long - Valor comprendido entre -100 y 100 que define el nivel de balance de la secuencia de video.
' Purpose: Devuelve el nivel del balance de la secuencia de video.
' Remarks:
'===============================================================================
Public Function VIDEO_GetPan() As Long
Attribute VIDEO_GetPan.VB_HelpID = 178
On Local Error Resume Next

If Not m_hWnd = 0 Then VIDEO_GetPan = VideoBuffer.DSAudio.Balance / 100

End Function

'===============================================================================
' Name: VIDEO_GetVolume
' Input:
' Output:
'   Long - Valor comprendido entre 0 y 100 que define el nivel de volumen de la secuencia de video.
' Purpose: Devuelve el nivel del volumen de la secuencia de video.
' Remarks:
'===============================================================================
Public Function VIDEO_GetVolume() As Long
Attribute VIDEO_GetVolume.VB_HelpID = 177
On Local Error Resume Next

If Not m_hWnd = 0 Then VIDEO_GetVolume = (VideoBuffer.DSAudio.volume / 100) + 100

End Function

'===============================================================================
' Name: VIDEO_GetCurrentPosition
' Input:
' Output:
'   Long - Valor que define la posición de lectura de la secuencia de video.
' Purpose: Devuelve la posición de lectura de la secuencia de video.
' Remarks:
'===============================================================================
Public Function VIDEO_GetCurrentPosition() As Long
Attribute VIDEO_GetCurrentPosition.VB_HelpID = 176
On Local Error Resume Next

If Not m_hWnd = 0 Then VIDEO_GetCurrentPosition = VideoBuffer.DSPosition.CurrentPosition * 1000

End Function

'===============================================================================
' Name: VIDEO_GetDuration
' Input:
'   Optional ByVal Video As Long - Identificador de la secuencia de video. Por defecto, si no se especifica valor alguno, se calcula la duración del video que esta actualmente en reproducción.
' Output:
'   Long - Valor que define en milisegundos la duración de una secuencia de video.
' Purpose: Devuelve en milisegundos la duración de una secuencia de video.
' Remarks:
'===============================================================================
Public Function VIDEO_GetDuration(Optional Video As Long = -1) As Long
Attribute VIDEO_GetDuration.VB_HelpID = 175
On Local Error Resume Next

If Not m_hWnd = 0 Then
    If Video = -1 Then
        VIDEO_GetDuration = VideoBuffer.DSPosition.Duration * 1000
        
    Else
        VIDEO_GetDuration = m_Video(Video).Cursor.Duration * 1000
        
    End If

End If

End Function

'===============================================================================
' Name: VIDEO_SetRate
' Input:
'   ByVal Value As Long - Valor comprendido entre 0 y 220 que establecerá la velocidad de reproduccion de la secuencia de sonido.
' Output:
' Purpose: Establece la velocidad de reproducción de la secuencia de sonido.
' Remarks: El valor para la velocidad normal de una secuencia de video es de 100.
'===============================================================================
Public Sub VIDEO_SetSpeed(Value As Long)
Attribute VIDEO_SetSpeed.VB_HelpID = 174
On Local Error Resume Next

If Not m_hWnd = 0 Then
    If Value > 220 Then Value = 220 Else If Value < 0 Then Value = 0

    VideoBuffer.DSPosition.Rate = CSng(Value / 100)
    
End If

End Sub

'===============================================================================
' Name: VIDEO_SetVolume
' Input:
'   ByVal Value As Long - Valor comprendido entre 0 y 100 que establecerá el nivel de volumen de la secuencia de video.
' Output:
' Purpose: Establece el nivel del volumen de la secuencia de video.
' Remarks:
'===============================================================================
Public Sub VIDEO_SetVolume(Value As Long)
Attribute VIDEO_SetVolume.VB_HelpID = 173
On Local Error Resume Next

If Not m_hWnd = 0 Then
    If Value > 100 Then Value = 100 Else If Value < 0 Then Value = 0

    VideoBuffer.DSAudio.volume = (Value * 100) + -4000
    
End If

End Sub

'===============================================================================
' Name: VIDEO_SetPan
' Input:
'   ByVal Value As Long - Valor comprendido entre -100 y 100 que establecerá el nivel de balance de la secuencia de video. El rango de valores negativos hasta representa el canal izquierdo y los valores positivos el canal derecho y 0 representa el valor neutro para el balance entre ambos canales.
' Output:
' Purpose: Establece el nivel del balance de la secuencia de video.
' Remarks:
'===============================================================================
Public Sub VIDEO_SetPan(Value As Long)
Attribute VIDEO_SetPan.VB_HelpID = 172
On Local Error Resume Next

If Not m_hWnd = 0 Then
    If Value > 100 Then Value = 100 Else If Value < -100 Then Value = -100

    VideoBuffer.DSAudio.Balance = Value * 100
    
End If

End Sub

'===============================================================================
' Name: VIDEO_SetCurrentPosition
' Input:
'   ByVal Value As Long - Valor que define la posición de lectura de la secuencia de video.
' Output:
' Purpose: Establece la posición de lectura de la secuencia de video.
' Remarks:
'===============================================================================
Public Sub VIDEO_SetCurrentPosition(Value As Long)
Attribute VIDEO_SetCurrentPosition.VB_HelpID = 171
On Local Error Resume Next

If Not m_hWnd = 0 Then
    If Value > VideoBuffer.DSPosition.Duration * 1000 Then Value = 0 Else If Value < 0 Then Value = 0

    VideoBuffer.DSPosition.CurrentPosition = Value / 1000
    
End If

End Sub

'===============================================================================
' Name: VIDEO_Unload
' Input:
'   ByVal Video As Long - Identificador de la secuencia de video.
' Output:
' Purpose: Descarga una secuencia de video de la memoria.
' Remarks:
'===============================================================================
Public Sub VIDEO_Unload(Video As Long)
Attribute VIDEO_Unload.VB_HelpID = 170
On Error GoTo ErrOut

If Not m_hWnd = 0 Then
    If Video = UBound(m_Video) Then
        If Video = 0 Then Erase m_Video Else ReDim Preserve m_Video(UBound(m_Video) - 1) As Video_Pool
    
    Else
        With m_Video(Video)
            Set .Buffer = Nothing
            Set .Cursor = Nothing
            Set .Video = Nothing
            .Free = True
        End With

    End If
    
End If

ErrOut:
End Sub

'===============================================================================
' Name: VIDEO_GetSize
' Input:
'   ByVal Video As Long - Identificador de la secuencia de video.
'   ByVal Width As Long - Ancho en pixeles del video.
'   ByVal Height As Long - Altura en pixeles del video.
' Output:
' Purpose: Devuelve las dimensiones originales del video.
' Remarks:
'===============================================================================
Public Sub VIDEO_GetSize(Video As Long, Width As Long, Height As Long)
    Call m_Video(Video).Video.GetVideoSize(Width, Height)
End Sub


'===============================================================================
' Name: VIDEO_GetViewPort
' Input:
'   ByVal Data As GFX_Rect - Argumento que nos devuelve la posición y la escala con la que se renderizara el video que esta reproduciéndose.
' Output:
' Purpose: Devuelve la posición y escala de la secuencia de video que se esta reproduciendo.
' Remarks:
'===============================================================================
Public Sub VIDEO_GetViewPort(Data As GFX_Rect)
Attribute VIDEO_GetViewPort.VB_HelpID = 169
On Error GoTo ErrOut
    
If Not m_hWnd = 0 Then
    With VideoBuffer.DSWindow
        Data.X = .Left
        Data.Y = .Top
        Data.Height = .Height
        Data.Width = .Width
    End With
    
End If

Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: VIDEO_SetViewPort
' Input:
'   ByVal Data As GFX_Rect - Argumento que define la posición y la escala con la que se renderizara el video que esta reproduciendo.
' Output:
' Purpose: Establece la posición y escala de la secuencia de video que se esta reproduciendo.
' Remarks:
'===============================================================================
Public Sub VIDEO_SetViewPort(Data As GFX_Rect)
Attribute VIDEO_SetViewPort.VB_HelpID = 168
On Local Error GoTo ErrOut
    
If Not m_hWnd = 0 Then
    With VideoBuffer.DSWindow
        .Left = Data.X
        .Top = Data.Y
        .Height = Data.Height
        .Width = Data.Width
        
    End With

End If

ErrOut:
End Sub

'Devuelve un identificador para el video en memoria:
Private Function Get_FreeID() As Long
On Local Error Resume Next

Dim i As Long, j As Long

i = UBound(m_Video)
        
If Err.Number = 9 Then
    ReDim m_Video(0) As Video_Pool
    Get_FreeID = 0
            
Else
    For j = 0 To i
        If Not m_Video(j).Free Then
            Get_FreeID = j
            Exit For
        End If
    Next j
            
    ReDim Preserve m_Video(i + 1) As Video_Pool
            
    Get_FreeID = UBound(m_Video)
            
End If
        
End Function

'===============================================================================
' Name: hWnd
' Input:
' Output:
'   Long - Identificador de la ventana.
' Purpose: Devuelve el identificador de la ventana que esta usando para ejecutar las rutinas de la clase.
' Remarks:
'===============================================================================
Public Property Get hWnd() As Long
Attribute hWnd.VB_HelpID = 167
hWnd = m_hWnd

End Property
