Attribute VB_Name = "DirectShowEvent_Mod"
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
' Purpose: Modulo para controlar los eventos de reproduccion de DirectShow
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 04/11/2005
' Modified: 11/11/2009
'===============================================================================

Option Explicit

Public Type Audio_Buffer
    DSAudio  As IBasicAudio             'Basic Audio Objectt
    DSEvent As IMediaEvent              'MediaEvent Object
    DSControl As IMediaControl          'MediaControl Object
    DSPosition As IMediaPosition        'MediaPosition Object
    Playing As Boolean
    Looping As Boolean
End Type

Public Type Video_Buffer
    DSAudio As IBasicAudio           'Basic Audio Object
    DSVideo As IBasicVideo           'Basic Video Object
    DSEvent As IMediaEvent           'VideoEvent Object
    DSWindow As IVideoWindow         'VideoWindow Object
    DSControl As IMediaControl       'VideoControl Object
    DSPosition As IMediaPosition     'VideoPosition Object
    Playing As Boolean
End Type

Public AudioBuffer() As Audio_Buffer   'Buffer primario y secundario para reproducir los archivos de musica.
Public VideoBuffer As Video_Buffer     'Buffer para reproducir los archivos de video.

'Esta funcion se encarga de controlar eventos del modulo de sonido mientras se ejecuta el programa:
Public Sub AudioEventControl()
Dim i As Long

For i = 0 To 1
    If Not AudioBuffer(i).DSPosition Is Nothing Then
        'Si Looping esta activo se reproduce la musica en bucle cerrado:
        If AudioBuffer(i).Playing And AudioBuffer(i).DSPosition.CurrentPosition = AudioBuffer(i).DSPosition.Duration Then
            If AudioBuffer(i).Looping Then
                AudioBuffer(i).DSPosition.CurrentPosition = 0
            Else
                AudioBuffer(i).Playing = False
            End If
        End If
    End If
Next i

End Sub

'Esta funcion se encarga de controlar eventos del modulo de video mientras se ejecuta el programa:
Public Sub VideoEventControl()
'Si Playing = True, Looping = False y la posicion de lectura ha llegado al final entonces Playing = False
If VideoBuffer.Playing And VideoBuffer.DSPosition.CurrentPosition = VideoBuffer.DSPosition.Duration Then VideoBuffer.Playing = False

End Sub
