VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dx_Sound_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 100
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
' Name: dx_Sound
' Purpose: Clase de funciones de audio y musica
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 07/08/2001
' Modified: 21/01/2010
'===============================================================================

Option Explicit

'===============================================================================
' Name: SND_ErrorCodes
'    SND_OK - La operación se ha realizado con éxito.
'    SND_FILENOTLOAD - Error al cargar el archivo.
'    SND_CHANNELUNAVAILABLE - No hay canales disponibles.
'    SND_UNKNOWNERROR - Error desconocido.
'    SND_UNKNOWNERROR - Error desconocido.
' Purpose: Códigos de error.
' Remarks: Definen errores o resultados de algunas funciones de esta clase.
'===============================================================================
Public Enum SND_ErrorCodes
    SND_OK = 0
    SND_FILENOTLOAD = (vbObjectError + 2000)
    SND_CHANNELUNAVAILABLE = (vbObjectError + 2001)
    SND_UNKNOWNERROR = (vbObjectError + 2999)

End Enum

'===============================================================================
' Name: Sound_Buffer
'   Primary_Buffer - Buffer primario de música.
'   Secondary_Buffer - Buffer secundario de música.
'   Secondary_Buffer - Buffer secundario de música.
' Purpose: Constantes que definen el buffer de música.
' Remarks:
'===============================================================================
Public Enum Sound_Buffer
    Primary_Buffer = 0
    Secondary_Buffer = 1
    
End Enum

'===============================================================================
' Name: Sound_Effects
'   FX_Chorus - Genera un efecto de coro en el sonido.
'   FX_Compressor - Aplica un efecto de compresion en las ondas del sonido.
'   FX_Echo - Genera eco en el sonido.
'   FX_Distortion - Aplica distorsion en las ondas del sonido.
'   FX_Flanger - Afecta a los tonos bajos.
'   FX_Gargle - Afecta a los tonos graves.
'   FX_Waves_Reverb - Aplica un efecto de reverberencia de ondas en el sonido.
' Purpose: Constantes que definen efectos especiales para sonidos.
' Remarks:
'===============================================================================
Public Enum Sound_Effects
    FX_Chorus = &H40 ' = DSFX_STANDARD_CHORUS
    FX_Compressor = &H100 ' = DSFX_STANDARD_COMPRESSOR
    FX_Echo = &H10 ' = DSFX_STANDARD_ECHO
    FX_Distortion = &H20 ' = DSFX_STANDARD_DISTORTION
    FX_Flanger = &H80 ' = DSFX_STANDARD_FLANGER
    FX_Gargle = &H200 ' = DSFX_STANDARD_GARGLE
    FX_Waves_Rerverb = &H800 ' = DSFX_STANDARD_WAVES_REVERB
End Enum

'Estructuras de datos:
Private Type Sound_Pool
    Buffer As DirectSoundSecondaryBuffer8
    BufferSize As Long
    Free As Boolean
End Type

Private Type Music_Pool
    Buffer As IMediaControl
    Cursor As IMediaPosition
    Free As Boolean
End Type
    
Private Type Channel_Pool
    Buffer As DirectSoundSecondaryBuffer8
    Looping As Boolean
End Type

'Variables:
Dim Dx As DirectX8                      'DirectX 8
Dim Ds As DirectSound8                  'Direct Sound 8
Dim DSEnum As DirectSoundEnum8          'Allows us to get information on available hardware/software devices.
Dim DS_Desc As DSBUFFERDESC             ' Describe un buffer no apot para aplicar efectos (se usara como temporal para cargar sonidos demasiado cortos para ser cargados en un buffer de efectos)
Dim DS_Desc_FX As DSBUFFERDESC          ' Describe un buffer apto para aplicar efectos.

Dim soundBuffer() As Byte               ' Buffer para realizar duplicados de las muestras de sonido en los canales (el metodo DuplicateBuffer no funciona con efectos activados).

Dim m_Sound() As Sound_Pool             'Array para almacenar los sonidos.
Dim m_Music() As Music_Pool             'Array para almacenar los modulos musicales.

Dim m_Channel() As Channel_Pool, Ch As Long 'Buffer de canales.

Dim CD_In_Pause As Boolean              'Flag que indica si el CD-Audio esta en pausa o no.
Dim CD_Open As Boolean                  'Flag que indica si la bandeja del CD-Audio esta abierta o no.

Dim m_hWnd As Long
Dim TimerID As Long

Dim DS_Init As Boolean
Dim CD_Init As Boolean                  'Indica si el cd se ha utilizado.

'===============================================================================
' Name: Init
' Input:
'   ByVal hWnd As Long - Identificador de la ventana.
'   ByVal hWnd As Long - Identificador de la ventana.
'   Optional ByVal Channels As Long - Numero de canales para efectos de sonido. Por defecto esta establecido a 16. Se pueden establecer un máximo de 64 y un mínimo de 8 canales. Los canales de efectos de sonido son independientes de los dos canales disponibles para música.
'   Optional ByVal Channels As Long - Numero de canales para efectos de sonido. Por defecto esta establecido a 16. Se pueden establecer un máximo de 64 y un mínimo de 8 canales. Los canales de efectos de sonido son independientes de los dos canales disponibles para música.
' Output:
'    Boolean - Devuelve Verdadero en caso de inicializarse correctamente.
' Purpose: Inicializa el sistema de audio y música.
' Remarks: Inicializa el sistema de audio y música y los subsistemas de la clase habilitando las rutinas y la memoria para trabajar con el audio y la música.
'===============================================================================
Public Function Init(hWnd As Long, Optional Channels As Long = 16) As Boolean
Attribute Init.VB_HelpID = 150
On Error GoTo ErrOut
Dim i As Long

If Not DS_Init Then
    
    Set Dx = New DirectX8
    Set DSEnum = Dx.GetDSEnum

    m_hWnd = hWnd

    If Channels <= 8 Then Channels = 8 Else If Channels >= 64 Then Channels = 64
    
    ReDim m_Channel(Channels - 1) As Channel_Pool
    
    ReDim AudioBuffer(1) As Audio_Buffer
        
    Ch = 0
    
    'Inicializamos el motor de sonido:
    Set Ds = Dx.DirectSoundCreate(vbNullString)
    Call Ds.SetCooperativeLevel(m_hWnd, DSSCL_PRIORITY)
    
    DS_Desc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_GLOBALFOCUS
    DS_Desc_FX.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_GLOBALFOCUS Or DSBCAPS_CTRLFX
    
    'Activamos la rutina que controla el la reproduccion en bucle:
    TimerID = Global_Mod.SetTimer(0, 0, 0, AddressOf AudioEventControl)
    
    DS_Init = True
    Init = True
    
    Exit Function
   
End If

ErrOut:

End Function

'===============================================================================
' Name: Terminate
' Purpose: Cierra el sistema de audio y música y los subsistemas de la clase.
' Remarks: Este método se encarga de liberar la memoria ocupada por los recursos de la clase y cerrar las referencias de la misma. Es aconsejable llamar a este método justo antes de terminar el programa para así asegurarnos que cualquier proceso de la clase se termina correctamente sin permanecer en memoria.
'===============================================================================
Public Sub Terminate()
Attribute Terminate.VB_HelpID = 149
On Error GoTo ErrOut

If DS_Init Then
    
    Erase m_Sound, m_Channel
        
    Set DSEnum = Nothing
    Set Ds = Nothing
    Set Dx = Nothing

    Erase m_Music, AudioBuffer

    'Cerramos el CD:
    If CD_Init Then
        Call CD_Stop
        Call SendCommand("set cd door closed")
        Call SendCommand("close cd")
    
        Call CDSoundVolume_Mod.CloseMixer
    
        CD_Init = False
    
    End If

    'Cerramos la entrada de sonido:
    Call SendCommand("close WavFile")
    
    'Si se activo el Timer se cierra:
    Call Global_Mod.KillTimer(0, TimerID)
    
End If

Exit Sub

ErrOut:
    
End Sub

'//DIRECT SOUND AUDIO:

'===============================================================================
' Name: SOUND_Load
' Input:
'   ByVal FileName As String - Nombre y ruta del archivo de sonido.
' Output:
'   Long - Devuelve el código identificador para el sonido en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar el sonido se devuelve -1.
' Purpose: Carga un archivo de audio desde disco como un efecto de sonido.
' Remarks: Esta función solo carga archivos de ondas (WAVE) *.WAV. Se pueden cargar tantos efectos de sonido como se deseen teniendo como limite la memoria del sistema.
'===============================================================================
Public Function SOUND_Load(Filename As String) As Long
Attribute SOUND_Load.VB_HelpID = 148
On Error GoTo ErrOut
Dim i As Long
Dim Caps As DSBCAPS
        
If DS_Init And Global_Mod.FileExists(Filename) Then
    i = -1
    i = Get_FreeID(0)
    
    On Error Resume Next
    Set m_Sound(i).Buffer = Ds.CreateSoundBufferFromFile(Filename, DS_Desc_FX)
    If Err.Number = DSERR_BUFFERTOOSMALL Then ' Si el archivo es demasiado pequeño le agregamos silencio al final:
        Call Err.Clear
        
        Dim dsBuffer As DirectSoundSecondaryBuffer8
        Set dsBuffer = Ds.CreateSoundBufferFromFile(Filename, DS_Desc)
        
        Call dsBuffer.GetCaps(Caps)
        
        Dim Buffer() As Byte
        ReDim Buffer(Caps.lBufferBytes)
        Call dsBuffer.ReadBuffer(0, 0, Buffer(0), DSBLOCK_ENTIREBUFFER)
        
        Dim Format As WAVEFORMATEX
        Call dsBuffer.GetFormat(Format)
        
        Dim fxLen As Long
        fxLen = Format.lAvgBytesPerSec * 0.15 ' Longitud minima que deberia tener el buffer.
        
        Dim currentLen As Long
        currentLen = UBound(Buffer) ' Longitud actual del buffer.
        
        Dim silentLen As Long
        silentLen = fxLen - currentLen ' Bytes que hay que rellenar con silencios.
        
        ReDim Preserve Buffer(fxLen) ' Modificamos el tamaño del buffer al tamaño minimo para los efectos.
        
        Dim Value As Byte
        If Format.nBitsPerSample = 8 Then Value = 128 Else Value = 0
        
        ' Rellenamos los nuevos bytes con el valor adecuado para el silencio:
        Call Global_Mod.FillMemory(Buffer(currentLen), silentLen, Value)
        
'        ' Creamos el buffer configurado para usar efectos:
'        Set m_Sound(i).Buffer = Ds.CreateSoundBuffer(DS_Desc_FX) ' ***** Si sigue fallando y no encontramos solucion se optara por guardar el buffer modificado como un
'        ' wav temporal que se cargara mediante CreateSoundBufferFromFile() con los flags correctos *****
'
'        ' Escribimos el buffer modificado con el silencio:
'        Call m_Sound(i).Buffer.WriteBuffer(0, 0, Buffer(0), DSBLOCK_ENTIREBUFFER)

        ' Creamos un nuevo buffer para escriir el buffer modificado:
        Dim saveBuffer As DirectSoundSecondaryBuffer8
        DS_Desc.lBufferBytes = fxLen ' Definimos el tamaño que tendra el buffer.
        Set saveBuffer = Ds.CreateSoundBuffer(DS_Desc)
        
        ' Escribimos el buffer modificado con el silencio:
        Call saveBuffer.WriteBuffer(0, 0, Buffer(0), DSBLOCK_ENTIREBUFFER)
        
        ' Informacion de depuracion:
'        Call saveBuffer.GetCaps(Caps)
'        Debug.Print Caps.lBufferBytes
        
        ' Guardamos el buffer modificado como un wav temporal:
        Dim tmpName As String
        tmpName = Global_Mod.GetTempFileName
        Call saveBuffer.SaveToFile(tmpName)
        
        ' Cargamos el wav temporal en el buffer con efectos:
        Set m_Sound(i).Buffer = Ds.CreateSoundBufferFromFile(tmpName, DS_Desc_FX)
        
        Call Kill(tmpName) ' Borramos el temporal.
        
        Set dsBuffer = Nothing
        Set saveBuffer = Nothing
        
    ElseIf Err.Number <> 0 Then
    
        If Not i = -1 Then Call Me.SOUND_Unload(i)
        SOUND_Load = SND_FILENOTLOAD
        Exit Function
        
    End If
    
    Call m_Sound(i).Buffer.GetCaps(Caps)
    m_Sound(i).BufferSize = Caps.lBufferBytes

    SOUND_Load = i

End If

Exit Function

ErrOut:
If Not i = -1 Then Call Me.SOUND_Unload(i)
SOUND_Load = SND_UNKNOWNERROR

End Function

'===============================================================================
' Name: SOUND_Unload
' Input:
'   ByVal Sound As Long - Identificador del sonido.
' Output:
' Purpose: Descarga un efecto de sonido de la memoria.
' Remarks:
'===============================================================================
Public Sub SOUND_Unload(Sound As Long)
Attribute SOUND_Unload.VB_HelpID = 147
On Error GoTo ErrOut

If DS_Init Then
    If Sound = UBound(m_Sound) Then
        If Sound = 0 Then Erase m_Sound Else ReDim Preserve m_Sound(UBound(m_Sound) - 1) As Sound_Pool
    
    Else
        With m_Sound(Sound)
            Set .Buffer = Nothing
            .Free = True
        End With

    End If
    
End If

Exit Sub

ErrOut:

End Sub

'===============================================================================
' Name: SOUND_Play
' Input:
'   ByVal Sound As Long - Identificador del efecto de sonido.
'   Optional ByVal Channel As Long - Canal donde se reproducirá el efecto de sonido. Por defecto se usa el primer canal libre que se encuentre.
'   Optional ByVal Looping As Boolean - Indica si se reproducirá el efecto de sonido en bucle.
' Output:
'   Long - Devuelve el canal donde se reproducirá el sonido. Esta valor es necesario al hora de modificar parámetros en la reproducción de un sonido, para ponerlo en pausa o detenerlo. En caso de no poder reproducir el efecto de sonido devuelve -1.
' Purpose: Reproduce un efecto de sonido.
' Remarks:
'===============================================================================
Public Function SOUND_Play(Sound As Long, Optional Channel As Long = -1, Optional Looping As Boolean = False) As Long
Attribute SOUND_Play.VB_HelpID = 146
On Error GoTo ErrOut

If DS_Init Then
    'Reproducimos el sonido y devolvemos el canal donde se esta reproduciendo:
    If Not m_Sound(Sound).Free Then
        If Channel = -1 Then Channel = Get_FreeID(2)
        
        'Si no se ha logrado un canal libre salimos de la funcion:
        If Channel = -1 Then
            SOUND_Play = SND_CHANNELUNAVAILABLE
            Exit Function
            
        End If
        
        ' Si el buffer es menor del tamaño minimo para efectos se redimensiona al tamaño minimo:
        If m_Sound(Sound).BufferSize < 10000 Then
            ReDim soundBuffer(10000)
        Else
            ReDim soundBuffer(m_Sound(Sound).BufferSize)
        End If
            
        Call m_Sound(Sound).Buffer.ReadBuffer(0, 0, soundBuffer(0), DSBLOCK_ENTIREBUFFER)
        
        DS_Desc_FX.lBufferBytes = m_Sound(Sound).BufferSize
        Set m_Channel(Channel).Buffer = Ds.CreateSoundBuffer(DS_Desc_FX)
        Call m_Channel(Channel).Buffer.WriteBuffer(0, 0, soundBuffer(0), DSBLOCK_ENTIREBUFFER)
        
        Call m_Channel(Channel).Buffer.Play(Abs(Looping))
        m_Channel(Channel).Looping = Looping

        SOUND_Play = Channel
        
        Exit Function
        
    End If

End If

ErrOut:
SOUND_Play = SND_UNKNOWNERROR
'MsgBox (Err.Number = DSERR_BUFFERTOOSMALL): MsgBox Err.Number: Clipboard.SetText CStr(Err.Number)
End Function

'===============================================================================
' Name: SOUND_IsPlaying
' Input:
'   ByVal Channel As Long - Canal al que se realizara la comprobación.
' Output:
'   Boolean - Devuelve verdadero si el canal esta en uso.
' Purpose: Comprueba si el canal esta siendo utilizado o no.
' Remarks: Esta función es útil a la hora de determinar si un efecto de sonido ha dejado de reproducirse o no.
'===============================================================================
Public Function SOUND_IsPlaying(Channel As Long) As Boolean
Attribute SOUND_IsPlaying.VB_HelpID = 145
On Local Error Resume Next
Dim i As Long

If DS_Init Then
    i = m_Channel(Channel).Buffer.GetStatus

    If i = DSBSTATUS_PLAYING Or i = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING Then SOUND_IsPlaying = True
    
End If

End Function

'===============================================================================
' Name: SOUND_Stop
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
' Output:
' Purpose: Detiene la reproducción de un efecto de sonido y resetea la posición de lectura.
' Remarks: Esta función resetea la posición de lectura al inicio del efecto de sonido y libera el canal en uso. Para realizar una pausa del canal utilizar SOUND_Pause.
'===============================================================================
Public Sub SOUND_Stop(Channel As Long)
Attribute SOUND_Stop.VB_HelpID = 144
On Local Error Resume Next
Dim i As Long

If DS_Init Then
    If Channel > -1 Then
        i = m_Channel(Channel).Buffer.GetStatus

        If i = DSBSTATUS_PLAYING Or i = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING Then
            Call m_Channel(Channel).Buffer.Stop
            Call m_Channel(Channel).Buffer.SetCurrentPosition(0)
    
        End If
        
    Else
        Call SetAllSound(0)
    
    End If
    
End If

End Sub

'===============================================================================
' Name: SOUND_Pause
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
' Output:
' Purpose: Detiene la reproducción de un efecto de sonido manteniendo la posición de lectura.
' Remarks:
'===============================================================================
Public Sub SOUND_Pause(Channel As Long)
Attribute SOUND_Pause.VB_HelpID = 143
On Local Error Resume Next
Dim i As Long

If DS_Init Then
    If Channel > -1 Then
        i = m_Channel(Channel).Buffer.GetStatus

        If i = DSBSTATUS_PLAYING Or i = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING Then
            Call m_Channel(Channel).Buffer.Stop
    
        Else
            Call m_Channel(Channel).Buffer.Play(Abs(m_Channel(Channel).Looping))
    
        End If
    
    Else
        Call SetAllSound(1)
    
    End If
    
End If
End Sub

'===============================================================================
' Name: SOUND_SetVolume
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
'   ByVal Value As Long - Valor comprendido entre 0 y 100 que establecerá el nivel de volumen del canal de audio.
' Output:
' Purpose: Establece el nivel del volumen del canal de audio.
' Remarks:
'===============================================================================
Public Sub SOUND_SetVolume(Channel As Long, Value As Long)
Attribute SOUND_SetVolume.VB_HelpID = 142
On Local Error Resume Next

If DS_Init Then
    If Value > 100 Then Value = 100 Else If Value < 0 Then Value = 0

    Call m_Channel(Channel).Buffer.SetVolume((Value * 100) + DSBVOLUME_MIN)
    
End If

End Sub

'===============================================================================
' Name: SOUND_GetVolume
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
' Output:
'   Long - Valor comprendido entre 0 y 100 que define el nivel de volumen del canal de audio.
' Purpose: Devuelve el nivel del volumen del canal de audio.
' Remarks:
'===============================================================================
Public Function SOUND_GetVolume(Channel As Long) As Long
Attribute SOUND_GetVolume.VB_HelpID = 141
On Local Error Resume Next

If DS_Init Then SOUND_GetVolume = (m_Channel(Channel).Buffer.GetVolume() / 100) + 100

End Function

'===============================================================================
' Name: SOUND_SetPan
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
'   ByVal Value As Long - Valor comprendido entre -100 y 100 que establecerá el nivel de balance del canal de audio. El rango de valores negativos hasta representa el canal izquierdo y los valores positivos el canal derecho y 0 representa el valor neutro para el balance entre ambos canales.
' Output:
' Purpose: Establece el nivel del balance del canal de audio.
' Remarks:
'===============================================================================
Public Sub SOUND_SetPan(Channel As Long, Value As Long)
Attribute SOUND_SetPan.VB_HelpID = 140
On Local Error Resume Next

If DS_Init Then
    If Value > 100 Then Value = 100 Else If Value < -100 Then Value = -100

    Call m_Channel(Channel).Buffer.SetPan(Value * 100)
    
End If

End Sub

'===============================================================================
' Name: SOUND_GetPan
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
' Output:
'   Long - Valor comprendido entre -100 y 100 que define el nivel de balance del canal de audio.
' Purpose: Devuelve el nivel del balance del canal de audio.
' Remarks:
'===============================================================================
Public Function SOUND_GetPan(Channel As Long) As Long
Attribute SOUND_GetPan.VB_HelpID = 139
On Local Error Resume Next

If DS_Init Then SOUND_GetPan = m_Channel(Channel).Buffer.GetPan() / 100

End Function

'===============================================================================
' Name: SOUND_SetFrequency
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el efecto de sonido.
'   ByVal Value As Long - Valor en hertzios que establecerá la frecuencia de reproduccion del canal de audio.
' Output:
' Purpose: Establece la frecuencia de reproducción del canal de audio.
' Remarks: Modificando la frecuencia de un sonido se puede alterar su velocidad.
'===============================================================================
Public Sub SOUND_SetFrequency(Channel As Long, Value As Long)
Attribute SOUND_SetFrequency.VB_HelpID = 138
On Local Error Resume Next

If DS_Init Then
    If Value > DSBFREQUENCY_MAX Then Value = DSBFREQUENCY_MAX Else If Value < DSBFREQUENCY_MIN Then Value = DSBFREQUENCY_MIN

    Call m_Channel(Channel).Buffer.SetFrequency(Value)

End If

End Sub

'===============================================================================
' Name: SOUND_GetFrequency
' Input:
'   ByVal Channel As Long - Canal donde se esta reproduciendo el sonido.
' Output:
'   Long - Valor en hertzios que define la frecuencia de reproducción del canal de audio.
' Purpose: Devuelve la frecuencia de reproducción del canal de audio.
' Remarks: Esta función devuelve la frecuencia de reproducción del canal donde se esta reproduciendo el sonido. Para obtener la frecuencia de reproducción del efecto de sonido original utilizar SOUND_GetSamplesPerSecond.
'===============================================================================
Public Function SOUND_GetFrequency(Channel As Long) As Long
Attribute SOUND_GetFrequency.VB_HelpID = 137
On Local Error Resume Next

If DS_Init Then SOUND_GetFrequency = m_Channel(Channel).Buffer.GetFrequency()
    
End Function

'===============================================================================
' Name: SOUND_GetSamplesPerSecond
' Input:
'   ByVal Sound As Long - Identificador del sonido.
' Output:
'   Long - Valor en hertzios que define la frecuencia de reproducción del sonido original.
' Purpose: Devuelve la frecuencia de reproducción del sonido original.
' Remarks: Esta función devuelve la frecuencia de reproducción del sonido original. Para obtener la frecuencia de reproducción del canal donde se esta reproduciendo un sonido utilizar SOUND_GetFrecuency.
'===============================================================================
Public Function SOUND_GetSamplesPerSecond(Sound As Long) As Long
Attribute SOUND_GetSamplesPerSecond.VB_HelpID = 136
On Local Error Resume Next

If DS_Init Then SOUND_GetSamplesPerSecond = m_Sound(Sound).Buffer.GetFrequency()

End Function

'//DIRECT SHOW AUDIO:

'===============================================================================
' Name: MUSIC_Load
' Input:
'   ByVal FileName As String - Nombre y ruta del archivo de música.
' Output:
'   Long - Devuelve el código identificador para la música en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar la música se devuelve -1.
' Purpose: Carga un archivo de audio desde disco como una secuencia musical.
' Remarks: Esta función soporta de forma nativa formatos *.MIDI, *.WAV y *.MP3. Se pueden procesar otros formatos de audio como el *.WMA o el *.OGG si se instalan los codecs necesarios para DirectShow. Se pueden cargar tantas secuencias de música como se deseen teniendo como limite la memoria del sistema.
'===============================================================================
Public Function MUSIC_Load(Filename As String) As Long
Attribute MUSIC_Load.VB_HelpID = 135
On Error GoTo ErrOut
Dim i As Long

If DS_Init And Global_Mod.FileExists(Filename) Then
    
    i = -1
    i = Get_FreeID(1)
    
    With m_Music(i)
        'Setup a filter graph for the file:
        Set .Buffer = New FilgraphManager
        Call .Buffer.RenderFile(Filename)
        Set .Cursor = .Buffer
        .Free = False
        
    End With

    MUSIC_Load = i

End If

Exit Function

ErrOut:
If Not i = -1 Then Call Me.MUSIC_Unload(i)
MUSIC_Load = SND_FILENOTLOAD

End Function

'===============================================================================
' Name: MUSIC_Unload
' Input:
'   ByVal Music As Long - Identificador de la secuencia de música.
' Output:
' Purpose: Descarga una secuencia de música de la memoria.
' Remarks:
'===============================================================================
Public Sub MUSIC_Unload(Music As Long)
Attribute MUSIC_Unload.VB_HelpID = 134
On Error GoTo ErrOut

If DS_Init Then
    If Music = UBound(m_Music) Then
        If Music = 0 Then Erase m_Music Else ReDim Preserve m_Music(UBound(m_Music) - 1) As Music_Pool
    
    Else
        With m_Music(Music)
            Set .Buffer = Nothing
            .Free = True
            
        End With

    End If
    
End If

Exit Sub

ErrOut:

End Sub

'===============================================================================
' Name: MUSIC_Play
' Input:
'   ByVal Music As Long - Identificador de la secuencia de música.
'   Optional ByVal Buffer As Long - Buffer donde se reproducirá la secuencia de musica. Por defecto se usa el buffer primario.
'   Optional ByVal Looping As Boolean - Indica si se reproducirá la secuencia de musica en bucle.
' Output:
'   Boolean - Devuelve verdadero si la operación se realiza con éxito.
' Purpose: Reproduce una secuencia de música.
' Remarks: Las funciones de secuencias de música están diseñadas para dar soporte a formatos avanzados de audio e implementado para dar posibilidad de reproducir una secuencia de música de fondo. Se permite reproducir dos secuencias de música como máximo, permitiendo así poder hacer efectos de transición entre dos secuencias de música (efecto 'Fade') por ejemplo.
'===============================================================================
Public Function MUSIC_Play(Music As Long, Optional Buffer As Sound_Buffer = Primary_Buffer, Optional Looping As Boolean) As Boolean
Attribute MUSIC_Play.VB_HelpID = 133
On Error GoTo ErrOut

If DS_Init Then
    
    With AudioBuffer(Buffer)
        'Setup a filter graph for the file:
        Set .DSControl = m_Music(Music).Buffer
    
        'Setup the basic audio object:
        Set .DSAudio = .DSControl
        .DSAudio.volume = 0
        .DSAudio.Balance = 0
    
        'Setup the media event and position objects:
        Set .DSEvent = .DSControl
        Set .DSPosition = m_Music(Music).Cursor
        If ObjPtr(.DSPosition) Then .DSPosition.Rate = 1#
        .DSPosition.CurrentPosition = 0
        
        .DSControl.Run
        
        .Playing = True
        
    End With
    
    AudioBuffer(Buffer).Looping = Looping
        
    MUSIC_Play = True

End If

Exit Function

ErrOut:
MUSIC_Play = SND_UNKNOWNERROR

End Function

'===============================================================================
' Name: MUSIC_IsPlaying
' Input:
'   ByVal Buffer As Sound_Buffer - Buffer al que se realizara la comprobación.
' Output:
'   Boolean - Devuelve verdadero si el buffer de sonido esta en uso.
' Purpose: Comprueba si el buffer de sonido esta siendo utilizado o no.
' Remarks: Esta función es útil a la hora de determinar si una secuencia de música ha dejado de reproducirse o no.
'===============================================================================
Public Function MUSIC_IsPlaying(Buffer As Sound_Buffer) As Boolean
Attribute MUSIC_IsPlaying.VB_HelpID = 132
On Local Error Resume Next

If DS_Init Then MUSIC_IsPlaying = AudioBuffer(Buffer).Playing

End Function

'===============================================================================
' Name: MUSIC_Stop
' Input:
'   ByVal Buffer As Long - Buffer de sonido donde se esta reproduciendo la secuencia de música.
' Output:
' Purpose: Detiene la reproducción de una secuencia de música y resetea la posicion de lectura.
' Remarks: Esta función resetea la posición de lectura al inicio de la secuencia de música y libera el buffer en uso. Para realizar una pausa del buffer de sonido utilizar MUSIC_Pause.
'===============================================================================
Public Sub MUSIC_Stop(Buffer As Sound_Buffer)
Attribute MUSIC_Stop.VB_HelpID = 131
On Local Error Resume Next

If DS_Init Then
    Call AudioBuffer(Buffer).DSControl.Stop
    AudioBuffer(Buffer).DSPosition.CurrentPosition = 0
    AudioBuffer(Buffer).Playing = False
            
End If

End Sub

'===============================================================================
' Name: MUSIC_Pause
' Input:
'   ByVal Buffer As Long - Buffer de sonido donde se esta reproduciendo la secuencia de música.
' Output:
' Purpose: Detiene la reproducción de una secuencia de música manteniendo la posición de lectura.
' Remarks:
'===============================================================================
Public Sub MUSIC_Pause(Buffer As Sound_Buffer)
Attribute MUSIC_Pause.VB_HelpID = 130
On Local Error Resume Next

If DS_Init Then
    If AudioBuffer(Buffer).Playing Then
        Call AudioBuffer(Buffer).DSControl.Pause
        AudioBuffer(Buffer).Playing = False
    
    Else
        Call AudioBuffer(Buffer).DSControl.Run
        AudioBuffer(Buffer).Playing = True
        
    End If
        
End If

End Sub

'===============================================================================
' Name: MUSIC_SetVolume
' Input:
'   ByVal Buffer As Long - Buffer de sonido donde se esta reproduciendo.
'   ByVal Value As Long - Valor comprendido entre 0 y 100 que establecerá el nivel de volumen del buffer de sonido.
' Output:
' Purpose: Establece el nivel del volumen del buffer de sonido.
' Remarks:
'===============================================================================
Public Sub MUSIC_SetVolume(Buffer As Sound_Buffer, Value As Long)
Attribute MUSIC_SetVolume.VB_HelpID = 129
On Local Error Resume Next

If DS_Init Then
    If Value > 100 Then Value = 100 Else If Value < 0 Then Value = 0

    AudioBuffer(Buffer).DSAudio.volume = (Value * 100) - 10000
    
End If

End Sub

'===============================================================================
' Name: MUSIC_GetVolume
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
' Output:
'   Long - Valor comprendido entre 0 y 100 que define el nivel de volumen del buffer de sonido.
' Purpose: Devuelve el nivel del volumen del buffer de sonido.
' Remarks:
'===============================================================================
Public Function MUSIC_GetVolume(Buffer As Sound_Buffer) As Long
Attribute MUSIC_GetVolume.VB_HelpID = 128
On Local Error Resume Next

If DS_Init Then MUSIC_GetVolume = (AudioBuffer(Buffer).DSAudio.volume / 100) + 100

End Function

'===============================================================================
' Name: MUSIC_SetPan
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
'   ByVal Value As Long - Valor comprendido entre -100 y 100 que establecerá el nivel de balance del buffer de sonido. El rango de valores negativos hasta representa el canal izquierdo y los valores positivos el canal derecho y 0 representa el valor neutro para el balance entre ambos canales.
' Output:
' Purpose: Establece el nivel del balance del buffer de sonido.
' Remarks:
'===============================================================================
Public Sub MUSIC_SetPan(Buffer As Sound_Buffer, Value As Long)
Attribute MUSIC_SetPan.VB_HelpID = 127
On Local Error Resume Next

If DS_Init Then
    If Value > 100 Then Value = 100 Else If Value < -100 Then Value = -100

    AudioBuffer(Buffer).DSAudio.Balance = Value * 100
    
End If

End Sub

'===============================================================================
' Name: MUSIC_GetPan
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
' Output:
'   Long - Valor comprendido entre -100 y 100 que define el nivel de balance del buffer de sonido.
' Purpose: Devuelve el nivel del balance del buffer de sonido.
' Remarks:
'===============================================================================
Public Function MUSIC_GetPan(Buffer As Sound_Buffer) As Long
Attribute MUSIC_GetPan.VB_HelpID = 126
On Local Error Resume Next

If DS_Init Then MUSIC_GetPan = AudioBuffer(Buffer).DSAudio.Balance / 100

End Function

'===============================================================================
' Name: MUSIC_SetRate
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
'   ByVal Value As Long - Valor comprendido entre 0 y 220 que establecerá la velocidad de reproducción del buffer de sonido.
' Output:
' Purpose: Establece la velocidad de reproducción del buffer de sonido.
' Remarks: El valor para la velocidad normal de una secuencia de música es de 100.
'===============================================================================
Public Sub MUSIC_SetRate(Buffer As Sound_Buffer, Value As Long)
Attribute MUSIC_SetRate.VB_HelpID = 125
On Local Error Resume Next

If DS_Init Then
    If Value > 220 Then Value = 220 Else If Value < 0 Then Value = 0

    AudioBuffer(Buffer).DSPosition.Rate = CSng(Value / 100)
    
End If

End Sub

'===============================================================================
' Name: MUSIC_GetRate
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
' Output:
'   Long - Valor comprendido entre 0 y 220 que define la velocidad de reproducción del buffer de sonido.
' Purpose: Devuelve la velocidad de reproducción del buffer de sonido.
' Remarks: Por defecto la velocidad de reproducción se establece a 100.
'===============================================================================
Public Function MUSIC_GetRate(Buffer As Sound_Buffer) As Long
Attribute MUSIC_GetRate.VB_HelpID = 124
On Local Error Resume Next

If DS_Init Then MUSIC_GetRate = CLng(AudioBuffer(Buffer).DSPosition.Rate * 100)

End Function

'===============================================================================
' Name: MUSIC_SetCurrentPosition
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
'   ByVal Value As Long - Valor que define la posición de lectura del buffer de sonido.
' Output:
' Purpose: Establece la posición de lectura del buffer de sonido.
' Remarks:
'===============================================================================
Public Sub MUSIC_SetCurrentPosition(Buffer As Sound_Buffer, Value As Long)
Attribute MUSIC_SetCurrentPosition.VB_HelpID = 123
On Local Error Resume Next

If DS_Init Then
    If Value > AudioBuffer(Buffer).DSPosition.Duration * 1000 Then Value = 0 Else If Value < 0 Then Value = 0

    AudioBuffer(Buffer).DSPosition.CurrentPosition = Value / 1000
    
End If

End Sub

'===============================================================================
' Name: MUSIC_GetCurrentPosition
' Input:
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica.
' Output:
'   Long - Valor que define la posición de lectura del buffer de sonido.
' Purpose: Devuelve la posición de lectura del buffer de sonido.
' Remarks:
'===============================================================================
Public Function MUSIC_GetCurrentPosition(Buffer As Sound_Buffer) As Long
Attribute MUSIC_GetCurrentPosition.VB_HelpID = 122
On Local Error Resume Next

If DS_Init Then MUSIC_GetCurrentPosition = AudioBuffer(Buffer).DSPosition.CurrentPosition * 1000

End Function

'===============================================================================
' Name: MUSIC_GetDuration
' Input:
'   ByVal Music As Long - Identificador de la secuencia de música. Si se pasa como valor un numero negativo se lee la duracion de la secuencia de musica que se esta reproduciendo en el buffer indicado en el siguiente parametro.
'   ByVal Buffer As Long - Buffer donde se esta reproduciendo la secuencia de musica. Este parámetro solo es necesario si el valor del parámetro Music es negativo.
' Output:
'   Long - Valor que define en milisegundos la duración de una secuencia de musica.
' Purpose: Devuelve en milisegundos la duración de una secuencia de música.
' Remarks:
'===============================================================================
Public Function MUSIC_GetDuration(Music As Long, Optional Buffer As Sound_Buffer) As Long
Attribute MUSIC_GetDuration.VB_HelpID = 121
On Local Error Resume Next

If DS_Init Then
    If Music < 0 Then
        MUSIC_GetDuration = AudioBuffer(Buffer).DSPosition.Duration * 1000
        
    Else
        MUSIC_GetDuration = m_Music(Music).Cursor.Duration * 1000
        
    End If

End If

End Function


'//CD-AUDIO:

'Envia los comandos al dispositivo MCI:
Private Sub SendCommand(Command As String, Optional ReturnString As String, Optional ReturnLength As Long)

If ReturnString = vbNullString Then ReturnString = 0
Call mciSendString(Command, ReturnString, ReturnLength, 0)

End Sub

'===============================================================================
' Name: CD_Play
' Input:
'   Optional ByVal Track As Integer - Numero de pista de audio del CD de música. Si no se especifica un valor se reproduce la pista numero 1.
' Output:
' Purpose: Reproduce una pista de audio de un CD de música.
' Remarks:
'===============================================================================
Public Sub CD_Play(Optional Track As Long)
Attribute CD_Play.VB_HelpID = 120
On Local Error Resume Next

If DS_Init Then
    If Not CD_Init Then
        'Inicializamos el CD Audio:
        Call SendCommand("open cdaudio alias cd wait shareable")
        Call SendCommand("set cd time format tmsf wait")
        Call SendCommand("set cd door closed")
        
        'Inicializamos el control de volumen del CD-Audio:
        Call CDSoundVolume_Mod.OpenMixer
        
    End If

    If Track <= 0 Then If Not CD_In_Pause Then Track = 1

    Call SendCommand("play cd from " & Track)

    CD_In_Pause = False
    CD_Init = True
    
End If

End Sub

'===============================================================================
' Name: CD_Stop
' Input:
' Output:
' Purpose: Detiene la reproducción del CD de música.
' Remarks: Esta función detiene la reproducción del CD de música reseteando la posición de lectura.
'===============================================================================
Public Sub CD_Stop()
Attribute CD_Stop.VB_HelpID = 119
On Local Error Resume Next

If DS_Init Then
    If CD_Init Then
        Call SendCommand("stop cd wait")
        Call SendCommand("seek cd to " & CD_GetTrack)
    
    End If
    
End If

End Sub

'===============================================================================
' Name: CD_Pause
' Input:
' Output:
' Purpose: Detiene la reproducción de un CD de música manteniendo la posición de lectura.
' Remarks:
'===============================================================================
Public Sub CD_Pause()
Attribute CD_Pause.VB_HelpID = 118
On Local Error Resume Next

If DS_Init Then
    If CD_Init Then
        If Not CD_In_Pause Then Call SendCommand("pause cd"): CD_In_Pause = True Else If CD_In_Pause Then Call CD_Play

    End If
    
End If
    
End Sub

'===============================================================================
' Name: CD_NextTrack
' Input:
' Output:
' Purpose: Reproduce la pista de audio siguiente a la actual del CD de música.
' Remarks: Si no hubiera mas pistas se pasa directamente a la primera.
'===============================================================================
Public Sub CD_NextTrack()
Attribute CD_NextTrack.VB_HelpID = 117
On Local Error Resume Next
Dim Track As Integer

If DS_Init Then
    If CD_Init Then
        Track = CD_GetTrack + 1
        If Track > CD_GetTracks Then Track = 1

        Call SendCommand("play cd from " & Track)
            
    End If
    
End If
    
End Sub

'===============================================================================
' Name: CD_PrevTrack
' Input:
' Output:
' Purpose: Reproduce la pista de audio anterior a la actual del CD de música.
' Remarks: Si no hubiera mas pistas se pasa directamente a la ultima.
'===============================================================================
Public Sub CD_PrevTrack()
Attribute CD_PrevTrack.VB_HelpID = 116
On Local Error Resume Next
Dim Track As Integer

If DS_Init Then
    If CD_Init Then
        Track = CD_GetTrack - 1
        If Track < 1 Then Track = CD_GetTracks

        Call SendCommand("play cd from " & Track)
    
    End If
    
End If
    
End Sub

'===============================================================================
' Name: CD_GetTrack
' Input:
' Output:
'   Long - Numero de la pista de audio del CD de música que se esta reproduciendo actualmente.
' Purpose: Devuelve el numero de la pista de audio que se esta reproduciendo.
' Remarks:
'===============================================================================
Public Function CD_GetTrack() As Long
Attribute CD_GetTrack.VB_HelpID = 115
On Local Error Resume Next
Dim Track As String * 30

If DS_Init Then
    If CD_Init Then
        Call SendCommand("status cd position", Track, Len(Track))
        CD_GetTrack = CLng(Mid(Track, 1, 2))
    
    End If

End If

End Function

'===============================================================================
' Name: CD_GetTracks
' Input:
' Output:
'   Long - Numero de pistas de audio que contiene el CD de música.
' Purpose: Devuelve el numero de pistas de audio que contiene el CD de música.
' Remarks:
'===============================================================================
Public Function CD_GetTracks() As Long
Attribute CD_GetTracks.VB_HelpID = 114
On Local Error Resume Next
Dim Tracks As String * 30

If DS_Init Then
    If CD_Init Then
        Call SendCommand("status cd number of tracks wait", Tracks, Len(Tracks))
        CD_GetTracks = CLng(Mid(Tracks, 1, 2))
    
    End If
    
End If
    
End Function

'===============================================================================
' Name: CD_GetTrackLength
' Input:
'   Optional ByVal Track As Long - Numero de pista de audio que mediremos.
' Output:
'   Long - Duración en milisegundos de la pista de audio.
' Purpose: Devuelve la duración en milisegundos de una pista de audio del CD de música.
' Remarks:
'===============================================================================
Public Function CD_GetTrackLength(Optional Track As Long) As Long
Attribute CD_GetTrackLength.VB_HelpID = 113
On Local Error Resume Next
Dim Length As String * 30
Dim A As Long, B As Long

If DS_Init Then
    If CD_Init Then
        If Track = 0 Then Track = CD_GetTrack

        Call SendCommand("status cd length track " & Track & " wait", Length, Len(Length))
        A = CLng(Mid(Length, 1, 2)) 'Minutos
        B = CLng(Mid(Length, 4, 2)) 'Segundos

        A = A * 60 'Minutos a Segundos
        A = A + B 'Sumamos los segundos
        A = A * 1000 'Milisegundos

        CD_GetTrackLength = A
    
    End If
    
End If
    
End Function

'===============================================================================
' Name: CD_GetTrackPosition
' Input:
' Output:
'   Integer - Duración en milisegundos de la posición de lectura.
' Purpose: Devuelve la posición en milisegundos de reproducción de la pista de audio actual del CD de música.
' Remarks:
'===============================================================================
Public Function CD_GetTrackPosition() As Long
Attribute CD_GetTrackPosition.VB_HelpID = 112
On Local Error Resume Next
Dim Val As String * 40
Dim A As Long, B As Long

If DS_Init Then
    If CD_Init Then
        Call SendCommand("status cd position wait", Val, Len(Val))
        A = CLng(Mid(Val, 4, 2)) 'Minutos
        B = CLng(Mid(Val, 7, 2)) 'Segundos

        A = A * 60 'Minutos a Segundos
        A = A + B 'Sumamos los segundos
        A = A * 1000 'Milisegundos

        CD_GetTrackPosition = A
    
    End If
    
End If
    
End Function

'===============================================================================
' Name: CD_SetVolume
' Input:
'   ByVal Value As Long - Valor comprendido entre 0 y 100 que define el nivel de volumen del CD de música.
' Output:
' Purpose: Establece el nivel de volumen del CD de música.
' Remarks:
'===============================================================================
Public Sub CD_SetVolume(Value As Long)
Attribute CD_SetVolume.VB_HelpID = 111
On Local Error Resume Next

If DS_Init Then
    If CD_Init Then
        If Value > 100 Then Value = 100 Else If Value < 0 Then Value = 0
        Call CDSoundVolume_Mod.Set_Volumen(CLng((65535 * Value) / 100))
    
    End If
    
End If
    
End Sub

'===============================================================================
' Name: CD_GetVolume
' Input:
' Output:
'   Long - Valor comprendido entre 0 y 100 que define el nivel de volumen del CD de música.
' Purpose: Devuelve el nivel de volumen del CD de música.
' Remarks:
'===============================================================================
Public Function CD_GetVolume() As Long
Attribute CD_GetVolume.VB_HelpID = 110
On Local Error Resume Next

If DS_Init Then If CD_Init Then CD_GetVolume = CLng((CDSoundVolume_Mod.Get_Volumen / 65535) * 100)
    
End Function

'===============================================================================
' Name: CD_Eject
' Input:
' Output:
' Purpose: Abre o cierra la bandeja del lector de CD's.
' Remarks:
'===============================================================================
Public Sub CD_Eject()
Attribute CD_Eject.VB_HelpID = 109
On Local Error Resume Next

If DS_Init Then
    If CD_Init Then
        If Not CD_Open Then Call SendCommand("set cd door open") Else If CD_Open Then Call SendCommand("set cd door closed")
        CD_Open = Not CD_Open

    End If
    
End If

End Sub


'//Captura de sonido:

'===============================================================================
' Name: RECORD_Start
' Input:
' Output:
' Purpose: Inicia la captura de la entrada de audio.
' Remarks: Inicia la captura de la entrada de audio, generalmente el micrófono, y la va recopilando en memoria.
'===============================================================================
Public Sub RECORD_Start()
Attribute RECORD_Start.VB_HelpID = 108
On Local Error Resume Next
    
If DS_Init Then
    
    Call SendCommand("open new type waveaudio alias WavFile")
    Call SendCommand("record WavFile insert")
    
End If

End Sub

'===============================================================================
' Name: RECORD_Finish
' Input:
' Output:
' Purpose: Termina la captura de la entrada de audio.
' Remarks:
'===============================================================================
Public Sub RECORD_Finish()
Attribute RECORD_Finish.VB_HelpID = 107
On Local Error Resume Next
    
If DS_Init Then
        
    Call SendCommand("stop WavFile wait")

End If

End Sub

'===============================================================================
' Name: RECORD_Save
' Input:
'   ByVal FileName As String - Nombre y ruta donde se generara el archivo con el contenido de la captura de audio.
' Output:
' Purpose: Crea un archivo de ondas (WAVE) *.WAV con el contenido de la captura de audio.
' Remarks:
'===============================================================================
Public Sub RECORD_Save(Filename As String)
Attribute RECORD_Save.VB_HelpID = 106
On Local Error Resume Next
    
If DS_Init Then
    Call SendCommand("save WavFile " & Filename)
    Call SendCommand("close WavFile")
    
End If

End Sub


'//Efectos de sonido:

'===============================================================================
' Name: SOUND_FX_SetEffects
' Input:
'   ByVal Channel As Long - Canal donde se activaran o desactivaran los efectos indicados.
'   ByVal Effects() As Long - Array que contiene las cosntantes de los efectos de sonido que se activaran.
' Output:
' Purpose: Activa los efectos de sonido definidos en el array sonido en el canal indicado y desactivara aquellos que estuviesen activos pero no esten en el array.
' Remarks: Si se indica un valor negativo se activan o desactivaran los efectos en todos los canales activos.
'===============================================================================
Public Sub SOUND_FX_SetEffects(Channel As Long, Effects() As Sound_Effects)
    On Local Error Resume Next
    Dim i As Long, stp As Boolean
    Dim DSEffects() As DSEFFECTDESC
    Dim lResults() As Long
    
    ReDim DSEffects(UBound(Effects))
    ReDim lResults(UBound(Effects))
    
    ' Si no se activo ningun efecto desactivamos todos los restantes que pudieran estar activos:
    If Err.Number = 9 Then
        ReDim DSEffects(0)
        ReDim lResults(0)
    Else
        For i = 0 To UBound(Effects)
            Select Case Effects(i)
                Case Sound_Effects.FX_Chorus: DSEffects(i).guidDSFXClass = DSFX_STANDARD_CHORUS
                Case Sound_Effects.FX_Compressor: DSEffects(i).guidDSFXClass = DSFX_STANDARD_COMPRESSOR
                Case Sound_Effects.FX_Distortion: DSEffects(i).guidDSFXClass = DSFX_STANDARD_DISTORTION
                Case Sound_Effects.FX_Echo: DSEffects(i).guidDSFXClass = DSFX_STANDARD_ECHO
                Case Sound_Effects.FX_Flanger: DSEffects(i).guidDSFXClass = DSFX_STANDARD_FLANGER
                Case Sound_Effects.FX_Gargle: DSEffects(i).guidDSFXClass = DSFX_STANDARD_GARGLE
                Case Sound_Effects.FX_Waves_Rerverb: DSEffects(i).guidDSFXClass = DSFX_STANDARD_WAVES_REVERB
            End Select
        Next i
    End If
    
    If DS_Init Then
        If Channel = -1 Then 'Si se especifica -1 se aplica a todos los canales activos:
            For i = 0 To UBound(m_Channel)
                stp = SOUND_IsPlaying(i)
                If (stp) Then Call SOUND_Pause(i)
                Call m_Channel(i).Buffer.SetFX(UBound(DSEffects) + 1, DSEffects, lResults)
                If (stp) Then Call SOUND_Pause(i)
                stp = False
            Next i

        Else
            stp = SOUND_IsPlaying(Channel)
            If (stp) Then Call SOUND_Pause(Channel)
            Call m_Channel(Channel).Buffer.SetFX(UBound(DSEffects) + 1, DSEffects, lResults)
            If (stp) Then Call SOUND_Pause(Channel)
        End If

    End If
End Sub

'===============================================================================
' Name: SOUND_FX_MakeEffectsArray
' Input:
'   ByVal Chorus As Boolean - Añade el efecto coro al array.
'   ByVal Compressor As Boolean - Añade el efecto compresor al array.
'   ByVal Echo As Boolean - Añade el efecto eco al array.
'   ByVal Distortion As Boolean - Añade el efecto distorsion al array.
'   ByVal Flanger As Boolean - Añade el efecto flanger al array.
'   ByVal Gargle As Boolean - Añade el efecto gargle al array.
'   ByVal WavesReverb As Boolean - Añade el efecto ondas de reverberencia al array.
' Output: Devuelve un array con los efectos seleccionados.
' Purpose: Genera un array con los efectos seleccionados.
' Remarks: Esta funcion es una pequeña utilidad para gestionar el array de los efectos de forma rapida y comoda. Tambien se puede definir el array manualmente a traves de los valores de la lista enumerada Sound_Effects.
'===============================================================================
Public Function SOUND_FX_MakeArrayEffects(Chorus As Boolean, Compressor As Boolean, Distortion As Boolean, Echo As Boolean, Flanger As Boolean, Gargle As Boolean, WavesReverb As Boolean) As Sound_Effects()
    Dim tmp() As Sound_Effects, count As Long
    
    If (Chorus) Then count = 1: ReDim tmp(count - 1): tmp(count - 1) = FX_Chorus
    If (Compressor) Then count = count + 1: ReDim Preserve tmp(count - 1): tmp(count - 1) = FX_Compressor
    If (Distortion) Then count = count + 1: ReDim Preserve tmp(count - 1): tmp(count - 1) = FX_Distortion
    If (Echo) Then count = count + 1: ReDim Preserve tmp(count - 1): tmp(count - 1) = FX_Echo
    If (Flanger) Then count = count + 1: ReDim Preserve tmp(count - 1): tmp(count - 1) = FX_Flanger
    If (Gargle) Then count = count + 1: ReDim Preserve tmp(count - 1): tmp(count - 1) = FX_Gargle
    If (WavesReverb) Then count = count + 1: ReDim Preserve tmp(count - 1): tmp(count - 1) = FX_Waves_Rerverb
    
    SOUND_FX_MakeArrayEffects = tmp
End Function

'===============================================================================
' Name: SOUND_FX_DisableAllEffects
' Input:
'   ByVal Channel As Long - Canal donde se desactivaran los efectos.
' Output:
' Purpose: Desactiva todos los efectos de sonido en el canal indicado.
' Remarks: Si se indica un valor negativo se desactivan los efectos en todos los canales activos.
'===============================================================================
Public Sub SOUND_FX_DisableAllEffects(Channel As Long)
    On Local Error Resume Next
'    Dim i As Long, stp As Boolean
'    Dim DSEffects(0) As DSEFFECTDESC
'    Dim lResults(0) As Long
    
'    If Not DS_Init Then
'        If Channel = -1 Then 'Si se especifica -1 se aplica a todos los canales activos:
'            For i = 0 To UBound(m_Channel)
'                stp = SOUND_IsPlaying(i)
'                If (stp) Then Call SOUND_Pause(i)
'                Call m_Channel(i).Buffer.SetFX(1, DSEffects, lResults)
'                If Not (stp) Then Call SOUND_Pause(i)
'                stp = False
'            Next i
'
'        Else
'            stp = SOUND_IsPlaying(Channel)
'            If (stp) Then Call SOUND_Pause(Channel)
'            Call m_Channel(Channel).Buffer.SetFX(1, DSEffects, lResults)
'            If Not (stp) Then Call SOUND_Pause(Channel)
'        End If
'
'    End If
    Call Me.SOUND_FX_SetEffects(Channel, Me.SOUND_FX_MakeArrayEffects(False, False, False, False, False, False, False))
End Sub


'//...

'Devuelve un identificador para el sonido o el modulo musical en memoria:
Private Function Get_FreeID(Flag As Byte) As Long
On Local Error Resume Next

Dim i As Long, j As Long

Select Case Flag
    Case 0 'sonidos:
        i = UBound(m_Sound)
        
        If Err.Number = 9 Then
            ReDim m_Sound(0) As Sound_Pool
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If m_Sound(j).Free Then
                    Get_FreeID = j
                    Exit For
                End If
            Next j
            
            ReDim Preserve m_Sound(i + 1) As Sound_Pool
            
            Get_FreeID = UBound(m_Sound)
            
        End If
        
    Case 1 'Musicas:
        i = UBound(m_Music)
        
        If Err.Number = 9 Then
            ReDim m_Music(0) As Music_Pool
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If m_Music(j).Free Then
                    Get_FreeID = j
                    Exit For
                End If
            Next j
            
            ReDim Preserve m_Music(i + 1) As Music_Pool
            
            Get_FreeID = UBound(m_Music)
            
        End If
    
    Case 2 'Canales:
        i = 0

        For i = 0 To UBound(m_Channel)
            'Si el canal esta en uso se busca otro que este libre:
            If (Not Me.SOUND_IsPlaying(i)) Then
                Get_FreeID = i
                Exit Function
                
            End If
            
        Next
        
        Get_FreeID = -1  'Si recorremos la lista entera y no encontramos ningun canal vacio salimos del bucle devolviendo -1.
        
End Select

End Function

Private Sub SetAllSound(Acction As Byte)
Dim i As Long, j As Long

For i = 0 To UBound(m_Channel)
    j = m_Channel(i).Buffer.GetStatus

    If j = DSBSTATUS_PLAYING Or j = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING Then
            
        Call m_Channel(i).Buffer.Stop
        If Acction = 0 Then Call m_Channel(i).Buffer.SetCurrentPosition(0)
            
    Else
        'Si el sonido ya estaba en pausa se vuelve a poner en marcha:
        If Acction = 1 Then Call m_Channel(i).Buffer.Play(Abs(m_Channel(i).Looping))
        
    End If
        
    DoEvents
    
Next i

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
Attribute hWnd.VB_HelpID = 103
hWnd = m_hWnd

End Property

'===============================================================================
' Name: SOUND_Count
' Input:
' Output:
'   Long - Numero de efectos de sonido en memoria.
' Purpose: Devuelve el numero de efectos de sonido en memoria.
' Remarks: Esta función no devuelve el numero de efectos de sonido en canales en reproducción si no los sonidos almacenados en memoria mediante SOUND_Load.
'===============================================================================
Public Property Get SOUND_Count() As Long
Attribute SOUND_Count.VB_HelpID = 102
On Local Error Resume Next
SOUND_Count = UBound(m_Sound) + 1

End Property

'===============================================================================
' Name: MUSIC_Count
' Input:
' Output:
'   Long - Numero de secuencias de música en memoria.
' Purpose: Devuelve el numero de secuencias de música en memoria.
' Remarks:
'===============================================================================
Public Property Get MUSIC_Count() As Long
Attribute MUSIC_Count.VB_HelpID = 101
On Local Error Resume Next
MUSIC_Count = UBound(m_Music) + 1

End Property

