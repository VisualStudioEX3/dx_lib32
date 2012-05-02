VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dx_GFX_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Name: dx_GFX
' Purpose: Clase de funciones graficas
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 07/08/2001
' Modified: 12/05/2010
'===============================================================================

Option Explicit

Private Const z_Levels = 8

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Enumeraciones:
'===============================================================================
' Name: GFX_ErrorCodes
'    GFX_OK - La operación se ha realizado con éxito.
'    GFX_DEVICELOST - Se ha perdido el control del dispositivo grafico.
'    GFX_DEVICENOTRESET - No es posible recuperar el control de dispositivo grafico.
'    GFX_FILENOTLOAD - Error al cargar el archivo.
'    GFX_DATANOTREAD - Error al leer los datos de la memoria.
'    GFX_NOTCREATED - No se pudo crear el objeto o recurso.
'    GFX_UNKNOWNERROR - Error desconocido.
'    GFX_UNKNOWNERROR - Error desconocido.
' Purpose: Codigos de error.
' Remarks: Definen errores o resultados de algunas funciones de esta clase.
'===============================================================================
Public Enum GFX_ErrorCodes
    GFX_OK = 0
    GFX_DEVICELOST = (vbObjectError + 1000)
    GFX_DEVICENOTRESET = (vbObjectError + 1001)
    GFX_FILENOTLOAD = (vbObjectError + 1002)
    GFX_DATANOTREAD = (vbObjectError + 1003)
    GFX_NOTCREATED = (vbObjectError + 1004)
    GFX_UNKNOWNERROR = (vbObjectError + 1999)

End Enum

'===============================================================================
' Name: Text_Align
'   Align_Left - Texto justificado a la izquierda (por defecto)
'   Align_Right - Texto justificado a la derecha.
'   Align_Center - Texto centrado.
'   Align_Center - Texto centrado.
' Purpose: Alineamiento del texto.
' Remarks:
'===============================================================================
Public Enum Text_Align              'Alineamiento del texto:
    Align_Left = DT_LEFT
    Align_Right = DT_RIGHT
    Align_Center = DT_CENTER

End Enum
    
'===============================================================================
' Name: Blit_Filter
'   Filter_None - No se aplica filtro alguno (por defecto)
'   Filter_Bilinear - Filtro bilinear.
'   Filter_Trilinear - Filtro trilinear.
' Purpose: Filtros de suavizado.
' Remarks:
'===============================================================================
Public Enum Blit_Filter             'Filtros de suavizado:
    Filter_None = 0                 'No se aplica filtro.
    Filter_Bilinear = 1             'Filtro Bilinear.
    Filter_Trilinear = 2            'Filtro Trilinear.
End Enum
    
'===============================================================================
' Name: Blit_Mirror
'   Mirror_None - No se aplica espejado alguno (por defecto)
'   Mirror_Horizontal - Espejado sobre el eje X.
'   Mirror_Vertical - Espejado sobre el eje Y.
'   Mirror_Both - Espejado en ambos ejes.
'   Mirror_Both - Espejado en ambos ejes.
' Purpose: Modos de espejado.
' Remarks:
'===============================================================================
Public Enum Blit_Mirror             'Modos de espejado:
    Mirror_None = 0                 'No se aplica espejado.
    Mirror_Horizontal = 1           'Espejado sobre eje X.
    Mirror_Vertical = 2             'Espejado sobre eje Y.
    Mirror_both = 3                 'Espejado sobre ambos ejes.
        
End Enum

'===============================================================================
' Name: Blit_Alpha
'   Blendop_Color - Canal Alpha de los colores.
'   Blendop_Aditive - Opacidad Aditiva.
'   Blendop_Sustrative - Opacidad sustractiva.
'   Blendop_Sustrative - Opacidad sustractiva.
' Purpose: Modos de opacidad o transparencia.
' Remarks:
'===============================================================================
Public Enum Blit_Alpha              'Modos de opacidades:
    Blendop_Color = 0               'Color a traves de alpha.
    Blendop_Aditive = 1             'Aditiva.
    Blendop_Sustrative = 2          'Sustrativa.
    Blendop_Inverse = 3             'Invertir colores.
    Blendop_XOR = 4                 'Realiza una exclusion logica entre el color de origen del mapa y el color destino del backbuffer.
    Blendop_Crystaline = 5          'Intercambia el color de destino con el origen generando un efecto de transparencia cristalina. No permite modular la opacidad mediante el componente Alpha del color.
    Blendop_GreyScale = 256         'Convierte los colores a escala de grises. No permite modular la opacidad mediante el componente Alpha del color y no se aplica transparencia del canal Alpha de la textura.
End Enum

'===============================================================================
' Name: Blit_Perspective
'    Caballera_Width             Modifica las coordenadas anchura.
'    Caballera_Height            Modifica las coordenadas de altura.
'    Caballera_Width_Negative    Modifica las coordenadas anchura con valores negativos.
'    Caballera_Height_Negative   Modifica las coordenadas de altura con valores negativos.
'    Isometric_Base              Modifica ambas coordenadas para perspectiva isométrica.
'    Isometric_Height            Modifica las coordenadas de altura para la vista isométrica.
'    Isometric_Height_Negative   Modifica las coordenadas de altura para la vista isométrica con valores negativos.
'    Isometric_Height_Negative   Modifica las coordenadas de altura para la vista isométrica con valores negativos.
' Purpose: Operaciones para dibujo de perspectiva.
' Remarks:
'===============================================================================
Public Enum Blit_Perspective       'Operaciones para dibujo de perspectiva:
    Caballera_Width = 0             'Modifica las coodenadas anchura.
    Caballera_Height = 1            'Modifica las coordenadas de altura.
    Caballera_Width_Negative = 2    'Modifica las coodenadas anchura con valores negativos.
    Caballera_Height_Negative = 3   'Modifica las coordenadas de altura con valores negativos.
    Isometric_Base = 4              'Modifica ambas coordenadas para perspectiva isometrica.
    Isometric_Height = 5            'Modifica las coordenadas de altura para la vista isometrica.
    Isometric_Height_Negative = 6   'Modifica las coordenadas de altura para la vista isometrica con valores negativos.

End Enum

'===============================================================================
' Name: Vertex
'   ByRef X As Long - Coordenada X. Define la posición en el eje horizontal.
'   ByRef Y As Long - Coordenada Y. Define la posición en el eje vertical.
'   ByRef Z As Long - Color ARGB que define el color del vértice.
'   ByRef Color As Long - Coordenada Z. Define la profundidad de dibujo.
' Purpose: Define los valores de un vértice 2D.
' Remarks: Algunas funciones de la clase grafica necesitan definir valores de los vertices para el dibujo de ciertos gráficos.
'===============================================================================
Public Type Vertex
    X As Long
    Y As Long
    Z As Long
    Color As Long
    
End Type
    
'===============================================================================
' Name: GFX_Rect
'   ByRef X As Long - Coordenada X. Define la posición en el eje horizontal de la esquina superior izquierda.
'   ByRef Y As Long - Coordenada Y. Define la posición en el eje vertical de la esquina superior izquierda.
'   ByRef Width As Long - Define la anchura del rectángulo.
'   ByRef Height As Long - Define la altura del rectángulo.
' Purpose: Define los valores de un rectángulo.
' Remarks: Esta estructura se utiliza para delimitar zonas en algunas operaciones de dibujo con gráficos.
'===============================================================================
Public Type GFX_Rect
    X As Long
    Y As Long
    Width As Long
    Height As Long
    
End Type
    
'===============================================================================
' Name: ARGB
'   ByRef Alpha As Integer - Define el valor para el canal alfa.
'   ByRef Red As Integer - Define el valor para el canal rojo.
'   ByRef Green As Integer - Define el valor para el canal verde.
'   ByRef Blue As Integer - Define el valor para el canal azul.
' Purpose: Define los valores de un color ARGB.
' Remarks: Esta estructura se utiliza para leer e interpretar por separado los componentes de un color ARGB.
'===============================================================================
Public Type ARGB
    Alpha As Integer
    Red As Integer
    Green As Integer
    Blue As Integer
    
End Type

Private Type D3D_Texture
    Texture As Direct3DTexture8
    Image_Width As Long
    Image_Height As Long
    Texture_Width As Long
    Texture_Height As Long
    ColorKey As Long
    Size As Long
    RenderTarget As Boolean

End Type

Private Type D3D_Surface
    Surface As Direct3DSurface8
    Width As Long
    Height As Long
    Size As Long
    
End Type

Private Type Blit_Structure
    BlitType As Long
    PrimitiveCount As Long
    PrimitiveType As Long
    Quad() As D3DTLVERTEX
    Texture As Direct3DTexture8
    AlphaMode As Long
    Filter As Long
    Text As String
    RectRegion As RECT
    FontDesc As D3DXFont
    FontColor As Long
    TextAlign As Text_Align
    ShaderSetup As Boolean          ' Indica que se va a aplicar un shader al nivel de render.
    ShaderProgram As String         ' Codigo del programa PixelShader.
    ShaderConstants() As Single     ' Lista de constantes (opcional).

End Type

Private Type Render_Buffer
    Z() As Blit_Structure
    
End Type

'===============================================================================
' Name: Display_Mode
'   ByRef Height As Long - Define la anchura en píxeles.
'   ByRef Width As Long - Define la altura en píxeles.
'   ByRef Bpp As Long - Define la profundidad de color en bits.
'   ByRef RefreshRate As Long - Define la frecuencia de refresco del monitor.
' Purpose: Define los valores de un modo de video.
' Remarks: Esta estructura se utiliza para leer e interpretar los valores de un modo de video.<br>
'El campo <b>Bpp</b> siempre devolverá o 16 bits o 32 bits.<br>
'El Campo <b>Refresh_Rate</b> es orientativo ya que no se usa a la hora de establecer un modo de video.
'===============================================================================
Public Type Display_Mode
    Height As Long
    Width As Long
    Bpp As Long
    RefreshRate As Long
    
End Type

'===============================================================================
' Name: Device_Info
'   ByRef AdapterName As String - Nombre del dispositivo.
'   ByRef DriverName As String - Nombre del controlador de video.
'   ByRef DriverVersionHigh As Long - Numero de versión del controlador de video.
'   ByRef DriverVersionLow As Long - Numero de subversión del controlador de video.
'   ByRef Revision As Long - Numero de revisión del controlador de video. Corresponde al 3º grupo de números de la versión. E.j.: Revisión 918 de la version 2.0 = 2.0.918
'   ByRef VendorID As Long - Código del proveedor.
'   ByRef MaxTextureHeight As Long - Altura máxima de una textura en memoria para este dispositivo.
'   ByRef MaxTextureWidth As Long - Anchura máxima de una textura en memoria para este dispositivo.
'   ByRef MaxBitsPerPixel As Long - Profundidad máxima de color (o bits por pixel) para este dispositivo.
' Purpose: Define los valores de un dispositivo de video.
' Remarks: Esta estructura se utiliza para leer e interpretar los valores de un dispositivo de video.
'===============================================================================
Public Type Device_Info
    AdapterName As String 'Identificador del adaptador.
    DriverName As String 'Driver utilizado.
    DriverVersionHigh As Long
    DriverVersionLow As Long
    Revision As Long
    VendorID As Long
    MaxTextureHeight As Long 'Altura maxima de una textura.
    MaxTextureWidth As Long 'Anchura maxima de una textura.
    MaxBitsPerPixel As Long 'Maximo de Bits de color soportados.
    
End Type

'===============================================================================
' Name: GFX_Info
'   ByRef Height As Long - Altura en píxeles del grafico en memoria.
'   ByRef Width As Long - Anchura en píxeles del grafico en memoria.
'   ByRef Size As Long - Tamaño en bytes del grafico en memoria.
'   ByRef ColorKey As Long - Valor que define el color transparente para este grafico.
'   ByRef Image_Height As Long - Altura en píxeles de la imagen original.
'   ByRef Image_Width As Long - Anchura en píxeles de la imagen original.
' Purpose: Define los valores de un grafico cargado en memoria.
' Remarks: Esta estructura se utiliza para leer e interpretar los valores de un grafico en memoria.
'Los valores de altura y anchura del grafico pueden diferir de los originales ya que la imagen se reescala en memoria a unas dimensiones cuyos valores sean potencia de 2. E.j.: 125x187 = 128x256...
'===============================================================================
Public Type GFX_Info
    Height As Long 'Altura del grafico.
    Width As Long 'Anchura del grafico.
    Size As Long 'Tamaño en bytes.
    ColorKey As Long 'Color de transparencia.
    Image_Height As Long 'Altura original de la imagen.
    Image_Width As Long 'Anchura original de la imagen.
    
End Type

'Tipo Circulos precalculados:
Private Type PreCalCircle_Data
    'SubArrays dinamicos.
    Circle_Border() As D3DTLVERTEX
    Circle_Fill() As D3DTLVERTEX
    NumVertex As Long
    MatchColor As Boolean
    Fill As Boolean
    Free As Boolean
    
End Type

Private Type PreCal_SpriteTransform
    Map As Long
    Vector As Vertex
    Angle As Single
    Perspective As Blit_Perspective
    Center As Boolean
    Factor As Long
End Type


'Tipo Pixel Array:
Private Type GFX_PixelData
    Surf As Long 'Superficie de origen.
    Height As Long 'Altura del grafico.
    Pitch As Long 'Anchura del grafico.
    BGRA() As Long 'Array dinamico de pixeles en formato BGRA.
    Free As Boolean

End Type

'Tipo Fuente de Sistema:
Private Type System_Font
    D3D_Font As D3DXFont
    Std_Font As New StdFont
    I_Font As IFont
    Free As Boolean

End Type

Private Type GFX_RenderTarget
    Target As D3DXRenderToSurface 'Render Target.
    ViewPort As D3DVIEWPORT8 'Viewport que usara el render target sobre la superifice.
    Surface As Direct3DSurface8 'Superficie que apunta a la memoria de la textura.
    Buffer(-z_Levels To z_Levels) As Render_Buffer
    
End Type

Dim Dx As DirectX8
Dim D3D As Direct3D8
Dim d3dDevice As Direct3DDevice8
Dim bRunning As Boolean

Dim D3DX As D3DX8
Dim D3DWindow As D3DPRESENT_PARAMETERS

Dim Texture() As D3D_Texture 'Array de texturas.
Dim Surface() As D3D_Surface 'Array de superficies.
Dim D3D_ImgInfo As D3DXIMAGE_INFO
Dim D3D_Desc As D3DSURFACE_DESC

'Listas de vertices para precalculos de circunferencias:
Dim PreCircle() As PreCalCircle_Data

'Lista de array de pixeles:
Dim Pixel_Array() As GFX_PixelData

'Lista de Fuentes de sistema:
Dim D3DFont() As System_Font
Dim ImportFont As Collection 'Lista de fuentes importadas desde archivo.

'Lista de Render Targets:
Dim RenderTarget() As GFX_RenderTarget

Dim TextRegion As RECT

Dim FrameRect As RECT
Dim FrameSet As Boolean

'Parametros nulos:
Dim NullQuad(0) As D3DTLVERTEX

Dim LastTimeCheckFPS As Long
Dim FramesDrawn As Long
Dim FrameRate As Long

Dim MainBuffer(-z_Levels To z_Levels) As Render_Buffer '16 niveles de profundidad para renderizar: -8 lo mas cercano a la camara, 8 lo mas lejano.
'Posiblemente se haga dinamico el array y se puedan especificar los niveles mediante un parametro en la funcion de inicio.

Dim WndStyle As Long

Dim D3D_Init As Boolean
Dim D3D_Screen As Display_Mode
Dim D3D_Antialiasing As Boolean
Dim D3D_TripleBuffer As Boolean
Dim D3D_VSync As Boolean
Dim D3D_AlphaBlending As Long
Dim D3D_Filter As Long

Dim TargetID As Long
Dim SetSpecular As Boolean, Specular(3) As Long 'Almacena los valores para el canal specular que se usara en la siguiente rutina llamada.
Dim NewCenterX As Long, NewCenterY As Long, SetCenter As Boolean

Dim SetVertexColor As Boolean, VertexColor(3) As Long ' Color de cada vertice que se aplicara en la proxima llamada a DRAW_Map(), DRAW_MapEx() y DRAW_AdvMap().
Dim ReadVertex As Boolean
Dim PrecalVertex(3) As Vertex ' Almacena los ultimos vertices transformados en las operaciones de dibujo.
Dim IsPrecalVertex As Boolean ' Indica se han precalculado los vertices de un sprite.
Dim PrecalSprite As Blit_Structure ' Almacena los parametros del ultimo sprite precalculado.

Dim FrameWait As Long

Dim GammaFactor As Single

Dim TotalVideoMem As Long

'===============================================================================
' Name: Init
' Input:
'   ByVal hWnd As Long - Identificador de la ventana.
'   ByVal Width As Long - Ancho en píxeles del modo de video.
'   ByVal Height As Long - Alto en píxeles del modo de video.
'   ByVal Bpp As Long - Profundidad de color del modo de video.
'   ByVal Windowed As Boolean - Indica si se inicializa en modo ventana.
'   ByVal TripleBuffer As Boolean - Indica si se inicializa con Triple Buffer activado.
'   ByVal vSync As Boolean - Indica si se inicializa con la espera de refresco vertical activada.
'   ByVal RefreshRate As Long - Refresco del modo de video en herzios. Si se pasa como valor 0 se usa el valor por defecto del sistema.
' Output:
'    Boolean - Devuelve Verdadero en caso de inicializarse correctamente.
' Purpose: Inicializa el modo de video.
' Remarks: Inicializa el modo de video y los subsistemas de la clase habilitando las rutinas y la memoria para trabajar con gráficos.
'===============================================================================
Public Function Init(hWnd As Long, Width As Long, Height As Long, Optional Bpp As Long = 16, Optional Windowed As Boolean, Optional TripleBuffer As Boolean, Optional VSync As Boolean, Optional RefreshRate As Long = 60) As Boolean
On Error GoTo ErrOut

Dim DispMode As D3DDISPLAYMODE, lstyle As Long

If Not D3D_Init Then
    
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D_TripleBuffer = TripleBuffer
    D3D_VSync = VSync
    
    If Bpp < 16 Then Bpp = 16
    If RefreshRate < 60 Then RefreshRate = 60
    
    'Guardamos los parametros actuales del estilo de la ventana:
    WndStyle = Global_Mod.GetWindowLong(hWnd, GWL_STYLE)

    'Leemos los valores actuales del modo de pantalla para almacenar informacion:
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    If Windowed Then
        D3DWindow.Windowed = 1
        If VSync Then
            D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        
        Else
            D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
            
        End If
        
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.BackBufferHeight = Height
        D3DWindow.BackBufferWidth = Width
        D3DWindow.hDeviceWindow = hWnd
    
        Global_Mod.D3D_FullScreen = False
    
    Else
        Select Case Bpp
            Case 16, Is < 16: DispMode.Format = D3DFMT_R5G6B5
            Case 32, Is > 32: DispMode.Format = D3DFMT_X8R8G8B8
        
        End Select
    
        DispMode.Width = Width
        DispMode.Height = Height

        If VSync Then
            D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
        
        Else
            D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
            D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
        End If
    
        D3DWindow.BackBufferCount = 2 + Abs(TripleBuffer)
        D3DWindow.BackBufferFormat = DispMode.Format
        D3DWindow.BackBufferHeight = DispMode.Height
        D3DWindow.BackBufferWidth = DispMode.Width
        'D3DWindow.FullScreen_RefreshRateInHz = RefreshRate
        D3DWindow.hDeviceWindow = hWnd
    
        'Quitamos cualquier estilo a la ventana para conseguir una ventana rectangular:
        lstyle = Global_Mod.GetWindowLong(Me.hWnd, GWL_STYLE)
        lstyle = lstyle And WGAME_STYLE
        Call Global_Mod.SetWindowLong(hWnd, GWL_STYLE, lstyle)
        
        Global_Mod.D3D_FullScreen = True
    
    End If
       
    Call Global_Mod.SystemParametersInfo(Global_Mod.SPI_SETSCREENSAVEACTIVE, False, CStr(1), 0)
    Call Global_Mod.Set_SizeWindow(hWnd, Height, Width, Windowed)
        
    Set d3dDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    Call Global_Mod.SetForegroundWindow(hWnd)

    Call SetRenderStates
    
    With D3D_Screen
        .Bpp = Bpp
        .Width = Width
        .Height = Height
        .RefreshRate = RefreshRate

    End With
    
    TargetID = -1
    GammaFactor = 1!
    
    Set ImportFont = New Collection
    
    Init = True
    D3D_Init = Init
    
    Exit Function

End If

ErrOut:

End Function

'===============================================================================
' Name: Frame
' Input:
'   Optional ByVal Color As Long - Valor ARGB que representa el color que se usara en el borrado de pantalla una vez finalizado el proceso de volcado de gráficos en pantalla. Si no se especifica valor alguno se toma el valor ARGB 0,0,0,0.
'   Optional ByVal MaxFrames As Integer - Valor que representa el numero máximo de cuadros por segundo que alcanzara el programa. Si el valor es 0 el programa ejecutara tantos cuadros por segundo como pueda. Un valor aconsejable para lograr fluidez en pantalla suele ser 60 cuadros por segundo.
'   Optional ByVal Clear As Boolean - Valor que indica si se limpia la pantalla antes de pintar.
' Output:
'    Long - Devuelve GFX_OK si la función ha finalizado correctamente, GFX_DEVICELOST si el control del dispositivo grafico se ha perdido, GFX_DEVICENOTRESET si el dispositivo grafico no se puede recuperar y GFX_UNKNOWNERROR si se produce cualquier otro error.
' Purpose: Muestra en pantalla el contenido del backbuffer.
' Remarks: En caso de que la ventana donde se esta dibujando pierda el foco en modo a pantalla completa, ya sea por hacer una combinación de teclas como Alt-Tab por ejemplo, dx_lib32 perderá el control de dispositivo grafico. En cuanto esto ocurre dx_lib32 trata de recuperar el control del dispositivo aunque a veces puede suceder que esto no ocurra. En cuyo caso se debería reinicializar la clase dx_GFX.
'===============================================================================

Public Function Frame(Optional Color As Long, Optional MaxFrames As Integer, Optional Clear As Boolean = True) As Long
On Local Error Resume Next

Dim hr As Long, i As Long, j As Long
Dim Ticks As Long
Static Wait As Long

If D3D_Init And (Not CBool(IsIconic(Me.hWnd))) Then
    Ticks = Global_Mod.GetTickCount()
    
    'Call TestCooperativeLevel to see what state the device is in.
    hr = d3dDevice.TestCooperativeLevel
    
    If hr <> D3D_OK Then
        Erase MainBuffer
        If hr = D3DERR_DEVICENOTRESET Then ' Device needs to be reset
            Call ResetDevice(D3DWindow)
        Else ' If the device is lost, exit and wait for it to come back.
            Frame = GFX_DEVICELOST
            DoEvents
            Exit Function
        End If
    End If
    
    ' Si esta activado el limpiado de pantalla se limpia con el valor establecido en la variable Color:
    If (Clear) Then Call d3dDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0)
    
    'Ejecutamos las llamadas sobre el backbuffer:
    Call ExecuteRenderBuffer(MainBuffer)
    
    'Procesamos todas las llamadas:
    Call d3dDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    
    'Control de cuadros por segundo:
    If MaxFrames >= 24 Then Call WaitFrame(MaxFrames)
        
    'Calculo de cuadros por segundo:
    If Global_Mod.GetTickCount - LastTimeCheckFPS >= 1000 Then
        LastTimeCheckFPS = GetTickCount
        FrameRate = FramesDrawn
        FramesDrawn = 0
    End If

    FramesDrawn = FramesDrawn + 1
    
    If Not Err.Number <> 0 Or Not Err.Number = D3DERR_INVALIDCALL Then 'Si no ocurrio ningun error salimos de la funcion:
        Frame = GFX_OK
    
    ElseIf Not Err.Number = D3DERR_DEVICELOST And Not Err.Number = D3DERR_DEVICENOTRESET Then
        Frame = GFX_UNKNOWNERROR
        bRunning = False
    
    End If

End If

DoEvents

End Function

'Espera hasta el siguiente frame:
Private Sub WaitFrame(MaxFrames As Integer)
    Static Old_Ticks As Currency
    Dim ts As Currency, frq As Currency
    
    Call Global_Mod.QueryPerformanceFrequency(frq)
    Call Global_Mod.QueryPerformanceCounter(ts)
    
    Do While ((ts - Old_Ticks) / frq) * 1000 < (1# / MaxFrames * 1000)
        Call Global_Mod.QueryPerformanceCounter(ts)
        'DoEvents
        
    Loop
    
    Call Global_Mod.QueryPerformanceCounter(Old_Ticks)

End Sub

'===============================================================================
' Name: Terminate
' Purpose: Cierra el modo de video y los subsistemas de la clase.
' Remarks: Este método se encarga de liberar la memoria ocupada por los recursos de la clase y cerrar las referencias de la misma. Es aconsejable llamar a este método justo antes de terminar el programa para así asegurarnos que cualquier proceso de la clase se termina correctamente sin permanecer en memoria.
'===============================================================================
Public Sub Terminate()
On Local Error Resume Next

Dim lstyle As Long
Dim fRet As Long
Dim i As Long

If D3D_Init Then

    Erase Texture, Surface
    
    'Descargamos de la memoria cualquier fuente de texto que se haya cargado desde archivo:
    Dim cfont As CFontPreview
    For Each cfont In ImportFont
        Set cfont = Nothing
    Next
    Set ImportFont = Nothing
    
    Set d3dDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing

    'Devolvemos cualquier estilo a la ventana:
    If Not Windowed Then
        Call Global_Mod.SetWindowLong(hWnd, GWL_STYLE, WndStyle)
        Call SetWindowRgn(hWnd, GetWindowRgn(hWnd, 0), False)

    End If
    
    Call Global_Mod.SystemParametersInfo(Global_Mod.SPI_SETSCREENSAVEACTIVE, True, CStr(1), 0)

    D3D_Init = False

    Global_Mod.D3D_FullScreen = False

End If

End Sub

Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tu As Single, tv As Single) As D3DTLVERTEX
    CreateTLVertex.sx = X
    CreateTLVertex.sy = Y
    CreateTLVertex.sz = Z
    CreateTLVertex.rhw = 1! 'rhw
    CreateTLVertex.Color = Color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

'===============================================================================
' Name: DRAW_Pixel
' Input:
'   ByVal X As Long - Coordenada horizontal del píxel.
'   ByVal Y As Long - Coordenada vertical del píxel.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo.
' Output:
' Purpose: Dibuja un punto en la pantalla.
' Remarks: <Things to take care of>
'===============================================================================
Public Sub DRAW_Pixel(X As Long, Y As Long, Z As Long, Color As Long)
On Error GoTo ErrOut

Dim point(0) As D3DTLVERTEX

    point(0) = CreateTLVertex(CSng(X), CSng(Y), 0, 1, Color, 0, 0, 0)

    If InScreen(point()) Then Call AddCallToRenderBuffer(0, Z, 1, D3DPT_POINTLIST, point(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)
    
Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: DRAW_Pixels
' Input:
'   ByVal List As Vertex - Lista de coordenadas y colores de los pixeles.
' Output:
' Purpose: Dibuja una lista de puntos en la pantalla.
' Remarks: Todos los pixeles se dibujaran en la profundidad definida por la coordenada Z del primer elemento de la lista.
'===============================================================================
Public Sub DRAW_Pixels(list() As Vertex)
On Error GoTo ErrOut

    Dim point() As D3DTLVERTEX, i As Long

    ReDim point(UBound(list))
    For i = 0 To UBound(list)
        point(i) = CreateTLVertex(CSng(list(i).X), CSng(list(i).Y), 0, 1, list(i).Color, 0, 0, 0)
    Next

    If InScreen(point()) Then Call AddCallToRenderBuffer(0, list(0).Z, UBound(list) + 1, D3DPT_POINTLIST, point(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)

Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: DRAW_Line
' Input:
'   ByVal X1 As Long - Coordenada horizontal de origen de la línea.
'   ByVal Y1 As Long - Coordenada vertical de origen de la línea.
'   ByVal X2 As Long - Coordenada horizontal de destino de la línea.
'   ByVal Y2 As Long - Coordenada vertical de destino de la línea.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo.
' Output:
' Purpose: Dibuja una línea en pantalla.
' Remarks: <Things to take care of>
'===============================================================================
Public Sub DRAW_Line(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Z As Long, Color As Long)
On Error GoTo ErrOut

Dim point(1) As D3DTLVERTEX

    'Generamos el poligono:
    point(0) = CreateTLVertex(CSng(X1), CSng(Y1), 0, 1, Color, Specular(0), 0, 0)
    point(1) = CreateTLVertex(CSng(X2), CSng(Y2), 0, 1, Color, Specular(1), 0, 0)

    If InScreen(point()) Then Call AddCallToRenderBuffer(0, Z, 1, D3DPT_LINELIST, point(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)
    
    Erase Specular

Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: DRAW_Box
' Input:
'   ByVal X1 As Long - Coordenada horizontal de la esquina superior izquierda de la caja.
'   ByVal Y1 As Long - Coordenada vertical de la esquina superior izquierda de la caja.
'   ByVal X2 As Long - Coordenada horizontal de la esquina inferior derecha de la caja.
'   ByVal Y2 As Long - Coordenada vertical de la esquina inferior derecha de la caja.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo del borde de la figura.
'   Optional ByVal Fill As Boolean - Indica si se dibujara la figura con relleno.
'   Optional ByVal FillColor As Long - Color ARGB que se aplicara al relleno de la figura.
' Output:
' Purpose: Dibuja una caja con o sin relleno en pantalla.
' Remarks: <Things to take care of>
'===============================================================================
Public Sub DRAW_Box(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Z As Long, Color As Long, Optional Fill As Boolean, Optional FillColor As Long)
On Error GoTo ErrOut

Dim Quad(4) As D3DTLVERTEX
Dim MatchColor As Boolean

    If Color = FillColor Then MatchColor = True
    
    If Fill Then
        'Generamos el relleno:
        Quad(0) = CreateTLVertex(CSng(X1), CSng(Y1), 0, 1, FillColor, Specular(0), 0, 0)
        Quad(1) = CreateTLVertex(CSng(X2), CSng(Y1), 0, 1, FillColor, Specular(1), 0, 0)
        Quad(2) = CreateTLVertex(CSng(X1), CSng(Y2), 0, 1, FillColor, Specular(2), 0, 0)
        Quad(3) = CreateTLVertex(CSng(X2), CSng(Y2), 0, 1, FillColor, Specular(3), 0, 0)
    
        If InScreen(Quad()) Then Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)
    
    End If
    
    'Dibujamos el borde:
    If Not MatchColor Then
        Quad(0) = CreateTLVertex(CSng(X1), CSng(Y1), 0, 1, Color, Specular(0), 0, 0)
        Quad(1) = CreateTLVertex(CSng(X2), CSng(Y1), 0, 1, Color, Specular(1), 0, 0)
        Quad(2) = CreateTLVertex(CSng(X2), CSng(Y2), 0, 1, Color, Specular(2), 0, 0)
        Quad(3) = CreateTLVertex(CSng(X1), CSng(Y2), 0, 1, Color, Specular(3), 0, 0)
        Quad(4) = CreateTLVertex(CSng(X1), CSng(Y1), 0, 1, Color, Specular(0), 0, 0)

        If InScreen(Quad()) Then Call AddCallToRenderBuffer(0, Z, 4, D3DPT_LINESTRIP, Quad(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)
        
    End If
    
    Erase Specular

Exit Sub

ErrOut:
End Sub

'===============================================================================
' Name: PRECAL_NewCircle
' Input:
'   ByVal X As Long - Coordenada horizontal del centro de la circunferencia.
'   ByVal Y As Long - Coordenada vertical del centro de la circunferencia.
'   ByVal Radius As Long - Radio de la circunferencia.
'   ByVal Segment As Long - Define un valor de longitud para dibujar la circunferencia por medio de segmentos. Si el valor es 0 la circunferencia se dibujara por medio de píxeles.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo del borde de la figura.
'   Optional ByVal Fill As Boolean - Indica si se dibujara la figura con relleno.
'   Optional ByVal FillColor As Long - Color ARGB que se aplicara al relleno de la figura.
' Output:
'   Long - Identificador de la lista precalculada.
' Purpose: Calcula los parámetros de dibujo de una circunferencia y los almacena en un lista en memoria.
' Remarks: Dado que el calculo de los parámetros de la circunferencia y su dibujo son procesos lentos para sistema de dibujo primero se calculan y se almacenan los parámetros en memoria para luego ser usados directamente en la función DRAW_Circle. Se recomienda para no perder mucho rendimiento calcular las circunferencias con un valor definido para el parámetro Segment. De esta forma perderá calidad la circunferencia pero le costara menos trabajo dibujarla y perder menos velocidad.
'===============================================================================
Public Function PRECAL_NewCircle(X As Long, Y As Long, Radius As Long, Segment As Long, Color As Long, Optional Fill As Boolean, Optional FillColor As Long) As Long
On Error GoTo ErrOut

Dim Angle As Single
Dim cX As Long, cY As Long
Dim i As Long, j As Long
Dim MatchColor As Boolean
Dim Id As Long

If D3D_Init Then

    Id = Get_FreeID(2)
    
    With PreCircle(Id)
    
        If Color = FillColor Then MatchColor = True

        'Creamos las coordenadas de los puntos que forman el circulo:
        If Segment <= 0 Then Segment = 1 Else If Segment > Radius Then Segment = Radius
    
        Do While Angle < 6.28
        
            cX = X + Radius * Cos(Angle)
            cY = Y + Radius * Sin(Angle) ' / 1.2
        
            If j Mod Segment = 0 Then
                If Not MatchColor Then
                    ReDim Preserve .Circle_Border(i + 1) As D3DTLVERTEX
                    .Circle_Border(i) = CreateTLVertex(CLng(cX), CLng(cY), 0, 1, Color, 0, 0, 0)
            
                End If
            
                If Fill Then
                    ReDim Preserve .Circle_Fill(i + 1) As D3DTLVERTEX
                    .Circle_Fill(i) = CreateTLVertex(CLng(cX), CLng(cY), 0, 1, FillColor, 0, 0, 0)
            
                End If
                i = i + 1
            
            End If
        
            Angle = Angle + 0.005
            j = j + 1
    
        Loop
    
        If Fill Then .Circle_Fill(i) = CreateTLVertex(.Circle_Fill(0).sx, .Circle_Fill(0).sy, 0, 1, FillColor, 0, 0, 0)
    
        'Dibujamos el borde:
        If Not MatchColor Then .Circle_Border(i) = CreateTLVertex(.Circle_Border(0).sx, .Circle_Border(0).sy, 0, 1, Color, 0, 0, 0)
        
        .NumVertex = i
        .Fill = Fill
        .MatchColor = MatchColor
        
    End With
        
    PRECAL_NewCircle = Id
        
End If

ErrOut:

End Function

'===============================================================================
' Name: PRECAL_DeleteCircle
' Input:
'   ByVal Id As Long - Identificador de la lista precalculada.
' Output:
' Purpose: Borra una lista precalculada de una circunferencia.
' Remarks:
'===============================================================================
Public Sub PRECAL_DeleteCircle(Id As Long)
On Error GoTo ErrOut

    If Id = UBound(PreCircle) Then
        If Id = 0 Then
            Erase PreCircle
        
        Else
            ReDim Preserve PreCircle(UBound(PreCircle) - 1) As PreCalCircle_Data
        
        End If
    
    Else
        With PreCircle(Id)
            .Free = True
            .Fill = False
            .MatchColor = False
            .NumVertex = 0
            
            Erase .Circle_Border, .Circle_Fill
        
        End With
        
    End If

ErrOut:

End Sub

'===============================================================================
' Name: DRAW_Circle
' Input:
'   ByVal Id As Long - Identificador de la lista precalculada.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
' Output:
' Purpose: Dibuja una circunferencia con o sin relleno.
' Remarks: Para dibujar una circunferencia antes se debe haber calculado mediante la función PRECAL_NewCircle.
'===============================================================================
Public Sub DRAW_Circle(Id As Long, Z As Long)
On Error GoTo ErrOut

If D3D_Init Then
    With PreCircle(Id)
        'Dibujamos el relleno:
        If .Fill Then Call AddCallToRenderBuffer(0, Z, .NumVertex, D3DPT_TRIANGLEFAN, .Circle_Fill(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)

        'Dibujamos el borde:
        If Not .MatchColor Then Call AddCallToRenderBuffer(0, Z, .NumVertex, D3DPT_LINESTRIP, .Circle_Border(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)

    End With
    
End If

ErrOut:
End Sub

'===============================================================================
' Name: DRAW_Trapezoid
' Input:
'   ByVal VertexArray As Vertex - Lista de vértices del trapezoide.
' Output:
' Purpose: Dibuja en pantalla un trapezoide relleno con efectos de degradado.
' Remarks: Esta función es similar a DRAW_AdvBox con la única diferencia de que esta permite publicar los vértices de la figura libremente mediante el argumento VertexArray, un array de 3 vectores que definen los 4 vértices de la figura y sus colores. Se pueden definir un color diferente a cada vértice creando así efectos de degradados. Esta función solo dibuja figuras sólidas.
'===============================================================================
Public Sub DRAW_Trapezoid(VertexArray() As Vertex)
On Error GoTo ErrOut

Dim Quad(3) As D3DTLVERTEX

    'Generamos el poligono:
    Quad(0) = CreateTLVertex(CSng(VertexArray(0).X), CSng(VertexArray(0).Y), 0, 1, VertexArray(0).Color, Specular(0), 0, 0)
    Quad(1) = CreateTLVertex(CSng(VertexArray(1).X), CSng(VertexArray(1).Y), 0, 1, VertexArray(1).Color, Specular(1), 0, 0)
    Quad(2) = CreateTLVertex(CSng(VertexArray(2).X), CSng(VertexArray(2).Y), 0, 1, VertexArray(2).Color, Specular(2), 0, 0)
    Quad(3) = CreateTLVertex(CSng(VertexArray(3).X), CSng(VertexArray(3).Y), 0, 1, VertexArray(3).Color, Specular(3), 0, 0)

    If InScreen(Quad()) Then Call AddCallToRenderBuffer(0, VertexArray(0).Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)
    
    Erase Specular

ErrOut:

End Sub

'===============================================================================
' Name: DRAW_AdvBox
' Input:
'   ByVal X1 As Long - Coordenada horizontal de la esquina superior izquierda de la caja.
'   ByVal Y1 As Long - Coordenada vertical de la esquina superior izquierda de la caja.
'   ByVal X2 As Long - Coordenada horizontal de la esquina inferior derecha de la caja.
'   ByVal Y2 As Long - Coordenada vertical de la esquina inferior derecha de la caja.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal ColorV1 As Long - Color ARGB que se aplicara al vértice X1 Y1.
'   ByVal ColorV2 As Long - Color ARGB que se aplicara al vértice X2 Y1.
'   ByVal ColorV3 As Long - Color ARGB que se aplicara al vértice X1 Y2.
'   ByVal ColorV4 As Long - Color ARGB que se aplicara al vértice X2 Y2.
' Output:
' Purpose: Versión avanzada de DRAW_Box. Dibuja en pantalla una caja rellena con efectos de degradado.
' Remarks: Esta función es similar a DRAW_Box con la diferencia de que esta permite aplicar un color individual a cada vértice pudiendo crear efectos de degradado entre los colores. Esta función solo dibuja figuras sólidas.
'===============================================================================
Public Sub DRAW_AdvBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Z As Long, ColorV1 As Long, ColorV2 As Long, ColorV3 As Long, ColorV4 As Long)
On Error GoTo ErrOut

Dim Quad(3) As D3DTLVERTEX
    
    'Generamos el poligono:
    Quad(0) = CreateTLVertex(CSng(X1), CSng(Y1), 0, 1, ColorV1, Specular(0), 0, 0)
    Quad(1) = CreateTLVertex(CSng(X2), CSng(Y1), 0, 1, ColorV2, Specular(1), 1, 0)
    Quad(2) = CreateTLVertex(CSng(X1), CSng(Y2), 0, 1, ColorV3, Specular(2), 0, 1)
    Quad(3) = CreateTLVertex(CSng(X2), CSng(Y2), 0, 1, ColorV4, Specular(3), 1, 1)

    If InScreen(Quad()) Then Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Nothing, 0, 0, vbNullString, TextRegion, Nothing, 0, 0)
    
    Erase Specular
    
ErrOut:

End Sub

'===============================================================================
' Name: DRAW_Map
' Input:
'   ByVal Map As Long - Identificador del grafico.
'   ByVal X As Long - Coordenada horizontal de dibujo.
'   ByVal Y As Long - Coordenada vertical de dibujo.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Width As Long - Anchura con la que se dibujara el grafico.
'   ByVal Height As Long - Altura con la que se dibujara el grafico.
' Output:
' Purpose: Dibuja un grafico en pantalla.
' Remarks:
'===============================================================================
Public Sub DRAW_Map(Map As Long, X As Long, Y As Long, Z As Long, Width As Long, Height As Long)

    Call Me.DRAW_MapEx(Map, X, Y, Z, Width, Height, 0, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_None, False)
    
End Sub

'===============================================================================
' Name: DRAW_MapEx
' Input:
'   ByVal Map As Long - Identificador del grafico.
'   ByVal X As Long - Coordenada horizontal de dibujo.
'   ByVal Y As Long - Coordenada vertical de dibujo.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Width As Long - Anchura con la que se dibujara el grafico. Si el valor es 0 se toma las dimensiones originales del grafico, si el valor es -1 se toma la dimensiones del grafico en memoria.
'   ByVal Height As Long - Altura con la que se dibujara el grafico. Si el valor es 0 se toma las dimensiones originales del grafico, si el valor es -1 se toma la dimensiones del grafico en memoria.
'   ByVal Angle As Single - Angulo de rotación con el que se dibujara el grafico. El rango permitido es de 0 a 359.99. Si un valor sobrepasa dicho rango se transforma automáticamente a un valor incluido en el rango. El eje de rotación del grafico se define según el valor de los parámetros Center, RotX y RotY.
'   ByVal AlphaBlendMode As Blit_Alpha - Modo de opacidad.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo. El componente Alpha del color no se aplicara en los modos de opacidad BlendOp_Aditive y BlendOp_Sustrative.
'   ByVal Mirror As Blit_Mirror - Modo de espejado.
'   ByVal Filter As Blit_Filter - Filtro de suavizado que se utilizara para dibujar el grafico.
'   ByVal Center As Boolean - Indica si se tomara como eje de dibujo y de calculo de rotación el centro del grafico.
' Output:
' Purpose: Función extendida de DRAW_Map. Dibuja un grafico en pantalla con efectos.
' Remarks: Esta función añade extras a la versión de DRAW_Map permitiendo operaciones de rotación, espejados, opacidad (Alpha Blending) e incluso aplicar un filtro para suavizar los píxeles del grafico en pantalla.
'===============================================================================
Public Sub DRAW_MapEx(Map As Long, X As Long, Y As Long, Z As Long, Width As Long, Height As Long, Angle As Single, AlphaBlendMode As Blit_Alpha, Color As Long, Mirror As Blit_Mirror, Filter As Blit_Filter, Center As Boolean)
    On Error GoTo ErrOut
    
    Dim S As GFX_Rect 'Copias locales de X, Y, Width y Height
    Dim A As Single ' Angulo en radianes
    Dim cX As Single, cY As Single ' Centro de rotacion.
    Dim Quad(3) As D3DTLVERTEX
    Dim matResult As D3DMATRIX, matScale As D3DMATRIX, matTras As D3DMATRIX, matRot As D3DMATRIX
    Dim Vec As D3DVECTOR2
    Dim tu(3) As Single, tv(3) As Single
    Dim tL As Single, tT As Single, tR As Single, tB As Single
    Dim i As Long

    'Coordenadas de dibujo:
    S.X = X
    S.Y = Y
    
    A = (PI / 180) * Angle
    
    ' Utilizamos color unificado o color independiente por vertice:
    If (Not SetVertexColor) Then
        VertexColor(0) = Color
        VertexColor(1) = Color
        VertexColor(2) = Color
        VertexColor(3) = Color
    End If
    
    ' Generamos el quad base:
    Quad(0) = CreateTLVertex(0, 0, 0, 1, VertexColor(0), Specular(0), 0, 0)
    Quad(1) = CreateTLVertex(CSng(Texture(Map).Texture_Width), 0, 0, 1, VertexColor(1), Specular(1), 1, 0)
    Quad(2) = CreateTLVertex(0, CSng(Texture(Map).Texture_Height), 0, 1, VertexColor(2), Specular(2), 0, 1)
    Quad(3) = CreateTLVertex(CSng(Texture(Map).Texture_Width), CSng(Texture(Map).Texture_Height), 0, 1, VertexColor(3), Specular(3), 1, 1)
    
    ' Definimos la porcion de textura a pintar:
    If FrameSet Then
    
        ' Las medidas van por porcentajes del tamaño real:
        tL = (FrameRect.Left / Texture(Map).Image_Width * 100) / 100
        tT = (FrameRect.Top / Texture(Map).Image_Height * 100) / 100
        tR = (FrameRect.Right / Texture(Map).Image_Width * 100) / 100
        tB = (FrameRect.bottom / Texture(Map).Image_Height * 100) / 100
            
        If Height < 1 Then S.Height = FrameRect.bottom - FrameRect.Top Else S.Height = Height
        If Width < 1 Then S.Width = FrameRect.Right - FrameRect.Left Else S.Width = Width
    
    Else
        tL = 0
        tT = 0
        tR = 1
        tB = 1
 
        If Height = 0 Then
            S.Height = Texture(Map).Image_Height
            
        ElseIf Height = -1 Then
            S.Height = Texture(Map).Texture_Height
            
        Else
            S.Height = Height
        
        End If
        
        If Width = 0 Then
            S.Width = Texture(Map).Image_Width
        
        ElseIf Width = -1 Then
            S.Width = Texture(Map).Texture_Width
            
        Else
            S.Width = Width
        
        End If
        
    End If

    ' Situamos el centro de rotacion:
    If Center Then
        cX = Texture(Map).Texture_Width / 2
        cY = Texture(Map).Texture_Height / 2
    
    ElseIf SetCenter Then
        cX = NewCenterX * (Texture(Map).Texture_Width / (FrameRect.Right - FrameRect.Left))
        cY = NewCenterY * (Texture(Map).Texture_Height / (FrameRect.bottom - FrameRect.Top))
        
        If Mirror = Mirror_Horizontal Then
            cX = Texture(Map).Texture_Width - cX
        ElseIf Mirror = Mirror_Vertical Then
            cY = Texture(Map).Texture_Height - cY
        ElseIf Mirror = Mirror_both Then
            cX = Texture(Map).Texture_Width - cX
            cY = Texture(Map).Texture_Height - cY
        End If
    
    Else
        cX = 0
        cY = 0
    
    End If

    ' Configuramos las coordenadas UV de la textura para aplicar el espejado:
    Select Case Mirror
        Case 0 'Normal
            tu(0) = tL: tv(0) = tT
            tu(1) = tR: tv(1) = tT
            tu(2) = tL: tv(2) = tB
            tu(3) = tR: tv(3) = tB
              
        Case 1 'Horizontal
            tu(0) = tR: tv(0) = tT
            tu(1) = tL: tv(1) = tT
            tu(2) = tR: tv(2) = tB
            tu(3) = tL: tv(3) = tB
        
        Case 2 'Vertical
            tu(0) = tL: tv(0) = tB
            tu(1) = tR: tv(1) = tB
            tu(2) = tL: tv(2) = tT
            tu(3) = tR: tv(3) = tT
                
        Case 3 'Horizontal/Vertical
            tu(0) = tR: tv(0) = tB
            tu(1) = tL: tv(1) = tB
            tu(2) = tR: tv(2) = tT
            tu(3) = tL: tv(3) = tT
            
    End Select
    
    ' Aplicamos las transformaciones:
    For i = 0 To 3
        Vec.X = Quad(i).sx
        Vec.Y = Quad(i).sy

        Call D3DXMatrixIdentity(matResult)

        '1 Traslacion para ubicar el centro de rotacion:
        Call D3DXMatrixTranslation(matTras, CSng(-cX), CSng(-cY), 0)
        Call D3DXMatrixMultiply(matResult, matResult, matTras)

        '2 Escala
        Call D3DXMatrixScaling(matScale, S.Width / Texture(Map).Texture_Width, S.Height / Texture(Map).Texture_Height, 0)
        Call D3DXMatrixMultiply(matResult, matResult, matScale)

        '3 Rotacion
        Call D3DXMatrixRotationZ(matRot, A)
        Call D3DXMatrixMultiply(matResult, matResult, matRot)
        
        '4 Traslacion para ubicar la posicion real:
        Call D3DXMatrixTranslation(matTras, CSng(S.X), CSng(S.Y), 0)
        Call D3DXMatrixMultiply(matResult, matResult, matTras)
        
        '5 Obtener vector
        Call D3DXVec2TransformCoord(Vec, Vec, matResult)

        Quad(i).sx = Vec.X
        Quad(i).sy = Vec.Y
        Quad(i).tu = tu(i)
        Quad(i).tv = tv(i)

    Next i

    If ReadVertex Then
        For i = 0 To 3
            PrecalVertex(i).X = Quad(i).sx: PrecalVertex(i).Y = Quad(i).sy: PrecalVertex(i).Z = Z
        Next
        
        'PrecalSprite
        
        'IsPrecalVertex = True
    Else
        If InScreen(Quad()) Then
            ' Añadimos los datos a la pila de dibujo:
            If (AlphaBlendMode = 128) Then AlphaBlendMode = Blendop_Color
            If (AlphaBlendMode = Blendop_GreyScale) Then Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Texture(Map).Texture, 128, Filter, vbNullString, FrameRect, Nothing, 0, 0)
            Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Texture(Map).Texture, AlphaBlendMode, Filter, vbNullString, FrameRect, Nothing, 0, 0)
        End If
    End If
    
ErrOut:
    Erase Specular
    Erase VertexColor

    SetVertexColor = False
    FrameSet = False
    SetCenter = False
    ReadVertex = False
End Sub

' Determina si un objeto esta dentro del area de la pantalla:
Private Function InScreen(Quad() As D3DTLVERTEX) As Boolean
    InScreen = True ' El algorritmo hay que perfeccionarlo.
'    Dim i As Long, res As Boolean
'    Dim scr As RECT
'
'    scr.Left = 0
'    scr.Top = 0
'    scr.Right = Me.Screen.Width
'    scr.bottom = Me.Screen.Height
'
'    For i = 0 To UBound(Quad)
'        res = res Or CBool(Global_Mod.PtInRect(scr, Quad(i).sx, Quad(i).sy))
'    Next
'
'    InScreen = res ' Se devolvera verdadero siempre y cuando haya algun vertice dentro del area.
End Function

'===============================================================================
' Name: DRAW_AdvMap
' Input:
'   ByVal Map As Long - Identificador del grafico.
'   ByVal X As Long - Coordenada horizontal de dibujo.
'   ByVal Y As Long - Coordenada vertical de dibujo.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Width As Long - Anchura con la que se dibujara el grafico. Si el valor es 0 se toma las dimensiones originales del grafico, si el valor es -1 se toma la dimensiones del grafico en memoria.
'   ByVal Height As Long - Altura con la que se dibujara el grafico. Si el valor es 0 se toma las dimensiones originales del grafico, si el valor es -1 se toma la dimensiones del grafico en memoria.
'   ByVal AlphaBlendMode As Blit_Alpha - Modo de opacidad.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operacion de dibujo. El componente Alpha del color no se aplicara en los modos de opacidad BlendOp_Aditive y BlendOp_Sustrative.
'   ByVal Mirror As Blit_Mirror - Modo de espejado.
'   ByVal Filter As Blit_Filter - Filtro de suavizado que se utilizara para dibujar el grafico.
'   ByVal Perspectve As Blit_Perspective - Especifica el modo de perspectiva que tendra el grafico en pantalla.
'   Optional ByVal Factor As Long - Parámetro opcional que permite introducir un valor para alterar la perspectiva final.
' Output:
' Purpose: Función avanzada de DRAW_MapEx. Dibuja un grafico con efectos aplicandole perspectiva.
' Remarks: Esta función añade extras a la versión de DRAW_Map permitiendo operaciones de espejados, opacidad (Alpha Blending), filtrado para suavizar los pixeles del grafico en pantalla y posibilidad de proyectar un grafico en perspectiva tanto caballera como isométrica.
'===============================================================================
Public Sub DRAW_AdvMap(Map As Long, X As Long, Y As Long, Z As Long, Width As Long, Height As Long, AlphaBlendMode As Blit_Alpha, Color As Long, Mirror As Blit_Mirror, Filter As Blit_Filter, Perspective As Blit_Perspective, Optional Factor As Long)
On Error GoTo ErrOut

Dim Quad(3) As D3DTLVERTEX, tu(3) As Single, tv(3) As Single, eH As Long, eW As Long
Dim tL As Single, tT As Single, tR As Single, tB As Single
Dim tWidth As Long
    
    'Generamos el poligono:
    If FrameSet Then
        
        'Las medidas van por porcentajes del tamaño real:
        tL = (FrameRect.Left / Texture(Map).Image_Width * 100) / 100
        tT = (FrameRect.Top / Texture(Map).Image_Height * 100) / 100
        tR = (FrameRect.Right / Texture(Map).Image_Width * 100) / 100
        tB = (FrameRect.bottom / Texture(Map).Image_Height * 100) / 100
            
        If Height < 1 Then Height = FrameRect.bottom - FrameRect.Top
        If Width < 1 Then Width = FrameRect.Right - FrameRect.Left
    
    Else
        tL = 0
        tT = 0
        tR = 1
        tB = 1
        
        If Height = 0 Then Height = Texture(Map).Image_Height Else If Height = -1 Then Height = Texture(Map).Texture_Height
        If Width = 0 Then Width = Texture(Map).Image_Width Else If Width = -1 Then Width = Texture(Map).Texture_Width
        
    End If
    
    Select Case Mirror
        Case 0 'Normal
            tu(0) = tL: tv(0) = tT
            tu(1) = tR: tv(1) = tT
            tu(2) = tL: tv(2) = tB
            tu(3) = tR: tv(3) = tB
            
        Case 1 'Horizontal
            tu(0) = tR: tv(0) = tT
            tu(1) = tL: tv(1) = tT
            tu(2) = tR: tv(2) = tB
            tu(3) = tL: tv(3) = tB
        
        Case 2 'Vertical
            tu(0) = tL: tv(0) = tB
            tu(1) = tR: tv(1) = tB
            tu(2) = tL: tv(2) = tT
            tu(3) = tR: tv(3) = tT
        
        Case 3 'Horizontal/Vertical
            tu(0) = tR: tv(0) = tB
            tu(1) = tL: tv(1) = tB
            tu(2) = tR: tv(2) = tT
            tu(3) = tL: tv(3) = tT
            
    End Select
    
    Debug.Print tL & ", " & tT & ", " & tR & ", " & tB
    
    Select Case Perspective
        Case Caballera_Width
            If Factor = 0 Then Factor = Width / 2
            
            eW = Factor
            eH = 0
            
        Case Caballera_Height
            If Factor = 0 Then Factor = Height / 4
            
            eW = 0
            eH = Factor
            
        Case Caballera_Width_Negative
            If Factor = 0 Then Factor = Width / 2
            
            eW = -Factor
            eH = 0
            
        Case Caballera_Height_Negative
            If Factor = 0 Then Factor = Height / 4
            
            eW = 0
            eH = -Factor
        
        Case Isometric_Height
            If Factor = 0 Then Factor = Width / 2
            
            eW = 0
            eH = Factor
            
        Case Isometric_Height_Negative
            If Factor = 0 Then Factor = Height / 2
            
            eW = 0
            eH = -Factor
            
    End Select
    
    ' Utilizamos color unificado o color independiente por vertice:
    If (Not SetVertexColor) Then
        VertexColor(0) = Color
        VertexColor(1) = Color
        VertexColor(2) = Color
        VertexColor(3) = Color
    End If
    
    If Perspective = Isometric_Base Then
        tWidth = Width * 2
        
        Quad(0) = CreateTLVertex(CSng(X) + (tWidth / 2), CSng(Y), 0, 1, VertexColor(0), Specular(0), tu(0), tv(0))
        Quad(1) = CreateTLVertex((CSng(X) + tWidth), CSng(Y) + (Height / 2), 0, 1, VertexColor(1), Specular(1), tu(1), tv(1))
        Quad(2) = CreateTLVertex(CSng(X), CSng(Y) + (Height / 2), 0, 1, VertexColor(2), Specular(2), tu(2), tv(2))
        Quad(3) = CreateTLVertex(CSng(X) + (tWidth / 2), (CSng(Y) + Height), 0, 1, VertexColor(3), Specular(3), tu(3), tv(3))
            
    Else
        Quad(0) = CreateTLVertex(CSng(X) + eW, CSng(Y), 0, 1, VertexColor(0), Specular(0), tu(0), tv(0))
        Quad(1) = CreateTLVertex((CSng(X) + Width) + eW, CSng(Y) + eH, 0, 1, VertexColor(1), Specular(1), tu(1), tv(1))
        Quad(2) = CreateTLVertex(CSng(X), (CSng(Y) + Height), 0, 1, VertexColor(2), Specular(2), tu(2), tv(2))
        Quad(3) = CreateTLVertex((CSng(X) + Width), (CSng(Y) + Height) + eH, 0, 1, VertexColor(3), Specular(3), tu(3), tv(3))
        
    End If
    
    If ReadVertex Then
        Dim i As Long
        For i = 0 To 3
            PrecalVertex(i).X = Quad(i).sx: PrecalVertex(i).Y = Quad(i).sy: PrecalVertex(i).Z = Z
        Next
    Else
        If InScreen(Quad()) Then
            If (AlphaBlendMode = 128) Then AlphaBlendMode = Blendop_Color
            If (AlphaBlendMode = Blendop_GreyScale) Then Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Texture(Map).Texture, 128, Filter, vbNullString, FrameRect, Nothing, 0, 0)
            Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Texture(Map).Texture, AlphaBlendMode, Filter, vbNullString, FrameRect, Nothing, 0, 0)
        End If
    End If
    
ErrOut:
    Erase Specular
    Erase VertexColor

    SetVertexColor = False
    FrameSet = False
    ReadVertex = False
End Sub

'===============================================================================
' Name: DRAW_VertexMap
' Input:
'   ByVal Map As Long - Identificador del grafico.
'   ByVal VertexData() As Vertex - Informacion de los vertices de la textura.
'   ByVal VertexSpecular() As Long - Valores specular para indicar la iluminacion a cada vertice.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal AlphaBlendMode As Blit_Alpha - Modo de opacidad.
'   ByVal Mirror As Blit_Mirror - Modo de espejado.
'   ByVal Filter As Blit_Filter - Filtro de suavizado que se utilizara para dibujar el grafico.
' Output:
' Purpose: Función avanzada de DRAW_Map. Dibuja un grafico en pantalla con efectos definiendo la posicion absoluta de los vertices.
' Remarks: Esta función esta diseñada para permitir optimizaciones en implementaciones externas dado que no realiza calculo alguno en cuanto a transformaciones de los vertices ni rotaciones ya que esta funcion esta diseñada para recibir toda la informacion lista para procesarse. Esto puede resultar util si implementaciones externas a dx_lib32 necesitan dibujar graficos en pantalla aplicando transformaciones de forma continuada permitiendo evitar calculos continuos en cada llamada a la funcion implementado tablas de valores que almacenen los cambios puntuales de dibujo, realizando estos calculos de forma externa y solo en el instante de aplicar las transformaciones. A su vez, esta funcion es la unica de las 4 versiones disponibles que permite dibujar un sprite posicionando sus vertices en cualquier posicion libre, pudiendo generar efectos de deformaciones si se aplica en conjunto, por ejemplo ondulaciones submarinas o efectos "lupa" en un mapa de tiles.
'===============================================================================
Public Sub DRAW_VertexMap(Map As Long, VertexData() As Vertex, Z As Long, VertexSpecular() As Long, AlphaBlendMode As Blit_Alpha, Mirror As Blit_Mirror, Filter As Blit_Filter)
    On Error GoTo ErrOut
    
    Dim Quad(3) As D3DTLVERTEX
    Dim tu(3) As Single, tv(3) As Single
    Dim tL As Single, tT As Single, tR As Single, tB As Single
    Dim i As Long
    
    ' Definimos la porcion de textura a pintar:
    If FrameSet Then
        ' Las medidas van por porcentajes del tamaño real:
        tL = (FrameRect.Left / Texture(Map).Image_Width * 100) / 100
        tT = (FrameRect.Top / Texture(Map).Image_Height * 100) / 100
        tR = (FrameRect.Right / Texture(Map).Image_Width * 100) / 100
        tB = (FrameRect.bottom / Texture(Map).Image_Height * 100) / 100
    Else
        tL = 0
        tT = 0
        tR = 1
        tB = 1
    End If

    ' Configuramos las coordenadas UV de la textura para aplicar el espejado:
    Select Case Mirror
        Case 0 'Normal
            tu(0) = tL: tv(0) = tT
            tu(1) = tR: tv(1) = tT
            tu(2) = tL: tv(2) = tB
            tu(3) = tR: tv(3) = tB
        Case 1 'Horizontal
            tu(0) = tR: tv(0) = tT
            tu(1) = tL: tv(1) = tT
            tu(2) = tR: tv(2) = tB
            tu(3) = tL: tv(3) = tB
        Case 2 'Vertical
            tu(0) = tL: tv(0) = tB
            tu(1) = tR: tv(1) = tB
            tu(2) = tL: tv(2) = tT
            tu(3) = tR: tv(3) = tT
        Case 3 'Horizontal/Vertical
            tu(0) = tR: tv(0) = tB
            tu(1) = tL: tv(1) = tB
            tu(2) = tR: tv(2) = tT
            tu(3) = tL: tv(3) = tT
    End Select

    ' Generamos el quad base:
    Quad(0) = CreateTLVertex(CSng(VertexData(0).X), CSng(VertexData(0).Y), 0, 1, VertexData(0).Color, VertexSpecular(0), tu(0), tv(0))
    Quad(1) = CreateTLVertex(CSng(VertexData(1).X), CSng(VertexData(1).Y), 0, 1, VertexData(1).Color, VertexSpecular(1), tu(1), tv(1))
    Quad(2) = CreateTLVertex(CSng(VertexData(2).X), CSng(VertexData(2).Y), 0, 1, VertexData(2).Color, VertexSpecular(2), tu(2), tv(2))
    Quad(3) = CreateTLVertex(CSng(VertexData(3).X), CSng(VertexData(3).Y), 0, 1, VertexData(3).Color, VertexSpecular(3), tu(3), tv(3))

    If ReadVertex Then
        For i = 0 To 3
            PrecalVertex(i).X = Quad(i).sx: PrecalVertex(i).Y = Quad(i).sy: PrecalVertex(i).Z = Z
        Next
    Else
        If InScreen(Quad()) Then
            ' Añadimos los datos a la pila de dibujo:
            If (AlphaBlendMode = 128) Then AlphaBlendMode = Blendop_Color
            If (AlphaBlendMode = Blendop_GreyScale) Then Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Texture(Map).Texture, 128, Filter, vbNullString, FrameRect, Nothing, 0, 0)
            Call AddCallToRenderBuffer(0, Z, 2, D3DPT_TRIANGLESTRIP, Quad(), Texture(Map).Texture, AlphaBlendMode, Filter, vbNullString, FrameRect, Nothing, 0, 0)
        End If
    End If
    
ErrOut:
    Erase Specular
    Erase VertexColor

    SetVertexColor = False
    FrameSet = False
    SetCenter = False
    ReadVertex = False
End Sub

'===============================================================================
' Name: PRECAL_WriteSpriteTransformVertex
' Input:
' Output:
' Purpose: Prepara las funciones de dibujo de graficos para calcular los vertices transformados del sprite.
' Remarks: Este metodo configura las funciones de dibujo DRAW_Map, DRAW_MapEx y DRAW_AdvMap para precalcular la posicion de los vertices de un sprite con las transformaciones afectadas por escala y rotacion. Si se llama a este metodo el grafico no se dibujara. Para obtener los vertices precalculados utilice el metodo PRECAL_ReadSpriteTransformVertex.
'===============================================================================
Public Sub PRECAL_WriteSpriteTransformVertex()
    ReadVertex = True
End Sub

'===============================================================================
' Name: PRECAL_ReadSpriteTransformVertex
' Input:
'   ByVal LeftTop As Vertex - Posicion del primer vertice. (esquina superior izquierda)
'   ByVal RightTop As Vertex - Posicion del segundo vertice. (esquina superior derecha)
'   ByVal LeftBottom As Vertex - Posicion del tercer vertice. (esquina inferior izquierda)
'   ByVal RightBottom As Vertex - Posicion del cuarto vertice. (esquina inferior derecha)
' Output:
' Purpose: Obtiene los vertices transformados del ultimo sprite precalculado.
' Remarks: Este metodo obtiene la posicion de los ultimos vertices precalculados de un sprite con las transformaciones afectadas por escala y rotacion.
'===============================================================================
Public Sub PRECAL_ReadSpriteTransformVertex(LeftTop As Vertex, RightTop As Vertex, LeftBottom As Vertex, RightBottom As Vertex)
    LeftTop = PrecalVertex(0)
    RightTop = PrecalVertex(1)
    LeftBottom = PrecalVertex(2)
    RightBottom = PrecalVertex(3)
End Sub

'===============================================================================
' Name: DEVICE_SetVertexColor
' Input:
'   ByVal A As Long - Color ARGB del primer vertice. (esquina superior izquierda)
'   ByVal B As Long - Color ARGB del segundo vertice. (esquina superior derecha)
'   ByVal C As Long - Color ARGB del tercer vertice. (esquina inferior izquierda)
'   ByVal D As Long - Color ARGB del cuarto vertice. (esquina inferior derecha)
' Output:
' Purpose: Establece el color de los vertices en una operacion de dibujo.
' Remarks: Este metodo permite establecer el color de los vertices en un sprite por separado. Si se llama a este metodo se omite el valor del parametro Color en DRAW_MapEx y DRAW_AdvMap.
' Para poder configurar los valores de cada vertice de un grafico antes se ha de llamar a este procedimiento para asignar los valores a cada vertice.
' Dichos valores se reiniciaran a 0 despues de terminar una llamada a cualquiera de las funciones funciones a las que se aplica los valores de este procedimiento:
' DRAW_Map, DRAW_MapEx, DRAW_AdvMap
'===============================================================================
Public Sub DEVICE_SetVertexColor(A As Long, B As Long, C As Long, D As Long)
    VertexColor(0) = A
    VertexColor(1) = B
    VertexColor(2) = C
    VertexColor(3) = D
    SetVertexColor = True
End Sub

'Ejecuta una pila de llamadas:
Private Sub ExecuteRenderBuffer(pBuffer() As Render_Buffer)
    Dim i As Long, j As Long, k As Long, l As Long, m As Long
    Dim currentTexture As Direct3DTexture8
    
    On Error Resume Next
    
    If TargetID = -1 Then
        Call d3dDevice.BeginScene
    
    Else
        With RenderTarget(TargetID)
            'Abrimos el Render Target:
            Call .Target.BeginScene(.Surface, .ViewPort)
        
        End With
                
        'Limpiamos la superficie:
        Call d3dDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0)
    
    End If
    
    For i = UBound(pBuffer) To LBound(pBuffer) Step -1 'Recorremos el array principal a la inversa:
        
        On Local Error Resume Next
        j = UBound(pBuffer(i).Z) 'Obtenemos el numero de elementos del array.
        
        If Not Err.Number = 9 Then 'Si hay elementos leemos el array:
        
            For k = 0 To j
                With pBuffer(i).Z(k)
                    'Configuramos los estados de renderizado para dibujar sprites o fuentes de texto:
                    If .BlitType >= 0 Or .BlitType <= 2 Then
                        'Si el modo de alpha es diferente al el establecido se aplican los cambios:
                        If Not .AlphaMode = D3D_AlphaBlending Then
                            ' Estados de textura por defecto si se cumple la condicion:
                            If Not ((.AlphaMode = 3) Or (.AlphaMode = 128) Or (.AlphaMode = 256)) Then Call SetDefaultTextureStages
                            
                            ' Estados de blending por defecto si se cumple la condicion:
                            If Not ((.AlphaMode = 1) Or (.AlphaMode = 2) Or (.AlphaMode = 4) Or (.AlphaMode = 256)) Then Call SetDefaultBlendingStages
                            
                            Select Case .AlphaMode
                                Case 1 ' Aditivo:
                                    d3dDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                                    d3dDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                                   
                                Case 2 ' Sustrativo:
                                    d3dDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
                                    d3dDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
                                
                                Case 3 ' Negativo: * Pogacha
                                    d3dDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG1, (D3DTA_TEXTURE Or D3DTA_COMPLEMENT)
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_CURRENT
                                    
                                    d3dDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
                                    d3dDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
                                    d3dDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
                                    
                                Case 4 ' Exclusion (XOR):
                                    d3dDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVDESTCOLOR
                                    d3dDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
                                
                                Case 5 ' Efecto cristal:
                                    d3dDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
                                    d3dDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCCOLOR
                                
                                Case 128 ' Escala grises (1º Pasada) * Pogacha
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_DISABLE
                                    d3dDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
                                    d3dDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1 Or D3DTOP_MODULATE

                                Case 256 ' Escala grises (2º Pasada) * Pogacha
                                    d3dDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorMake(0.299, 0.587, 0.114, 1#)
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
                                    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG2, (D3DTA_TEXTURE Or D3DTA_ALPHAREPLICATE)

                                    d3dDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_CURRENT
                                    d3dDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_DOTPRODUCT3
                                    d3dDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_TFACTOR

                                    d3dDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
                                    d3dDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                            End Select
                            
                            D3D_AlphaBlending = .AlphaMode
                        
                        End If
                        
                        'Si el filtro es diferente al el establecido se aplican los cambios:
                        If Not .Filter = D3D_Filter Then
                            Select Case .Filter
                                
                                Case 1 'Bilinear:
                                    d3dDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
                                    d3dDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
                                            
                                Case 2 'Trilinear:
                                    d3dDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
                                    d3dDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
                        
                                Case Else 'Sin filtro:
                                    d3dDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
                                    d3dDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
                                                
                            End Select
                            
                            D3D_Filter = .Filter
                        
                        End If
                    
                    End If
                    
                    Select Case .BlitType
                        Case 0 'Primitivas graficas y texturas:
                            If (Not .Texture Is currentTexture) Or (currentTexture Is Nothing) And (Not .Texture Is Nothing) Then
                                Set currentTexture = .Texture
                                d3dDevice.SetTexture 0, currentTexture
                            End If
                            
                            d3dDevice.DrawPrimitiveUP .PrimitiveType, .PrimitiveCount, .Quad(0), Len(.Quad(0))
                            
                        Case 1 'Texto:
                            'Calculamos el espacio que ocupara el texto en pantalla:
                            Call D3DX.DrawText(.FontDesc, .FontColor, .Text, .RectRegion, DT_CALCRECT)
                            
                            'Corregimos la coordenada X:
                            Select Case .TextAlign
                                Case Align_Left
                                    .RectRegion.Left = .PrimitiveCount
                                
                                Case Align_Center
                                    .RectRegion.Left = ((Screen.Width / 2) - (.RectRegion.Right / 2))
                                    
                                Case Align_Right
                                    .RectRegion.Left = (Screen.Width - .RectRegion.Right) - .PrimitiveCount
                                
                            End Select
                            
                            'Corregimos el resto de valores:
                            .RectRegion.Top = .PrimitiveType
                            .RectRegion.Right = .RectRegion.Left + .RectRegion.Right
                            .RectRegion.bottom = .RectRegion.Top + .RectRegion.bottom
                                                    
                            'Dibujamos el texto:
                            Call D3DX.DrawText(.FontDesc, .FontColor, .Text, .RectRegion, .TextAlign)
                                                
                    End Select
                
                End With
                
                'DoEvents
            
            Next k
        
        End If
    
    Next i
    
    If TargetID = -1 Then
        Call d3dDevice.EndScene
        Erase MainBuffer
    
    Else
        With RenderTarget(TargetID)
            'Cerramos el Render Target:
            Call .Target.EndScene
            
            Erase .Buffer
        
        End With
    
    End If
    
    'Reiniciamos las variables de comprobacion de estado:
    D3D_AlphaBlending = -1
    D3D_Filter = -1

End Sub

Private Sub SetDefaultBlendingStages()
    d3dDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    d3dDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Private Sub SetDefaultTextureStages()
    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    d3dDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_CURRENT
    d3dDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    d3dDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    d3dDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    
    d3dDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
    d3dDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
    d3dDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
    
    d3dDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP
    d3dDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP
End Sub

Private Sub AddCallToRenderBuffer(BlitType As Long, Z As Long, PrimitiveCount As Long, PrimitiveType As Long, Quad() As D3DTLVERTEX, Texture As Direct3DTexture8, AlphaMode As Long, Filter As Long, Text As String, RectRegion As RECT, FontDesc As D3DXFont, TextColor As Long, TextAlign As Text_Align)
On Local Error Resume Next
    
Dim i As Long
Dim pBuffer As Blit_Structure
    
If Z < -z_Levels Then
    Z = -z_Levels
    
ElseIf Z > z_Levels Then
    Z = z_Levels
    
End If
        
If TargetID = -1 Then
    i = UBound(MainBuffer(Z).Z)  'Obtenemos el numero de elementos del array.
    
    If Err.Number = 9 Then 'Si no hay elementos se añade el primero:
        ReDim MainBuffer(Z).Z(0) As Blit_Structure
        i = 0
        
    Else 'Si hay elementos se añade otro al array:
        ReDim Preserve MainBuffer(Z).Z(i + 1) As Blit_Structure
        i = i + 1
           
    End If

Else
    i = UBound(RenderTarget(TargetID).Buffer(Z).Z)  'Obtenemos el numero de elementos del array.
    
    If Err.Number = 9 Then 'Si no hay elementos se añade el primero:
        ReDim RenderTarget(TargetID).Buffer(Z).Z(0) As Blit_Structure
        i = 0
        
    Else 'Si hay elementos se añade otro al array:
        ReDim Preserve RenderTarget(TargetID).Buffer(Z).Z(i + 1) As Blit_Structure
        i = i + 1
           
    End If

End If
    
'Se introducen los datos:
With pBuffer
    .Quad = Quad
    
    Set .Texture = Texture

    .AlphaMode = AlphaMode
    .Filter = Filter
    
    .PrimitiveCount = PrimitiveCount
    .PrimitiveType = PrimitiveType
    
    .BlitType = BlitType
    
    .RectRegion = RectRegion
    
    .FontColor = TextColor
    Set .FontDesc = FontDesc
    .Text = Text
    .TextAlign = TextAlign

End With
    
If TargetID = -1 Then
    MainBuffer(Z).Z(i) = pBuffer

Else
    RenderTarget(TargetID).Buffer(Z).Z(i) = pBuffer

End If
    
End Sub

Private Function TemporalRenderBufferElement(BlitType As Long, Z As Long, PrimitiveCount As Long, PrimitiveType As Long, Quad() As D3DTLVERTEX, Texture As Direct3DTexture8, AlphaMode As Long, Filter As Long, Text As String, RectRegion As RECT, FontDesc As D3DXFont, TextColor As Long, TextAlign As Text_Align) As Blit_Structure
    On Local Error Resume Next
        
    Dim pBuffer As Blit_Structure
        
    'Se introducen los datos:
    With pBuffer
        .Quad = Quad
        
        Set .Texture = Texture
    
        .AlphaMode = AlphaMode
        .Filter = Filter
        
        .PrimitiveCount = PrimitiveCount
        .PrimitiveType = PrimitiveType
        
        .BlitType = BlitType
        
        .RectRegion = RectRegion
        
        .FontColor = TextColor
        Set .FontDesc = FontDesc
        .Text = Text
        .TextAlign = TextAlign
    
    End With
    
    TemporalRenderBufferElement = pBuffer
    
End Function

Private Sub AddCallToRenderBufferFromPrecalData(Data As Blit_Structure, Z As Long)
    Call AddCallToRenderBuffer(Data.BlitType, Z, Data.PrimitiveCount, Data.PrimitiveType, Data.Quad(), Data.Texture, Data.AlphaMode, Data.Filter, Data.Text, Data.RectRegion, Data.FontDesc, Data.FontColor, Data.TextAlign)
End Sub

'===============================================================================
' Name: DRAW_Text
' Input:
'   ByVal Font As Long - Identificador de la fuente.
'   ByVal Text As String - Cadena de texto que se va a dibujar en pantalla.
'   ByVal X As Long - Coordenada horizontal de dibujo. Este valor no altera la posicion de dibujo del texto cuando el valor del parámetro Align es Align_Center.
'   ByVal Y As Long - Coordenada vertical de dibujo.
'   ByVal Z As Long - Coordenada de profundidad de dibujo.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo.
'   ByVal Align As Text_Align - Define el alineamiento del texto en pantalla.
' Output:
' Purpose: Dibuja una cadena de texto con formato en pantalla.
' Remarks: Para usar esta función previamente se debe haber cargado una fuente en memoria mediante FONT_LoadSystemFont.<br>
'Esta función es lenta por lo que no conviene abusar de ella ya que el rendimiento de nuestro programa se reducirá considerablemente.
'===============================================================================
Public Sub DRAW_Text(Font As Long, Text As String, X As Long, Y As Long, Z As Long, Color As Long, Align As Text_Align)
Call AddCallToRenderBuffer(1, Z, X, Y, NullQuad, Nothing, 0, 0, Text, TextRegion, D3DFont(Font).D3D_Font, Color, Align)

End Sub

'===============================================================================
' Name: DEVICE_GetDisplayModeCount
' Input:
' Output:
'   Long - Numero de modos de video disponibles.
' Purpose: Devuelve el numero de modos de video de la tarjeta.
' Remarks: Esta función no requiere inicializar la clase para ser ejecutada.
'===============================================================================
Public Function DEVICE_GetDisplayModeCount() As Long
On Error GoTo ErrOut
Dim hDx As DirectX8
Dim hD3D As Direct3D8

If Not D3D_Init Then
    Set hDx = New DirectX8
    Set hD3D = hDx.Direct3DCreate
    
Else
    Set hDx = Dx
    Set hD3D = D3D
    
End If

DEVICE_GetDisplayModeCount = hD3D.GetAdapterModeCount(D3DADAPTER_DEFAULT)

ErrOut:
Set hDx = Nothing
Set hD3D = Nothing

End Function

'===============================================================================
' Name: DEVICE_GetDisplayMode
' Input:
'   ByVal Mode As Long - Modo de video que leeremos.
'   ByVal Data As Display_Mode - Parámetro de salida que nos devolverá la informacion sobre el modo de video.
' Output:
' Purpose: Devuelve la información del modo de video seleccionado.
' Remarks: Esta función no requiere inicializar la clase para ser ejecutada.
'===============================================================================
Public Sub DEVICE_GetDisplayMode(Mode As Long, Data As Display_Mode)
On Error GoTo ErrOut

Dim hDx As DirectX8
Dim hD3D As Direct3D8

Dim ModeTemp As D3DDISPLAYMODE

If Not D3D_Init Then
    Set hDx = New DirectX8
    Set hD3D = hDx.Direct3DCreate
    
Else
    Set hDx = Dx
    Set hD3D = D3D
    
End If

Call hD3D.EnumAdapterModes(D3DADAPTER_DEFAULT, Mode, ModeTemp)
    
'Comprobamos si el dispositivo es valido:
If hD3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
    With Data
        .Height = ModeTemp.Height
        .Width = ModeTemp.Width
        .RefreshRate = ModeTemp.RefreshRate
            
    'Modos de 32 Bits:
    If ModeTemp.Format = D3DFMT_R8G8B8 Or ModeTemp.Format = D3DFMT_X8R8G8B8 Or ModeTemp.Format = D3DFMT_A8R8G8B8 Then
        .Bpp = 32

    'Modos de 16 Bits:
    Else
        .Bpp = 16
        
    End If
    
    End With
    
End If

ErrOut:
Set hDx = Nothing
Set hD3D = Nothing

End Sub

'===============================================================================
' Name: DEVICE_ExistsDisplayMode
' Input:
'   ByVal Width As Long - Anchura del modo de video.
'   ByVal Height As Long - Altura del modo de video.
'   ByVal Bpp As Long - Profundidad de color.
'   ByVal RefreshRate As Long - Refresco del modo de video en herzios. Si se pasa como valor 0 se usa el valor por defecto del sistema.
' Output:
'   Boolean - Devuelve verdadero si el modo de video existe y es soportado por el dispositivo de video.
' Purpose: Comprueba que la resolución indicada es soportada por el dispositivo de video.
' Remarks: Esta función no requiere inicializar la clase para ser ejecutada.
'===============================================================================
Public Function DEVICE_ExistsDisplayMode(Width As Long, Height As Long, Optional Bpp As Long = 16, Optional RefreshRate As Long = 60) As Boolean
On Error GoTo ErrOut

Dim hDx As DirectX8
Dim hD3D As Direct3D8
Dim Color As Long
Dim i As Integer, ModeTemp As D3DDISPLAYMODE, nModes As Long

If Not D3D_Init Then
    Set hDx = New DirectX8
    Set hD3D = hDx.Direct3DCreate
    
Else
    Set hDx = Dx
    Set hD3D = D3D
    
End If

Select Case Bpp
    Case 16: Color = D3DFMT_R5G6B5
    Case 32: Color = D3DFMT_X8R8G8B8
    
End Select

If RefreshRate < 60 Then RefreshRate = 60

nModes = hD3D.GetAdapterModeCount(D3DADAPTER_DEFAULT)

For i = 0 To nModes - 1
    
    Call hD3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, ModeTemp)
    
    If ModeTemp.Height = Height And ModeTemp.Width = Width And ModeTemp.Format = Color And ModeTemp.RefreshRate = RefreshRate Then
        DEVICE_ExistsDisplayMode = True
        Exit For
        
    End If
    
    'DoEvents
Next i

ErrOut:

Set hDx = Nothing
Set hD3D = Nothing
End Function

'===============================================================================
' Name: DEVICE_GetInfo
' Input:
'   ByVal Data As Device_Info - Parámetro de salida que nos devolverá la informacion sobre el dispositivo de video.
' Output:
' Purpose: Devuelve información sobre el dispositivo de video.
' Remarks: Esta función no requiere inicializar la clase para ser ejecutada.
'===============================================================================
Public Sub DEVICE_GetInfo(Data As Device_Info)
On Error GoTo ErrOut

Dim hDx As DirectX8
Dim hD3D As Direct3D8
Dim Caps As D3DCAPS8 'Informacion sobre el dispositivo.
Dim AdapterInfo As D3DADAPTER_IDENTIFIER8 'Informacion sobre el adaptador grafico.

Dim sTemp As String
Dim j As Long
Dim i As Integer, ModeTemp As D3DDISPLAYMODE, nModes As Long

If Not D3D_Init Then
    Set hDx = New DirectX8
    Set hD3D = hDx.Direct3DCreate
    
Else
    Set hDx = Dx
    Set hD3D = D3D
    
End If

Call hD3D.GetAdapterIdentifier(D3DADAPTER_DEFAULT, 0, AdapterInfo)
    
Call hD3D.GetDeviceCaps(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps)
    
With Data
    
    'Descripcion del adaptador:
    sTemp = vbNullString
    For j = 0 To 511
        If Chr$(AdapterInfo.Description(j)) = Chr$(0) Then Exit For
        
        sTemp = sTemp & Chr$(AdapterInfo.Description(j))
        'DoEvents
        
    Next j
    
    .AdapterName = Trim(sTemp)
    
    'Driver del adaptador:
    sTemp = vbNullString
    For j = 0 To 511
        If Chr$(AdapterInfo.Driver(j)) = Chr$(0) Then Exit For
        
        sTemp = sTemp & Chr$(AdapterInfo.Driver(j))
        'DoEvents
        
    Next j
    
    .DriverName = Trim(sTemp)
    
    'Version primaria del driver:
    .DriverVersionHigh = AdapterInfo.DriverVersionHighPart And &HFFFF&
    
    'Version secundaria del driver:
    .DriverVersionLow = AdapterInfo.DriverVersionLowPart And &HFFFF&
    
    'Revision:
    .Revision = AdapterInfo.Revision And &HFFFF&
    
    'Vendor ID:
    .VendorID = AdapterInfo.VendorID And &HFFFF&
    
    'Altura maxima de una textura:
    .MaxTextureHeight = Caps.MaxTextureHeight
    
    'Anchura maxima de una textura:
    .MaxTextureWidth = Caps.MaxTextureWidth
    
    'Maximos Bits de color soportados:
    nModes = hD3D.GetAdapterModeCount(D3DADAPTER_DEFAULT)

    For i = 0 To nModes - 1
    
        Call hD3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, ModeTemp)
    
        If ModeTemp.Format = D3DFMT_X8R8G8B8 Then
            .MaxBitsPerPixel = 32
            
            Exit For
        
        Else
            .MaxBitsPerPixel = 16
        
        End If
        
        'DoEvents
        
    Next i

End With

ErrOut:
Set hDx = Nothing
Set hD3D = Nothing

End Sub

'===============================================================================
' Name: DEVICE_GetFreeVideoMem
' Input:
' Output:
'   Long - Memoria de video libre en bytes.
' Purpose: Devuelve la cantidad de memoria de Video libre.
' Remarks: Esta función no requiere inicializar la clase para ser ejecutada.
'===============================================================================
Public Function DEVICE_GetFreeVideoMem() As Long
    If d3dDevice Is Nothing Then
        DEVICE_GetFreeVideoMem = DEVICE_GetTotalVideoMem
    Else
        DEVICE_GetFreeVideoMem = d3dDevice.GetAvailableTextureMem(D3DPOOL_MANAGED)
    End If
End Function

'===============================================================================
' Name: DEVICE_GetTotalVideoMem
' Input:
' Output:
'   Long - Memoria de video disponible en bytes.
' Purpose: Devuelve la cantidad total de memoria de Video disponible.
' Remarks: Esta función no requiere inicializar la clase para ser ejecutada.
'===============================================================================
Public Function DEVICE_GetTotalVideoMem() As Long
    DEVICE_GetTotalVideoMem = TotalVideoMem
End Function

'===============================================================================
' Name: DEVICE_SetAntialiasing
' Input:
'   ByVal Enable As Boolean - Activa el modo de suavizado de contorno de pixeles.
' Output:
' Purpose: Activa o desactiva el suavizado de contorno de píxeles en las operaciones de dibujo de primitivas.
' Remarks: Si su dispositivo de video no soporta el modo de Antialiasing esta función no surtirá efecto alguno.
'===============================================================================
Public Sub DEVICE_SetAntialiasing(Enable As Boolean)
If Not Me.Antialiasing = Enable Then
    d3dDevice.SetRenderState D3DRS_EDGEANTIALIAS, Enable
    d3dDevice.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, Enable
    D3D_Antialiasing = Enable
    
End If

End Sub

'===============================================================================
' Name: DEVICE_SetDisplayMode
' Input:
'   ByVal Width As Long - Ancho en píxeles del modo de video.
'   ByVal Height As Long - Alto en píxeles del modo de video.
'   ByVal Bpp As Long - Profundidad de color del modo de video.
'   ByVal Windowed As Boolean - Indica si se inicializa en modo ventana.
'   ByVal TripleBuffer As Boolean - Indica si se inicializa con Triple Buffer activado.
'   ByVal vSync As Boolean - Indica si se inicializa con la espera de refresco vertical activada.
'   ByVal RefreshRate As Long - Refresco del modo de video en herzios. Si se pasa como valor 0 se usa el valor por defecto del sistema.
' Output:
'   Boolean - Devuelve verdadero si se ha realizado con éxito el cambio de modo de video.
' Purpose: Cambia el modo de video en tiempo de ejecución.
' Remarks: Asegúrese de que el modo de video exista y sea valido y soportado por el dispositivo de video, de lo contrario se producirá un fallo en el sistema de video. En caso de ocurrir se puede volver a llamar a esta función para aplicar un modo de video correcto, por lo tanto no debe preocuparse por los datos cargados en memoria.
'===============================================================================
Public Function DEVICE_SetDisplayMode(Width As Long, Height As Long, Optional Bpp As Long = 16, Optional Windowed As Boolean, Optional TripleBuffer As Boolean, Optional VSync As Boolean, Optional RefreshRate As Long = 60) As Boolean
On Error GoTo ErrOut

Dim DispMode As D3DDISPLAYMODE, lstyle As Long
Dim TempD3DWindow As D3DPRESENT_PARAMETERS
Dim lErrNum As Long

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

If RefreshRate < 60 Then RefreshRate = 60

If Windowed Then
    TempD3DWindow.Windowed = 1
    If VSync Then TempD3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC Else TempD3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    TempD3DWindow.BackBufferFormat = DispMode.Format
    TempD3DWindow.BackBufferHeight = Height
    TempD3DWindow.BackBufferWidth = Width
    TempD3DWindow.hDeviceWindow = Me.hWnd
    
    'Devolvemos cualquier estilo a la ventana:
    Call Global_Mod.SetWindowLong(hWnd, GWL_STYLE, WndStyle)

    Global_Mod.D3D_FullScreen = False
    
Else
    Select Case Bpp
        Case 16, Is < 16: DispMode.Format = D3DFMT_R5G6B5
        Case 32, Is > 32: DispMode.Format = D3DFMT_X8R8G8B8
        
    End Select
    
    DispMode.Width = Width
    DispMode.Height = Height

    If VSync Then
        TempD3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
        
    Else
        TempD3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
        TempD3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    End If
    
    TempD3DWindow.BackBufferCount = 2 + Abs(TripleBuffer)
    TempD3DWindow.BackBufferFormat = DispMode.Format
    TempD3DWindow.BackBufferHeight = DispMode.Height
    TempD3DWindow.BackBufferWidth = DispMode.Width
    TempD3DWindow.FullScreen_RefreshRateInHz = RefreshRate
    TempD3DWindow.hDeviceWindow = Me.hWnd
    
'    'Quitamos cualquier estilo a la ventana para conseguir una ventana rectangular:
    lstyle = Global_Mod.GetWindowLong(Me.hWnd, GWL_STYLE)
    lstyle = lstyle And WGAME_STYLE
    Call Global_Mod.SetWindowLong(hWnd, GWL_STYLE, lstyle)
    
    Global_Mod.D3D_FullScreen = True
    
End If

'Call Global_Mod.ShowWindow(Me.hwnd, SW_RESTORE)

Call ResetDevice(TempD3DWindow)

Call Global_Mod.Set_SizeWindow(Me.hWnd, Height, Width, Windowed)

D3DWindow = TempD3DWindow

With D3D_Screen
    .Bpp = Bpp
    .Width = Width
    .Height = Height
    .RefreshRate = RefreshRate

End With

DEVICE_SetDisplayMode = True
D3D_TripleBuffer = TripleBuffer
D3D_VSync = VSync

Exit Function

ErrOut:

End Function

'===============================================================================
' Name: MAP_Create
' Input:
'   ByVal Width As String - Anchura en pixeles.
'   ByVal Height As String - Altura en pixeles.
'   ByVal RenderTarget As Boolean - Define si el grafico sera usado como destino de un Render Target. Esa opcion optimiza ligeramente las operaciones de Render Targets con los graficos.
'   ByVal GrayScale As Boolean - Define si el grafico utilizara escala de grises.
' Output:
'   Long - Devuelve el código identificador para este grafico en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar el grafico se devuelve GFX_FILENOTLOAD.
' Purpose: Crea un grafico vacio en memoria.
' Remarks: Se pueden crear tantos graficos como se deseen teniendo como limite la memoria del dispositivo grafico.
'===============================================================================
Public Function MAP_Create(Width As Long, Height As Long, Optional RenderTarget As Boolean, Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long
Dim Format As CONST_D3DFORMAT

If D3D_Init Then
    i = -1
    i = Get_FreeID(0) 'D3DX_FILTER_TRIANGLE Or D3DX_FILTER_MIRROR

    If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8
    
    Set Texture(i).Texture = D3DX.CreateTexture(d3dDevice, Width, Height, _
                     D3DX_DEFAULT, Abs(RenderTarget), Format, _
                     Not RenderTarget)

    'Obtenemos los valores de la imagen original:
    Texture(i).Image_Height = Height
    Texture(i).Image_Width = Width

    'Obtenemos los valores de la textura en memoria:
    Call Texture(i).Texture.GetLevelDesc(0, D3D_Desc)
    Texture(i).Texture_Height = D3D_Desc.Height
    Texture(i).Texture_Width = D3D_Desc.Width
    Texture(i).ColorKey = &H0
    Texture(i).RenderTarget = RenderTarget
    
    MAP_Create = i
    Exit Function
    
End If

ErrOut:
If Not i = -1 Then Call Me.MAP_Unload(i)
MAP_Create = GFX_NOTCREATED

End Function

'===============================================================================
' Name: MAP_Load
' Input:
'   ByVal FileName As String - Nombre y ruta de la imagen.
'   ByVal ColorKey As Boolean - Valor que define el color transparente para este grafico.
'   ByVal Smooth As Boolean - Valor que define si el grafico se cargara con suavizado de píxeles o no.
'   ByVal GrayScale As Boolean - Define si el grafico se cargara en escala de grises.
' Output:
'   Long - Devuelve el código identificador para este grafico en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar el grafico se devuelve GFX_FILENOTLOAD.
' Purpose: Carga un archivo de imagen desde disco como un grafico.
' Remarks: Esta función soporta los siguientes formatos gráficos: *.DDS, *.TGA, *.JPG, *.JPEG, *.JFIF, *.JPE, *.BMP, *.DIB, *.TIF, *.TIFF, *.PNG, *.PPM. Se pueden cargar tantos gráficos como se deseen teniendo como limite la memoria del dispositivo grafico.
'===============================================================================
Public Function MAP_Load(Filename As String, ColorKey As Long, Optional Smooth As Boolean = True, Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long
Dim Filter As Long
Dim Format As CONST_D3DFORMAT

If D3D_Init And Global_Mod.FileExists(Filename) Then
    i = -1
    
    i = Get_FreeID(0)
    
    If Smooth Then
        Filter = D3DX_FILTER_TRIANGLE Or D3DX_FILTER_MIRROR
        
    Else
        Filter = D3DX_FILTER_TRIANGLE Or D3DX_FILTER_NONE
    
    End If
    
    If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8

    Set Texture(i).Texture = D3DX.CreateTextureFromFileEx(d3dDevice, Filename, D3DX_DEFAULT, D3DX_DEFAULT, _
                     D3DX_DEFAULT, 0, Format, _
                     D3DPOOL_MANAGED, Filter, _
                     Filter, ColorKey, _
                     D3D_ImgInfo, ByVal 0)

    'Obtenemos los valores de la imagen original:
    Texture(i).Image_Height = D3D_ImgInfo.Height
    Texture(i).Image_Width = D3D_ImgInfo.Width

    'Obtenemos los valores de la textura en memoria:
    Call Texture(i).Texture.GetLevelDesc(0, D3D_Desc)
    Texture(i).Texture_Height = D3D_Desc.Height
    Texture(i).Texture_Width = D3D_Desc.Width
    Texture(i).ColorKey = ColorKey
    Texture(i).Size = D3D_Desc.Size

    MAP_Load = i
    Exit Function
    
End If

ErrOut:
If Not i = -1 Then Call Me.MAP_Unload(i)
MAP_Load = GFX_FILENOTLOAD

End Function

'===============================================================================
' Name: MAP_LoadFromMemory
' Input:
'   ByVal FileData() As Byte - Array de datos que contiene la información del archivo. El argumento se debe pasar apuntando al primer vector del array, el 0.
'   ByVal ColorKey As Boolean - Valor que define el color transparente para este grafico.
'   ByVal Smooth As Boolean - Valor que define si el grafico se cargara con suavizado de píxeles o no.
'   ByVal GrayScale As Boolean - Define si el grafico se cargara en escala de grises.
' Output:
'   Long - Devuelve el código identificador para este grafico en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar el grafico se devuelve GFX_FILENOTLOAD.
' Purpose: Carga un archivo de imagen alojado en memoria como un grafico.
' Remarks: Esta función soporta los siguientes formatos gráficos: *.DDS, *.TGA, *.JPG, *.JPEG, *.JFIF, *.JPE, *.BMP, *.DIB, *.TIF, *.TIFF, *.PNG, *.PPM. Se pueden cargar tantos gráficos como se deseen teniendo como limite la memoria del dispositivo grafico.
'===============================================================================
Public Function MAP_LoadFromMemory(FileData() As Byte, ColorKey As Long, Optional Smooth As Boolean = True, Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long
Dim Filter As Long
Dim Format As CONST_D3DFORMAT

If D3D_Init Then
    If (IsPNG(FileData)) Then
        i = -1
        i = Get_FreeID(0)
        
        If Smooth Then
            Filter = D3DX_FILTER_TRIANGLE Or D3DX_FILTER_MIRROR
    
        Else
            Filter = D3DX_FILTER_NONE Or D3DX_FILTER_MIRROR
    
        End If
        
        If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8
        
        Set Texture(i).Texture = D3DX.CreateTextureFromFileInMemoryEx(d3dDevice, FileData(0), UBound(FileData), D3DX_DEFAULT, D3DX_DEFAULT, _
                         D3DX_DEFAULT, 0, Format, _
                         D3DPOOL_MANAGED, Filter, _
                         Filter, ColorKey, _
                         D3D_ImgInfo, ByVal 0)
                             
        'Obtenemos los valores de la imagen original:
        Texture(i).Image_Height = D3D_ImgInfo.Height
        Texture(i).Image_Width = D3D_ImgInfo.Width
    
        'Obtenemos los valores de la textura en memoria:
        Call Texture(i).Texture.GetLevelDesc(0, D3D_Desc)
        Texture(i).Texture_Height = D3D_Desc.Height
        Texture(i).Texture_Width = D3D_Desc.Width
        Texture(i).ColorKey = ColorKey
        Texture(i).Size = D3D_Desc.Size

    Else ' Si no es una imagen PNG la guardamos en un archivo temporal y la cargamos mediante MAP_Load():
        Dim file As Integer, tmpFile As String
        file = FreeFile
        tmpFile = Global_Mod.GetTempFileName
        
        ' Guardamos en un archivo temporal el contenido del array:
        Open tmpFile For Binary Access Write As file
            Put #file, , FileData
        Close file
        
        ' Cargamos el archivo mediante MAP_Load():
        i = MAP_Load(tmpFile, ColorKey, Smooth, GrayScale)
        
        ' Borramos el archivo temporal:
        Call Kill(tmpFile)
        
        ' Si falla la carga se provoca manualmente un error para ejecutar el codigo de control de errores del metodo:
        If (Not i = GFX_OK) Then Call Err.Raise(SubError)
    End If

    MAP_LoadFromMemory = i
    Exit Function

End If

ErrOut:
If Not i = -1 Then Call Me.MAP_Unload(i)
MAP_LoadFromMemory = GFX_FILENOTLOAD

End Function
' Comprueba si los datos del array pertenecen al contenido de un archivo de imagen con formato PNG:
Private Function IsPNG(Buffer() As Byte) As Boolean
    IsPNG = (Buffer(0) = 137) And (Buffer(1) = 80) And (Buffer(2) = 78) And (Buffer(3) = 71) And _
            (Buffer(4) = 13) And (Buffer(5) = 10) And (Buffer(6) = 26) And (Buffer(7) = 10)
End Function

'===============================================================================
' Name: MAP_Unload
' Input:
'   ByVal Map As Long - Identificador del grafico.
' Output:
' Purpose: Descarga un grafico en memoria.
' Remarks:
'===============================================================================
Public Sub MAP_Unload(Map As Long)
On Error GoTo ErrOut

    If Map = UBound(Texture) Then
        If Map = 0 Then
            Erase Texture
        
        Else
            ReDim Preserve Texture(UBound(Texture) - 1) As D3D_Texture
        
        End If
    
    Else
        Set Texture(Map).Texture = Nothing
        
        With Texture(Map)
            .Image_Height = 0
            .Image_Width = 0
            .Texture_Height = 0
            .Texture_Width = 0
            .ColorKey = 0
            
        End With

    End If

ErrOut:
End Sub

'===============================================================================
' Name: MAP_GetInfo
' Input:
'   ByVal Map As Long - Identificador del grafico.
'   ByVal Info As GFX_Info - Parámetro de salida que nos devuelve la información del grafico.
' Output:
' Purpose: Devuelve información sobre un grafico.
' Remarks:
'===============================================================================
Public Sub MAP_GetInfo(Map As Long, Info As GFX_Info)
On Error GoTo ErrOut

With Info
    .ColorKey = Texture(Map).ColorKey
    .Height = Texture(Map).Texture_Height
    .Width = Texture(Map).Texture_Width
    .Image_Height = Texture(Map).Image_Height
    .Image_Width = Texture(Map).Image_Width
    .Size = Texture(Map).Size
    
End With
    
ErrOut:

End Sub

'===============================================================================
' Name: MAP_SetRegion
' Input:
'   ByVal Map As Long - Identificador del grafico.
'   ByVal Region As GFX_Rect - Región del grafico que dibujaremos.
' Output:
' Purpose: Prepara una región del grafico para ser dibujada con DRAW_Map, DRAW_MapEx y DRAW_AdvMap.
' Remarks: Este procedimiento debe llamarse siempre antes de DRAW_Map, DRAW_MapEx o DRAW_AdvMap para que funcione.
'===============================================================================
Public Sub MAP_SetRegion(Map As Long, Region As GFX_Rect)
On Local Error Resume Next

Call Texture(Map).Texture.GetLevelDesc(0, D3D_Desc)

With Region
    If .X < 0 Then .X = 0 Else If .X > D3D_Desc.Width Then .X = D3D_Desc.Width
    If .Y < 0 Then .Y = 0 Else If .Y > D3D_Desc.Height Then .Y = D3D_Desc.Height
    If .Height < 1 Or .Height > D3D_Desc.Height Then .Height = D3D_Desc.Height
    If .Width < 1 Or .Width > D3D_Desc.Width Then .Width = D3D_Desc.Width
    
    FrameRect.Left = .X
    FrameRect.Top = .Y
    FrameRect.Right = .X + .Width
    FrameRect.bottom = .Y + .Height
    
End With

FrameSet = True

End Sub

'===============================================================================
' Name: MAP_CopyRects
' Input:
'   ByVal SrcMap As Long - Identificador del grafico origen.
'   ByVal DestMap As Long - Identificador del grafico destino.
'   ByVal SrcRect() As GFX_Rect - Lista de regiones que se copiaran del grafico origen al grafico destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
'   ByVal DestPixel() As Vertex - Lista de coordenadas donde se copiaran las regiones en el grafico destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Copia una serie de regiones de un grafico en otro.
' Remarks: No se respeta el ColorKey de los gráficos en esta operación.
'===============================================================================
Public Function MAP_CopyRects(SrcMap As Long, DestMap As Long, SrcRect() As GFX_Rect, DestPixel() As Vertex) As Boolean
On Error GoTo ErrOut

Dim rctSource() As RECT
Dim ptDest() As point
Dim tmpSurf(1) As Direct3DSurface8
Dim i As Long

If D3D_Init Then

    Set tmpSurf(0) = Texture(SrcMap).Texture.GetSurfaceLevel(0)
    Set tmpSurf(1) = Texture(DestMap).Texture.GetSurfaceLevel(0)
    
    ReDim rctSource(UBound(SrcRect)) As RECT
    ReDim pdest(UBound(DestPixel)) As point
    
    For i = 0 To UBound(rctSource)
        With rctSource(i)
            .Left = SrcRect(i).X
            .Top = SrcRect(i).Y
            .Right = SrcRect(i).Width + SrcRect(i).X
            .bottom = SrcRect(i).Height + SrcRect(i).Y
        End With
            
        With ptDest(i)
            .X = DestPixel(i).X
            .Y = DestPixel(i).Y
        
        End With
        
        'DoEvents
        
    Next i

    Call d3dDevice.CopyRects(tmpSurf(0), rctSource(0), i + 1, tmpSurf(1), ptDest(0))

    Set tmpSurf(0) = Nothing
    Set tmpSurf(1) = Nothing

    MAP_CopyRects = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: MAP_AdvCopyRects
' Input:
'   ByVal SrcMap As Long - Identificador del grafico origen.
'   ByVal DestMap As Long - Identificador del grafico destino.
'   ByVal SrcRect As GFX_Rect - Región que se copiara del grafico origen al grafico destino.
'   ByVal DestRect As GFX_Rect - Región donde se copiara la región del grafico origen.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Versión avanzada de MAP_CopyRects. Copia una región de un grafico en otro con opción de reescalado.
' Remarks: La región destino indica las dimensiones que tendrá la región origen cuando se pinte sobre el grafico destino. No se respeta el ColorKey de los gráficos en esta operación.
'===============================================================================
Public Function MAP_AdvCopyRects(SrcMap As Long, DestMap As Long, SrcRect As GFX_Rect, DestRect As GFX_Rect) As Boolean
On Error GoTo ErrOut

Dim rct(1) As RECT
Dim Palette(1) As PALETTEENTRY
Dim tmpSurf(1) As Direct3DSurface8

If D3D_Init Then

    Set tmpSurf(0) = Texture(SrcMap).Texture.GetSurfaceLevel(0)
    Set tmpSurf(1) = Texture(DestMap).Texture.GetSurfaceLevel(0)

    With rct(0)
        .Left = SrcRect.X
        .Top = SrcRect.Y
        .Right = SrcRect.Width + SrcRect.X
        .bottom = SrcRect.Height + SrcRect.Y
    End With

    With rct(1)
        .Left = DestRect.X
        .Top = DestRect.Y
        .Right = DestRect.Width + DestRect.X
        .bottom = DestRect.Height + DestRect.Y
    End With

    Call D3DX.LoadSurfaceFromSurface(tmpSurf(1), ByVal 0, rct(1), tmpSurf(0), ByVal 0, rct(0), D3DX_DEFAULT, 0)

    Set tmpSurf(0) = Nothing
    Set tmpSurf(1) = Nothing

    MAP_AdvCopyRects = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: MAP_CopyRectsToSurf
' Input:
'   ByVal SrcMap As Long - Identificador del grafico origen.
'   ByVal DestSurf As Long - Identificador de la superficie destino.
'   ByVal SrcRect() As GFX_Rect - Lista de regiones que se copiaran del grafico origen a la superficie destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
'   ByVal DestPixel() As Vertex - Lista de coordenadas donde se copiaran las regiones en la superficie destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Copia una serie de regiones de un grafico en una superficie.
' Remarks: No se respeta el ColorKey de los gráficos en esta operación.
'===============================================================================
Public Function MAP_CopyRectsToSurf(SrcMap As Long, DestSurf As Long, SrcRect() As GFX_Rect, DestPixel() As Vertex) As Boolean
On Error GoTo ErrOut

Dim rctSource() As RECT
Dim ptDest() As point
Dim tmpSurf(1) As Direct3DSurface8
Dim i As Long

Dim TexSurf As Direct3DSurface8 'Superficie de la textura.

If D3D_Init Then
    Set TexSurf = Texture(SrcMap).Texture.GetSurfaceLevel(0)
    
    ReDim rctSource(UBound(SrcRect)) As RECT
    ReDim ptDest(UBound(DestPixel)) As point
    
    For i = 0 To UBound(rctSource)
        With rctSource(i)
            .Left = SrcRect(i).X
            .Top = SrcRect(i).Y
            .Right = SrcRect(i).Width + SrcRect(i).X
            .bottom = SrcRect(i).Height + SrcRect(i).Y
        End With
            
        With ptDest(i)
            .X = DestPixel(i).X
            .Y = DestPixel(i).Y
        
        End With
        
        'DoEvents
        
    Next i

    Call d3dDevice.CopyRects(TexSurf, rctSource(0), 1, Surface(DestSurf).Surface, ptDest(0))
    
    Set TexSurf = Nothing

    MAP_CopyRectsToSurf = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: MAP_AdvCopyRectsToSurf
' Input:
'   ByVal SrcMap As Long - Identificador del grafico origen.
'   ByVal DestSurf As Long - Identificador de la superficie destino.
'   ByVal SrcRect As GFX_Rect - Región que se copiara del grafico origen a la superficie destino.
'   ByVal DestRect As GFX_Rect - Región donde se copiara la región del grafico origen.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Versión avanzada de MAP_CopyRectsToSurf. Copia un grafico o parte de el en una superficie con posibilidad de reescalado.
' Remarks: La región destino indica las dimensiones que tendrá la región origen cuando se pinte sobre la superficie destino. No se respeta el ColorKey de los gráficos en esta operación.
'===============================================================================
Public Function MAP_AdvCopyRectsToSurf(SrcMap As Long, DestSurf As Long, SrcRect As GFX_Rect, DestRect As GFX_Rect) As Boolean
On Error GoTo ErrOut
Dim rct(1) As RECT
Dim TexSurf As Direct3DSurface8 'Superficie de la textura.

If D3D_Init Then
    Set TexSurf = Texture(SrcMap).Texture.GetSurfaceLevel(0)

    With rct(0)
        .Left = SrcRect.X
        .Top = SrcRect.Y
        .Right = SrcRect.Width + SrcRect.X
        .bottom = SrcRect.Height + SrcRect.Y
    End With

    With rct(1)
        .Left = DestRect.X
        .Top = DestRect.Y
        .Right = DestRect.Width + DestRect.X
        .bottom = DestRect.Height + DestRect.Y
    End With

    Call D3DX.LoadSurfaceFromSurface(Surface(DestSurf).Surface, ByVal 0, rct(1), TexSurf, ByVal 0, rct(0), D3DX_DEFAULT, 0)

    Set TexSurf = Nothing

    MAP_AdvCopyRectsToSurf = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: SURF_Create
' Input:
'   ByVal Width As Long - Anchura en pixeles.
'   ByVal Height As Long - Altura en pixeles.
'   ByVal GrayScale As Boolean - Define si la superficie utilizara escala de grises.
' Output:
'   Long - Devuelve el código identificador para la superficie en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede crear la superficie se devuelve GFX_NOTCREATED.
' Purpose: Crea una superficie vacia en memoria.
' Remarks: Se pueden crear tantas superficies como se deseen teniendo como limite la memoria del dispositivo grafico.
'===============================================================================
Public Function SURF_Create(Width As Long, Height As Long, Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long
Dim Format As CONST_D3DFORMAT

If D3D_Init Then
    i = -1
    i = Get_FreeID(1)

    If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8
    
    Set Surface(i).Surface = d3dDevice.CreateImageSurface(Width, Height, Format)

    Surface(i).Width = Width
    Surface(i).Height = Height
        
    SURF_Create = i
    
End If
Exit Function

ErrOut:
If Not i = -1 Then Call Me.SURF_Unload(i)
SURF_Create = GFX_NOTCREATED

End Function

'===============================================================================
' Name: SURF_Load
' Input:
'   ByVal FileName As String - Nombre y ruta de la imagen.
'   ByVal GrayScale As Boolean - Define si la imagen se cargara en escala de grises.
' Output:
'   Long - Devuelve el código identificador para la superficie en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar la superficie se devuelve GFX_FILENOTLOAD.
' Purpose: Carga un archivo de imagen desde disco como una superficie.
' Remarks: Esta función solo soporta archivos de Mapas de Bits de Windows *.BMP. Se pueden crear tantas superficies como se deseen teniendo como limite la memoria del dispositivo grafico.
'===============================================================================
Public Function SURF_Load(Filename As String, Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long, Id As Long, Width As Long, Height As Long
Dim bmFh As BITMAPFILEHEADER
Dim bmIh As BITMAPINFOHEADER
Dim Format As CONST_D3DFORMAT

If D3D_Init And Global_Mod.FileExists(Filename) Then

    'Leemos datos BMP:
    Open Filename For Binary Access Read As #1
        Get #1, , bmFh
        Get #1, Len(bmFh) + 1, bmIh
    Close #1
    
    'Si se trata de un Mapa de Bits de Windows...
    If bmFh.bfType = 19778 Then
        i = -1
        i = Get_FreeID(1)
        Surface(i).Width = bmIh.biWidth
        Surface(i).Height = bmIh.biHeight
        
        If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8
        
        Set Surface(i).Surface = d3dDevice.CreateImageSurface(Surface(i).Width, Surface(i).Height, Format)
        Call D3DX.LoadSurfaceFromFile(Surface(i).Surface, ByVal 0, ByVal 0, Filename, ByVal 0, D3DX_FILTER_TRIANGLE Or D3DX_FILTER_MIRROR, &HFF000000, ByVal 0)
        
        'Obtenemos los valores de la textura en memoria:
        Call Surface(i).Surface.GetDesc(D3D_Desc)
        Surface(i).Height = D3D_Desc.Height
        Surface(i).Width = D3D_Desc.Width
        Surface(i).Size = D3D_Desc.Size
        
        SURF_Load = i
        Exit Function
    End If
End If

ErrOut:
Close #1
If Not i = -1 Then Call Me.SURF_Unload(i)
SURF_Load = GFX_FILENOTLOAD

End Function

'===============================================================================
' Name: SURF_LoadFromMemory
' Input:
'   ByVal FileData() As Byte - Array de datos que contiene la información del archivo. El argumento se debe pasar apuntando al primer vector del array, el 0.
'   ByVal GrayScale As Boolean - Define si la imagen se cargara en escala de grises.
' Output:
'   Long - Devuelve el código identificador para la superficie en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar la superficie se devuelve -1.
' Purpose: Carga un archivo de imagen alojado en memoria como una superficie.
' Remarks: Esta función solo soporta archivos Mapas de Bits de Windows *.BMP. Se pueden crear tantas superficies como se deseen teniendo como limite la memoria del dispositivo grafico.
'===============================================================================
Public Function SURF_LoadFromMemory(FileData() As Byte, Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long
Dim Format As CONST_D3DFORMAT

If D3D_Init Then
    i = -1
    i = Get_FreeID(1)

    If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8
    
    Set Surface(i).Surface = d3dDevice.CreateImageSurface(Surface(i).Width, Surface(i).Height, Format)

    Call D3DX.LoadSurfaceFromFileInMemory(Surface(i).Surface, ByVal 0, ByVal 0, FileData(0), UBound(FileData), ByVal 0, D3DX_FILTER_TRIANGLE Or D3DX_FILTER_MIRROR, &HFF000000, ByVal 0)

    'Obtenemos los valores de la textura en memoria:
    Call Surface(i).Surface.GetDesc(D3D_Desc)
    Surface(i).Width = D3D_Desc.Width
    Surface(i).Height = D3D_Desc.Height
    Surface(i).Size = D3D_Desc.Size
    
    SURF_LoadFromMemory = i

    Exit Function

End If

ErrOut:
If Not i = -1 Then Call Me.SURF_Unload(i)
SURF_LoadFromMemory = -1

End Function

'===============================================================================
' Name: SURF_Unload
' Input:
'   ByVal Surf As Long - Identificador de la superficie.
' Output:
' Purpose: Descarga una superficie de la memoria.
' Remarks:
'===============================================================================
Public Sub SURF_Unload(Surf As Long)
On Error GoTo ErrOut

    If Surf = UBound(Surface) Then
        If Surf = 0 Then
            Erase Surface
        
        Else
            ReDim Preserve Surface(UBound(Surface) - 1) As D3D_Surface
        
        End If
    
    Else
        Set Surface(Surf).Surface = Nothing
        
        With Surface(Surf)
            .Height = 0
            .Width = 0
            
        End With

    End If

ErrOut:
End Sub

'===============================================================================
' Name: SURF_Save
' Input:
'   ByVal Surf As Long - Identificador de la superficie.
'   ByVal FileName As String - Nombre y ruta donde se creara el archivo.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Guarda una superficie como archivo de imagen.
' Remarks: La superficie se guarda en formato Mapa de Bits de Windows *.BMP.
'===============================================================================
Public Function SURF_Save(Surf As Long, Filename As String) As Boolean
On Error GoTo ErrOut

Dim Pal As PALETTEENTRY
Dim rct As RECT

If D3D_Init Then

    With rct
        .Left = 0
        .Top = 0
        .Right = Surface(Surf).Width
        .bottom = Surface(Surf).Height
    End With

    Call D3DX.SaveSurfaceToFile(Filename, D3DXIFF_BMP, Surface(Surf).Surface, Pal, rct)
                       
    SURF_Save = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: SURF_ScreenCapture
' Input:
'   ByVal GrayScale As Boolean - Define si la superficie utilizara escala de grises.
' Output:
'   Long - Devuelve el identificador de la superficie en memoria donde esta almacenada la captura.
' Purpose: Genera una captura del buffer de pantalla.
' Remarks: La función genera automáticamente una superficie nueva en memoria para almacenar la captura. En modo ventana solo se realiza la captura del área de cliente de la ventana, sin los bordes ni la barra de titulo de la ventana.
'===============================================================================
Public Function SURF_ScreenCapture(Optional GrayScale As Boolean) As Long
On Error GoTo ErrOut

Dim oSurface As Direct3DSurface8
Dim Pal As PALETTEENTRY
Dim SrcRect As RECT
Dim DispMode As D3DDISPLAYMODE
Dim ClientRect As POINTAPI

Dim Surf As Long
Dim Format As CONST_D3DFORMAT

If D3D_Init Then
    'Creamos la superficie donde almacenaremos la captura:
    Surf = SURF_Create(D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, GrayScale)
    
    If Surf < 0 Then GoTo ErrOut
    
    If Not Windowed Then 'FullScreen:
    
        'Leemos el contenido de la pantalla (FrontBuffer) y lo guardamos en la superficie:
        d3dDevice.GetFrontBuffer Surface(Surf).Surface
    
    Else 'Windowed:
        'Obtenemos las dimensiones de la resolucion de pantalla del adaptador:
        d3dDevice.GetDisplayMode DispMode
        
        If GrayScale Then Format = D3DFMT_A4L4 Else Format = D3DFMT_A8R8G8B8
        
        'Creamos una superficie con la resolucion seleccionada del adaptador grafico a
        '32 Bit de color por que la funcion siempre devuelve el resultado a 32 Bits (D3DFMT_A8R8G8B8):
        Set oSurface = d3dDevice.CreateImageSurface(DispMode.Width, _
                DispMode.Height, _
                Format)

        'Leemos el contenido de la pantalla (FrontBuffer) y lo guardamos en la superficie:
        d3dDevice.GetFrontBuffer oSurface
        
        'En modo ventana la captura es de toda la pantalla, escritorio incluido, asi que calculamos
        'la posicion y las dimensiones del area de cliente de la ventana donde dibujamos para salvar
        'unicamente el contenido del programa:
        
        Call GetClientRect(Me.hWnd, SrcRect)
        Call ClientToScreen(Me.hWnd, ClientRect)
    
        With SrcRect
            .Left = ClientRect.X
            .Right = .Left + .Right
            .Top = ClientRect.Y
            .bottom = .Top + .bottom
        End With
        
        'Copiamos la seccion de captura del programa a la superficie en memoria:
        Call D3DX.LoadSurfaceFromSurface(Surface(Surf).Surface, Pal, ByVal 0, oSurface, Pal, SrcRect, D3DX_DEFAULT, 0)
   
    End If
    
    SURF_ScreenCapture = Surf
    
End If
Exit Function

ErrOut:
SURF_ScreenCapture = -1
End Function

'===============================================================================
' Name: SURF_CopyRects
' Input:
'   ByVal SrcSurf As Long - Identificador de la superficie origen.
'   ByVal DestSurf As Long - Identificador de la superficie destino.
'   ByVal SrcRect() As GFX_Rect - Lista de regiones que se copiaran de la superficie origen a la superficie destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
'   ByVal DestPixel() As Vertex - Lista de coordenadas donde se copiaran las regiones en la superficie destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Copia una serie de regiones de una superficie en otra.
' Remarks:
'===============================================================================
Public Function SURF_CopyRects(SrcSurf As Long, DestSurf As Long, SrcRect() As GFX_Rect, DestPixel() As Vertex) As Boolean
On Error GoTo ErrOut

Dim rctSource() As RECT
Dim ptDest() As point
Dim tmpSurf(1) As Direct3DSurface8
Dim i As Long

If D3D_Init Then
    ReDim rctSource(UBound(SrcRect)) As RECT
    ReDim pdest(UBound(DestPixel)) As point
    
    For i = 0 To UBound(rctSource)
        With rctSource(i)
            .Left = SrcRect(i).X
            .Top = SrcRect(i).Y
            .Right = SrcRect(i).Width + SrcRect(i).X
            .bottom = SrcRect(i).Height + SrcRect(i).Y
        End With
            
        With ptDest(i)
            .X = DestPixel(i).X
            .Y = DestPixel(i).Y
        
        End With
        
        'DoEvents
        
    Next i

    Call d3dDevice.CopyRects(Surface(SrcSurf).Surface, rctSource(0), i + 1, Surface(DestSurf).Surface, ptDest(0))

    SURF_CopyRects = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: SURF_AdvCopyRects
' Input:
'   ByVal SrcSurf As Long - Identificador de la superficie origen.
'   ByVal DestSurf As Long - Identificador de la superficie destino.
'   ByVal SrcRect As GFX_Rect - Región que se copiara de la superficie origen a la superficie destino.
'   ByVal DestRect As GFX_Rect - Región donde se copiara la región de la superficie origen.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Versión avanzada de SURF_CopyRects. Copia una región de una superficie en otro con opción de reescalado.
' Remarks: La región destino indica las dimensiones que tendrá la región origen cuando se pinte sobre la superficie destino.
'===============================================================================
Public Function SURF_AdvCopyRects(SrcSurf As Long, DestSurf As Long, SrcRect As GFX_Rect, DestRect As GFX_Rect) As Boolean
On Error GoTo ErrOut
Dim rct(1) As RECT

If D3D_Init Then
    With rct(0)
        .Left = SrcRect.X
        .Top = SrcRect.Y
        .Right = SrcRect.Width + SrcRect.X
        .bottom = SrcRect.Height + SrcRect.Y
    End With

    With rct(1)
        .Left = DestRect.X
        .Top = DestRect.Y
        .Right = DestRect.Width + DestRect.X
        .bottom = DestRect.Height + DestRect.Y
    End With

    Call D3DX.LoadSurfaceFromSurface(Surface(DestSurf).Surface, ByVal 0, rct(1), Surface(SrcSurf).Surface, ByVal 0, rct(0), D3DX_DEFAULT, 0)

    SURF_AdvCopyRects = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: Surf_GetInfo
' Input:
'   ByVal Surf As Long - Identificador de la superficie.
'   ByVal Info As GFX_Info - Parámetro de salida que nos devuelve la información de la superficie.
' Output:
' Purpose: Devuelve información sobre una superficie.
' Remarks:
'===============================================================================
Public Sub SURF_GetInfo(Surf As Long, Info As GFX_Info)
On Error GoTo ErrOut

Call Surface(Surf).Surface.GetDesc(D3D_Desc)

With Info
    .ColorKey = 0
    .Height = D3D_Desc.Height
    .Width = D3D_Desc.Width
    .Image_Height = .Height
    .Image_Width = .Width
    .Size = D3D_Desc.Size
    
End With
    
ErrOut:

End Sub

'===============================================================================
' Name: SURF_GetPixelData
' Input:
'   ByVal Surf As Long - Identificador de la superficie que vamos a leer.
' Output:
'   Long - Devuelve el código identificador para esta lista de píxeles en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede copiar la lista de píxeles se devuelve -1.
' Purpose: Crea una copia en memoria de la lista de los píxeles de una superficie que contiene el valor ARGB de cada píxel.
' Remarks: Las copias realizadas con esta función sirven para poder leer y/o escribir los valores de los colores de los píxeles de una superficie. Se pueden realizar tantas copias como se precisen.
'===============================================================================
Public Function SURF_GetPixelData(Surf As Long) As Long
On Error GoTo ErrOut

Dim i As Long, pData As D3DLOCKED_RECT, pxArr() As Byte

If D3D_Init Then

    'Bloqueamos la superficie:
    Call Surface(Surf).Surface.LockRect(pData, ByVal 0, 0)

        i = Get_FreeID(3)
        With Pixel_Array(i)

            .Surf = Surf
            .Height = Surface(Surf).Height
            .Pitch = pData.Pitch

            'Redimensionamos el array para almacenar los varlores de los pixeles:
            ReDim pxArr(pData.Pitch * Surface(Surf).Height)
            ReDim .BGRA(UBound(pxArr) \ 4)

            'Copiamos la informacion en el array:
            If DXCopyMemory(pxArr(0), ByVal pData.pBits, pData.Pitch * Surface(Surf).Height) = D3D_OK Then
                ' Convertimos el array de grupos de 4 bytes a enteros de 32 bits (Long):
                Call Global_Mod.CopyMem(.BGRA(0), pxArr(0), UBound(pxArr) + 1)
                SURF_GetPixelData = i
            End If
        End With

    'Desbloqueamos la superficie:
    Call Surface(Surf).Surface.UnlockRect

End If

Exit Function

ErrOut:
SURF_GetPixelData = GFX_DATANOTREAD

End Function

'===============================================================================
' Name: SURF_ReadPixelArray
' Input:
'   ByVal Id As Long - Identificador de la lista de píxeles de la superficie.
'   ByRef pxArr() As Long - Array de enteros de 32 bits que representan los valores ARGB de los pixeles.
' Output:
' Purpose: Lee el array de píxeles de una lista de pixeles existente.
' Remarks: Utilice este metodo para obtener un array con los valores ARGB de los pixeles para realizar modificaciones en masa.
'===============================================================================
Public Sub SURF_ReadPixelArray(Id As Long, pxArr() As Long)
    ReDim pxArr(UBound(Pixel_Array(Id).BGRA))
    pxArr = Pixel_Array(Id).BGRA
End Sub

'===============================================================================
' Name: SURF_WritePixelArray
' Input:
'   ByVal Id As Long - Identificador de la lista de píxeles de la superficie.
'   ByRef pxArr() As Long - Array de enteros de 32 bits que representan los valores ARGB de los pixeles.
' Output: Devuelve false si array contiene un numero diferente de pixeles que la lista original.
' Purpose: Escribe el array de píxeles de una lista de pixeles existente.
' Remarks: Utilice este metodo para modificar los valores de los pixeles de una superficie de forma rapida y directa.
'===============================================================================
Public Function SURF_WritePixelArray(Id As Long, pxArr() As Long) As Boolean
    If (UBound(pxArr) = UBound(Pixel_Array(Id).BGRA)) Then
        Pixel_Array(Id).BGRA = pxArr
        SURF_WritePixelArray = True
    End If
End Function

'===============================================================================
' Name: SURF_DeletePixelData
' Input:
'   ByVal Id As Long - Identificador de la lista de píxeles de la superficie.
' Output:
' Purpose: Elimina una lista de píxeles de una superficie de la memoria.
' Remarks:
'===============================================================================
Public Sub SURF_DeletePixelData(Id As Long)
On Error GoTo ErrOut

    If Id = UBound(Pixel_Array) Then
        If Id = 0 Then
            Erase Pixel_Array

        Else
            ReDim Preserve Pixel_Array(UBound(Pixel_Array) - 1) As GFX_PixelData

        End If

    Else
        With Pixel_Array(Id)
            .Free = True

            .Surf = -1
            .Height = 0
            .Pitch = 0

            Erase .BGRA

        End With

    End If

ErrOut:

End Sub

'===============================================================================
' Name: SURF_SetPixelData
' Input:
'   ByVal Id As Long - Identificador de la lista de píxeles de la superficie.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con exito.
' Purpose: Aplica los valores de una lista de píxeles en la superficie a la que pertenece.
' Remarks: Esta función modifica los valores de los píxeles de la superficie por los contenidos en la copia en memoria.
'===============================================================================
Public Function SURF_SetPixelData(Id As Long) As Boolean
On Error GoTo ErrOut

Dim pData As D3DLOCKED_RECT

If D3D_Init Then

    With Pixel_Array(Id)

        'Bloqueamos la superficie:
        Call Surface(.Surf).Surface.LockRect(pData, ByVal 0, 0)
        
        'Copiamos los datos del array de pixeles en la superficie:
        SURF_SetPixelData = (DXCopyMemory(ByVal pData.pBits, .BGRA(0), pData.Pitch * Surface(.Surf).Height) = D3D_OK)
        
        'Desbloqueamos la superficie:
        Call Surface(.Surf).Surface.UnlockRect

    End With

End If

ErrOut:
End Function

'===============================================================================
' Name: SURF_GetPixel
' Input:
'   ByVal Id As Long - Identificador de la lista de píxeles de la superficie.
'   ByVal X As Long - Coordenada horizontal de lectura.
'   ByVal Y As Long - Coordenada vertical de lectura.
' Output:
'   Long - Devuelve el valor ARGB del píxel leído.
' Purpose: Lee un píxel de una lista de píxeles en memoria.
' Remarks: Para poder leer un píxel de una superficie primero se ha de crear una lista de píxeles con los valores de los píxeles de la superficie original.
'===============================================================================
Public Function SURF_GetPixel(Id As Long, X As Long, Y As Long) As Long
On Local Error Resume Next

SURF_GetPixel = Pixel_Array(Id).BGRA((Surface(Pixel_Array(Id).Surf).Width * Y) + X)

End Function

'===============================================================================
' Name: SURF_PutPixel
' Input:
'   ByVal PixelData As Long - Identificador de la lista de píxeles de la superficie.
'   ByVal X As Long - Coordenada horizontal de lectura.
'   ByVal Y As Long - Coordenada vertical de lectura.
'   ByVal Color As Long - Color ARGB que se aplicara para realizar la operación de dibujo.
' Purpose: Modifica el valor de un píxel de una lista de píxeles en memoria.
' Remarks: Para poder modificar un píxel de una superficie primero se ha de crear una lista de píxeles con los valores de los píxeles de la superficie original.
'===============================================================================
Public Sub SURF_PutPixel(Id As Long, X As Long, Y As Long, Color As Long)
On Error Resume Next

Pixel_Array(Id).BGRA((Surface(Pixel_Array(Id).Surf).Width * Y) + X) = Color

End Sub

'===============================================================================
' Name: SURF_CopyRectsToMap
' Input:
'   ByVal SrcSurf As Long - Identificador de la superficie origen.
'   ByVal DestMap As Long - Identificador del grafico destino.
'   ByVal SrcRect() As GFX_Rect - Lista de regiones que se copiaran de la superficie origen al grafico destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
'   ByVal DestPixel() As Vertex - Lista de coordenadas donde se copiaran las regiones en el grafico destino. El argumento se debe pasar apuntando al primer vector del array, el 0.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Copia una serie de regiones de una superficie en un grafico.
' Remarks:
'===============================================================================
Public Function SURF_CopyRectsToMap(SrcSurf As Long, DestMap As Long, SrcRect() As GFX_Rect, DestPixel() As Vertex) As Boolean
On Error GoTo ErrOut

Dim rctSource() As RECT
Dim ptDest() As point
Dim tmpSurf(1) As Direct3DSurface8
Dim i As Long

Dim TexSurf As Direct3DSurface8 'Superficie de la textura.

If D3D_Init Then
    Set TexSurf = Texture(DestMap).Texture.GetSurfaceLevel(0)
    
    ReDim rctSource(UBound(SrcRect)) As RECT
    ReDim ptDest(UBound(DestPixel)) As point
    
    For i = 0 To UBound(rctSource)
        With rctSource(i)
            .Left = SrcRect(i).X
            .Top = SrcRect(i).Y
            .Right = SrcRect(i).Width + SrcRect(i).X
            .bottom = SrcRect(i).Height + SrcRect(i).Y
        End With
            
        With ptDest(i)
            .X = DestPixel(i).X
            .Y = DestPixel(i).Y
        
        End With
        
        'DoEvents
        
    Next i

    Call d3dDevice.CopyRects(Surface(SrcSurf).Surface, rctSource(0), 1, TexSurf, ptDest(0))
    
    Set TexSurf = Nothing

    SURF_CopyRectsToMap = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: SURF_AdvCopyRectsToMap
' Input:
'   ByVal SrcSurf As Long - Identificador de la superficie origen.
'   ByVal DestMap As Long - Identificador del grafico destino.
'   ByVal SrcRect As GFX_Rect - Región que se copiara de la superficie origen al grafico destino.
'   ByVal DestRect As GFX_Rect - Región donde se copiara la región del grafico origen.
' Output:
'   Boolean - Devuelve verdadero si la operación se ha realizado con éxito.
' Purpose: Versión avanzada de SURF_CopyRectsToSurf. Copia una superficie o parte de ella en un grafico con posibilidad de reescalado.
' Remarks: La región destino indica las dimensiones que tendrá la región origen cuando se pinte sobre el grafico destino.
'===============================================================================
Public Function SURF_AdvCopyRectsToMap(SrcSurf As Long, DestMap As Long, SrcRect As GFX_Rect, DestRect As GFX_Rect) As Boolean
On Error GoTo ErrOut
Dim rct(1) As RECT
Dim TexSurf As Direct3DSurface8 'Superficie de la textura.

If D3D_Init Then
    Set TexSurf = Texture(DestMap).Texture.GetSurfaceLevel(0)

    With rct(0)
        .Left = SrcRect.X
        .Top = SrcRect.Y
        .Right = SrcRect.Width + SrcRect.X
        .bottom = SrcRect.Height + SrcRect.Y
    End With

    With rct(1)
        .Left = DestRect.X
        .Top = DestRect.Y
        .Right = DestRect.Width + DestRect.X
        .bottom = DestRect.Height + DestRect.Y
    End With

    Call D3DX.LoadSurfaceFromSurface(TexSurf, ByVal 0, rct(1), Surface(SrcSurf).Surface, ByVal 0, rct(0), D3DX_DEFAULT, 0)

    Set TexSurf = Nothing

    SURF_AdvCopyRectsToMap = True
    
End If
Exit Function

ErrOut:
End Function

'===============================================================================
' Name: FONT_LoadSystemFont
' Input:
'   ByVal Font As String - Nombre de la fuente de texto.
'   ByVal Size As Long - Tamaño con el se cargara la fuente.
'   ByVal Bold As Boolean - Indica si se cargara la fuente con estilo Negrita.
'   ByVal Italic As Boolean - Indica si se cargara la fuente con estilo Cursiva.
'   ByVal Underline As Boolean - Indica si se cargara la fuente con estilo Subrayado.
'   ByVal Strikethrough As Boolean - Indica si se cargara la fuente con estilo Tachado.
' Output:
'   Long - Devuelve el código identificador para esta fuente en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar la fuente se devuelve GFX_FILENOTLOAD.
' Purpose: Carga una fuente de texto TrueType de Windows y la prepara para ser usada con DRAW_Text.
' Remarks: Para poder cargar una fuente de texto TrueType de Windows esta debe estar instalada en el sistema, en el directorio Fonts de Windows (C:\Windows\Fonts por ejemplo) Se pueden cargar tantas fuentes de texto como se deseen teniendo como limite la memoria del sistema.
'===============================================================================
Public Function FONT_LoadSystemFont(Font As String, Size As Long, Bold As Boolean, Italic As Boolean, UnderLine As Boolean, Strikethrough As Boolean) As Long
On Error GoTo ErrOut
Dim i As Long

If D3D_Init Then
    i = -1
    i = Get_FreeID(5)
    
    With D3DFont(i).Std_Font
        .Name = Font
        .Size = Size
        .Bold = Bold
        .Italic = Italic
        .UnderLine = UnderLine
        .Strikethrough = Strikethrough
    
    End With
    
    With D3DFont(i)
        Set .I_Font = .Std_Font
    
        Set .D3D_Font = D3DX.CreateFont(d3dDevice, .I_Font.hFont)
    
    End With
    
    FONT_LoadSystemFont = i

    Exit Function

End If

ErrOut:
If Not i = -1 Then Call Me.FONT_UnloadSystemFont(i)
FONT_LoadSystemFont = GFX_FILENOTLOAD

End Function

'===============================================================================
' Name: FONT_LoadSystemFontFromFile
' Input:
'   ByVal FileName As String - Nombre del archivo de fuente de texto.
'   ByVal FontName As String - Devuelve el nombre de la fuente de texto.
'   ByVal Size As Long - Tamaño con el se cargara la fuente.
'   ByVal Bold As Boolean - Indica si se cargara la fuente con estilo Negrita.
'   ByVal Italic As Boolean - Indica si se cargara la fuente con estilo Cursiva.
'   ByVal Underline As Boolean - Indica si se cargara la fuente con estilo Subrayado.
'   ByVal Strikethrough As Boolean - Indica si se cargara la fuente con estilo Tachado.
' Output:
'   Long - Devuelve el código identificador para esta fuente en memoria, un valor comprendido entre 0 y el máximo permitido por el tipo Long. Si no se puede cargar la fuente se devuelve GFX_FILENOTLOAD.
' Purpose: Carga una fuente de texto TrueType de Windows desde archivo y la prepara para ser usada con DRAW_Text.
' Remarks: Esta funcion permite importar una fuente de texto TrueType de Windows desde archivo y añadirla a la lista de fuentes del sistema. La fuente, una vez Se pueden cargar tantas fuentes de texto como se deseen teniendo como limite la memoria del sistema.
'===============================================================================
Public Function FONT_LoadSystemFontFromFile(Filename As String, FontName As String, Size As Long, Bold As Boolean, Italic As Boolean, UnderLine As Boolean, Strikethrough As Boolean) As Long
    On Error GoTo ErrOut
    Dim fName As String
    Dim cfont As New CFontPreview
    
    If D3D_Init Then
        cfont.FontFile = Filename
        'If cfont. Then
            FONT_LoadSystemFontFromFile = Me.FONT_LoadSystemFont(cfont.FaceName, Size, Bold, Italic, UnderLine, Strikethrough)
            FontName = cfont.FaceName
            Call ImportFont.Add(cfont)
            Exit Function
        
        'End If
    
    End If
    
ErrOut:
    FONT_LoadSystemFontFromFile = GFX_FILENOTLOAD

End Function

'===============================================================================
' Name: Font_UnloadSystemFont
' Input:
'   ByVal Font As Long - Identificador de la fuente.
' Output:
' Purpose: Descarga una fuente en memoria.
' Remarks:
'===============================================================================
Public Sub FONT_UnloadSystemFont(Font As Long)
On Error GoTo ErrOut

    If Font = UBound(D3DFont) Then
        If Font = 0 Then
            Erase D3DFont
        
        Else
            ReDim Preserve D3DFont(UBound(D3DFont) - 1) As System_Font
        
        End If
    
    Else
        
        With D3DFont(Font)
            Set .D3D_Font = Nothing
            Set .Std_Font = Nothing
            Set .I_Font = Nothing
            .Free = True
            
        End With
        
    End If

ErrOut:
End Sub

'===============================================================================
' Name: FONT_SystemGetTextWidth
' Input:
'   ByVal Font As Long - Identificador de la fuente.
'   ByVal Text As String - Cadena de texto a medir.
' Output:
'   Long - Anchura en pixeles de la cadena de texto.
' Purpose: Devuelve la anchura en pixeles de una cadena de texto.
' Remarks: Utilice esta funcion para medir la anchura que ocupara una cadena de texto con una fuente determinada que previamente haya cargado.
'===============================================================================
Public Function FONT_SystemGetTextWidth(Font As Long, Text As String) As Long
Dim TextSize As RECT

'Calculamos el espacio que ocupara el texto en pantalla:
With D3DFont(Font)
    Call D3DX.DrawText(.D3D_Font, 0, Text, TextSize, DT_CALCRECT)

End With

FONT_SystemGetTextWidth = TextSize.Left + TextSize.Right

End Function

'===============================================================================
' Name: FONT_SystemGetTextHeight
' Input:
'   ByVal Font As Long - Identificador de la fuente.
'   ByVal Text As String - Cadena de texto a medir.
' Output:
'   Long - Altura en pixeles de la cadena de texto.
' Purpose: Devuelve la altura en pixeles de una cadena de texto.
' Remarks: Utilice esta funcion para medir la altura que ocupara una cadena de texto con una fuente determinada que previamente haya cargado.
'===============================================================================
Public Function FONT_SystemGetTextHeight(Font As Long, Text As String) As Long
Dim TextSize As RECT

'Calculamos el espacio que ocupara el texto en pantalla:
With D3DFont(Font)
    Call D3DX.DrawText(.D3D_Font, 0, Text, TextSize, DT_CALCRECT)

End With

FONT_SystemGetTextHeight = TextSize.Top + TextSize.bottom

End Function

'===============================================================================
' Name: ARGB_Set
' Input:
'   ByVal Alpha As Integer - Valor para el canal Alpha.
'   ByVal Red As Integer - Valor para el canal Rojo.
'   ByVal Green As Integer - Valor para el canal Verde.
'   ByVal Blue As Integer - Valor para el canal Azul.
' Output:
'   Long - Devuelve un valor ARGB.
' Purpose: Genera un color ARGB a partir de los valores definidos en sus canales.
' Remarks: Esta función ayuda en la tarea de definir valores para los colores pudiendo generar cualquier gama tanto en 16 como en 32 bits color. La función es independiente de la profundidad de color establecida para el modo de video transformando el valor final en uno valido para el modo establecido. También se pueden definir valore ARGB directamente en formato Hexadecimal, como por ejemplo el color ARGB definido según el orden de sus canales 255, 62, 189 y 255 seria en Hexadecimal &HFF3EBDFF.
'===============================================================================
Public Function ARGB_Set(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer) As Long
If Alpha < 0 Then Alpha = 0 Else If Alpha > 255 Then Alpha = 255
If Red < 0 Then Red = 0 Else If Red > 255 Then Red = 255
If Green < 0 Then Green = 0 Else If Green > 255 Then Green = 255
If Blue < 0 Then Blue = 0 Else If Blue > 255 Then Blue = 255

ARGB_Set = D3DColorARGB(Alpha, Red, Green, Blue)

End Function

'===============================================================================
' Name: ARGB_Get
' Input:
'   ByVal Color As Long - Valor ARGB a descomponer.
'   ByVal Data As ARGB - Parámetro de salida que nos devolverá el valor de cada canal del color.
' Output:
' Purpose: Descompone un valor ARGB extrayendo el valor de cada canal.
' Remarks:
'===============================================================================
Public Sub ARGB_Get(Color As Long, Data As ARGB)
    Dim A As Long, r As Long, g As Long, B As Long
        
    If Color < 0 Then
        A = ((Color And (&H7F000000)) / (2 ^ 24)) Or &H80&
    Else
        A = Color / (2 ^ 24)
    End If
    
    r = (Color And &HFF0000) / (2 ^ 16)
    g = (Color And &HFF00&) / (2 ^ 8)
    B = (Color And &HFF&)
    
    With Data
        .Alpha = A
        .Red = r
        .Green = g
        .Blue = B
        
    End With
        
End Sub

Private Function Get_FreeID(Flag As Byte) As Long
On Local Error Resume Next

Dim i As Long, j As Long

Select Case Flag
    Case 0 'Texturas:
        i = UBound(Texture)
        
        If Err.Number = 9 Then
            ReDim Texture(0) As D3D_Texture
            
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If Texture(j).Texture Is Nothing Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve Texture(i + 1) As D3D_Texture
            Get_FreeID = i + 1
            
        End If
                
    Case 1 'Superficies:
        i = UBound(Surface)
        
        If Err.Number = 9 Then
            ReDim Surface(0) As D3D_Surface
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If Surface(j).Surface Is Nothing Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve Surface(i + 1) As D3D_Surface
            Get_FreeID = i + 1
           
        End If
    
    Case 2 'Listas precalculadas para circunferencias:
        i = UBound(PreCircle)
        
        If Err.Number = 9 Then
            ReDim PreCircle(0) As PreCalCircle_Data
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If PreCircle(j).Free Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve PreCircle(i + 1) As PreCalCircle_Data
            Get_FreeID = i + 1
           
        End If
        
    Case 3 'Listas de arrays de pixeles:
        i = UBound(Pixel_Array)
        
        If Err.Number = 9 Then
            ReDim Pixel_Array(0) As GFX_PixelData
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If Pixel_Array(j).Free Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve Pixel_Array(i + 1) As GFX_PixelData
            Get_FreeID = i + 1
           
        End If
    
    Case 5 'Lista de Fuentes de sistema:
        i = UBound(D3DFont)
        
        If Err.Number = 9 Then
            ReDim D3DFont(0) As System_Font
            
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If D3DFont(j).Free Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve D3DFont(i + 1) As System_Font
            Get_FreeID = i + 1
            
        End If
        
    Case 6 'Lista de Render Targets:
        i = UBound(RenderTarget)
        
        If Err.Number = 9 Then
            ReDim RenderTarget(0) As GFX_RenderTarget
            
            Get_FreeID = 0
            
        Else
            For j = 0 To i
                If RenderTarget(j).Target Is Nothing Then
                    Get_FreeID = j
                    Exit For
                    
                End If
                
            Next j
            
            ReDim Preserve RenderTarget(i + 1) As GFX_RenderTarget
            Get_FreeID = i + 1
            
        End If
    
'     Case 7 'Lista de fuentes importadas desde archivo:
'        i = UBound(ImportFont)
'
'        If Err.Number = 9 Then
'            ReDim ImportFont(0) As String
'
'            Get_FreeID = 0
'
'        Else
'            For j = 0 To i
'                If ImportFont(j) = vbNullString Then
'                    Get_FreeID = j
'                    Exit For
'
'                End If
'
'            Next j
'
'            ReDim Preserve ImportFont(i + 1) As String
'            Get_FreeID = i + 1
'
'        End If
        
End Select

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
    hWnd = D3DWindow.hDeviceWindow

End Property

'===============================================================================
' Name: FPS
' Input:
' Output:
'   Long - Numero de cuadros por segundo.
' Purpose: Devuelve el numero de cuadros por segundo (Frames Per Second) de actualización de renderizado.
' Remarks: Este valor define la velocidad de actualización de la ejecución de las rutinas de dibujo y demás llamadas que realiza el programa entre la actualización de la pantalla y la anterior. Una velocidad recomendada para presentar gráficos en pantalla con un movimiento suave es de 60 cuadros por segundo y la mínima aconsejable suelen ser 30. Si la propiedad vSync fue activada al establecer el modo de video los cuadros por segundo de la aplicación se limitaran a un máximo definido por la actualización del monitor (si el monitor tiene un refresco de 60hz la actualización será de 60 cuadros por segundo como máximo) ya que el dispositivo grafico no ejecutara mas rutinas hasta que no se realice un refresco de pantalla por completo.
'===============================================================================
Public Property Get FPS() As Long
    FPS = FrameRate

End Property

'===============================================================================
' Name: Windowed
' Input:
' Output:
'   Boolean - Devuelve verdadero si el modo de video esta establecido en modo ventana.
' Purpose: Indica si se esta ejecutando el modo de video en modo ventana.
' Remarks:
'===============================================================================
Public Property Get Windowed() As Boolean
    If D3DWindow.Windowed = 1 Then Windowed = True

End Property

'===============================================================================
' Name: TripleBuffer
' Input:
' Output:
'   Boolean - Devuelve verdadero si el Triple Buffer esta activado.
' Purpose: Indica si esta activado el modo de TripleBuffer.
' Remarks: El modo de TripleBuffer solo esta disponible en modos a pantalla completa.
'===============================================================================
Public Property Get TripleBuffer() As Boolean
    TripleBuffer = D3D_TripleBuffer

End Property

'===============================================================================
' Name: Antialiasing
' Input:
' Output:
'   Boolean - Devuelve verdadero si el modo de suavizado de contorno de pixeles esta activado.
' Purpose: Indica si esta activado el modo Antialiasing.
' Remarks: Si su dispositivo grafico no soporta modos de Antialiasing esta propiedad siempre devolverá falso.
'===============================================================================
Public Property Get Antialiasing() As Boolean
    Antialiasing = D3D_Antialiasing

End Property

'===============================================================================
' Name: MAP_Count
' Input:
' Output:
'   Long - Numero de gráficos en memoria.
' Purpose: Devuelve el numero de gráficos en memoria.
' Remarks:
'===============================================================================
Public Property Get MAP_Count() As Long
    On Local Error Resume Next
    MAP_Count = UBound(Texture) + 1

End Property

'===============================================================================
' Name: SURF_Count
' Input:
' Output:
'   Long - Numero de superficies en memoria.
' Purpose: Devuelve el numero de superficies en memoria.
' Remarks:
'===============================================================================
Public Property Get SURF_Count() As Long
    On Local Error Resume Next
    SURF_Count = UBound(Surface) + 1

End Property

'===============================================================================
' Name: PRECAL_CircleCount
' Input:
' Output:
'   Long - Numero de listas precalculadas de circunferencias en memoria.
' Purpose: Devuelve el numero de listas precalculadas de circunferencias en memoria.
' Remarks:
'===============================================================================
Public Property Get PRECAL_CircleCount() As Long
    On Local Error Resume Next
    PRECAL_CircleCount = UBound(PreCircle) + 1

End Property

'===============================================================================
' Name: FONT_SystemCount
' Input:
' Output:
'   Long - Numero de fuentes de texto TrueType en memoria.
' Purpose: Devuelve el numero de fuentes de texto TrueType en memoria.
' Remarks: No confundir con el numero de fuentes de texto TrueType instaladas en el sistema.
'===============================================================================
Public Property Get FONT_SystemCount() As Long
    On Local Error Resume Next
    FONT_SystemCount = UBound(D3DFont) + 1

End Property

'===============================================================================
' Name: VSync
' Input:
' Output:
'   Boolean - Devuelve verdadero si la espera de refresco vertical del monitor esta activada.
' Purpose: Indica si esta activado el VSync.
' Remarks:
'===============================================================================
Public Property Get VSync() As Boolean
    VSync = D3D_VSync

End Property

'Reset the Device:
Private Function ResetDevice(DeviceParams As D3DPRESENT_PARAMETERS) As Long
    On Local Error GoTo ErrOut
    
    Dim Result As Long, i As Long
        
    'Clear the font from memory
    If Not Me.FONT_SystemCount = 0 Then
        For i = 0 To UBound(D3DFont)
            Call D3DFont(i).D3D_Font.OnLostDevice
        Next i
    End If

    'Reset the device
    DoEvents
    Call d3dDevice.Reset(DeviceParams)
    Global_Mod.SetForegroundWindow Me.hWnd
    
    ' Reiniciamos y recreamos las fuentes:
    If Not Me.FONT_SystemCount = 0 Then
        For i = 0 To UBound(D3DFont)
            'Restore the font
            Call D3DFont(i).D3D_Font.OnResetDevice
            'Volvemos a configurar la fuente de texto actual:
            Set D3DFont(i).I_Font = D3DFont(i).Std_Font
            Set D3DFont(i).D3D_Font = D3DX.CreateFont(d3dDevice, D3DFont(i).I_Font.hFont)
        Next
    End If
    
    Call SetRenderStates
    Exit Function
    
ErrOut:
    ResetDevice = Err.Number
    Debug.Print "Error al reiniciar hardware: " & Err.Number & " " & Err.Description
End Function

'Reset the textures and device settings
Private Sub SetRenderStates()
    On Error GoTo ErrOut
    
    d3dDevice.SetVertexShader FVF
    
    d3dDevice.SetRenderState D3DRS_LIGHTING, 0
    d3dDevice.SetRenderState D3DRS_SPECULARENABLE, 1
    d3dDevice.SetRenderState D3DRS_CULLMODE, 0

    d3dDevice.SetRenderState D3DRS_LASTPIXEL, 1

    d3dDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    d3dDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
    d3dDevice.SetRenderState D3DRS_ALPHAREF, &H1
    d3dDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL

    d3dDevice.SetRenderState D3DRS_EDGEANTIALIAS, Me.Antialiasing
    d3dDevice.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, Me.Antialiasing
    
    Exit Sub
    
ErrOut:
    Debug.Print "Error al configurar estados de hardware: " & Err.Number & " " & Err.Description
End Sub

'===============================================================================
' Name: Screen
' Input:
' Output:
'   Display_Mode - Devuelve la información del modo de video establecido.
' Purpose: Información del modo de video establecido.
' Remarks: El objeto Screen contiene la información del modo de video establecido ya sea con Init o con DEVICE_SetDisplayMode.
'===============================================================================
Public Property Get Screen() As Display_Mode
    With Screen
        .Bpp = D3D_Screen.Bpp
        .Height = D3D_Screen.Height
        .RefreshRate = D3D_Screen.RefreshRate
        .Width = D3D_Screen.Width
        
    End With

End Property

'// Version 1.0.1

'Crea un render target a partir de una textura:
Public Function TARGET_Create(Map As Long) As Long
    On Error GoTo ErrOut
    Dim i As Long
    
    If D3D_Init And Not Texture(Map).Texture Is Nothing Then
        i = -1
        i = Get_FreeID(6)
        
        With RenderTarget(i)
            'Creamos el Render Target:
            Set .Target = D3DX.CreateRenderToSurface(d3dDevice, Texture(Map).Texture_Width, Texture(Map).Texture_Height, _
                D3DFMT_A8R8G8B8, 0, D3DFMT_UNKNOWN)
            
            'Asociamos la textura como destino del Render Target:
            Set .Surface = Texture(Map).Texture.GetSurfaceLevel(0)
            
            'Configuramos los parametros del puerto de vision acorde las dimensiones de la textura destino:
            With .ViewPort
                .Height = Texture(Map).Image_Height
                .Width = Texture(Map).Image_Width
                
            End With
        
        End With
        
        TARGET_Create = i
        Exit Function
        
    End If
    
ErrOut:
    If Not i = -1 Then Call Me.TARGET_Destroy(i)
    TARGET_Create = GFX_NOTCREATED

End Function

'Destruye un render target:
Public Function TARGET_Destroy(Target As Long)
    On Error GoTo ErrOut
    
        If Target = UBound(RenderTarget) Then
            If Target = 0 Then
                Erase RenderTarget
            
            Else
                ReDim Preserve RenderTarget(UBound(RenderTarget) - 1) As GFX_RenderTarget
            
            End If
        
        Else
            With RenderTarget(Target)
                Set .Target = Nothing
                Set .Surface = Nothing
                
            End With
    
        End If
    
ErrOut:

End Function

'Abre la textura como destino de las operaciones graficas:
Public Sub TARGET_Open(Target As Long)
    On Local Error Resume Next
    Dim i As Long
    
    i = UBound(RenderTarget)
    
    If Not Err.Number = 9 And Not Target > i And Not Target < 0 Then TargetID = Target

End Sub

'Cierra y actualiza la memoria de la textura:
Public Sub TARGET_Close()
    On Local Error Resume Next
    Call ExecuteRenderBuffer(RenderTarget(TargetID).Buffer)
    TargetID = -1

End Sub

'===============================================================================
' Name: TARGET_Count
' Input:
' Output:
'   Long - Numero de Render Targets en memoria.
' Purpose: Devuelve el numero de Render Targets en memoria.
' Remarks:
'===============================================================================
Public Property Get TARGET_Count() As Long
    On Local Error Resume Next
    TARGET_Count = UBound(RenderTarget) + 1

End Property

'===============================================================================
' Name: DEVICE_SetSpecularChannel
' Input:
'   ByVal A As Long - Color ARGB del primer vertice. (esquina superior izquierda)
'   ByVal B As Long - Color ARGB del segundo vertice. (esquina superior derecha)
'   ByVal C As Long - Color ARGB del tercer vertice. (esquina inferior izquierda)
'   ByVal D As Long - Color ARGB del cuarto vertice. (esquina inferior derecha)
' Output:
' Purpose: Configura los valores del canal specular de cada vertice en una operacion de dibujo.
' Remarks: Para poder configurar los valores del canal specular de un grafico o primitiva grafica antes se ha de llamar a este procedimiento para asignar los valores a cada vertice.
' Dichos valores se reiniciaran a 0 despues de terminar una llamada a cualquiera de las funciones funciones a las que se aplica los valores de este procedimiento:
' DRAW_Line, DRAW_Box, DRAW_AdvBox, DRAW_Trapezoid, DRAW_Map, DRAW_MapEx, DRAW_AdvMap
'===============================================================================
Public Sub DEVICE_SetSpecularChannel(A As Long, B As Long, C As Long, D As Long)
    Specular(0) = A
    Specular(1) = B
    Specular(2) = C
    Specular(3) = D
    SetSpecular = True
End Sub

'===============================================================================
' Name: DEVICE_SetGamma
' Input:
'   Optional ByVal Value As Single - Valor del vertice 1.
' Output:
' Purpose: Corrige el brillo de la pantalla.
' Remarks: La correccion gamma solo esta disponible en modo pantalla completa. En modo ventana no surte efecto alguno.
'===============================================================================
Public Sub DEVICE_SetGamma(Optional ByVal Value As Single = 1!)
    On Error GoTo ErrOut
    
    Dim d3dCaps As D3DCAPS8
    Dim NewRamp As D3DGAMMARAMP
    Dim NewValue As Long
    Dim ui As Long
    Dim i As Long
    Static CapsRead As Boolean
    Static GammaSupport As Boolean

    If Not CapsRead Then
        'see if this device can do fullscreen gamma
        Call d3dDevice.GetDeviceCaps(d3dCaps)
        GammaSupport = ((d3dCaps.Caps2 And D3DCAPS2_FULLSCREENGAMMA) = D3DCAPS2_FULLSCREENGAMMA)
        CapsRead = True
    
    End If
        
    If GammaSupport And Not Windowed Then
        'create linear gamma ramp
        For i = 0 To 255
            NewValue = i * Value * 255: ui = 0
            If NewValue > 32767 Then NewValue = NewValue - 32767: ui = 1
            If NewValue > 32767 Then NewValue = 32767
    
            'manipulate bits to handle unsigned integers
            NewRamp.Red(i) = NewValue Or (&H8000 * ui)
            NewRamp.Green(i) = NewValue Or (&H8000 * ui)
            NewRamp.Blue(i) = NewValue Or (&H8000 * ui)
            
        Next i
    
        'send gamma ramp to device
        Call d3dDevice.SetGammaRamp(D3DSGR_NO_CALIBRATION, NewRamp)
        
        GammaFactor = Value
            
    End If

ErrOut:
    
End Sub

'===============================================================================
' Name: Gamma
' Input:
' Output:
'   Single - Nivel actual de brillo.
' Purpose: Devuelve el nivel de brillo establecido.
' Remarks:
'===============================================================================
Public Property Get Gamma() As Single
    Gamma = GammaFactor

End Property

'===============================================================================
' Name: DEVICE_SetDrawCenter
' Input:
'   ByVal X As Long - Nuevo origen de dibujo para X.
'   ByVal Y As Long - Nuevo origen de dibujo para Y.
' Output:
' Purpose: Configura una nueva posicion de dibujo en una operacion de dibujo.
' Remarks: Este procedimiento permite definir un nuevo origen de dibujo y eje de rotacion en una operacion de dibujo de imagenes.
' Dichos valores se reiniciaran despues de terminar una llamada a cualquiera de las funciones funciones a las que se aplica los valores de este procedimiento:
' DRAW_Map, DRAW_MapEx
' Este metodo no afecta a DRAW_MapEx si el parametro Center es Verdadero.
'===============================================================================
Public Sub DEVICE_SetDrawCenter(X As Long, Y As Long)
    NewCenterX = X
    NewCenterY = Y
    SetCenter = True

End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrOut
    
    ' Obtenemos la cantidad de memoria de video del hardware:
    Dim DispMode As D3DDISPLAYMODE
    Dim tmpD3DWindow As D3DPRESENT_PARAMETERS
    
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate
    
    With tmpD3DWindow
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
        .BackBufferFormat = DispMode.Format
        .BackBufferHeight = DispMode.Height
        .BackBufferWidth = DispMode.Width
        .hDeviceWindow = Global_Mod.GetDesktopWindow()
    End With
    
    Set d3dDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Global_Mod.GetDesktopWindow(), D3DCREATE_SOFTWARE_VERTEXPROCESSING, tmpD3DWindow)
    
    TotalVideoMem = d3dDevice.GetAvailableTextureMem(D3DPOOL_MANAGED)
    
    Set d3dDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing
    
    Exit Sub

ErrOut:
    Dim exCod As Long: exCod = Err.Number
    Dim exDesc As String: exDesc = Err.Description

    Call Err.Clear
    On Error GoTo 0
    
    Call Err.Raise(GFX_ErrorCodes.GFX_UNKNOWNERROR, , "Error al crear la instancia de la clase dx_GFX_Class. Posiblemente falte un componente de DirectX 8.1 (revise el archivo dx8vb.dll) o el dispositivo grafico no tenga soporte de aceleracion 3D." & vbNewLine & vbNewLine & _
                                                      "Codigo de la excepcion: " & exCod & " (" & Hex(exCod) & ")" & vbNewLine & "Descripcion: " & exDesc)
End Sub
