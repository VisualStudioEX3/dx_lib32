VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_GFX -Efectos avanzados en sprites"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2070
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Graphics As dx_GFX_Class ' Instancia del objeto grafico de dx_lib32.
Private Render As Boolean ' Controla el bucle de renderizado.
Private Texture As Long ' Identificador de la textura.
Private Background As Long ' Identificador de la textura de fondo.

Private Angle As Single
Private Color As Long, Alpha As Integer, Incr As Boolean

Private Sub Form_Load()
    Me.Show ' Forzamos al formulario a mostrarse.
    Set Graphics = New dx_GFX_Class ' Creamos la instancia del objeto grafico.
    Render = Graphics.Init(Me.hWnd, 600, 480, 32, True) ' Inicializamos el objeto grafico y el modo de video.
    Texture = Graphics.MAP_Load(App.Path & "\texture.png", 0) ' Cargamos la textura para el texture.
    Background = Graphics.MAP_Load(App.Path & "\background.png", 0) ' Cargamos la textura para el fondo.
    
    Do While Render
        Color = Graphics.ARGB_Set(Alpha, 255, 255, 255)
        
        ' Dibujamos una imagen como fondo para poder apreciar los efectos:
        Graphics.DRAW_Map Background, 0, 0, 0, 0, 0
        
        ' Opacidad a traves del canal alfa del color:
        Graphics.DRAW_MapEx Texture, 128, 160, 0, 128, 128, Angle, Blendop_Color, Color, Mirror_None, Filter_Bilinear, True
    
        ' Opacidad aditiva:
        Graphics.DRAW_MapEx Texture, 256 + 64, 80, 0, 128, 128, Angle, Blendop_Aditive, &HFFFFFFFF, Mirror_None, Filter_Bilinear, True
    
        ' Opacidad sustrativa:
        Graphics.DRAW_MapEx Texture, 384 + 128, 160, 0, 128, 128, Angle, Blendop_Sustrative, &HFFFFFFFF, Mirror_None, Filter_Bilinear, True
    
        ' Efecto exclusion:
        Graphics.DRAW_MapEx Texture, 256 + 64, 240, 0, 128, 128, Angle, Blendop_XOR, &HFFFFFFFF, Mirror_None, Filter_Bilinear, True
    
        ' Efecto invertir colores:
        Graphics.DRAW_MapEx Texture, 128, 320, 0, 128, 128, Angle, Blendop_Inverse, Color, Mirror_None, Filter_Bilinear, True
    
        ' Opacidad cristalina:
        Graphics.DRAW_MapEx Texture, 256 + 64, 400, 0, 128, 128, Angle, Blendop_Crystaline, &HFFFFFFFF, Mirror_None, Filter_Bilinear, True
    
        ' Efecto escala de grises:
        Graphics.DRAW_MapEx Texture, 384 + 128, 320, 0, 128, 128, Angle, Blendop_GreyScale, Color, Mirror_None, Filter_Bilinear, True
        
        Graphics.Frame ' Renderizamos la escena.
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Descargamos las texturas de memoria:
    Graphics.MAP_Unload Texture
    Graphics.MAP_Unload Background
    Render = False ' Termina el bucle de renderizado.
    Graphics.Terminate ' Terminamos la ejecucion de la clase grafica y liberamos los recursos utilizados.
    Set Graphics = Nothing ' Destruimos la instancia del objeto grafico.
End Sub

' El control Timer calcula el angulo para los sprites y el valor del canal alfa del color para la funcion ARGB_Set():
Private Sub Timer1_Timer()
    Angle = Angle + 0.5: If Angle > 360 Then Angle = 0
    
    If Incr Then
        If Alpha = 255 Then Incr = False
        Alpha = Alpha + 5
    Else
        If Alpha = 0 Then Incr = True
        Alpha = Alpha - 5
    End If
End Sub

