VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_GFX -Efectos basicos en sprites"
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

Private Angle As Single

Private Sub Form_Load()
    Me.Show ' Forzamos al formulario a mostrarse.
    Set Graphics = New dx_GFX_Class ' Creamos la instancia del objeto grafico.
    Render = Graphics.Init(Me.hWnd, 500, 100, 32, True) ' Inicializamos el objeto grafico y el modo de video.
    Texture = Graphics.MAP_Load(App.Path & "\texture.png", 0) ' Cargamos la textura para el sprite.
    
    Do While Render
        Angle = Angle + 0.025: If Angle > 360 Then Angle = 0
        
        ' Rotacion con centro de giro y origen de dibujo definido.
        Graphics.DRAW_MapEx Texture, 50, 50, 0, 100, 100, Angle, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, True
        
        ' Tintado:
        Graphics.DRAW_MapEx Texture, 100, 0, 0, 100, 100, 0, Blendop_Color, Graphics.ARGB_Set(255, 255, 0, 0), Mirror_None, Filter_Bilinear, False
        
        ' Espejado horizontal:
        Graphics.DRAW_MapEx Texture, 200, 0, 0, 100, 100, 0, Blendop_Color, &HFFFFFFFF, Mirror_Horizontal, Filter_Bilinear, False
        
        ' Espejado vertical:
        Graphics.DRAW_MapEx Texture, 300, 0, 0, 100, 100, 0, Blendop_Color, &HFFFFFFFF, Mirror_Vertical, Filter_Bilinear, False
        
        ' Espejado en ambos ejes:
        Graphics.DRAW_MapEx Texture, 400, 0, 0, 100, 100, 0, Blendop_Color, &HFFFFFFFF, Mirror_Both, Filter_Bilinear, False
        
        Graphics.Frame ' Renderizamos la escena.
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Graphics.MAP_Unload Texture ' Descargamos la textura de memoria.
    Render = False ' Termina el bucle de renderizado.
    Graphics.Terminate ' Terminamos la ejecucion de la clase grafica y liberamos los recursos utilizados.
    Set Graphics = Nothing ' Destruimos la instancia del objeto grafico.
End Sub
