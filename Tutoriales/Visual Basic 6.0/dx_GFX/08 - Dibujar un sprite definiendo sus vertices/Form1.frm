VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_GFX - Dibujar un sprite definiendo sus vertices"
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
Private VertexData(3) As Vertex ' Array que define la informacion de la posicion de los vertices y su color.
Private VertexSpecular(3) As Long ' Array que define el valor del canal Specular de cada vertices (en el ejemplo este array tienes sus valor por defecto a 0)

Private Sub Form_Load()
    Me.Show ' Forzamos al formulario a mostrarse.
    Set Graphics = New dx_GFX_Class ' Creamos la instancia del objeto grafico.
    Render = Graphics.Init(Me.hWnd, 640, 480, 32, True) ' Inicializamos el objeto grafico y el modo de video.
    Texture = Graphics.MAP_Load(App.Path & "\texture.png", 0) ' Cargamos la textura para el sprite.
    
    ' Definimos los vertices en orden de las agujas del reloj:
    
    ' 0 -------------- 1
    ' |                |
    ' |                |
    ' |                |
    ' |                |
    ' |                |
    ' |                |
    ' |                |
    ' 3 -------------- 2
    
    VertexData(0).X = 32: VertexData(0).Y = 96: VertexData(0).Color = &HFFFFFFFF
    VertexData(1).X = 128: VertexData(1).Y = 0: VertexData(1).Color = &HFFFFFFFF
    VertexData(2).X = 256: VertexData(2).Y = 256: VertexData(2).Color = &HFFFFFFFF
    VertexData(3).X = 256: VertexData(3).Y = 64: VertexData(3).Color = &HFFFFFFFF
    
    Do While Render
        Graphics.DRAW_VertexMap Texture, VertexData(), 0, VertexSpecular(), Blendop_Color, Mirror_None, Filter_Trilinear ' Dibujamos la textura.
        Graphics.Frame ' Renderizamos la escena.
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Graphics.MAP_Unload Texture ' Descargamos la textura de memoria.
    Render = False ' Termina el bucle de renderizado.
    Graphics.Terminate ' Terminamos la ejecucion de la clase grafica y liberamos los recursos utilizados.
    Set Graphics = Nothing ' Destruimos la instancia del objeto grafico.
End Sub
