VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_GFX - Dibujar sprites en perspectiva"
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

Private Sub Form_Load()
    Me.Show ' Forzamos al formulario a mostrarse.
    Set Graphics = New dx_GFX_Class ' Creamos la instancia del objeto grafico.
    Render = Graphics.Init(Me.hWnd, 640, 480, 32, True) ' Inicializamos el objeto grafico y el modo de video.
    Texture = Graphics.MAP_Load(App.Path & "\texture.png", 0) ' Cargamos la textura para el sprite.
    
    Do While Render
        'Dibujamos un cubo en perspectiva caballera:
        
        '   Para que las proyecciones coincidan hay que tener en cuenta la anchura y altura de cada
        '   grafico ya que la funcion simplemente aplica una deformacion de coordenadas segun los
        '   valores de altura y anchura que se le pasan como argumentos. A continuacion se detalla
        '   un pequeño esquema que explica como obtener los valores correctos para los dos argumentos:
        
        ' > Caballera_Width
        '   _____
        '  /    /   La altura del grafico transformado es igual a la anchura del grafico dividido
        ' /____/    entre 4:
        '               Altura = Anchura_Grafico / 4
        '               Anchura = Anchura_Grafico
        '
        
        ' > Caballera_Height_Negative
        '
        '   /|      La anchura del grafico transformado es igual a mitad de la altura del grafico:
        '  / |          Anchura = Altura_Grafico / 2
        ' |  |          Altura = Altura_Grafico
        ' | /
        ' |/
        '
        '   Estos valores son validos para cuando se van a dibujar graficos con la misma altura y
        '   anchura. Se pueden tambien dibujar graficos con alturas o anchuras diferentes, para ello hay
        '   calcular los valores correctos para no proyectar angulos incorrectos. Tambien se puede
        '   el parametro 'Factor' para aplicar una correccion a la proyeccion de los vertices.
        
        Graphics.DRAW_MapEx Texture, 64, 96, 0, 128, 128, 0, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, False
        Graphics.DRAW_AdvMap Texture, 64, 64, 0, 128, 32, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, Caballera_Width
        Graphics.DRAW_AdvMap Texture, 192, 96, 0, 64, 128, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, Caballera_Height_Negative
        
        'Dibujamos un cubo en perspectiva isometrica:
        
        '   En este caso no sera necesario calcular un valor para la altura y anchura ya que en la
        '   proyeccion isometrica la proyeccion se calcula dividiendo entre 2 el valor a deducir
        '   del elemento opuesto:
        '       Anchura = Altura / 2
        '       Altura = Anchura / 2
        
        '   Estos valores ya los calcula automaticamente la funcion DRAW_AdvMap().
        
        Graphics.DRAW_AdvMap Texture, 288, 160, 0, 128, 128, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, Isometric_Base
        Graphics.DRAW_AdvMap Texture, 288, 224, 0, 128, 128, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, Isometric_Height
        Graphics.DRAW_AdvMap Texture, 416, 288, 0, 128, 128, Blendop_Color, &HFFFFFFFF, Mirror_None, Filter_Bilinear, Isometric_Height_Negative
        
        Graphics.Frame ' Renderizamos la escena.
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Graphics.MAP_Unload Texture ' Descargamos la textura de memoria.
    Render = False ' Termina el bucle de renderizado.
    Graphics.Terminate ' Terminamos la ejecucion de la clase grafica y liberamos los recursos utilizados.
    Set Graphics = Nothing ' Destruimos la instancia del objeto grafico.
End Sub
