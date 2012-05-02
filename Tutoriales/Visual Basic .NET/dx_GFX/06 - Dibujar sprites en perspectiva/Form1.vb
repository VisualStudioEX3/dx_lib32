Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Graphics As dx_lib32.dx_GFX_Class ' Instancia del objeto grafico de dx_lib32.
	Private Render As Boolean ' Controla el bucle de renderizado.
	Private Texture As Integer ' Identificador de la textura.
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Show() ' Forzamos al formulario a mostrarse.
		Graphics = New dx_lib32.dx_GFX_Class ' Creamos la instancia del objeto grafico.
		Render = Graphics.Init(Me.Handle.ToInt32, 640, 480, 32, True) ' Inicializamos el objeto grafico y el modo de video.
		Texture = Graphics.MAP_Load(My.Application.Info.DirectoryPath & "\texture.png", 0) ' Cargamos la textura para el sprite.
		
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
			
			Graphics.DRAW_MapEx(Texture, 64, 96, 0, 128, 128, 0, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, False)
			Graphics.DRAW_AdvMap(Texture, 64, 64, 0, 128, 32, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, dx_lib32.Blit_Perspective.Caballera_Width)
			Graphics.DRAW_AdvMap(Texture, 192, 96, 0, 64, 128, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, dx_lib32.Blit_Perspective.Caballera_Height_Negative)
			
			'Dibujamos un cubo en perspectiva isometrica:
			
			'   En este caso no sera necesario calcular un valor para la altura y anchura ya que en la
			'   proyeccion isometrica la proyeccion se calcula dividiendo entre 2 el valor a deducir
			'   del elemento opuesto:
			'       Anchura = Altura / 2
			'       Altura = Anchura / 2
			
			'   Estos valores ya los calcula automaticamente la funcion DRAW_AdvMap().
			
			Graphics.DRAW_AdvMap(Texture, 288, 160, 0, 128, 128, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, dx_lib32.Blit_Perspective.Isometric_Base)
			Graphics.DRAW_AdvMap(Texture, 288, 224, 0, 128, 128, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, dx_lib32.Blit_Perspective.Isometric_Height)
			Graphics.DRAW_AdvMap(Texture, 416, 288, 0, 128, 128, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, dx_lib32.Blit_Perspective.Isometric_Height_Negative)
			
			Graphics.Frame() ' Renderizamos la escena.
		Loop 
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Graphics.MAP_Unload(Texture) ' Descargamos la textura de memoria.
		Render = False ' Termina el bucle de renderizado.
		Graphics.Terminate() ' Terminamos la ejecucion de la clase grafica y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Graphics no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Graphics = Nothing ' Destruimos la instancia del objeto grafico.
	End Sub
End Class