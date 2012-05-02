Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Graphics As dx_lib32.dx_GFX_Class ' Instancia del objeto grafico de dx_lib32.
	Private Render As Boolean ' Controla el bucle de renderizado.
	Private Texture As Integer ' Identificador de la textura.
	Private VertexData(3) As dx_lib32.Vertex ' Array que define la informacion de la posicion de los vertices y su color.
	Private VertexSpecular(3) As Integer ' Array que define el valor del canal Specular de cada vertices (en el ejemplo este array tienes sus valor por defecto a 0)
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Show() ' Forzamos al formulario a mostrarse.
		Graphics = New dx_lib32.dx_GFX_Class ' Creamos la instancia del objeto grafico.
		Render = Graphics.Init(Me.Handle.ToInt32, 640, 480, 32, True) ' Inicializamos el objeto grafico y el modo de video.
		Texture = Graphics.MAP_Load(My.Application.Info.DirectoryPath & "\texture.png", 0) ' Cargamos la textura para el sprite.
		
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
		
		VertexData(0).X = 32 : VertexData(0).Y = 96 : VertexData(0).Color = &HFFFFFFFF
		VertexData(1).X = 128 : VertexData(1).Y = 0 : VertexData(1).Color = &HFFFFFFFF
		VertexData(2).X = 256 : VertexData(2).Y = 256 : VertexData(2).Color = &HFFFFFFFF
		VertexData(3).X = 256 : VertexData(3).Y = 64 : VertexData(3).Color = &HFFFFFFFF
		
		Do While Render
			Graphics.DRAW_VertexMap(Texture, VertexData, 0, VertexSpecular, dx_lib32.Blit_Alpha.Blendop_Color, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Trilinear) ' Dibujamos la textura.
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