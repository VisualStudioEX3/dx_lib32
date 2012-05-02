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
			' Definir el color de cada vertice del sprite:
			Graphics.DEVICE_SetVertexColor(&HFFFFFFFF, &HFFFF0000, &HFF00FF00, &HFF0000FF)
			Graphics.DRAW_Map(Texture, 0, 0, 0, 200, 200)
			
			' Definir el tono de iluminacion de cada vertice:
			Graphics.DEVICE_SetSpecularChannel(&HFFFFFFFF, &HFFFF0000, &HFF00FF00, &HFF0000FF)
			Graphics.DRAW_MapEx(Texture, Graphics.Screen.Width \ 2, Graphics.Screen.Height \ 2, 0, 200, 200, 0, dx_lib32.Blit_Alpha.Blendop_Color, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Combinar ambas tecnicas:
			Graphics.DEVICE_SetVertexColor(&HFFFFFFFF, &HFFFF0000, &HFF00FF00, &HFF0000FF)
			Graphics.DEVICE_SetSpecularChannel(&HFFFFFFFF, 0, 0, &HFFFF0000)
			Graphics.DRAW_Map(Texture, Graphics.Screen.Width - 200, Graphics.Screen.Height - 200, 0, 200, 200)
			
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