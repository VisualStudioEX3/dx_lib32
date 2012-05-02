Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Graphics As dx_lib32.dx_GFX_Class ' Instancia del objeto grafico de dx_lib32.
	Private Render As Boolean ' Controla el bucle de renderizado.
	Private Texture As Integer ' Identificador de la textura.
	Private Background As Integer ' Identificador de la textura de fondo.
	
	Private Angle As Single
	Private Color As Integer
	Private Alpha As Short
	Private Incr As Boolean
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Show() ' Forzamos al formulario a mostrarse.
		Graphics = New dx_lib32.dx_GFX_Class ' Creamos la instancia del objeto grafico.
		Render = Graphics.Init(Me.Handle.ToInt32, 600, 480, 32, True) ' Inicializamos el objeto grafico y el modo de video.
		Texture = Graphics.MAP_Load(My.Application.Info.DirectoryPath & "\texture.png", 0) ' Cargamos la textura para el texture.
		Background = Graphics.MAP_Load(My.Application.Info.DirectoryPath & "\background.png", 0) ' Cargamos la textura para el fondo.
		
		Do While Render
			Color = Graphics.ARGB_Set(Alpha, 255, 255, 255)
			
			' Dibujamos una imagen como fondo para poder apreciar los efectos:
			Graphics.DRAW_Map(Background, 0, 0, 0, 0, 0)
			
			' Opacidad a traves del canal alfa del color:
			Graphics.DRAW_MapEx(Texture, 128, 160, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_Color, Color, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Opacidad aditiva:
			Graphics.DRAW_MapEx(Texture, 256 + 64, 80, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_Aditive, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Opacidad sustrativa:
			Graphics.DRAW_MapEx(Texture, 384 + 128, 160, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_Sustrative, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Efecto exclusion:
			Graphics.DRAW_MapEx(Texture, 256 + 64, 240, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_XOR, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Efecto invertir colores:
			Graphics.DRAW_MapEx(Texture, 128, 320, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_Inverse, Color, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Opacidad cristalina:
			Graphics.DRAW_MapEx(Texture, 256 + 64, 400, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_Crystaline, &HFFFFFFFF, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			' Efecto escala de grises:
			Graphics.DRAW_MapEx(Texture, 384 + 128, 320, 0, 128, 128, Angle, dx_lib32.Blit_Alpha.Blendop_GreyScale, Color, dx_lib32.Blit_Mirror.Mirror_None, dx_lib32.Blit_Filter.Filter_Bilinear, True)
			
			Graphics.Frame() ' Renderizamos la escena.
		Loop 
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		' Descargamos las texturas de memoria:
		Graphics.MAP_Unload(Texture)
		Graphics.MAP_Unload(Background)
		Render = False ' Termina el bucle de renderizado.
		Graphics.Terminate() ' Terminamos la ejecucion de la clase grafica y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Graphics no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Graphics = Nothing ' Destruimos la instancia del objeto grafico.
	End Sub
	
	' El control Timer calcula el angulo para los sprites y el valor del canal alfa del color para la funcion ARGB_Set():
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Angle = Angle + 0.5 : If Angle > 360 Then Angle = 0
		
		If Incr Then
			If Alpha = 255 Then Incr = False
			Alpha = Alpha + 5
		Else
			If Alpha = 0 Then Incr = True
			Alpha = Alpha - 5
		End If
	End Sub
End Class