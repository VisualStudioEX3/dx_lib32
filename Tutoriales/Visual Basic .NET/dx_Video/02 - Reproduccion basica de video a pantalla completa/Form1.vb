Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	' Este tutorial utiliza el objeto grafico de dx_lib32 para mostrar el video a pantalla completa.
	Private Graphics As New dx_lib32.dx_GFX_Class ' Instancia del objeto grafico de dx_lib32.
	
	Private Video As New dx_lib32.dx_Video_Class ' Instancia del objeto de video de dx_lib32.
	Private clip As Integer ' Guarda el identificador de la pelicula de video.
	Private clipWidth, clipHeight As Integer ' Almacenaran las dimensiones de la pelicula de video.
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Show()
		Graphics.Init(Me.Handle.ToInt32, 800, 600, 32) ' Inicializamos el modo de video a pantalla completa a 800x600x32.
		Video.Init(Me.Handle.ToInt32)
		clip = Video.VIDEO_Load(My.Application.Info.DirectoryPath & "\clock.avi") ' Carga la pelicula de video en memoria.
		Video.VIDEO_GetSize(clip, clipWidth, clipHeight) ' Obtenemos las dimensiones originales de la pelicula de video.
		' Reproducimos el video centrado en la pantalla:
		Video.VIDEO_Play(clip, (Graphics.Screen.Width \ 2) - (clipWidth \ 2), (Graphics.Screen.Height \ 2) - (clipHeight \ 2), 0, 0)
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Graphics.Terminate() ' Termina la ejecucion de la clase grafica y liberamos los recursos utilizados.
		Video.VIDEO_Unload(clip) ' Descarga la pelicula de video de la memoria.
		Video.Terminate() ' Terminamos la ejecucion de la clase de video y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Video no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Video = Nothing
	End Sub
	
	Private Sub Form1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Me.Close() ' Cerramos la aplicacion pulsando cualquier tecla.
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class