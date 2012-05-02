Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Audio As dx_lib32.dx_Sound_Class ' Instancia del objeto de audio de dx_lib32.
	Private Sample As Integer ' Guarda el identificador de la muestra de sonido.
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Audio = New dx_lib32.dx_Sound_Class ' Creamos la instancia del objeto.
		Audio.Init(Me.Handle.ToInt32, 64) ' Inicializamos el motor de audio con 64 canales para efectos de sonido.
		Sample = Audio.SOUND_Load(My.Application.Info.DirectoryPath & "\sample.wav") ' Cargamos en memoria la muestra de sonido.
		HScroll3.Value = Audio.SOUND_GetSamplesPerSecond(Sample) / 10 ' Obtenemos las muestras por segundo de la muestra de audio.
		Audio.SOUND_Play(Sample, 0, True) ' Reproducimos la muestra en bucle en el canal 0.
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Audio.SOUND_Unload(Sample) ' Descargamos de memoria la muestra de sonido.
		Audio.Terminate() ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Audio no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Audio = Nothing ' Destruimos la instancia del objeto de audio.
	End Sub
	
	'UPGRADE_NOTE: HScroll1.Change pasó de ser un evento a un procedimiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar evento HScroll1.Change tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub HScroll1_Change(ByVal newScrollValue As Integer)
		Audio.SOUND_SetVolume(0, (newScrollValue)) ' Modificamos el nivel de volumen del canal 0.
	End Sub
	
	'UPGRADE_NOTE: HScroll2.Change pasó de ser un evento a un procedimiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar evento HScroll2.Change tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub HScroll2_Change(ByVal newScrollValue As Integer)
		Audio.SOUND_SetPan(0, (newScrollValue)) ' Modificamos el nivel de balance del canal 0.
	End Sub
	
	'UPGRADE_NOTE: HScroll3.Change pasó de ser un evento a un procedimiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar evento HScroll3.Change tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub HScroll3_Change(ByVal newScrollValue As Integer)
		Dim value As Integer
		value = newScrollValue
		Audio.SOUND_SetFrequency(0, value * 10) ' Modificamos el nivel de velocidad del canal 0.
	End Sub
	Private Sub HScroll1_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles HScroll1.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				HScroll1_Change(eventArgs.newValue)
		End Select
	End Sub
	Private Sub HScroll2_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles HScroll2.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				HScroll2_Change(eventArgs.newValue)
		End Select
	End Sub
	Private Sub HScroll3_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles HScroll3.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				HScroll3_Change(eventArgs.newValue)
		End Select
	End Sub
End Class