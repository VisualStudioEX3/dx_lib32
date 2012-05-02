Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Audio As dx_lib32.dx_Sound_Class ' Instancia del objeto de audio de dx_lib32.
	Private Sample As Integer ' Guarda el identificador de la muestra de sonido.
	Private FX() As dx_lib32.Sound_Effects
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Audio = New dx_lib32.dx_Sound_Class ' Creamos la instancia del objeto.
		Audio.Init(Me.Handle.ToInt32, 64) ' Inicializamos el motor de audio con 64 canales para efectos de sonido.
		Sample = Audio.SOUND_Load(My.Application.Info.DirectoryPath & "\sample.wav") ' Cargamos en memoria la muestra de sonido.
		Audio.SOUND_Play(Sample, 0, True) ' Reproducimos la muestra de sonido en bucle en el canal 0.
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Audio.SOUND_Unload(Sample) ' Descargamos de memoria la muestra de sonido.
		Audio.Terminate() ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Audio no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Audio = Nothing ' Destruimos la instancia del objeto de audio.
	End Sub
	
	'UPGRADE_WARNING: El evento Check1.CheckStateChanged se puede desencadenar cuando se inicializa el formulario. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
		Dim Index As Short = Check1.GetIndex(eventSender)
		' Utilizamos la funcion auxiliar para generar el array con los valores de los efectos seleccionados:
		FX = Audio.SOUND_FX_MakeArrayEffects(CBool(Check1(0).CheckState), CBool(Check1(1).CheckState), CBool(Check1(2).CheckState), CBool(Check1(3).CheckState), CBool(Check1(4).CheckState), CBool(Check1(5).CheckState), CBool(Check1(6).CheckState))
		
        Audio.SOUND_FX_SetEffects(0, CType(FX, Array)) ' Aplicamos los efectos al canal.
	End Sub
End Class