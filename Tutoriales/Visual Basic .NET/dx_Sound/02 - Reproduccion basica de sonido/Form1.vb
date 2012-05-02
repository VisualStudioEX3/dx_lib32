Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Audio As dx_lib32.dx_Sound_Class ' Instancia del objeto de audio de dx_lib32.
	Private Sample As Integer ' Guarda el identificador de la muestra de sonido.
	Private Channel As Integer ' Guarda el indice del canal donde se esta reproduciendo.
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Audio = New dx_lib32.dx_Sound_Class ' Creamos la instancia del objeto.
		Audio.Init(Me.Handle.ToInt32, 64) ' Inicializamos el motor de audio con 64 canales para efectos de sonido.
		Sample = Audio.SOUND_Load(My.Application.Info.DirectoryPath & "\sample.wav") ' Cargamos en memoria la muestra de sonido.
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Audio.SOUND_Unload(Sample) ' Descargamos de memoria la muestra de sonido.
		Audio.Terminate() ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Audio no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Audio = Nothing ' Destruimos la instancia del objeto de audio.
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Channel = Audio.SOUND_Play(Sample) ' Reproducimos la muestra en el primer canal disponible.
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Audio.SOUND_Pause(Channel) ' Pausamos la reproduccion del canal. Si el canal esta en pausa se reanuda la reproduccion.
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Audio.SOUND_Stop(Channel) ' Detenemos la reproduccion del canal y lo dejamos libre para ser utilizado por otra muestra.
	End Sub
End Class