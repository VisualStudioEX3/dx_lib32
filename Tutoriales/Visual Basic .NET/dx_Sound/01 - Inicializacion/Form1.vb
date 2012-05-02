Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Audio As dx_lib32.dx_Sound_Class ' Instancia del objeto de audio de dx_lib32.
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Audio = New dx_lib32.dx_Sound_Class ' Creamos la instancia del objeto.
		Audio.Init(Me.Handle.ToInt32, 64) ' Inicializamos el motor de audio con 64 canales para efectos de sonido.
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Audio.Terminate() ' Terminamos la ejecucion de la clase de audio y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto Audio no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Audio = Nothing ' Destruimos la instancia del objeto de audio.
	End Sub
End Class