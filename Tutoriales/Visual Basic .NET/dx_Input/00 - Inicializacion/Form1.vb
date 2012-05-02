Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private GameInput As New dx_lib32.dx_Input_Class ' Instancia del objeto de entrada de dx_lib32.
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		GameInput.Init(Me.Handle.ToInt32)
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		GameInput.Terminate() ' Terminamos la ejecucion de la clase de entrada y liberamos los recursos utilizados.
		'UPGRADE_NOTE: El objeto GameInput no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GameInput = Nothing
	End Sub
End Class