Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	Private Sys As dx_lib32.dx_System_Class 'Objeto que accede a dx_System.
	
	Private PAKFile As String 'Guarda la ruta de acceso al archivo paquete.
	
	Private List() As dx_lib32.PAK_FileInfo 'Esta variable guarda el listado de archivos del paquete.
	Private Files As Integer 'Guarda el numero de archivos que hay en el paquete.
	'UPGRADE_NOTE: Size se actualizó a Size_Renamed. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Size_Renamed As Integer 'Guarda el tamaño en KiloBytes del conjunto de archivos del paquete.
	
	Private Ret As dx_lib32.SYS_ErrorCodes 'Guarda el valor devuelto por las funciones.
	Private ErrDesc As String 'En caso de error se almacenara aqui la descripcion del mismo.
	
	Private i As Integer
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrOut
		
		'Creamos la instancia de la clase dx_System:
		Sys = New dx_lib32.dx_System_Class
		
		'Utilizamos la ventana de abrir archivo para buscar un paquete para su lectura:
		PAKFile = Sys.DLG_OpenFile(Me.Handle.ToInt32, "Archivos PAK (*.pak)|*.pak|Todos los archivos (*.*)|*.*", "Abrir archivo PAK", My.Application.Info.DirectoryPath)
		If PAKFile = vbNullString Then End
		
		'Leemos el contenido del paquete:
		Ret = Sys.PAK_Load(PAKFile, List, Files, Size_Renamed)
		
		If Not Ret = dx_lib32.SYS_ErrorCodes.SYS_OK Then Err.Raise(vbObjectError)
		
		'Mostramos el tamaño en memoria del paquete y el numero de archivos:
		If (Size_Renamed / 1024) < 1024 Then
			Label1.Text = Files & " archivos en " & Int(Size_Renamed / 1024) & " Kb"
			
		Else
			Label1.Text = Files & " archivos en " & Int(Size_Renamed / 1024 ^ 2) & " Mb (" & Int(Size_Renamed / 1024) & " Kb)"
			
		End If
		
		'Metemos los datos de la lista en el listbox:
		For i = 0 To UBound(List)
			List1.Items.Insert(i, List(i).FileName)
			System.Windows.Forms.Application.DoEvents()
		Next i
		
		Exit Sub
		
ErrOut: 
		Call ProcessError()
		Me.Close()
	End Sub
	
	'Extrae el archivo del paquete:
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		On Error GoTo ErrOut
		
		Dim FileName As String
		Dim SaveName As String
		
		'Obtenemos la ruta y nombre de acceso al archivo dentro del paquete:
		FileName = List(List1.SelectedIndex).FileName
		
		'Obtenemos solo el nombre del archivo:
		FileName = Get_FileName(FileName)
		
		'Utilizamos la ventana de guardar archivo para indicar donde extraemos el archivo:
		SaveName = Sys.DLG_SaveFile(Me.Handle.ToInt32, "Todos los archivos (*.*)|*.*", "Extraer archivo", My.Application.Info.DirectoryPath, FileName)
		If SaveName = vbNullString Then Exit Sub
		
		'Extraemos al disco el archivo:
		If Not Sys.PAK_ExtractFile(PAKFile, List(List1.SelectedIndex), SaveName) Then Err.Raise(vbObjectError)
		
		Exit Sub
		
ErrOut: 
		Call ProcessError()
	End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'Destruimos la instancia de la clase:
		'UPGRADE_NOTE: El objeto Sys no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Sys = Nothing
	End Sub
	
	'UPGRADE_WARNING: El evento List1.SelectedIndexChanged se puede desencadenar cuando se inicializa el formulario. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub List1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles List1.SelectedIndexChanged
		Command1.Enabled = True
	End Sub
	
	'Extrae el nombre del archivo de la ruta de acceso del paquete:
	Private Function Get_FileName(ByRef FileName As String) As String
        Dim i As Short, tmp As String = ""
		
		For i = Len(FileName) To 1 Step -1
			If Mid(FileName, i, 1) = "/" Then Exit For
            tmp = Mid(FileName, i, 1) & tmp
        Next i

        Return tmp
	End Function
	
	'Procesa el valor devuelto por las funciones y genera un mensaje de error:
	Private Sub ProcessError()
		Select Case Ret
			Case dx_lib32.SYS_ErrorCodes.SYS_EMPTYLIST
				ErrDesc = "El paquete no contiene archivos."
			Case dx_lib32.SYS_ErrorCodes.SYS_INVALIDFORMAT
				ErrDesc = "El formato del archivo no es correcto."
			Case dx_lib32.SYS_ErrorCodes.SYS_UNKNOWNERROR
				ErrDesc = "Error desconocido."
			Case Else
				ErrDesc = Err.Number & ": " & Err.Description
		End Select
		
		MsgBox(ErrDesc, MsgBoxStyle.Critical, Me.Text)
	End Sub
End Class