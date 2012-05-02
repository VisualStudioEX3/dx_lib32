VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_lib32 - Trabajar con archivos PAK"
   ClientHeight    =   3945
   ClientLeft      =   2580
   ClientTop       =   5595
   ClientWidth     =   5805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5805
   Begin VB.CommandButton Command1 
      Caption         =   "&Extraer"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Left            =   4380
      TabIndex        =   2
      Top             =   3600
      Width           =   1410
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   3645
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sys As dx_System_Class          'Objeto que accede a dx_System.

Private PAKFile As String               'Guarda la ruta de acceso al archivo paquete.

Private List() As PAK_FileInfo          'Esta variable guarda el listado de archivos del paquete.
Private Files As Long                   'Guarda el numero de archivos que hay en el paquete.
Private Size As Long                    'Guarda el tamaño en KiloBytes del conjunto de archivos del paquete.

Private Ret As SYS_ErrorCodes           'Guarda el valor devuelto por las funciones.
Private ErrDesc As String               'En caso de error se almacenara aqui la descripcion del mismo.

Private i As Long

Private Sub Form_Load()
    On Error GoTo ErrOut
    
    'Creamos la instancia de la clase dx_System:
    Set Sys = New dx_System_Class
    
    'Utilizamos la ventana de abrir archivo para buscar un paquete para su lectura:
    PAKFile = Sys.DLG_OpenFile(Me.hWnd, "Archivos PAK (*.pak)|*.pak|Todos los archivos (*.*)|*.*", "Abrir archivo PAK", App.Path)
    If PAKFile = vbNullString Then End
    
    'Leemos el contenido del paquete:
    Ret = Sys.PAK_Load(PAKFile, List(), Files, Size)
    
    If Not Ret = SYS_OK Then Err.Raise vbObjectError
    
    'Mostramos el tamaño en memoria del paquete y el numero de archivos:
    If (Size / 1024) < 1024 Then
        Label1.Caption = Files & " archivos en " & Int(Size / 1024) & " Kb"
    
    Else
        Label1.Caption = Files & " archivos en " & Int(Size / 1024 ^ 2) & " Mb (" & Int(Size / 1024) & " Kb)"
    
    End If
    
    'Metemos los datos de la lista en el listbox:
    For i = 0 To UBound(List)
        List1.AddItem List(i).FileName, i
        DoEvents
    Next i
    
    Exit Sub
    
ErrOut:
    Call ProcessError
    Unload Me
End Sub

'Extrae el archivo del paquete:
Private Sub Command1_Click()
    On Error GoTo ErrOut
    
    Dim FileName As String
    Dim SaveName As String
    
    'Obtenemos la ruta y nombre de acceso al archivo dentro del paquete:
    FileName = List(List1.ListIndex).FileName
    
    'Obtenemos solo el nombre del archivo:
    FileName = Get_FileName(FileName)
    
    'Utilizamos la ventana de guardar archivo para indicar donde extraemos el archivo:
    SaveName = Sys.DLG_SaveFile(Me.hWnd, "Todos los archivos (*.*)|*.*", "Extraer archivo", App.Path, FileName)
    If SaveName = vbNullString Then Exit Sub
    
    'Extraemos al disco el archivo:
    If Not Sys.PAK_ExtractFile(PAKFile, List(List1.ListIndex), SaveName) Then Err.Raise vbObjectError
    
    Exit Sub
    
ErrOut:
    Call ProcessError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Destruimos la instancia de la clase:
    Set Sys = Nothing
End Sub

Private Sub List1_Click()
    Command1.Enabled = True
End Sub

'Extrae el nombre del archivo de la ruta de acceso del paquete:
Private Function Get_FileName(FileName As String) As String
    Dim i As Integer
       
    For i = Len(FileName) To 1 Step -1
        If Mid(FileName, i, 1) = "/" Then Exit For
        Get_FileName = Mid(FileName, i, 1) & Get_FileName
    Next i
End Function

'Procesa el valor devuelto por las funciones y genera un mensaje de error:
Private Sub ProcessError()
    Select Case Ret
        Case SYS_EMPTYLIST
            ErrDesc = "El paquete no contiene archivos."
        Case SYS_INVALIDFORMAT
            ErrDesc = "El formato del archivo no es correcto."
        Case SYS_UNKNOWNERROR
            ErrDesc = "Error desconocido."
        Case Else
            ErrDesc = Err.Number & ": " & Err.Description
    End Select
    
    MsgBox ErrDesc, vbCritical, Me.Caption
End Sub
