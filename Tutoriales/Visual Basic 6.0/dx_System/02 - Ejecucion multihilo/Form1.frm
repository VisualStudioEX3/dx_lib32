VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   2580
   ClientTop       =   5595
   ClientWidth     =   2715
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   58
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   181
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1110
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sys As dx_System_Class      ' Objeto que accede a dx_System.
Private TimerID As Long             ' Guarda el identificador del cronometro.

Private Sub Form_Load()
    On Error GoTo ErrOut
    
    Me.Show
    
    ' Creamos la instancia de la clase dx_System:
    Set Sys = New dx_System_Class
    
    i = 1
    
    ' Creamos el cronometro y le asociamos el metodo Evento() de Modulo1.bas para que se ejecute periodicamente:
    TimerID = Sys.TIMER_CreateProcess(Me.hWnd, 100, AddressOf Evento)
    
    ' Mostramos un cuadro de mensaje (MsgBox) para demostrar la capacidad proceso en paralelo con dx_lib32.
    ' Las cajas de mensajes detienen la ejecucion del hilo principal de ejecucion hasta que la caja
    ' de mensaje es cerrada:
    Call MsgBox("Este ejemplo muestra la capacidad de proceso en paralelo de dx_lib32 " & _
                "mediante cronometros. Estos cronometros serian similares a los controles " & _
                "Timer de Visual Basic con la excepcion de que no precisan la instancia" & _
                "de un control en un formulario." & vbNewLine & _
                vbNewLine & _
                "Este metodo se puede considerar programacion " & _
                "multihilo ya que el cronometro ejecuta el procedimiento en un hilo " & _
                "de ejecucion independiente al del programa principal. Este metodo " & _
                "permitiria ejecutar varios procesos simultaneamente sin interrumpir " & _
                "el hilo de ejecucion principal del programa." & vbNewLine & _
                vbNewLine & _
                "Los cronometros, aun siendo algo mas estables que la programacion " & _
                "multihilo real, no dejan de ser minimamente inestables. Si se " & _
                "produce el menor error mientras un cronometro se esta ejecutando " & _
                "es posible que el programa termine precipitadamente su ejecucion " & _
                "y sin aviso de error o dejar colgado al propio programa. Por ello " & _
                "se aconseja ser precavidos al usar los cronometros y tratar de " & _
                "evitarlos si no son realmente necesarios.", vbExclamation, "" & _
                "dx_lib32 - Ejecucion multihilo")
    
    ' El codigo no seguira ejecutandose hasta que se cierre el cuadro de mensaje modal pero el codigo del metodo
    ' Evento() del Module1.bas seguira ejecutandose hasta que destruyamos la instancia del cronometro.
    
ErrOut:
    ' Destruimos el cronometro:
    Call Sys.TIMER_KillProcess(Me.hWnd, TimerID)
    
    ' Destruimos la instancia de la clase:
    Set Sys = Nothing
    
    End
End Sub
