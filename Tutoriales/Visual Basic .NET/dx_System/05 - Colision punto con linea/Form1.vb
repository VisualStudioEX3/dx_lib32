Option Strict Off
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Drawing2D

Friend Class Form1
    Inherits System.Windows.Forms.Form

    Private Gfx As Graphics
    Private Sys As New dx_lib32.dx_System_Class ' Objeto que hace referencia a dx_System.
    Private A, B As Point
    Private Result As String

    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Gfx = Me.CreateGraphics

        ' Definimos la posicion y dimensiones de la linea:
        A = New Point(10, 30)
        B = New Point(Me.Size.Width - 30, Me.Size.Height - 40)
        
    End Sub

    Private Sub Form1_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        ' Obtenemos el resultado de la colsion entre el puntero del raton y la caja:
        Result = Sys.MATH_PointInLine(A.X, A.Y, B.X, B.Y, eventArgs.X, eventArgs.Y).ToString()
        Form1_Paint(eventSender, Nothing)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Gfx.Dispose()
    End Sub

    Private Sub Form1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        ' Limpiamos el formulario:
        Gfx.Clear(Color.Black)

        ' Actualizamos la posicion del control Shape en el formulario:
        Gfx.DrawLine(Pens.Red, A, B)

        ' Mostramos en pantalla el resultado de comprobar si existe o no colision entre las cajas A y B:
        Gfx.DrawString(Result, Me.Font, Brushes.White, 10, 10)
    End Sub
End Class