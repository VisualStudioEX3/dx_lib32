VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_lib32 - Colision punto con linea"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4620
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   12
      X2              =   289
      Y1              =   48
      Y2              =   183
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sys As New dx_System_Class          ' Objeto que hace referencia a dx_System.

Private Sub Form_Load()
    Me.AutoRedraw = True
    Me.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cls
    ' Mostramos el resultado de comprobar si existe colision entre vector definido
    ' por la linea roja y el punto donde se encuentra el cursor del raton:
    Print Sys.MATH_PointInLine(CLng(Line1.X1), CLng(Line1.Y1), CLng(Line1.X2), CLng(Line1.Y2), CLng(X), CLng(Y))
End Sub
