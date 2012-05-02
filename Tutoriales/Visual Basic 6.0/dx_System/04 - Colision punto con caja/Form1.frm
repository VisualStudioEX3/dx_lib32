VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_lib32 - Colision punto con caja"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1125
      Left            =   1590
      Top             =   915
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sys As New dx_System_Class      ' Objeto que hace referencia a dx_System.
Private A As GFX_Rect                   ' Variable que contiene los valores de la caja.

Private Sub Form_Load()
    Me.AutoRedraw = True
    Me.ForeColor = vbWhite
    
    ' Definimos la posicion y dimensiones de la caja A:
    A.X = Shape1.Left
    A.Y = Shape1.Top
    A.Width = Shape1.Width
    A.Height = Shape1.Height

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cls
    ' Mostramos el resultado de comprobar si existe colision entre la caja A
    ' y el punto donde se encuentra el cursor del raton:
    Print Sys.MATH_PointInRect(CLng(X), CLng(Y), A)

End Sub
