VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_lib32 - Colision caja con caja"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   1365
      Top             =   2085
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   1575
      Top             =   630
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sys As New dx_System_Class          ' Objeto que hace referencia a dx_System.
Private A As GFX_Rect, B As GFX_Rect        ' Variables que contendran los valores de las dos cajas.

Private Sub Form_Load()
    Me.AutoRedraw = True
    Me.ForeColor = vbWhite
    Me.BackColor = vbBlack
    
    ' Definimos la posicion y dimensiones de la caja A:
    A.X = Shape1.Left
    A.Y = Shape1.Top
    A.Width = Shape1.Width
    A.Height = Shape1.Height

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Actualizamos la posicion del control Shape en el formulario:
    Shape2.Left = CLng(X)
    Shape2.Top = CLng(Y)
    
    ' Definimos o actualizamos la posicion y dimensiones de la caja B:
    B.X = Shape2.Left
    B.Y = Shape2.Top
    B.Width = Shape2.Width
    B.Height = Shape2.Height
    
    Cls
    
    ' Mostramos en pantalla el resultado de comprobar si existe o no
    ' colision entre las cajas A y B:
    Print Sys.MATH_IntersectRect(A, B)

End Sub
