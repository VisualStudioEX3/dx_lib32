VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_Input - Inicializacion"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GameInput As New dx_Input_Class ' Instancia del objeto de entrada de dx_lib32.

Private Sub Form_Load()
    GameInput.Init Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GameInput.Terminate  ' Terminamos la ejecucion de la clase de entrada y liberamos los recursos utilizados.
    Set GameInput = Nothing
End Sub
