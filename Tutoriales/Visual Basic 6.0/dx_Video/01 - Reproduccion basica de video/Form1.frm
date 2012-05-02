VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dx_Video -Reproduccion basica de video"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Video As New dx_Video_Class ' Instancia del objeto de video de dx_lib32.
Private clip As Long ' Guarda el identificador de la pelicula de video.

Private Sub Form_Load()
    Video.Init Me.hWnd
    clip = Video.VIDEO_Load(App.Path & "\clock.avi") ' Carga la pelicula de video en memoria.
    Video.VIDEO_Play clip, 0, 0, 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Video.VIDEO_Unload clip ' Descarga la pelicula de video de la memoria.
    Video.Terminate ' Terminamos la ejecucion de la clase de video y liberamos los recursos utilizados.
    Set Video = Nothing
End Sub
