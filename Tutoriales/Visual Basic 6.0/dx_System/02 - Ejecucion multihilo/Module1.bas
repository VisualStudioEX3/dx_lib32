Attribute VB_Name = "Module1"
Option Explicit

Public i As Long, j As Long, k As Long, l As Long   'Variables contador.

' Evento que ejecutara el cronometro periodicamente:
Public Sub Evento()
    On Error Resume Next
    
    If i > 8 Then
        i = 1
        j = j + 1
        
        If j = 10 Then
            k = k + 1
            l = l + j
            j = 0
            
            Call MsgBox("La luna ha dado " & l & " vueltas y sigue rotando " & _
                        "mientras lees este mensaje!", vbExclamation, "Mensaje desde el cronometro")
        End If
    
    End If
    
    Form1.Image1.Picture = LoadPicture(App.Path & "\MOON0" & i & ".ICO")
    Form1.Caption = l + j & " vueltas"
    
    i = i + 1
End Sub

