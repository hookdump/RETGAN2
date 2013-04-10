VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RETGAN 2"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function procesarLinea(s As String) As String
    procesarLinea = UCase(s) + "ZZZ"
End Function
    
Sub procesarArchivo(x As Integer)
    Dim FS As FileSystemObject
    Dim TS As TextStream
    
    Dim FSnew As FileSystemObject
    Dim TSnew As TextStream
    
    Dim filename As String, newfilename As String
    Dim fullpath As String, newpath As String
    filename = Me.File1.List(x)
    newfilename = filename + ".ASD.txt"
    fullpath = Me.File1.Path + "\" + filename
    newpath = Me.File1.Path + "\" + newfilename
    
    Set FS = New FileSystemObject
    Set TS = FS.OpenTextFile(fullpath, ForReading, False)
    
    Set FSnew = New FileSystemObject
    Set TSnew = FS.CreateTextFile(newpath, True)
    
    Dim line As String, newline As String
    Do While Not (TS.AtEndOfStream)
        line = TS.ReadLine()
        newline = procesarLinea(line)
        TSnew.WriteLine (newline)
    Loop
    
    TSnew.Close
End Sub

Sub procesarLista()
    Dim i As Integer
    For i = 0 To Me.File1.ListCount - 1
        procesarArchivo i
    Next
End Sub

Private Sub cmdProcesar_Click()
    procesarLista
End Sub

Private Sub Form_Load()
    File1.Path = "C:\RETGAN"
End Sub
