Imports System
Imports System.Type
Imports System.Activator
Imports System.Threading
Imports System.Runtime.InteropServices
Imports Inventor
Imports System.IO

Module Module1
    Public _invApp As Inventor.Application
    Dim _Doc As Inventor.PartDocument
    Dim _OrigDoc As Inventor.PartDocument
    Dim _started As Boolean = False

    Public Sub Main()

        Try
            _invApp = Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            Try
                Dim invAppType As Type = GetTypeFromProgID("Inventor.Application")
                _invApp = CreateInstance(invAppType)
                _invApp.Visible = True
                _started = True
                MsgBox("Inventor Started")

            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("unable to start Inventor")
            End Try
        End Try

        If _invApp.Documents.Count = 0 Then
            MsgBox("Need to open an Assembly Document")
            Return
        End If

        If _invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MsgBox("Need to have an Assembly Document active")
            Return
        End If

    End Sub



End Module
