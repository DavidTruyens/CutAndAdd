Imports System
Imports System.Type
Imports System.Activator
Imports System.Threading
Imports System.Runtime.InteropServices
Imports Inventor
Imports System.IO

Module Module1
    Public _invApp As Inventor.Application
    Dim _Doc As Inventor.AssemblyDocument
    Dim _OrigDoc As Inventor.PartDocument
    Dim _started As Boolean = False
    Dim _CompDef As AssemblyComponentDefinition

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

        _Doc = _invApp.ActiveDocument
        _CompDef = _Doc.ComponentDefinition
        FindCutObjects()

        AssociativeBodyCopy()



    End Sub

    Private Sub FindCutObjects()

        ' Get all of the leaf occurrences of the assembly. 
        Dim oLeafOccs As ComponentOccurrencesEnumerator
        oLeafOccs = _CompDef.Occurrences.AllLeafOccurrences


        ' Iterate through the occurrences and print the name. 
        Dim oOcc As ComponentOccurrence
        For Each oOcc In oLeafOccs

            ' Check to see if this is a part. 
            If oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Dim oPartDoc As PartDocument
                oPartDoc = oOcc.ReferencedDocumentDescriptor.ReferencedDocument
                If GetCustomPropertyValue(oPartDoc, "CutAdd") = "True" Then
                    'MsgBox(oPartDoc.DisplayName)
                    CheckInterference(oOcc)
                End If
            End If
        Next

    End Sub

    Private Sub CheckInterference(ByVal AddCutBody As ComponentOccurrence)

        'Add all occurences to a collecion
        Dim CheckSet As ObjectCollection = _invApp.TransientObjects.CreateObjectCollection

        For Each Occ As ComponentOccurrence In _Doc.ComponentDefinition.Occurrences
            CheckSet.Add(Occ)
        Next

        'Check for interference
        Dim InterResults As InterferenceResults
        InterResults = _Doc.ComponentDefinition.AnalyzeInterference(CheckSet)

        If InterResults.Count >= 1 Then

        End If
    End Sub

    Sub AssociativeBodyCopy()

        ' Select the body to copy.  This will be a proxy since it's 
        ' being selected in the context of the assembly. 
        Dim sourceBody As SurfaceBodyProxy
        sourceBody = _invApp.CommandManager.Pick(SelectionFilterEnum.kPartBodyFilter, "Select a body to copy.")

        ' Get the occurrence to create the new body within. 
        Dim targetOcc As ComponentOccurrence
        targetOcc = _invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select an occurrence to copy the body into.")
        Dim targetDef As PartComponentDefinition
        targetDef = targetOcc.Definition

        ' Create a base feature definition.  This is used to define the 
        ' various inputs needed to create a base feature. 
        Dim nonPrmFeatures As NonParametricBaseFeatures
        nonPrmFeatures = targetDef.Features.NonParametricBaseFeatures
        Dim featureDef As NonParametricBaseFeatureDefinition
        featureDef = nonPrmFeatures.CreateDefinition
        Dim transObjs As TransientObjects
        transObjs = _invApp.TransientObjects
        Dim col As ObjectCollection

        ' Ask if an associative or non-associative copy should be made. 
        Dim answer As MsgBoxResult
        answer = MsgBox("Choose the type of copy to create." &
                    vbCrLf & vbCrLf &
                    "   Yes - Associative surface" & vbCrLf &
                    "   No - Non-associative solid", vbYesNoCancel)
        If answer = vbYes Then
            ' Define the geometry to be copied.  In this case, 
            ' it's the selected body. 
            col = transObjs.CreateObjectCollection
            col.Add(sourceBody)

            ' This creates an associative copy of the model.  To 
            ' create an associative copy inventor only supports creating 
            ' a surface of composite result, not a solid. 
            featureDef.BRepEntities = col
            featureDef.OutputType = BaseFeatureOutputTypeEnum.kSurfaceOutputType
            featureDef.TargetOccurrence = targetOcc
            featureDef.IsAssociative = True

            Dim baseFeature As NonParametricBaseFeature
            baseFeature = nonPrmFeatures.AddByDefinition(featureDef)
        ElseIf answer = vbNo Then
            ' The selected body is a body proxy in the context of   
            ' the assembly. However, there's a problem with the
            ' TransientBrep.Copy method and it creates a copy of the 
            ' body that ignores the transorm.  The code below creates 
            ' the copy and then performs an extra step to apply the 
            ' transform. 
            Dim newBody As SurfaceBody
            newBody = _invApp.TransientBRep.Copy(sourceBody)
            Call _invApp.TransientBRep.Transform(newBody,
             sourceBody.ContainingOccurrence.Transformation)

            ' Transform the body into the parts space of the 
            ' target occurrence. 
            Dim trans As Matrix
            trans = targetOcc.Transformation
            trans.Invert()
            Call _invApp.TransientBRep.Transform(newBody, trans)

            col = transObjs.CreateObjectCollection
            col.Add(newBody)

            ' This creates an non-associative copy that is a solid. 
            featureDef.BRepEntities = col
            featureDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType
            featureDef.TargetOccurrence = targetOcc
            featureDef.IsAssociative = False

            nonPrmFeatures.AddByDefinition(featureDef)

        End If
    End Sub

    Sub RunSculptDemo()
        'reference to the part document
        Dim oDoc As PartDocument = TryCast(_invApp.ActiveDocument, PartDocument)
        If oDoc Is Nothing Then Exit Sub
        Dim oDef As PartComponentDefinition = oDoc.ComponentDefinition

        'features collection
        Dim oFeatures As PartFeatures = oDef.Features
        'surfaces collection
        Dim oSurfaces As ObjectCollection = _invApp.TransientObjects.CreateObjectCollection()
        For Each osurface As WorkSurface In oDef.WorkSurfaces
            oSurfaces.Add(oFeatures.SculptFeatures.CreateSculptSurface(osurface, PartFeatureExtentDirectionEnum.kNegativeExtentDirection))
        Next
        'create sculpt feature
        Dim oSculpt As SculptFeature = oFeatures.SculptFeatures.Add(oSurfaces, PartFeatureOperationEnum.kCutOperation)
    End Sub

    Private Function GetCustomPropertyValue(ByVal oDocument As Inventor.Document, ByVal PropertyName As String) As String
        Dim Result As String = Nothing
        Try
            Dim oProperty As Inventor.Property = Nothing
            For Each oProperty In oDocument.PropertySets.Item(4)
                If oProperty.Name = PropertyName Then
                    Result = oProperty.Value
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "GetCustomPropertyValue")
        End Try
        Return Result
    End Function

End Module
