Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor

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

        RemoveCutAddOperations()
        FindCutObjects()

    End Sub

    Private Sub RemoveCutAddOperations()
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

                If GetCustomPropertyValue(oPartDoc, "CutAddTarget") = "True" Then
                    For Each oFeature As Object In oPartDoc.ComponentDefinition.Features
                        If oFeature.Name.contains("Cuttool") Or oFeature.Name.contains("Addtool") Then
                            oFeature.Delete()
                        End If
                    Next

                    For Each obody As Object In oPartDoc.ComponentDefinition.Features
                        If obody.Name.contains("Cut") Or obody.Name.contains("Add") Then
                            obody.Delete()
                        End If
                    Next

                End If
            End If
        Next
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
                    Debug.Print(oPartDoc.DisplayName)
                    CheckInterference(oOcc)
                End If
            End If
        Next

    End Sub

    Private Sub CheckInterference(ByVal AddCutBody As ComponentOccurrence)
        'Make sure the body is visible
        AddCutBody.Visible = True

        'Add all occurences to a collecion
        Dim CheckSet As ObjectCollection = _invApp.TransientObjects.CreateObjectCollection
        Dim OtherOcc As ObjectCollection = _invApp.TransientObjects.CreateObjectCollection

        CheckSet.Add(AddCutBody)

        For Each Occ As ComponentOccurrence In _Doc.ComponentDefinition.Occurrences
            If Not AddCutBody Is Occ Then
                OtherOcc.Add(Occ)
            End If
        Next

        'Check for interference
        Dim InterResults As InterferenceResults
        InterResults = _Doc.ComponentDefinition.AnalyzeInterference(CheckSet, OtherOcc)

        If InterResults.Count = 1 Then
            Debug.Print(InterResults.Count)
            CutandAdd(InterResults.Item(1).OccurrenceOne, InterResults.Item(1).OccurrenceTwo)
        ElseIf InterResults.Count > 1 Then

            MsgBox("Multiple intersections found...")
        Else
            MsgBox("No intersections found...")
        End If

        AddCutBody.Visible = False
    End Sub

    Sub CutandAdd(ByVal CutAddOcc As ComponentOccurrence, ByVal TargetOcc As ComponentOccurrence)

        For Each body As SurfaceBody In CutAddOcc.SurfaceBodies
            If body.Name = "Cut" Or body.Name = "Add" Then
                CopyBody(body, TargetOcc)
                AddProperty(TargetOcc)
            End If
        Next

    End Sub

    Sub CopyBody(ByVal sourcebody As SurfaceBody, ByVal targetocc As ComponentOccurrence)

        Dim targetDef As PartComponentDefinition
        targetDef = targetocc.Definition
        ' The selected body is a body proxy in the context of   
        ' the assembly. However, there's a problem with the
        ' TransientBrep.Copy method and it creates a copy of the 
        ' body that ignores the transorm.  The code below creates 
        ' the copy and then performs an extra step to apply the 
        ' transform. 
        Dim newBody As SurfaceBody
        newBody = _invApp.TransientBRep.Copy(sourcebody)
        Call _invApp.TransientBRep.Transform(newBody,
             sourcebody.ContainingOccurrence.Transformation)

        ' Transform the body into the parts space of the 
        ' target occurrence. 
        Dim trans As Matrix
        trans = targetocc.Transformation
        trans.Invert()
        Call _invApp.TransientBRep.Transform(newBody, trans)

        ' Create a base feature definition.  This is used to define the 
        ' various inputs needed to create a base feature. 
        Dim nonPrmFeatures As NonParametricBaseFeatures
        nonPrmFeatures = targetDef.Features.NonParametricBaseFeatures
        Dim featureDef As NonParametricBaseFeatureDefinition
        featureDef = nonPrmFeatures.CreateDefinition
        Dim transObjs As TransientObjects
        transObjs = _invApp.TransientObjects
        Dim col As ObjectCollection

        col = transObjs.CreateObjectCollection
        col.Add(newBody)

        ' This creates an non-associative copy that is a solid. 
        featureDef.BRepEntities = col
        featureDef.OutputType = BaseFeatureOutputTypeEnum.kSolidOutputType
        featureDef.TargetOccurrence = targetocc
        featureDef.IsAssociative = False

        nonPrmFeatures.AddByDefinition(featureDef)
        ' Get operation number
        Dim iCut As Integer = 1
        Dim iAdd As Integer = 1
        For Each oFeature As Object In targetDef.Features
            Dim featurename As String = oFeature.name

            If featurename.Contains("Cut") Then
                iCut = iCut + 1
            ElseIf featurename.Contains("Add") Then
                iAdd = iAdd + 1
            End If
        Next

        Dim cutoradd As PartFeatureOperationEnum
        If sourcebody.Name = "Cut" Then
            cutoradd = PartFeatureOperationEnum.kCutOperation
            targetDef.Features.Item(targetDef.Features.Count).Name = sourcebody.Name & iCut
        Else
            cutoradd = PartFeatureOperationEnum.kJoinOperation
            targetDef.Features.Item(targetDef.Features.Count).Name = sourcebody.Name & iAdd
        End If

        Dim toolcol As ObjectCollection = _invApp.TransientObjects.CreateObjectCollection
        toolcol.Add(targetDef.SurfaceBodies.Item(2))

        targetDef.Features.CombineFeatures.Add(targetDef.SurfaceBodies.Item(1), toolcol, cutoradd)
        If sourcebody.Name = "Cut" Then
            targetDef.Features.Item(targetDef.Features.Count).Name = sourcebody.Name & "tool" & iCut
        Else
            targetDef.Features.Item(targetDef.Features.Count).Name = sourcebody.Name & "tool" & iAdd
        End If


    End Sub

    Sub AddProperty(ByVal ComponentOcc As ComponentOccurrence)

        Dim doc As PartDocument = ComponentOcc.ReferencedDocumentDescriptor.ReferencedDocument

        If Not GetCustomPropertyValue(doc, "CutAddTarget") = "True" Then
            Dim customPropSet As PropertySet
            customPropSet = doc.PropertySets.Item("Inventor User Defined Properties")

            ' Create a new boolean property. 
            Dim yesNoValue As Boolean
            yesNoValue = True
            customPropSet.Add(yesNoValue, "CutAddTarget")
        End If

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
