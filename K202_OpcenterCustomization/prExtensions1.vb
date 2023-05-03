Imports System.Runtime.CompilerServices
Imports Preactor

Module Extensions

    <Extension()> _
    Public Function FindClassificationString(ByVal preactor As IPreactor, ByVal classificationString As String, ByVal formatName As String) As List(Of Preactor.FormatFieldPair)
        Return preactor.FindClassificationString(classificationString, preactor.GetFormatNumber(formatName))
    End Function

    <Extension()> _
    Public Function FindClassificationString(ByVal preactor As IPreactor, ByVal classificationString As String, ByVal formatNumber As Integer) As List(Of Preactor.FormatFieldPair)

        Dim result = New List(Of Preactor.FormatFieldPair)

        Dim list = preactor.FindClassificationString(classificationString)

        For Each item In list
            If item.FormatNumber = formatNumber Then
                result.Add(item)
            End If
        Next

        Return result

    End Function


    <Extension()> _
    Public Function FindFirstClassificationString(ByVal preactor As IPreactor, ByVal classificationString As String, ByVal formatName As String) As Preactor.FormatFieldPair?
        Return preactor.FindFirstClassificationString(classificationString, preactor.GetFormatNumber(formatName))
    End Function

    <Extension()> _
    Public Function FindFirstClassificationString(ByVal preactor As IPreactor, ByVal classificationString As String, ByVal formatNumber As Integer) As Preactor.FormatFieldPair?
        Dim list = preactor.FindClassificationString(classificationString)

        For Each item In list
            If item.FormatNumber = formatNumber Then
                Return item
            End If
        Next

        Return Nothing
    End Function


End Module
