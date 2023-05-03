Option Strict On
Option Explicit On
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)> _
<Microsoft.VisualBasic.ComClass("8d82c8bf-3a13-4354-b892-89cf180d7a35", "8868fa0d-6562-4c1f-9d72-3dde809f56fc")> _
Public Class APITrainingFPT
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        'TODO : Your code goes here

        Return 0
    End Function


    Public Function ASCLRuleExample(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningBoard As IPlanningBoard = preactor.PlanningBoard
        Dim BestResRec As Integer = 0
        Dim BestEndTime As DateTime = DateTime.Now
        Dim BestChangeStart As DateTime = DateTime.Now

        If planningBoard Is Nothing Then
            MsgBox("This Rule must be run from the Sequencer")
            Return 0
        End If

        Try
            Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
            Dim operationRecord As Integer = 0
            Dim ResourceRecords As IEnumerable(Of Integer)
            Dim operationTimes As Preactor.OperationTimes?
            CreateRankedParentQueue(preactor, planningBoard, ordersTable, "JobsQueue")

            While planningBoard.GetOperationInQueue("JobsQueue", 1, operationRecord)

                While operationRecord > 0
                    ResourceRecords = planningBoard.FindResources(operationRecord)
                    BestEndTime = planningBoard.ScheduleHorizon.[End]
                    BestChangeStart = planningBoard.TerminatorTime

                    For Each ResourceRecord As Integer In ResourceRecords
                        operationTimes = planningBoard.TestOperationOnResource(operationRecord, ResourceRecord, planningBoard.TerminatorTime)

                        If (operationTimes.HasValue) AndAlso (operationTimes.Value.ProcessEnd < BestEndTime.AddDays(-planningBoard.SchedulingAccuracy)) Then
                            BestResRec = ResourceRecord
                            BestEndTime = operationTimes.Value.ProcessEnd
                            BestChangeStart = operationTimes.Value.ChangeStart
                        End If
                    Next

                    If BestResRec > 0 Then
                        planningBoard.PutOperationOnResource(operationRecord, BestResRec, BestChangeStart)
                    End If

                    planningBoard.RemoveOperationFromQueue("JobsQueue", operationRecord)
                    operationRecord = planningBoard.GetNextOperation(operationRecord, 1)
                End While
            End While

        Catch ex As Exception
            Dim preactorException As Exception = Nothing

            If ex.InnerException IsNot Nothing AndAlso TypeOf ex.InnerException Is PreactorException Then
                preactorException = ex.InnerException
            ElseIf TypeOf ex Is PreactorException Then
                preactorException = ex
            End If

            If preactorException IsNot Nothing Then
                MsgBox(preactorException.Message, , "Runtime Error in " & ex.Source)
            End If
            Throw
        End Try

        Return 0
    End Function

    Private Shared Sub CreateRankedParentQueue(ByVal preactor As IPreactor, ByVal planningBoard As IPlanningBoard, ByVal ordersTable As Integer, ByVal queueName As String)
        Dim ordersParent As FormatFieldPair = Nothing
        Dim dueDateField As FormatFieldPair? = Nothing
        Dim priorityField As FormatFieldPair? = Nothing
        Dim ordersParentffp As IEnumerable(Of FormatFieldPair) = preactor.FindClassificationString("FAMILY(Order No.)")
        Dim dueDateffp As IEnumerable(Of FormatFieldPair) = preactor.FindClassificationString("DUE DATE")
        Dim priorityFieldffp As IEnumerable(Of FormatFieldPair) = preactor.FindClassificationString("PRIORITY")

        For Each ffp As FormatFieldPair In ordersParentffp
            If preactor.GetFormatName(ffp.FormatNumber) = "Orders" Then ordersParent = ffp
        Next

        For Each ffp As FormatFieldPair In dueDateffp
            If preactor.GetFormatName(ffp.FormatNumber) = "Orders" Then dueDateField = ffp
        Next

        For Each ffp As FormatFieldPair In priorityFieldffp
            If preactor.GetFormatName(ffp.FormatNumber) = "Orders" Then priorityField = ffp
        Next

        planningBoard.CreateQueue(queueName)
        Dim parentRecord As Integer = 0
        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)

        While parentRecord > 0
            planningBoard.AddOperationToQueue(queueName, parentRecord, QueuePosition.[End])
            parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        End While

        Dim SequenceMode As SequenceMode = planningBoard.SequenceMode

        Select Case SequenceMode.Priority
            Case SequencePriority.DueDate

                If dueDateField.HasValue Then
                    planningBoard.RankQueueByFieldName(queueName, preactor.GetFieldName(dueDateField.Value), QueueRanking.Ascending)
                End If

            Case SequencePriority.Priority

                If priorityField.HasValue Then
                    planningBoard.RankQueueByFieldName(queueName, preactor.GetFieldName(priorityField.Value), QueueRanking.Ascending)
                End If

            Case SequencePriority.ReversePriority

                If priorityField.HasValue Then
                    planningBoard.RankQueueByFieldName(queueName, preactor.GetFieldName(priorityField.Value), QueueRanking.Descending)
                End If

            Case Else
        End Select
    End Sub

End Class
