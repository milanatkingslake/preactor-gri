Option Strict On
Option Explicit On
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("ab88ee06-8db8-4bcc-903d-b9c16e1cee6c", "57893313-88d5-4035-8d69-923ef6dd212e")>
Public Class CustomRuleTest

    'Defulat sample execution code from preactor API
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Return 0
    End Function
    '' CustomRuleExample_1 this event given by Mark
    Public Function CustomRule_Kes_1(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
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
            Dim ResourceRecords As IEnumerable(Of Integer)
            Dim operationTimes As Preactor.OperationTimes?
            Dim num As Integer = preactor.RecordCount("Orders")
            Dim y As Integer = 1
            ''define table variable and assign columns
            Dim orderTbl As DataTable = New DataTable()
            Dim orderNo As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
            Dim dueDate As DataColumn = New DataColumn("DueDate", Type.[GetType]("System.DateTime"))
            Dim operationRecordNo As DataColumn = New DataColumn("OperationRecordNo", Type.[GetType]("System.Int16"))
            Dim recordNo As Integer
            orderTbl.Columns.Add(orderNo)
            orderTbl.Columns.Add(dueDate)
            orderTbl.Columns.Add(operationRecordNo)

            For i = 1 To preactor.RecordCount("Orders")
                Dim dr As DataRow = orderTbl.NewRow()

                dr("OrderNo") = preactor.ReadFieldString("Orders", "Order No.", i)
                dr("DueDate") = preactor.ReadFieldDateTime("Orders", "Due Date", i)

                recordNo = preactor.FindMatchingRecord("Orders", "Order No.", recordNo, dr("OrderNo").ToString())
                dr("OperationRecordNo") = recordNo

                orderTbl.Rows.Add(dr)
            Next

            Dim orderTblShort As DataTable = New DataTable()
            orderTbl.DefaultView.Sort = "DueDate"
            orderTblShort = orderTbl.DefaultView.ToTable()
            Dim operationRecord_ As Integer

            For Each ordeshr As DataRow In orderTblShort.Rows
                operationRecord_ = CInt(ordeshr("OperationRecordNo"))
                ResourceRecords = planningBoard.FindResources(operationRecord_)
                BestEndTime = planningBoard.ScheduleHorizon.[End]
                BestChangeStart = planningBoard.TerminatorTime

                For Each ResourceRecord As Integer In ResourceRecords
                    operationTimes = planningBoard.TestOperationOnResource(operationRecord_, ResourceRecord, planningBoard.TerminatorTime)

                    If (operationTimes.HasValue) AndAlso (operationTimes.Value.ProcessEnd < BestEndTime.AddDays(-planningBoard.SchedulingAccuracy)) Then
                        BestResRec = ResourceRecord
                        BestEndTime = operationTimes.Value.ProcessEnd
                        BestChangeStart = operationTimes.Value.ChangeStart
                    End If
                Next

                If BestResRec > 0 Then
                    planningBoard.PutOperationOnResource(operationRecord_, BestResRec, BestChangeStart)
                End If
            Next
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
    Public Function CustomRule_Kes_2_Algorithmic(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningBoard As IPlanningBoard = preactor.PlanningBoard
        Dim BestResRec As Integer = 0
        Dim BestEndTime As DateTime = DateTime.Now
        Dim BestChangeStart As DateTime = DateTime.Now

        If planningBoard Is Nothing Then
            MsgBox("This Rule must be run from the Sequencer")
            Return 0
        End If
        Dim ordersParent As FormatFieldPair = Nothing
        Dim dueDateField As FormatFieldPair? = Nothing
        Dim priorityField As FormatFieldPair? = Nothing

        Dim queueName As String = "JobsQueue"

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
        Dim orderby As Integer = 0
        Select Case SequenceMode.Priority
            Case SequencePriority.DueDate

                If dueDateField.HasValue Then
                    orderby = 1
                    ''planningBoard.RankQueueByFieldName(queueName, preactor.GetFieldName(dueDateField.Value), QueueRanking.Ascending)
                End If

            Case SequencePriority.Priority

                If priorityField.HasValue Then
                    orderby = 2
                    ''planningBoard.RankQueueByFieldName(queueName, preactor.GetFieldName(priorityField.Value), QueueRanking.Ascending)
                End If

            Case SequencePriority.ReversePriority

                If priorityField.HasValue Then
                    orderby = 3
                    ''planningBoard.RankQueueByFieldName(queueName, preactor.GetFieldName(priorityField.Value), QueueRanking.Descending)
                End If

            Case Else
        End Select



        Try
            Dim ordersTable As Integer = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
            Dim ResourceRecords As IEnumerable(Of Integer)
            Dim operationTimes As Preactor.OperationTimes?
            Dim num As Integer = preactor.RecordCount("Orders")
            Dim y As Integer = 1
            Dim recordNo As Integer
            Dim orderQue As List(Of Order) = New List(Of Order)()

            For i = 1 To preactor.RecordCount("Orders")
                Dim newOrder As Order = New Order()
                newOrder.OrderNo = preactor.ReadFieldString("Orders", "Order No.", i).ToString
                newOrder.DueDate = Convert.ToDateTime(preactor.ReadFieldDateTime("Orders", "Due Date", i))
                newOrder.OperationRecordNo = Convert.ToInt32(preactor.FindMatchingRecord("Orders", "Order No.", recordNo, preactor.ReadFieldString("Orders", "Order No.", i).ToString))
                newOrder.Priority = preactor.ReadFieldString("Orders", "Priority", i).ToString

                orderQue.Add(newOrder)
            Next
            '' Dim orderQue1 As List(Of Order) = New List(Of Order)()
            Dim orderQue_ As List(Of Order) = New List(Of Order)()

            Select Case orderby
                Case 1 ''DueDate
                    orderQue_ = orderQue.OrderBy(Function(o) o.DueDate).ToList()
                Case 2 ''Priority
                    'orderQue_ = orderQue.OrderBy(Function(o) o.pri).ToList()
                Case 3 ''ReversePriority
                    orderQue_ = orderQue.OrderBy(Function(o) o.OperationRecordNo).ToList()
                Case Else
                    Debug.WriteLine("NotFound")
            End Select

            Dim operationRecord_ As Integer
            For Each ordeshr In orderQue_
                operationRecord_ = CInt(ordeshr.OperationRecordNo)
                ResourceRecords = planningBoard.FindResources(operationRecord_)
                BestEndTime = planningBoard.ScheduleHorizon.[End]
                BestChangeStart = planningBoard.TerminatorTime

                For Each ResourceRecord As Integer In ResourceRecords
                    operationTimes = planningBoard.TestOperationOnResource(operationRecord_, ResourceRecord, planningBoard.TerminatorTime)

                    If (operationTimes.HasValue) AndAlso (operationTimes.Value.ProcessEnd < BestEndTime.AddDays(-planningBoard.SchedulingAccuracy)) Then
                        BestResRec = ResourceRecord
                        BestEndTime = operationTimes.Value.ProcessEnd
                        BestChangeStart = operationTimes.Value.ChangeStart

                        Dim str_Resource_Name As String = preactor.ReadFieldString("Resources", "Name", ResourceRecord)
                        Dim matchRecord = 0
                        matchRecord = preactor.FindMatchingRecord("Orders", "Resource", matchRecord, str_Resource_Name)

                        While matchRecord > 0
                            Dim dte_Start_Time As DateTime = preactor.ReadFieldDateTime("Orders", "Start Time", matchRecord)
                            Dim dte_End_Time As DateTime = preactor.ReadFieldDateTime("Orders", "End Time", matchRecord)
                            ''If str_Resource_Name = "ST3-CP-04" Then
                            If BestChangeStart < dte_End_Time Then
                                ''  MsgBox("Test")
                                BestChangeStart = dte_End_Time
                            End If
                            ''End If
                            matchRecord = preactor.FindMatchingRecord("Orders", "Resource", matchRecord, str_Resource_Name)

                        End While

                    End If
                Next

                If BestResRec > 0 Then
                    planningBoard.PutOperationOnResource(operationRecord_, BestResRec, BestChangeStart)
                End If
            Next
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
    '' CreateRankedParentQueue this event given by Mark
    Public Function ASCLRuleExample_Mark(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
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

            If preactorException IsNot Nothing Then MsgBox(preactorException.Message,, "Runtime Error in ")
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

    Public Class Order
        Public Property OrderNo As String
        Public Property DueDate As DateTime
        Public Property OperationRecordNo As Integer
        Public Property Priority As String

    End Class
End Class
