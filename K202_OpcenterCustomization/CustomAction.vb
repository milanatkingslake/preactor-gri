Option Strict On
Option Explicit On
Imports System.Collections.Specialized
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Windows.Forms
Imports K202_OpcenterCustomization.CustomRuleTest
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("6d100a35-3708-4f40-92df-b73992e5ce19", "c1aa5ecb-fb8a-4bf5-ba0f-bf261cc0b3a4")>
Public Class CustomAction

#Region "AppRunTest"
    Public Function AppRunTest(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        MsgBox("Application done")
        'TODO : Your code goes here

        Return 0
    End Function
#End Region

#Region "Demand date Calculation"
    ''DemandDateCalculation this programme will caclulate demand and update order table
    Public Function DemandDateCalculation(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim strOrderNo As String ''= preactor.ReadFieldString("Demand", "Order No.", 3)
        Dim demandDate As String
        Dim demandDateTbl As DataTable
        Dim result As DialogResult = MessageBox.Show("Please run the Sequencer to load the Gantt chart befor you run the demand date update.",
                              "Information",
                              MessageBoxButtons.YesNo)

        If (result = DialogResult.Yes) Then
            ''Read Demand table
            Dim num As Integer = preactor.RecordCount("Demand")
            Dim i As Integer = 1
            Do
                Dim matchRecord = 0
                Dim strDemandOrderNo As String = preactor.ReadFieldString("Demand", "Order No.", i)
                matchRecord = preactor.FindMatchingRecord("Demand", "Order No.", matchRecord, strDemandOrderNo)
                preactor.WriteField("Demand", "K202_ReadyForShipmentDate", matchRecord, "Unspecified")
                i = i + 1
            Loop While i <= num
            preactor.Commit("Demand")

            demandDateTbl = GetDemandDateByOrderNo(connetionString, strOrderNo)
            Try
                If demandDateTbl.Rows.Count > 0 Then
                    Try
                        For Each dimtbl As DataRow In demandDateTbl.Rows
                            strOrderNo = dimtbl("OrderNo").ToString()
                            demandDate = dimtbl("DemanDate").ToString()
                            If Not (String.IsNullOrEmpty(demandDate) And String.IsNullOrEmpty(strOrderNo)) Then

                                Dim matchRecord = 0
                                matchRecord = preactor.FindMatchingRecord("Demand", "Order No.", matchRecord, strOrderNo)

                                While matchRecord > 0
                                    ''Update Demand
                                    preactor.WriteField("Demand", "K202_ReadyForShipmentDate", matchRecord, demandDate)
                                    ''Check Sub Demands
                                    matchRecord = preactor.FindMatchingRecord("Demand", "Order No.", matchRecord, strOrderNo)
                                    If matchRecord > 0 Then
                                        preactor.WriteField("Demand", "K202_ReadyForShipmentDate", matchRecord, demandDate)
                                    End If
                                End While
                            End If
                        Next
                        preactor.Commit("Demand")
                    Catch ex As Exception
                        MsgBox("Demand no not found",, "error")
                    Finally
                        MsgBox("Demand updated",, "information")
                    End Try
                Else
                    MsgBox("Demand record not found",, "error")
                End If
            Catch ex As Exception
                MsgBox("Demand Update error Plese contact Admin",, "error")
            End Try
        End If
        Return 0
    End Function
    ''GetDemandDateByOrderNo Method will execute K202_GetDemandDate_Sp procedure and get all demand dates
    Public Function GetDemandDateByOrderNo(ByRef connetionString As String, ByRef demandNo As String) As DataTable

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetDemandDate_Sp"

            adapter = New SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            adapter.Fill(ds)

            Return ds.Tables(0)

            connection.Close()
        Catch ex As Exception
            MsgBox("Extra formers not define",, "error")
        Finally

        End Try

    End Function
    ''this programe will Backup Import File from live server file location
    Public Function BackupImportedFile(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim connection As DbConnection = New SqlConnection(connetionString)
        Dim DbName = connection.Database

        Dim sFilefromPath As String = "Z:\Import"
        Dim sFileToPath As String = "Z:\Import\Backup"

        Dim text As String = ""
        Dim files() As String = IO.Directory.GetFiles(sFilefromPath)
        Try
            For i = 0 To UBound(files)
                sFilefromPath = files(i)
                If File.Exists(sFilefromPath) Then
                    Dim filename As String = Path.GetFileName(sFilefromPath)
                    Dim destinationPath As String = Path.Combine(sFileToPath, filename)
                    If File.Exists(destinationPath) Then
                        File.Delete(destinationPath)
                    End If
                    File.Move(sFilefromPath, destinationPath)
                    If File.Exists(destinationPath) Then
                        Dim filenameOnly As String = Path.GetFileNameWithoutExtension(destinationPath)
                        Dim fileExtention As String = Path.GetExtension(destinationPath)
                        ''After backup the file rename the actual file name
                        My.Computer.FileSystem.RenameFile(destinationPath, filenameOnly & Format(Date.Now, "yyyyMMddhhmmss") & fileExtention)
                    End If
                    ''MsgBox("File Moved")
                Else
                    MsgBox("File Not move",, "Error")
                End If
            Next
        Catch ex As Exception

        End Try
        Return 0
    End Function
    ''this programme will execute when user click curing combile manue button
    Public Function CuringJobCombine(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim status As Boolean

        status = CuringJobUpdate(connetionString)
        If (status) Then
            MsgBox("Import Successful")
        Else
            MsgBox("Import Fail Please contact Administrator",, MessageBoxIcon.Error)
        End If
        Return 0
    End Function

    ''CuringJobUpdate method will execute K202_CuringJobUpdate_Sp
    Public Function CuringJobUpdate(ByRef connetionString As String) As Boolean
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim status As Boolean
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_CuringJobUpdate_Sp"
            command.CommandTimeout = 600
            command.ExecuteNonQuery()
            Return True
            connection.Close()
        Catch ex As Exception
            Return False
            MsgBox("Extra formers not define",, "error")
        Finally
        End Try
    End Function

    ''TbmOrderPriority_2p2o will execute the K202_TbmOrderPrioritizing_sp_2p2o procedure
    Public Function TbmOrderPriority_2p2o(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_TbmOrderPrioritizing_sp_2p2o"
            command.CommandTimeout = 340
            command.ExecuteNonQuery()

            connection.Close()
        Catch ex As Exception
            MsgBox("TBM Order Priority update Fail Please contact Administrator",, "error")
            stat = -1
        Finally
            IIf(stat = 1, MsgBox("TBM Order Priority Update Complted",, "Information"), stat = 0)
        End Try

        Return 0
    End Function
    ''TbmOrderPriority_2p3o will execute the K202_TbmOrderPrioritizing_sp_2p3o procedure
    Public Function TbmOrderPriority_2p3o(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_TbmOrderPrioritizing_sp_2p3o"
            command.CommandTimeout = 340
            command.ExecuteNonQuery()

            connection.Close()
        Catch ex As Exception
            MsgBox("TBM Order Priority update Fail Please contact Administrator",, "error")
            stat = -1
        Finally
            IIf(stat = 1, MsgBox("TBM Order Priority Update Complted",, "Information"), stat = 0)
            ''MsgBox("TBM Order Priority Update Complted",, "Information")
        End Try

        Return 0
    End Function
    ''TbmOrderPriority_3p2o will execute the K202_TbmOrderPrioritizing_sp_3p2o procedure
    Public Function TbmOrderPriority_3p2o(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_TbmOrderPrioritizing_sp_3p2o"
            command.CommandTimeout = 340
            command.ExecuteNonQuery()

            connection.Close()
        Catch ex As Exception
            MsgBox("TBM Order Priority update Fail Please contact Administrator",, "error")
            stat = -1
        Finally
            IIf(stat = 1, MsgBox("TBM Order Priority Update Complted",, "Information"), stat = 0)
        End Try

        Return 0
    End Function
    '' TbmOrderPriority_All method will execute K202_TBMOrderPrioritizing_SP
    Public Function TbmOrderPriority_All(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_TBMOrderPrioritizing_SP"
            command.CommandTimeout = 340
            command.ExecuteNonQuery()

            connection.Close()
        Catch ex As Exception
            MsgBox("TBM Order Priority update Fail Please contact Administrator",, "error")
            stat = -1
        Finally
            IIf(stat = 1, MsgBox("TBM Order Priority Update Complted",, "Information"), stat = 0)
        End Try
        Return 0
    End Function
    ''ExportDemandFile method will execute K202_GetDemandOrderDetails_Sp execute and convert data table to  CSV file and save to given location
    Public Function ExportDemandFile(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetDemandOrderDetails_Sp"
            Dim param As SqlParameter
            command.CommandTimeout = 340
            adapter = New SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            adapter.Fill(ds)
            connection.Close()

            Dim strData As String
            Dim fileloc As String = "Z:\Export\DemandExport.csv"
            Dim datatbl As DataTable = ds.Tables(0)
            Dim isfirstrow As Integer = 1
            For Each row As DataRow In datatbl.Rows
                Dim line As String = ""
                If isfirstrow = 1 Then
                    For Each column As DataColumn In datatbl.Columns
                        line += "," & (column.ColumnName).ToString()
                    Next
                    strData += line.Substring(1) & vbCrLf
                    isfirstrow = 0
                    line = ""
                End If

                For Each column As DataColumn In datatbl.Columns
                    line += "," & row(column.ColumnName).ToString()
                Next
                strData += line.Substring(1) & vbCrLf
            Next

            If File.Exists(fileloc) Then
                File.Delete(fileloc)
            End If
            Using sw As StreamWriter = New StreamWriter(fileloc)
                sw.WriteLine(strData)
            End Using
        Catch ex As Exception
            stat = -1
            MsgBox("Demand date export error plase contact administrator...",, "error")
        Finally
            IIf(stat = 1, stat = 0, MsgBox("Demand Export Complted",, "Information"))
        End Try

        Return 0
    End Function
    ''========================Aamir/2121-12-18=======================

    ''ExportTBMOrderFile method will execute K202_GetTbmOrderDetails_Sp and convert data table to  CSV file and save to given location
    Public Function ExportTBMOrderFile(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetTbmOrderDetails_Sp"
            Dim param As SqlParameter
            command.CommandTimeout = 340
            adapter = New SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            adapter.Fill(ds)
            connection.Close()

            Dim strData As String
            Dim fileloc As String = "Z:\Export\TBM Order Report.csv"
            Dim datatbl As DataTable = ds.Tables(0)
            Dim isfirstrow As Integer = 1
            For Each row As DataRow In datatbl.Rows
                Dim line As String = ""
                If isfirstrow = 1 Then
                    For Each column As DataColumn In datatbl.Columns
                        line += "," & (column.ColumnName).ToString()
                    Next
                    strData += line.Substring(1) & vbCrLf
                    isfirstrow = 0
                    line = ""
                End If

                For Each column As DataColumn In datatbl.Columns
                    line += "," & row(column.ColumnName).ToString()
                Next
                strData += line.Substring(1) & vbCrLf
            Next

            If File.Exists(fileloc) Then
                File.Delete(fileloc)
            End If
            Using sw As StreamWriter = New StreamWriter(fileloc)
                sw.WriteLine(strData)
            End Using
        Catch ex As Exception
            stat = -1
            System.Diagnostics.Debug.WriteLine(ex)
            MsgBox("TBM report export error plase contact administrator...",, "error")
        Finally
            IIf(stat = 1, stat = 0, MsgBox("Tbm Order File Export Completed",, "Information"))
        End Try

        Return 0
    End Function

    ''ExportMarangooniOrderFile method will execute K202_GetMarangooniOrderDetails_Sp and convert data table to  CSV file and save to given location
    Public Function ExportMarangooniOrderFile(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetMarangooniOrderDetails_Sp"
            Dim param As SqlParameter
            command.CommandTimeout = 340
            adapter = New SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            adapter.Fill(ds)
            connection.Close()

            Dim strData As String
            Dim fileloc As String = "Z:\Export\Marangooni Order Report.csv"
            Dim datatbl As DataTable = ds.Tables(0)
            Dim isfirstrow As Integer = 1
            For Each row As DataRow In datatbl.Rows
                Dim line As String = ""
                If isfirstrow = 1 Then
                    For Each column As DataColumn In datatbl.Columns
                        line += "," & (column.ColumnName).ToString()
                    Next
                    strData += line.Substring(1) & vbCrLf
                    isfirstrow = 0
                    line = ""
                End If

                For Each column As DataColumn In datatbl.Columns
                    line += "," & row(column.ColumnName).ToString()
                Next
                strData += line.Substring(1) & vbCrLf
            Next

            If File.Exists(fileloc) Then
                File.Delete(fileloc)
            End If
            Using sw As StreamWriter = New StreamWriter(fileloc)
                sw.WriteLine(strData)
            End Using
        Catch ex As Exception
            stat = -1
            System.Diagnostics.Debug.WriteLine(ex)
            MsgBox("Marangooni report export error plase contact administrator...",, "error")
        Finally
            IIf(stat = 1, stat = 0, MsgBox("Marangooni Order File Export Completed",, "Information"))
        End Try

        Return 0
    End Function



    ''TBMPriorityMapping method
    Public Function TBMPriorityMapping(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim oForm As New TBMpriorityMapForm()

        oForm.connetionString = connetionString
        oForm.ShowDialog()
    End Function
    ''CreateRankedParentQueue method for rul bulder
    Public Function CreateRankedParentQueue(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object
                                            ) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim ordersParent As Preactor.FormatFieldPair
        Dim partNo As String
        Dim orderNo As String

        Dim x As Integer = 1

        Dim orderCount As Integer = preactor.RecordCount("Orders")
        Dim orderRcNum As Integer
        Dim strCustomQueue As String = "CustomQueue"
        planningboard.CreateQueue(strCustomQueue)

        partNo = preactor.ReadFieldString("Orders", "Part No.", x)
        Do

            If (partNo = "PAB1263") Then

                For i = 0 To 4
                    orderNo = preactor.ReadFieldString("Orders", "Order No.", x)
                    orderRcNum = preactor.FindMatchingRecord("Orders", "Order No.", orderRcNum, orderNo)
                    If ("115" = CStr(preactor.ReadFieldString("Orders", "Op. No.", orderRcNum))) Then
                        i = 4
                    Else
                        orderRcNum = 0
                        i = i + 1
                    End If
                Next i
                planningboard.AddOperationToQueue("CustomQueue", orderRcNum, QueuePosition.End)
            End If

            x = x + 1
        Loop While x <= orderCount
    End Function
    ''CreateRankedParentQueue method for rul bulder
    Private Function CreateRankedParentQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
                                             ByVal ordersTable As Integer, ByVal QName As String) As Integer

        Dim ordersParent As Preactor.FormatFieldPair
        Dim dueDateField As Nullable(Of Preactor.FormatFieldPair)
        Dim priorityField As Nullable(Of Preactor.FormatFieldPair)
        Dim parentRecord As Integer
        Dim SequenceMode As Preactor.SequenceMode

        ordersParent = preactor.FindFirstClassificationString("FAMILY", ordersTable).Value
        dueDateField = preactor.FindFirstClassificationString("DUE DATE", ordersTable)
        priorityField = preactor.FindFirstClassificationString("PRIORITY", ordersTable)

        planningboard.CreateQueue(QName)
        parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
        While (parentRecord > 0)
            If (planningboard.GetOperationLocateState(parentRecord)) Then
                planningboard.AddOperationToQueue(QName, parentRecord, QueuePosition.End)
                parentRecord = preactor.FindMatchingRecord(ordersParent, parentRecord, -1)
            End If
        End While

        SequenceMode = planningboard.SequenceMode
        Select Case SequenceMode.Priority

            Case SequencePriority.DueDate

                If (dueDateField.HasValue) Then
                    planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(dueDateField.Value), QueueRanking.Ascending)
                End If
            Case SequencePriority.Priority
                If (priorityField.HasValue) Then
                    planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(priorityField.Value), QueueRanking.Ascending)
                End If
            Case SequencePriority.ReversePriority
                If (priorityField.HasValue) Then
                    planningboard.RankQueueByFieldName(QName, preactor.GetFieldName(priorityField.Value), QueueRanking.Descending)
                End If

            Case Else
        End Select
        Return 0
    End Function

    ''SimpleAlgorithmicRule1 method for rule development 
    Public Function SimpleAlgorithmicRule1(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        If planningboard Is Nothing Then
            MessageBox.Show("This Rule must be run from the Sequencer")
            Return 0
        End If ' if the planning board wasn't available
        Dim ordersTable As Integer
        Dim operationRecord As Integer
        Dim ResourceRecord As Integer
        Dim ResourceRecords As IEnumerable(Of Integer)
        Dim operationTimes As Nullable(Of Preactor.OperationTimes)

        ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        operationRecord = 0
        CreateRankedParentQueue(preactor, planningboard, ordersTable, "JobsQueue")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, operationRecord))

            planningboard.RemoveOperationFromQueue("JobsQueue", operationRecord)

            While (operationRecord > 0) ' inner loop for operations of the same family

                ResourceRecords = planningboard.FindResources(operationRecord)
                For Each ResourceRecord In ResourceRecords
                    operationTimes = planningboard.TestOperationOnResource(operationRecord, ResourceRecord, planningboard.TerminatorTime)
                    If (operationTimes.HasValue) Then
                        planningboard.PutOperationOnResource(operationRecord, ResourceRecord, operationTimes.Value.ChangeStart)
                        ' if the operation times had a value
                    End If
                    Exit For ' only do this for the first resource in this simple example
                Next ' for each resource record

                operationRecord = planningboard.GetNextOperation(operationRecord, 1)

            End While ' whilst there is another operation
        End While ' whilst there is another operation in the queue

        Return 0
    End Function
    ''SimpleAlgorithmicRule2 method for rule development 
    Public Function SimpleAlgorithmicRule2(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        If planningboard Is Nothing Then
            MessageBox.Show("This Rule must be run from the Sequencer")
            Return 0
        End If ' if the planning board wasn't available
        Dim ordersTable As Integer
        Dim operationRecord As Integer
        Dim ResourceRecord As Integer
        Dim ResourceRecords As IEnumerable(Of Integer)
        Dim operationTimes As Nullable(Of Preactor.OperationTimes)
        Dim dtBestEndTime As Date
        Dim intbestResRec As Integer
        Dim dtbestChangeStart As Date
        ordersTable = preactor.FindFirstClassificationString("LAUNCH TIME").Value.FormatNumber
        operationRecord = 0
        CreateRankedParentQueue(preactor, planningboard, ordersTable, "JobsQueue")

        While (planningboard.GetOperationInQueue("JobsQueue", 1, operationRecord))

            planningboard.RemoveOperationFromQueue("JobsQueue", operationRecord)

            While (operationRecord > 0) ' inner loop for operations of the same family
                dtBestEndTime = DateAdd(DateInterval.Day, 300, planningboard.TerminatorTime)
                intbestResRec = 0
                ResourceRecords = planningboard.FindResources(operationRecord)
                For Each ResourceRecord In ResourceRecords
                    operationTimes = planningboard.TestOperationOnResource(operationRecord, ResourceRecord, planningboard.TerminatorTime)
                    If (operationTimes.HasValue) Then
                        If operationTimes.Value.ProcessEnd < dtBestEndTime Then
                            dtbestChangeStart = operationTimes.Value.ChangeStart
                            dtBestEndTime = operationTimes.Value.ProcessEnd
                            intbestResRec = ResourceRecord
                        End If

                        ' if the operation times had a value
                    End If
                    ''Exit For ' only do this for the first resource in this simple example
                Next ' for each resource record
                If intbestResRec > 0 Then
                    planningboard.PutOperationOnResource(operationRecord, intbestResRec, dtbestChangeStart)
                End If
                operationRecord = planningboard.GetNextOperation(operationRecord, 1)

            End While ' whilst there is another operation
        End While ' whilst there is another operation in the queue

        Return 0
    End Function
    ''=======================Shreyas==========================================





    'Public Function CustomSchduling(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard
    '    Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")


    '    planningboard.CreateQueue("CustomQueue")
    '    Dim orderRcNum As Integer
    '    Dim orderCount As Integer = preactor.RecordCount("Orders")
    '    Dim partNo As String
    '    Dim partNo_ As String
    '    Dim orderNo As String


    '    Dim x As Integer = 1
    '    Dim y As Integer = 1

    '    Do
    '        partNo = preactor.ReadFieldString("Orders", "Part No.", x)
    '        Do
    '            partNo_ = preactor.ReadFieldString("Orders", "Part No.", y)
    '            If partNo = partNo_ Then
    '                For i = 0 To 4
    '                    orderNo = preactor.ReadFieldString("Orders", "Order No.", y)
    '                    orderRcNum = preactor.FindMatchingRecord("Orders", "Order No.", orderRcNum, orderNo)
    '                    If ("115" = CStr(preactor.ReadFieldString("Orders", "Op. No.", orderRcNum))) Then
    '                        i = 4
    '                    Else
    '                        orderRcNum = 0
    '                        i = i + 1
    '                    End If

    '                    planningboard.QuickPutOperationOnResource(orderRcNum,)

    '                Next i
    '            End If
    '            y = y + 1
    '        Loop While y <= orderCount



    '        'For i = 0 To 4
    '        '    orderRcNum = preactor.FindMatchingRecord("Orders", "Order No.", orderRcNum, OrderNo_)
    '        '    If ("115" = CStr(preactor.ReadFieldString("Orders", "Op. No.", orderRcNum))) Then
    '        '        i = 4
    '        '    Else
    '        '        orderRcNum = 0
    '        '        i = i + 1
    '        '    End If
    '        'Next i

    '        ''planningboard.QuickPutOperationOnResource(orderRcNum,)






    '        ''orderRcNum = 0
    '        x = x + 1
    '    Loop While x <= orderCount



    'End Function

    ''Private Function CreateRankedParentQueue(ByRef preactor As IPreactor, ByVal planningboard As IPlanningBoard,
    ''                                        ByVal ordersTable As Integer, ByVal QName As String) As Integer

    'Public Function CreateRankedParentQueue(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object,
    '                                        ByVal ordersTable As Integer, ByVal CustomQueue As String) As Integer
    '''=======================Shreyas==========================================



#End Region
    Public Function FormDataLoad(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        ''define variable and assign
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")

        Dim form1 As Form1 = New Form1()
        form1.form_sp_call(connetionString)

        form1.Show()

        Return 0
    End Function

    Public Function SwitchBetweenClampAndCompression(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef orderNumber As Integer) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim recordNumber As Integer = orderNumber
        Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)

        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim connection As SqlConnection = New SqlConnection(connetionString)
        Dim command As New SqlCommand()
        command.Connection = connection
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "K203_SwitchBetweenClampAndCompression"

        Dim param As SqlParameter
        param = New SqlParameter("@orderNo", orderNo)

        param.Direction = ParameterDirection.Input
        param.DbType = DbType.String
        command.Parameters.Add(param)


        connection.Open()
        command.ExecuteNonQuery()

        connection.Close()
        Return 0
    End Function

    Public Function SwitchBetweenClampAndCompressionNew(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef orderNumber As Integer) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim recordNumber As Integer = orderNumber
        Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)

        Dim operationNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", recordNumber)

        Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", recordNumber)

        Dim num As Integer = preactor.RecordCount("Orders")

        'Dim dt As DataTable = New DataTable()
        'Dim c_orderno As DataColumn = New DataColumn("order No.", Type.[GetType]("System.String"))
        'Dim c_opno As DataColumn = New DataColumn("Op. No.", Type.[GetType]("System.Int32"))
        'Dim c_opname As DataColumn = New DataColumn("Op. Name", Type.[GetType]("System.String"))
        'dt.Columns.Add(c_orderno)
        'dt.Columns.Add(c_opno)
        'dt.Columns.Add(c_opname)


        'If strBelongsToOrderNo = "PARENT" Then
        '    Do
        '        Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
        '        If newOrderNmb = orderNo Then
        '            'Dim dr_int As DataRow = dt.NewRow()

        '            MsgBox("child record")
        '            MsgBox(recordNumber)
        '        Else
        '            Exit Do
        '        End If
        '        recordNumber = recordNumber + 1
        '    Loop While recordNumber <= num
        'Else
        '    MsgBox("Only can do for the parents")
        'End If

        Dim i As Integer = 1
        Do
            Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
            If newOrderNmb = orderNo Then
                'Dim dr_int As DataRow = dt.NewRow()
                'dr_int("order No.") = preactor.ReadFieldString("Orders", "Order No.", i)
                'dr_int("Op. No.") = preactor.ReadFieldString("Orders", "Op. No.", i)
                'dr_int("Op. Name") = preactor.ReadFieldString("Orders", "Operation Name", i)
                'dt.Rows.Add(dr_int)
                Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)

                If OperationName = "ST1-OUTSIDE CURING" Then
                    Dim j As Integer = 1
                    Do
                        Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                        Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                        If secondNewOrderNmb = newOrderNmb Then
                            If secondOperationName = "ST1-CLAMP CURING" Then
                                Dim clampCuringCurrentVal2 As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", j)
                                Dim clampCuringOpNo2 As Integer = preactor.ReadFieldInt("Orders", "Op. No.", j)

                                If clampCuringOpNo2 = 40 Then
                                    preactor.WriteField("Orders", "Disable Operation", j, Not clampCuringCurrentVal2)
                                End If
                            End If
                            If secondOperationName = "ST1-COMPRESSION CURING" Then
                                Dim compressionCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", j)
                                Dim compressionCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", j)
                                Dim newCompressionCuringVal As Boolean = Not compressionCuringCurrentVal

                                If compressionCuringOpNo = 40 Then
                                    MsgBox(compressionCuringCurrentVal)
                                    MsgBox(newCompressionCuringVal)
                                    preactor.WriteField("Orders", "Disable Operation", j, newCompressionCuringVal)
                                    preactor.WriteField("Orders", "Disable Operation", i, newCompressionCuringVal)

                                End If
                            End If

                        End If
                        j = j + 1
                    Loop While j <= num

                    Exit Do
                    'If secondNewOrderNmb = newOrderNmb Then
                    '    If secondOperationName = "ST1-CLAMP CURING" Then
                    '        Dim clampCuringCurrentVal2 As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", j)
                    '    End If

                    'End If
                Else
                    If OperationName = "ST1-COMPRESSION CURING" Then
                        Dim operationNumber40OrdersCount As Integer = 0
                        Dim compressionCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", i)
                        Dim compressionCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)

                        Dim h As Integer = 1
                        Do
                            Dim newOrderNo2 As String = preactor.ReadFieldString("Orders", "Order No.", h)
                            Dim operationNoNew As Integer = preactor.ReadFieldInt("Orders", "Op. No.", h)
                            If newOrderNmb = newOrderNo2 Then
                                If operationNoNew = 40 Then
                                    operationNumber40OrdersCount = operationNumber40OrdersCount + 1
                                End If
                            End If
                            h = h + 1
                        Loop While h <= num

                        If compressionCuringOpNo = 40 Then
                            If operationNumber40OrdersCount >= 2 Then
                                preactor.WriteField("Orders", "Disable Operation", i, Not compressionCuringCurrentVal)
                            End If
                        End If

                    End If

                    If OperationName = "ST1-CLAMP CURING" Then
                        Dim operationNumber40OrdersCount As Integer = 0
                        Dim clampCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", i)
                        Dim clampCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                        Dim compressionCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", i)
                        Dim compressionCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)

                        Dim h As Integer = 1
                        Do
                            Dim newOrderNo2 As String = preactor.ReadFieldString("Orders", "Order No.", h)
                            Dim operationNoNew As Integer = preactor.ReadFieldInt("Orders", "Op. No.", h)
                            If newOrderNmb = newOrderNo2 Then
                                If operationNoNew = 40 Then
                                    operationNumber40OrdersCount = operationNumber40OrdersCount + 1
                                End If
                            End If
                            h = h + 1
                        Loop While h <= num

                        If clampCuringOpNo = 40 Then
                            If operationNumber40OrdersCount >= 2 Then
                                preactor.WriteField("Orders", "Disable Operation", i, Not clampCuringCurrentVal)
                            End If
                        End If

                    End If
                End If


            End If
            i = i + 1
        Loop While i <= num
        preactor.Commit("Orders")

        Return 0
    End Function

    Public Function SelectMultipleRecords(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim y As Integer = 1
        Dim list As New List(Of String)
        Do
            If (planningboard.GetOperationLocateState(y)) Then
                Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", y)
                list.Add(orderNo)
            End If
            y = y + 1
        Loop While y <= num
        MsgBox(list.Count)
        list = list.Distinct.ToList()

        For Each element As String In list
            MsgBox(element)

            Dim i As Integer = 1
            Do
                Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
                If newOrderNmb = element Then

                    Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)

                    If OperationName = "ST1-OUTSIDE CURING" Then
                        MsgBox("Have outside curing")
                        Dim j As Integer = 1
                        Do
                            Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                            Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                            If secondNewOrderNmb = newOrderNmb Then
                                If secondOperationName = "ST1-CLAMP CURING" Then
                                    Dim clampCuringCurrentVal2 As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", j)
                                    Dim clampCuringOpNo2 As Integer = preactor.ReadFieldInt("Orders", "Op. No.", j)

                                    If clampCuringOpNo2 = 40 Then
                                        preactor.WriteField("Orders", "Disable Operation", j, Not clampCuringCurrentVal2)
                                    End If
                                End If
                                If secondOperationName = "ST1-COMPRESSION CURING" Then
                                    Dim compressionCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", j)
                                    Dim compressionCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", j)
                                    Dim newCompressionCuringVal As Boolean = Not compressionCuringCurrentVal
                                    MsgBox("looped")
                                    If compressionCuringOpNo = 40 Then
                                        preactor.WriteField("Orders", "Disable Operation", j, newCompressionCuringVal)
                                        preactor.WriteField("Orders", "Disable Operation", i, newCompressionCuringVal)

                                    End If
                                End If

                            End If
                            j = j + 1
                        Loop While j <= num

                        Exit Do
                    Else
                        If OperationName = "ST1-COMPRESSION CURING" Then
                            Dim compressionCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", i)
                            Dim compressionCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)

                            If compressionCuringOpNo = 40 Then
                                preactor.WriteField("Orders", "Disable Operation", i, Not compressionCuringCurrentVal)
                            End If

                        End If

                        If OperationName = "ST1-CLAMP CURING" Then
                            Dim clampCuringCurrentVal As Boolean = preactor.ReadFieldBool("Orders", "Disable Operation", i)
                            Dim clampCuringOpNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)

                            If clampCuringOpNo = 40 Then
                                preactor.WriteField("Orders", "Disable Operation", i, Not clampCuringCurrentVal)
                            End If

                        End If
                    End If


                End If
                i = i + 1
            Loop While i <= num
        Next
        preactor.Commit("Orders")
        Return 0
    End Function


    Public Function JobSplit(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef RecordNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim strOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", RecordNumber)
        Dim strOrderOprName As String = preactor.ReadFieldString("Orders", "Operation Name", RecordNumber)
        Dim decOrderQty As Double = preactor.ReadFieldDouble("Orders", "Quantity", RecordNumber)
        Dim strSplitJobResource As String = preactor.ReadFieldString("Orders", "Resource", RecordNumber)
        Dim decNewOrderQty As Double = 0
        Dim decBalanceOrderQty As Double = 0
        Dim num As Integer = preactor.RecordCount("Orders")


        If strOrderOprName.Contains("ST3") Then
            Dim K202_ErpOrderNo As String = preactor.ReadFieldString("Orders", "K202_ErpOrderNo", RecordNumber)
            Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
            Dim oForm As New K202_JobSplitDetails()
            Dim serialNumber As Integer = K202_GetSerialsForJobSplit(connetionString, strOrderNo.Substring(1, 1), strOrderNo.Substring(2, 1))
            Dim K203_OrderSerialRecordNumber As Integer = 0

            Dim strSuffix As String = ""
            Dim intSuffixNo As Integer = 0

            '' ----------20-03-2022 Start - To get prefix of split jobs ---------------
            Dim intLengthOfJobOrderNo As Integer = strOrderNo.Length()

            Dim strSearchWithinOrderNo As String = strOrderNo
            Dim strSearchThis As String = "#"
            Dim intCharacterEndIndex As Integer = strOrderNo.IndexOf(strSearchThis)

            Dim strAfterHashOrderNo As String = strOrderNo.Substring(intCharacterEndIndex + 1, intLengthOfJobOrderNo - intCharacterEndIndex - 1)

            '' -------------------Sahan-------------- ''
            Dim newOrderNo As String = strOrderNo.Substring(0, 4)
            'Dim i As Integer = 1
            'Dim max As Integer = 1

            'Do
            '    Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
            '    Dim strFirstFourLetters As String = newOrderNmb.Substring(0, 4)
            '    If strFirstFourLetters = strOrderNo.Substring(0, 4) Then
            '        Dim lengthOfOrderNmb As Integer = strOrderNo.Length
            '        Dim lastFourCharacters As String = newOrderNmb.Substring(lengthOfOrderNmb - 4, 4)
            '        Dim lastFourCharactersToInteger As Integer = Convert.ToInt32(lastFourCharacters)
            '        If lastFourCharactersToInteger >= max Then
            '            max = lastFourCharactersToInteger
            '        End If
            '    End If

            '    i = i + 1
            'Loop While i <= num


            Dim strOrderNoBeforeHash = strOrderNo.Substring(0, intCharacterEndIndex + 1)

            Dim strParentOrderNumber As String = strOrderNoBeforeHash + strAfterHashOrderNo.Substring(0, 1) + "00"

            Dim strFirstCharOfAfterHashOrderNo = strAfterHashOrderNo.Substring(0, 1)

            strSearchThis = "-"
            Dim intCharEndIndex As Integer = strAfterHashOrderNo.IndexOf(strSearchThis)

            Dim strOrderNoAfterHashBeforeSuffix As String = ""
            If intCharEndIndex > 0 Then
                strOrderNoAfterHashBeforeSuffix = strAfterHashOrderNo.Substring(0, intCharEndIndex)
            Else
                strOrderNoAfterHashBeforeSuffix = strAfterHashOrderNo
            End If

            strOrderNo = strOrderNoBeforeHash + strOrderNoAfterHashBeforeSuffix

            K203_OrderSerialRecordNumber = preactor.FindMatchingRecord("K202_OrderSerial", "Order No.", K203_OrderSerialRecordNumber, strOrderNo)
            Dim intSerial As Integer = 0

            If K203_OrderSerialRecordNumber <= 0 Then
                MsgBox("Previous split : 0 times")
            Else
                Dim orderSerialCount As Integer = preactor.RecordCount("K202_OrderSerial")
                Dim a As Integer = 1
                If orderSerialCount > 0 Then
                    Do
                        Dim orderNumberInOrderSerial As String = preactor.ReadFieldString("K202_OrderSerial", "Order No.", a)
                        If orderNumberInOrderSerial = strOrderNo Then
                            Dim serialInOrderSerial As Integer = preactor.ReadFieldInt("K202_OrderSerial", "Serial", a)
                            MsgBox("Previous split : " & serialInOrderSerial & " times")
                        End If
                        a = a + 1
                    Loop While a <= orderSerialCount
                End If
            End If


            intSerial = GetOrderSerialSeq(connetionString, strOrderNo)

            If K203_OrderSerialRecordNumber > 0 Then
                intSerial = intSerial + 1
            Else
                intSerial = 1
            End If

            Dim strNewOrderNo As String = ""

            If intSerial > 99 Then
                MsgBox("Can't split this job. You have already splited 99 times.")
            Else
                If intSerial > 9 Then
                    strSuffix = strFirstCharOfAfterHashOrderNo + intSerial.ToString
                Else
                    strSuffix = strFirstCharOfAfterHashOrderNo + "0" + intSerial.ToString
                End If

                Dim decimalLength As Integer
                decimalLength = 4

                Dim maxStr As String = (serialNumber + 1).ToString("D" + decimalLength.ToString())
                strNewOrderNo = newOrderNo + maxStr
                'strNewOrderNo = strOrderNoBeforeHash + strSuffix

                '' 03-04-2022- Open job order Form
                oForm.strJobOrderNo = strNewOrderNo
                oForm.decJobOrderQty = decOrderQty
                oForm.isOkClick = False

                '' Open the windows form (K203_JobSplitDetails)
                oForm.ShowDialog()
                Try

                    If oForm.isOkClick Then
                        If Not oForm.JobQtyTxt.Text = "" Then
                            decNewOrderQty = CDec(oForm.JobQtyTxt.Text)
                            decBalanceOrderQty = decOrderQty - decNewOrderQty


                            Dim j As Integer = 0



                            'called to function
                            Dim returnNewReferenceOrderNo As String = ChangeCuringRecords(connetionString, decNewOrderQty, strOrderNo, strNewOrderNo)

                            '' Create record for new order number
                            Dim newBlock As Integer = preactor.CreateRecord("Orders")
                            Dim newRecordNum As Integer = preactor.ReadFieldInt("Orders", "Number", newBlock)
                            preactor.CopyRecord("Orders", RecordNumber, newBlock)
                            preactor.WriteField("Orders", "Number", newBlock, newRecordNum)
                            preactor.WriteField("Orders", "Order No.", newBlock, strNewOrderNo)
                            preactor.WriteField("Orders", "Quantity", newBlock, decNewOrderQty)
                            preactor.WriteField("Orders", "Selected Constraint 1", newBlock, -1)
                            preactor.WriteField("Orders", "K202_APSParentNumber", newBlock, strOrderNo)
                            preactor.WriteField("Orders", "K202_ErpOrderNo", newBlock, returnNewReferenceOrderNo)
                            preactor.WriteField("Orders", "Mid Batch Quantity", newBlock, 0)
                            '' Update the Operation sequence - 03-04-2022
                            '' 23-06-2022 - preactor.WriteField("Orders", "Opr. Sequence", newBlock, strSuffix)

                            '' Update the Order Type - 03-04-2022
                            preactor.WriteField("Orders", "Order Type", newBlock, "SPLIT")

                            '' Update the resource for splited job -03-04-2022
                            preactor.WriteField("Orders", "Resource", newBlock, strSplitJobResource)
                            Dim strPartNo As String = preactor.ReadFieldString("Orders", "Part No.", RecordNumber)
                            Dim intOPNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", RecordNumber)

                            '' Update original order quantity (Balance)
                            preactor.WriteField("Orders", "Quantity", RecordNumber, decBalanceOrderQty)


                            K202_UpdateSerialsForJobSplit(connetionString, strOrderNo.Substring(1, 1), strOrderNo.Substring(2, 1))

                            If K203_OrderSerialRecordNumber <= 0 Then
                                '' Create record for K203_OrderSerial
                                Dim newBlock_K203_OrderSerial As Integer = preactor.CreateRecord("K202_OrderSerial")
                                Dim K203_OrderSerialRecordNum As Integer = preactor.ReadFieldInt("K202_OrderSerial", "RecordId", newBlock_K203_OrderSerial)
                                preactor.WriteField("K202_OrderSerial", "Op. No.", newBlock_K203_OrderSerial, intOPNo)
                                preactor.WriteField("K202_OrderSerial", "Order No.", newBlock_K203_OrderSerial, strOrderNo)
                                preactor.WriteField("K202_OrderSerial", "Serial", newBlock_K203_OrderSerial, intSerial)
                                preactor.WriteField("K202_OrderSerial", "ERP Job Num", newBlock_K203_OrderSerial, K202_ErpOrderNo)
                            Else
                                '' Update record for K203_OrderSerial
                                preactor.WriteField("K202_OrderSerial", "Serial", K203_OrderSerialRecordNumber, intSerial)
                            End If

                            ''29-04-2022
                            preactor.Commit("K202_OrderSerial")
                            preactor.Commit("Orders")
                        End If

                        preactor.Redraw()
                    End If

                Catch ex As Exception
                    MsgBox("Error")
                    MsgBox(ex.Message)
                End Try

            End If
        ElseIf strOrderOprName.Contains("ST1") Then

            Dim k As Integer = 1
            Dim max As Integer = 1
            'Dim K202_ErpOrderNo As String = preactor.ReadFieldString("Orders", "K202_ErpOrderNo", RecordNumber)
            Dim K202_APSParentNumber As String = preactor.ReadFieldString("Orders", "K202_APSParentNumber", RecordNumber)
            Do

                Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", k)
                Dim strFirstFourLetters As String = newOrderNmb.Substring(0, 3)
                If strFirstFourLetters = strOrderNo.Substring(0, 3) Then
                    Dim lengthOfOrderNmb As Integer = strOrderNo.Length
                    Dim lastFourCharacters As String = newOrderNmb.Substring(lengthOfOrderNmb - 3, 3)
                    Dim lastFourCharactersToInteger As Integer = Convert.ToInt32(lastFourCharacters)
                    If lastFourCharactersToInteger >= max Then
                        max = lastFourCharactersToInteger
                    End If
                End If

                k = k + 1
            Loop While k <= num

            Dim strNewOrderNo As String = ""
            Dim decimalLength As Integer
            decimalLength = 3
            Dim newOrderNo As String = strOrderNo.Substring(0, 2)

            Dim maxStr As String = (max + 1).ToString("D" + decimalLength.ToString())
            strNewOrderNo = newOrderNo + maxStr

            Dim i As Integer = 1
            Dim oForm As New K202_JobSplitDetails()
            oForm.strJobOrderNo = strNewOrderNo
            oForm.decJobOrderQty = decOrderQty
            oForm.isOkClick = False

            '' Open the windows form (K203_JobSplitDetails)
            oForm.ShowDialog()

            If oForm.isOkClick = True Then

                Dim x As Integer = 1
                Dim belongsToOrderNo As Integer
                Do
                    Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
                    If newOrderNmb = strOrderNo Then

                        Try
                            If Not oForm.JobQtyTxt.Text = "" Then

                                decNewOrderQty = CDec(oForm.JobQtyTxt.Text)
                                decBalanceOrderQty = decOrderQty - decNewOrderQty
                                Dim j As Integer = 0

                                Dim newBlock As Integer = preactor.CreateRecord("Orders")
                                Dim newRecordNum As Integer = preactor.ReadFieldInt("Orders", "Number", newBlock)

                                preactor.CopyRecord("Orders", i, newBlock)

                                preactor.WriteField("Orders", "Number", newBlock, newRecordNum)
                                preactor.WriteField("Orders", "Order No.", newBlock, strNewOrderNo)
                                preactor.WriteField("Orders", "Quantity", newBlock, decNewOrderQty)
                                preactor.WriteField("Orders", "Selected Constraint 1", newBlock, -1)
                                preactor.WriteField("Orders", "Mid Batch Quantity", newBlock, 0)
                                preactor.WriteField("Orders", "K202_APSParentNumber", newBlock, K202_APSParentNumber)
                                preactor.WriteField("Orders", "Order Type", newBlock, "SPLIT")
                                If x = 1 Then
                                    belongsToOrderNo = newRecordNum
                                    x = x + 1
                                Else
                                    preactor.WriteField("Orders", "Belongs to Order No.", newBlock, belongsToOrderNo)
                                End If
                                '' Update the resource for splited job -03-04-2022
                                preactor.WriteField("Orders", "Resource", newBlock, strSplitJobResource)
                                Dim strPartNo As String = preactor.ReadFieldString("Orders", "Part No.", RecordNumber)
                                Dim intOPNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", RecordNumber)
                                preactor.WriteField("Orders", "Quantity", i, decBalanceOrderQty)
                                preactor.Commit("Orders")
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                    i = i + 1
                Loop While i <= num
            End If

            preactor.Commit("Orders")
        End If



        Return 0
    End Function

    Public Function GetOrderSerialSeq(ByRef connetionString As String, ByRef strOrderNo As String) As Integer

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetOrderSerialSeq_Sp"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderNo", strOrderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim OrderSerialSeq As Integer
            param = New SqlParameter("@OrderSerialSeq", OrderSerialSeq)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Double
            '' param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            ''MsgBox("CDec(param.Value) " & CDec(param.Value))
            ''MsgBox("param.Value.ToString " & param.Value.ToString)
            If Not (param.Value.ToString = "") Then
                '' Return CInt(param.Value)
                Return CInt(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Function

    Public Function ChangeCuringRecords(ByRef connetionString As String, ByRef quantity As Double, ByRef orderNo As String, ByRef newOrderNo As String) As String

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_ChangeCuringRecords_Sp"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@quantity", quantity)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int32
            command.Parameters.Add(param)

            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@newOrderNo", newOrderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)


            ''Dim refOrderNo As String = ""
            param = New SqlParameter("@newReferenceOrderNo", SqlDbType.NVarChar, 99)
            param.Direction = ParameterDirection.Output
            '' param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            ''MsgBox(refOrderNo)
            ''MsgBox("CDec(param.Value) " & CDec(param.Value))
            ''MsgBox("param.Value.ToString " & param.Value.ToString)
            If Not (param.Value.ToString = "") Then
                '' Return CInt(param.Value)
                Return param.Value.ToString
            Else
                Return ""
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Function
    ''Befor Milan Modification 20231003
    Function ForceInsideAndOutsideFunction(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef orderNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim recordNumber As Integer = orderNumber
        Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)

        Dim operationNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", recordNumber)

        Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", recordNumber)

        Dim num As Integer = preactor.RecordCount("Orders")

        Dim i As Integer = 1
        Do
            Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
            If newOrderNmb = orderNo Then
                Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)

                If OperationName = "ST1-OUTSIDE CURING" Then
                    Dim OperationNoOfOutsideCuring As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                    Dim j As Integer = 1

                    If OperationNoOfOutsideCuring = 50 Then
                        Dim OperationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", i)
                        Do
                            Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                            Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                            Dim OperationTimePerItemSec As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", j)
                            If secondNewOrderNmb = newOrderNmb Then


                                If secondOperationName = "ST1-COMPRESSION CURING" Then
                                    Dim toggleAttributeOne As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 1", j)
                                    If toggleAttributeOne = 0 Then

                                        Dim newOperationTimePerItem As Double = OperationTimePerItem + OperationTimePerItemSec
                                        preactor.WriteField("Orders", "Duration Attribute 1", j, OperationTimePerItem)
                                        preactor.Commit("Orders")
                                        preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                        MsgBox("Operation Forced inside")
                                        preactor.WriteField("Orders", "Toggle Attribute 1", j, 1)
                                        preactor.WriteField("Orders", "Disable Operation", i, 1)
                                    ElseIf toggleAttributeOne = 1 Then
                                        Dim newOperationTimePerItem As Double = OperationTimePerItemSec - OperationTimePerItem
                                        preactor.WriteField("Orders", "Duration Attribute 1", j, -1)
                                        preactor.Commit("Orders")
                                        preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                        MsgBox("Operation Forced outside")
                                        preactor.WriteField("Orders", "Toggle Attribute 1", j, 0)
                                        preactor.WriteField("Orders", "Disable Operation", i, 0)
                                    End If

                                End If

                            End If
                            j = j + 1
                        Loop While j <= num
                    Else
                        MsgBox("Operation No. of that outside curing is not 50")
                    End If
                    Exit Do
                End If
            End If
            i = i + 1
        Loop While i <= num
        preactor.Commit("Orders")
        Return 0
    End Function
    ''Befor Milan Modification 20231003 End

    ''After Milan Modification 20231003
    Function ForceInsideAndOutsideFunctionBulk(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef orderNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim forceInsideOutsideLookAheadWindowInDays As Integer = preactor.ReadFieldInt("K202_Parameters", "ForceInsideOutsideLookAheadWindowInDays", 1)

        Dim recordNumber As Integer = orderNumber
        Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
        Dim operationNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", recordNumber)
        Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", recordNumber)

        Dim erpOrderNo As String = preactor.ReadFieldString("Orders", "K202_ErpOrderNo", recordNumber)
        Dim partNo As String = preactor.ReadFieldString("Orders", "Part No.", recordNumber)
        Dim startTime As String = preactor.ReadFieldString("Orders", "Start Time", recordNumber)

        Dim num As Integer = preactor.RecordCount("Orders")
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim i As Integer = 1
        Do
            ''If (planningboard.GetOperationLocateState(i)) Then

            Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
            Dim newErpOrderNo As String = preactor.ReadFieldString("Orders", "K202_ErpOrderNo", i)
            Dim newPartNo As String = preactor.ReadFieldString("Orders", "Part No.", i)
            Dim newStartTime As DateTime = preactor.ReadFieldDateTime("Orders", "Start Time", i)
            Dim newResource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)
            If newOrderNmb = orderNo Then
                Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)
                If OperationName = "ST1-OUTSIDE CURING" Then
                    Dim OperationNoOfOutsideCuring As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                    Dim j As Integer = 1

                    If OperationNoOfOutsideCuring = 50 Then
                        Dim OperationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", i)
                        Do
                            Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                            Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                            Dim OperationTimePerItemSec As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", j)

                            Dim secondNewErpOrderNo As String = preactor.ReadFieldString("Orders", "K202_ErpOrderNo", j)
                            Dim secondNewPartNo As String = preactor.ReadFieldString("Orders", "Part No.", j)
                            Dim secondNewStartTime As DateTime = preactor.ReadFieldDateTime("Orders", "Start Time", j)
                            Dim secondNewResource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)


                            ''If secondNewOrderNmb = newOrderNmb Then ''old conndition

                            ''If (newErpOrderNo = secondNewErpOrderNo) And (newResource = secondNewResource) And (secondNewPartNo = secondNewPartNo) And (newStartTime < secondNewStartTime) And (secondNewStartTime <= DateAdd(DateInterval.Month, forceInsideOutsideLookAheadWindowInDays * 24, newStartTime)) Then
                            If (newErpOrderNo = secondNewErpOrderNo) And (newResource = secondNewResource) And (secondNewPartNo = secondNewPartNo) And (newStartTime < secondNewStartTime) And (secondNewStartTime <= DateAdd(DateInterval.Month, forceInsideOutsideLookAheadWindowInDays * 24, newStartTime)) Then
                                If secondOperationName = "ST1-COMPRESSION CURING" Then
                                    Dim toggleAttributeOne As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 1", j)
                                    If toggleAttributeOne = 0 Then

                                        Dim newOperationTimePerItem As Double = OperationTimePerItem + OperationTimePerItemSec
                                        preactor.WriteField("Orders", "Duration Attribute 1", j, OperationTimePerItem)
                                        preactor.Commit("Orders")
                                        preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                        ''MsgBox("Operation Forced inside")
                                        preactor.WriteField("Orders", "Toggle Attribute 1", j, 1)
                                        preactor.WriteField("Orders", "Disable Operation", i, 1)
                                    ElseIf toggleAttributeOne = 1 Then
                                        Dim newOperationTimePerItem As Double = OperationTimePerItemSec - OperationTimePerItem
                                        preactor.WriteField("Orders", "Duration Attribute 1", j, -1)
                                        preactor.Commit("Orders")
                                        preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                        ''MsgBox("Operation Forced outside")
                                        preactor.WriteField("Orders", "Toggle Attribute 1", j, 0)
                                        preactor.WriteField("Orders", "Disable Operation", i, 0)
                                    End If
                                End If
                            End If
                            j = j + 1
                        Loop While j <= num
                    Else
                        MsgBox("Operation No. of that outside curing is not 50")
                    End If
                    Exit Do
                End If
                Exit Do
            End If
            ''End If
            ''End If

            i = i + 1
        Loop While i <= num
        preactor.Commit("Orders")
        Return 0
    End Function
    ''After Milan Modification 20231003 End

    Function ForceInsideMultipleSelection(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim num As Integer = preactor.RecordCount("Orders")

        Dim s As Integer = 1
        Do
            If (planningboard.GetOperationLocateState(s)) Then
                Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", s)
                '' Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", s)

                Dim i As Integer = 1
                Do
                    Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
                    If newOrderNmb = orderNo Then
                        Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)

                        If OperationName = "ST1-OUTSIDE CURING" Then
                            Dim OperationNoOfOutsideCuring As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                            Dim j As Integer = 1

                            If OperationNoOfOutsideCuring = 50 Then
                                Dim OperationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", i)
                                Do
                                    Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                                    Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                                    Dim OperationTimePerItemSec As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", j)
                                    If secondNewOrderNmb = newOrderNmb Then


                                        If secondOperationName = "ST1-COMPRESSION CURING" Then
                                            Dim toggleAttributeOne As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 1", j)
                                            If toggleAttributeOne = 0 Then

                                                Dim newOperationTimePerItem As Double = OperationTimePerItem + OperationTimePerItemSec
                                                preactor.WriteField("Orders", "Duration Attribute 1", j, OperationTimePerItem)
                                                preactor.Commit("Orders")
                                                preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                                ''MsgBox("Operation Forced inside")
                                                preactor.WriteField("Orders", "Toggle Attribute 1", j, 1)
                                                preactor.WriteField("Orders", "Disable Operation", i, 1)
                                                'ElseIf toggleAttributeOne = 1 Then
                                                '    Dim newOperationTimePerItem As Double = OperationTimePerItemSec - OperationTimePerItem
                                                '    preactor.WriteField("Orders", "Duration Attribute 1", j, -1)
                                                '    preactor.Commit("Orders")
                                                '    preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                                '    MsgBox("Operation Forced outside")
                                                '    preactor.WriteField("Orders", "Toggle Attribute 1", j, 0)
                                                '    preactor.WriteField("Orders", "Disable Operation", i, 0)
                                            End If

                                        End If

                                    End If
                                    j = j + 1
                                Loop While j <= num
                            Else
                                MsgBox("Operation No. of that outside curing is not 50")
                            End If
                            Exit Do
                        End If
                    End If
                    i = i + 1
                Loop While i <= num

            End If
            s = s + 1
        Loop While s <= num


        '    End If
        '    y = y + 1
        'Loop While y <= num


        preactor.Commit("Orders")
        Return 0
    End Function

    Function ForceOutsideMultipleSelection(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim num As Integer = preactor.RecordCount("Orders")

        Dim s As Integer = 1
        Do
            If (planningboard.GetOperationLocateState(s)) Then
                Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", s)
                '' Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", s)

                Dim i As Integer = 1
                Do
                    Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
                    If newOrderNmb = orderNo Then
                        Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)

                        If OperationName = "ST1-OUTSIDE CURING" Then
                            Dim OperationNoOfOutsideCuring As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                            Dim j As Integer = 1

                            If OperationNoOfOutsideCuring = 50 Then
                                Dim OperationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", i)
                                Do
                                    Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                                    Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                                    Dim OperationTimePerItemSec As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", j)
                                    If secondNewOrderNmb = newOrderNmb Then


                                        If secondOperationName = "ST1-COMPRESSION CURING" Then
                                            Dim toggleAttributeOne As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 1", j)
                                            'If toggleAttributeOne = 0 Then

                                            '    Dim newOperationTimePerItem As Double = OperationTimePerItem + OperationTimePerItemSec
                                            '    preactor.WriteField("Orders", "Duration Attribute 1", j, OperationTimePerItem)
                                            '    preactor.Commit("Orders")
                                            '    preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                            '    MsgBox("Operation Forced inside")
                                            '    preactor.WriteField("Orders", "Toggle Attribute 1", j, 1)
                                            '    preactor.WriteField("Orders", "Disable Operation", i, 1)
                                            'Else
                                            If toggleAttributeOne = 1 Then
                                                    Dim newOperationTimePerItem As Double = OperationTimePerItemSec - OperationTimePerItem
                                                    preactor.WriteField("Orders", "Duration Attribute 1", j, -1)
                                                    preactor.Commit("Orders")
                                                    preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                                '' MsgBox("Operation Forced outside")
                                                preactor.WriteField("Orders", "Toggle Attribute 1", j, 0)
                                                    preactor.WriteField("Orders", "Disable Operation", i, 0)
                                                End If

                                            End If

                                    End If
                                    j = j + 1
                                Loop While j <= num
                            Else
                                MsgBox("Operation No. of that outside curing is not 50")
                            End If
                            Exit Do
                        End If
                    End If
                    i = i + 1
                Loop While i <= num

            End If
            s = s + 1
        Loop While s <= num


        '    End If
        '    y = y + 1
        'Loop While y <= num


        preactor.Commit("Orders")
        Return 0
    End Function


    Function ForceInsideAndOutsideFunctionForMultipleRecords(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim y As Integer = 1
        Dim list As New List(Of String)
        Do
            If (planningboard.GetOperationLocateState(y)) Then
                Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", y)
                list.Add(orderNo)
            End If
            y = y + 1
        Loop While y <= num
        list = list.Distinct.ToList()

        For Each element As String In list

            Dim i As Integer = 1
            Do
                Dim newOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", i)
                If newOrderNmb = element Then
                    Dim OperationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)

                    If OperationName = "ST1-OUTSIDE CURING" Then
                        Dim OperationNoOfOutsideCuring As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                        Dim j As Integer = 1

                        If OperationNoOfOutsideCuring = 50 Then
                            Dim OperationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", i)
                            Do
                                Dim secondNewOrderNmb As String = preactor.ReadFieldString("Orders", "Order No.", j)
                                Dim secondOperationName As String = preactor.ReadFieldString("Orders", "Operation Name", j)
                                Dim OperationTimePerItemSec As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", j)
                                If secondNewOrderNmb = newOrderNmb Then


                                    If secondOperationName = "ST1-COMPRESSION CURING" Then
                                        Dim toggleAttributeOne As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 1", j)
                                        If toggleAttributeOne = 0 Then

                                            Dim newOperationTimePerItem As Double = OperationTimePerItem + OperationTimePerItemSec
                                            preactor.WriteField("Orders", "Duration Attribute 1", j, OperationTimePerItem)
                                            preactor.Commit("Orders")
                                            preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                            MsgBox("Operation Forced inside")
                                            preactor.WriteField("Orders", "Toggle Attribute 1", j, 1)
                                            preactor.WriteField("Orders", "Disable Operation", i, 1)
                                        ElseIf toggleAttributeOne = 1 Then
                                            Dim newOperationTimePerItem As Double = OperationTimePerItemSec - OperationTimePerItem
                                            preactor.WriteField("Orders", "Duration Attribute 1", j, -1)
                                            preactor.Commit("Orders")
                                            preactor.WriteField("Orders", "Op. Time per Item", j, newOperationTimePerItem)
                                            MsgBox("Operation Forced outside")
                                            preactor.WriteField("Orders", "Toggle Attribute 1", j, 0)
                                            preactor.WriteField("Orders", "Disable Operation", i, 0)
                                        End If

                                    End If

                                End If
                                j = j + 1
                            Loop While j <= num
                        Else
                            MsgBox("Operation No. of that outside curing is not 50")
                        End If
                        Exit Do
                    End If
                End If
                i = i + 1
            Loop While i <= num
            preactor.Commit("Orders")

        Next

        Return 0
    End Function

    Function AdjustCuringStartAndEndForMultipleRecords(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        MsgBox("AdjustCuringStartAndEndForMultipleRecords")
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim y As Integer = 1
        Dim maxEndDate As Date = #9/23/1899 01:00 AM#
        Dim recordNumber As Integer
        Dim selectedOperationsCount As Integer = 0
        Dim selectedOperations As New List(Of Integer)
        Do
            If (planningboard.GetOperationLocateState(y)) Then
                Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", y)
                Dim endDateOfY As Date = preactor.ReadFieldDateTime("Orders", "End Time", y)
                If operationName.Contains("COMPRESSION") Then
                    selectedOperations.Add(y)
                    selectedOperationsCount = selectedOperationsCount + 1
                    If endDateOfY > maxEndDate Then
                        maxEndDate = endDateOfY
                        recordNumber = y
                    End If
                End If
            End If
            y = y + 1
        Loop While y <= num

        Dim strOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
        Dim strOrderOprName As String = preactor.ReadFieldString("Orders", "Operation Name", recordNumber)
        Dim endDate As Date = preactor.ReadFieldDateTime("Orders", "End Time", recordNumber)
        Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)

        Dim newStartDate As Date = startDate.AddMinutes(-30)
        Dim newEndDate As Date = endDate.AddMinutes(30)
        Dim resource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)
        Dim i As Integer = 1

        Dim table As DataTable = New DataTable()
        table.Columns.Add("ID", GetType(Integer))
        table.Columns.Add("Start_date", GetType(Date))
        Do
            Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)
            If operationName.Contains("COMPRESSION") Then
                Dim endDate2 As Date = preactor.ReadFieldDateTime("Orders", "End Time", i)
                Dim startDate2 As Date = preactor.ReadFieldDateTime("Orders", "Start Time", i)
                Dim resource2 As String = preactor.ReadFieldString("Orders", "Resource", i)
                If (endDate2 <= newEndDate And resource2 = resource) Then
                    table.Rows.Add(i, startDate2)
                End If
            End If
            i = i + 1
        Loop While i <= num
        table.DefaultView.Sort = "Start_date ASC"
        table = table.DefaultView.ToTable
        Dim rowsCount As Integer = table.Rows.Count

        For Each row As DataRow In table.Rows
            Dim rowId As Integer = Convert.ToInt32(row("ID"))
            Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", rowId)
            If orderNo.Substring(0, 5) = strOrderNo.Substring(0, 5) Then

                ' MsgBox("ID :" & row("ID").ToString & " Order No :" & orderNo & "  Start_date :" & row("Start_date").ToString)
                Dim strOrderNo2 As String = preactor.ReadFieldString("Orders", "Order No.", rowId)
                Dim strOrderOprName2 As String = preactor.ReadFieldString("Orders", "Operation Name", rowId)
                Dim num2 As Integer = preactor.RecordCount("Orders")

                Dim endDate2 As Date = preactor.ReadFieldDateTime("Orders", "End Time", rowId)
                Dim startDate2 As Date = preactor.ReadFieldDateTime("Orders", "Start Time", rowId)

                Dim newStartDate2 As Date = startDate2.AddMinutes(-30)
                Dim newEndDate2 As Date = endDate2.AddMinutes(30)
                Dim resource2 As String = preactor.ReadFieldString("Orders", "Resource", rowId)
                Dim j As Integer = 1
                Dim listOfOperations As New List(Of Integer)
                Do
                    Dim endDateOfJ As Date = preactor.ReadFieldDateTime("Orders", "End Time", j)
                    Dim startDateOfJ As Date = preactor.ReadFieldDateTime("Orders", "Start Time", j)
                    Dim resourceOfJ As String = preactor.ReadFieldString("Orders", "Resource", j)
                    If endDateOfJ >= newStartDate2 And endDateOfJ <= newEndDate2 And resource2 = resource Then
                        listOfOperations.Add(j)
                    End If
                    j = j + 1
                Loop While j <= num

                Dim maxEndDateNew As Date = #9/23/1899 01:00 AM#
                For Each record As Integer In listOfOperations
                    Dim endDateOfRecord As Date = preactor.ReadFieldDateTime("Orders", "End Time", record)
                    If maxEndDateNew < endDateOfRecord Then
                        maxEndDateNew = endDateOfRecord
                    End If
                Next

                For Each record As Integer In listOfOperations
                    Dim endDateOfRecordd As Date = preactor.ReadFieldDateTime("Orders", "End Time", record)
                    Dim timeDifference As TimeSpan = maxEndDateNew.Subtract(endDateOfRecordd)
                    Dim setupTime As Double = preactor.ReadFieldDouble("Orders", "Setup Time", record)
                    Dim totalSetupTime As Double = setupTime + timeDifference.TotalDays
                    preactor.WriteField("Orders", "Setup Time", record, totalSetupTime)
                Next
            End If
        Next
        preactor.Commit("Orders")
        Return 0
    End Function

    Function ImportOrdersDataToTable(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As String
        Try
            Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "InsertDataToOrdersTable"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            ''Dim refOrderNo As String = ""
            param = New SqlParameter("@Difference", SqlDbType.Int)
            param.Direction = ParameterDirection.Output
            '' param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            MsgBox("Successfully added " & param.Value.ToString & " records.")
            ''MsgBox(refOrderNo)
            ''MsgBox("CDec(param.Value) " & CDec(param.Value))
            ''MsgBox("param.Value.ToString " & param.Value.ToString)
            If Not (param.Value.ToString = "") Then
                '' Return CInt(param.Value)
                Return param.Value.ToString
            Else
                Return ""
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try
    End Function

    Function GetSplitQuantity(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef RecordNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim customClickTime As DateTime
        customClickTime = planningboard.CustomClickTime

        Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", RecordNumber)
        Dim endDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", RecordNumber)
        Dim midBatchTime As Date = preactor.ReadFieldDateTime("Orders", "Mid Batch Time", RecordNumber)
        Dim midBatchQuantity As Integer = preactor.ReadFieldInt("Orders", "Mid Batch Quantity", RecordNumber)
        Dim quantity As Integer = preactor.ReadFieldInt("Orders", "quantity", RecordNumber)

        MsgBox("Mid BatchTime " & midBatchTime)
        MsgBox("Mid Batch Quantity " & midBatchQuantity)

        Dim OrderQty As Double = preactor.ReadFieldDouble("Orders", "Quantity", RecordNumber)
        Dim splitQuantity As Double = (OrderQty - midBatchQuantity)
        MsgBox("Start Time " & startDate)
        MsgBox("Custom Click Time " & customClickTime)

        Return 0
    End Function


    Function ShowDetails(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim customClickTime As DateTime
        customClickTime = planningboard.CustomClickTime
        Dim demandCount As Integer = preactor.RecordCount("Demand")
        Dim ordersCount As Integer = preactor.RecordCount("Orders")
        Dim productCount As Integer = preactor.RecordCount("Products")
        Dim supplyCount As Integer = preactor.RecordCount("Supply")
        Dim k202_ActualProductionRecordsCount As Integer = preactor.RecordCount("K202_ActualProductionRecords")
        Dim secondaryConstraintGroupsCount As Integer = preactor.RecordCount("Secondary Constraint Groups")
        Dim secondaryConstraintsCount As Integer = preactor.RecordCount("Secondary Constraints")
        Dim secondaryCalenderPeriodsCount As Integer = preactor.RecordCount("Secondary Calendar Periods")
        Dim pb As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim i As Integer = 1
        Dim dt As DataTable = New DataTable()

        Dim currentDate As Date = Now
        Dim daysInMonth As Integer = Date.DaysInMonth(currentDate.Year, currentDate.Month)
        Dim currentMonthName As String = DateTime.Now.ToString("MMMM")
        Dim currentDateWithNumber As Integer = CInt(DateTime.Now.ToString("dd"))
        Dim currentYear As Integer = CInt(DateTime.Now.ToString("yyyy"))

        'Create SKU as first row.
        dt.Columns.Add(New DataColumn("Description", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("SKU", Type.[GetType]("System.String")))
        Dim k As Integer = 1
        Dim demandList As New List(Of String)
        Dim dt_row As DataRow
        dt_row = dt.NewRow()
        dt.Rows.Add(dt_row)
        If demandCount > 0 Then
            Do
                Dim partNo As String = preactor.ReadFieldString("Demand", "Part No.", k)
                Dim description As String = ""
                Dim a As Integer = 1
                Do
                    Dim partN = preactor.ReadFieldString("Products", "Part No.", a)
                    Dim isParent = preactor.ReadFieldInt("Products", "Parent Part", a)
                    If partN = partNo And isParent = -1 Then
                        Dim desc As String = preactor.ReadFieldString("Products", "Product", a)
                        description = desc
                    End If
                    a = a + 1
                Loop While a <= productCount

                If description = "" Then
                    description = "Product not defined."
                End If

                If Not demandList.Contains(partNo) Then
                    dt_row = dt.NewRow()
                    dt_row("Description") = description
                    dt_row("SKU") = partNo
                    dt.Rows.Add(dt_row)
                    demandList.Add(partNo)
                End If
                k = k + 1
            Loop While k <= demandCount
        End If

        Try
            For Each element As String In demandList
                Dim isPartNoIncludeInProductTable As Boolean = False
                Dim s As Integer = 1
                Do
                    Dim partNoInProductTable As String = preactor.ReadFieldString("Products", "Part No.", s)
                    If partNoInProductTable = element Then
                        isPartNoIncludeInProductTable = True
                    End If
                    s = s + 1
                Loop While s <= productCount
                If isPartNoIncludeInProductTable = False Then
                    MsgBox("No relavant " & element & " record defined in product table.")
                End If
            Next
        Catch ex As Exception

        End Try



        'demandList = demandList.Distinct.ToList()

        'Create orders list as column
        Dim j As Integer = 1
        Dim orderList As New List(Of String)
        If demandCount > 0 Then
            Do
                Dim OrderNo As String = preactor.ReadFieldString("Demand", "Order No.", j)
                Dim isParent As Integer = preactor.ReadFieldInt("Demand", "Belongs to Order No.", j)
                If isParent = -1 Then
                    orderList.Add(OrderNo)
                End If
                j = j + 1
            Loop While j <= demandCount
        End If

        orderList = orderList.Distinct.ToList()
        Dim markStatus As New NameValueCollection

        For Each element As String In orderList
            dt.Columns.Add(New DataColumn(element, Type.[GetType]("System.String")))
        Next

        For Each row As DataRow In dt.Rows
            For Each ele In orderList
                Dim maxDemandDate As Date = #9/23/1899 01:00 AM#
                Dim m As Integer = 1
                Dim u As Integer = 1
                Dim sumOfSupplyQty As Integer = 0
                Dim sumOfDemandQty As Integer = 0

                If demandCount > 0 Then
                    Do
                        Dim salesOrderNo As String = preactor.ReadFieldString("Demand", "Order No.", m)
                        If salesOrderNo = ele Then
                            Dim demandQty As Integer = preactor.ReadFieldInt("Demand", "Quantity", m)
                            sumOfDemandQty = sumOfDemandQty + demandQty
                        End If
                        m = m + 1

                    Loop While m <= demandCount
                End If
                If supplyCount > 0 Then
                    Do
                        Dim salesOrderNo As String = preactor.ReadFieldString("Supply", "Order No.", u)
                        If salesOrderNo = ele Then
                            Dim supplyQty As Integer = preactor.ReadFieldInt("Supply", "Quantity", u)
                            sumOfSupplyQty = sumOfSupplyQty + supplyQty
                        End If
                        u = u + 1
                    Loop While u <= supplyCount
                End If

                If sumOfSupplyQty = sumOfDemandQty Then
                    row(ele) = GetShipmentDate(connetionString, ele)
                Else
                    Dim endOfYear As String = currentYear & "/12" & "/31"
                    row(ele) = endOfYear
                End If

            Next

            Exit For
        Next
        dt.Columns.Add(New DataColumn("Rejected Qty", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("On Hand", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("On Plan", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Running Qty", Type.[GetType]("System.String")))


        dt.DefaultView.ToTable()

        For c As Integer = 0 To dt.Columns.Count - 1
            If Not (dt.Columns(c).ColumnName = "SKU" Or dt.Columns(c).ColumnName = "Description") Then
                Dim skucol As String = ""
                Dim skuRow As String = ""
                'For Each row As DataRow In dt.Rows
                '    skucol = dt.Columns(c).ColumnName

                '    Exit For
                'Next
                For Each row As DataRow In dt.Rows
                    skuRow = row("SKU").ToString
                    skucol = dt.Columns(c).ColumnName
                    Dim d As Integer = 1
                    If demandCount > 0 Then
                        Do
                            Dim partNoDemand As String = preactor.ReadFieldString("Demand", "Part No.", d)
                            Dim orderNoDemand As String = preactor.ReadFieldString("Demand", "Order No.", d)
                            Dim qty As Integer = 0
                            If Not row(skucol) Is DBNull.Value Then
                                If skuRow = partNoDemand And skucol = orderNoDemand Then
                                    qty = preactor.ReadFieldInt("Demand", "Quantity", d)
                                    row(skucol) = (CInt(row(skucol)) + qty).ToString
                                End If
                            Else
                                If skuRow = partNoDemand And skucol = orderNoDemand Then
                                    qty = preactor.ReadFieldInt("Demand", "Quantity", d)
                                    row(skucol) = qty.ToString
                                End If
                            End If
                            d = d + 1
                        Loop While d <= demandCount
                    End If

                    'a = a + 1
                Next
            End If
        Next
        dt.Columns.Add("Monthly Demand", Type.[GetType]("System.String"))


        Dim answer = currentDate.AddDays(-(currentDateWithNumber) + 1)
        Dim n As Integer = 1
        Do
            Dim datesInMonth As String = currentMonthName + " " + n.ToString
            dt.Columns.Add(New DataColumn(datesInMonth, Type.[GetType]("System.String")))
            n = n + 1
        Loop While n <= currentDateWithNumber

        'Put quantity under the each date in the month.
        For c As Integer = 0 To dt.Rows.Count - 1
            If c >= 1 Then
                Dim p As Integer = 1
                Do

                    Dim datesInMonth As String = currentMonthName + " " + p.ToString
                    Dim newDate = currentDate.AddDays(-(currentDateWithNumber) + p)
                    Dim eachSku As String = dt.Rows(c)("SKU").ToString
                    Dim q As Integer = 1
                    Dim sumOfQuantity As Integer = 0
                    If ordersCount > 0 Then
                        Do
                            Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", q)
                            Dim partNo As String = preactor.ReadFieldString("Orders", "Part No.", q)
                            Dim startTime As Date = preactor.ReadFieldDateTime("Orders", "Start Time", q)
                            If (startTime.Date = newDate.Date And partNo = eachSku) Then
                                Dim quantityOfOrder As Integer = preactor.ReadFieldInt("Orders", "Quantity", q)
                                sumOfQuantity = sumOfQuantity + quantityOfOrder
                            End If
                            q = q + 1
                        Loop While q <= ordersCount
                    End If


                    dt.Rows(c)(datesInMonth) = sumOfQuantity
                    p = p + 1
                Loop While p <= currentDateWithNumber
            End If
        Next

        dt.Columns.Add(New DataColumn("Access Qty", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Planning Qty", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("No Of CLM", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Available CLM", Type.[GetType]("System.Double")))
        dt.Columns.Add(New DataColumn("No Of CMP", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Available CMP", Type.[GetType]("System.Double")))
        dt.Columns.Add(New DataColumn("Order Qty CLM", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Order Qty CMP", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Balance Qty", Type.[GetType]("System.String")))
        dt.Columns.Add(New DataColumn("Urgency", Type.[GetType]("System.String")))
        'Dim checkBox As New DataGridViewCheckBoxColumn()
        'checkBox.HeaderText = "Is Updated"
        'dt.Columns.Add(checkBox)
        'dt.Columns.Add(New DataColumn("Is Updated", Type.[GetType]("System.Boolean")))

        For c As Integer = 0 To dt.Rows.Count - 1
            Dim sumQty As Integer = 0
            If c >= 1 Then
                dt.Rows(c)("Access Qty") = 0
                For d As Integer = 0 To dt.Columns.Count - 1
                    If Not (dt.Columns(d).ColumnName = "SKU" Or dt.Columns(d).ColumnName = "Description" Or dt.Columns(d).ColumnName = "Rejected Qty" Or
                        dt.Columns(d).ColumnName = "Planning Qty" Or dt.Columns(d).ColumnName = "No Of CLM" Or dt.Columns(d).ColumnName = "Access Qty" Or
                        dt.Columns(d).ColumnName = "No Of CMP" Or dt.Columns(d).ColumnName = "Order Qty CLM" Or
                        dt.Columns(d).ColumnName = "Order Qty CMP" Or dt.Columns(d).ColumnName = "On Hand" Or
                        dt.Columns(d).ColumnName = "On Plan" Or dt.Columns(d).ColumnName = "Running Qty" Or dt.Columns(d).ColumnName.Contains(currentMonthName)) Then
                        If Not dt.Rows(c)(dt.Columns(d)) Is DBNull.Value Then

                            sumQty = sumQty + CInt(dt.Rows(c)(dt.Columns(d).ColumnName))
                        End If

                    End If
                Next
                dt.Rows(c)("Monthly Demand") = sumQty
            End If
        Next


        For c As Integer = 0 To dt.Rows.Count - 1
            If c >= 1 Then
                Dim sku As String = dt.Rows(c).Item("SKU").ToString
                Dim b As Integer = 1
                Dim d As Integer = 1
                Dim e As Integer = 1

                If productCount > 0 Then
                    Do
                        Dim partNoNew As String = preactor.ReadFieldString("Products", "Part No.", b)
                        If partNoNew = sku Then
                            Dim OpNo As Integer = preactor.ReadFieldInt("Products", "Op. No.", b)
                            Dim OpName As String = preactor.ReadFieldString("Products", "Operation Name", b)
                            Dim constraintsGroup1 As Integer = 0
                            constraintsGroup1 = preactor.ReadFieldInt("Products", "Constraint Group 1", b)
                            If constraintsGroup1 = -1 Then
                                constraintsGroup1 = 0
                            End If
                            '' planningboard.gets
                            If OpNo = 40 And OpName.Contains("CLAMP CURING") Then
                                Dim maxMold As Integer = GetMaximumMoldNumber(connetionString, partNoNew, "CLAMP")
                                dt.Rows(c)("No Of CLM") = maxMold
                            End If

                            If OpNo = 40 And OpName.Contains("COMPRESSION CURING") Then
                                Dim maxMold As Integer = GetMaximumMoldNumber(connetionString, partNoNew, "COMPRESSION")
                                dt.Rows(c)("No Of CMP") = maxMold
                            End If
                        End If
                        b = b + 1
                    Loop While b <= productCount
                End If

                Dim sumOnPlanQty As Integer = 0
                If ordersCount > 0 Then
                    Do
                        Dim partNoNew As String = preactor.ReadFieldString("Orders", "Part No.", d)
                        If partNoNew = sku Then
                            Dim qty As Integer = preactor.ReadFieldInt("Orders", "Quantity", d)
                            Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", d)
                            Dim opNoOfOrder As Integer = preactor.ReadFieldInt("Orders", "Op. No.", d)
                            Dim resourceOfOrder As Integer = preactor.ReadFieldInt("Orders", "Resource", d)
                            If opNoOfOrder = 40 And Not resourceOfOrder = -1 Then
                                sumOnPlanQty = sumOnPlanQty + qty
                            End If
                        End If
                        d = d + 1
                    Loop While d <= ordersCount
                End If

                dt.Rows(c)("On Plan") = sumOnPlanQty

                Dim sumOnHand As Integer = 0
                If supplyCount > 0 Then
                    Do
                        Dim partNoNew As String = preactor.ReadFieldString("Supply", "Part No.", e)
                        If partNoNew = sku Then
                            Dim qty As Integer = preactor.ReadFieldInt("Supply", "Quantity", e)
                            sumOnHand = sumOnHand + qty
                        End If
                        e = e + 1
                    Loop While e <= supplyCount
                End If

                e = 1
                Dim sumRejectedQty As Integer = 0
                If k202_ActualProductionRecordsCount > 0 Then
                    Do
                        Dim partNoNew As String = preactor.ReadFieldString("K202_ActualProductionRecords", "Part No.", e)
                        Dim OpNo As Integer = preactor.ReadFieldInt("K202_ActualProductionRecords", "Op. No.", e)
                        Dim actualProdQty As Integer = preactor.ReadFieldInt("K202_ActualProductionRecords", "Quantity", e)
                        Dim Grade As String = preactor.ReadFieldString("K202_ActualProductionRecords", "Grade", e)
                        If partNoNew = sku And OpNo = 40 And Grade = "C2" Then
                            sumRejectedQty = sumRejectedQty + actualProdQty
                        End If
                        e = e + 1
                    Loop While e <= k202_ActualProductionRecordsCount
                End If

                If sumRejectedQty = 0 Then
                    dt.Rows(c)("Rejected Qty") = 0
                Else
                    dt.Rows(c)("Rejected Qty") = sumRejectedQty
                End If

                dt.Rows(c)("On Hand") = sumOnHand

                Dim runningQty As Integer = CInt(dt.Rows(c)("On Plan")) + CInt(dt.Rows(c)("On Hand"))
                dt.Rows(c)("Running Qty") = runningQty

                Dim planningQty As Integer = CInt(dt.Rows(c)("Monthly Demand")) - CInt(dt.Rows(c)("Running Qty")) - CInt(dt.Rows(c)("Rejected Qty"))
                dt.Rows(c)("Planning Qty") = planningQty
            End If
        Next

        For c As Integer = 0 To dt.Rows.Count - 1
            If c >= 1 Then
                Dim sku As String = dt.Rows(c).Item("SKU").ToString
                Dim f As Integer = 1
                If productCount > 0 Then
                    Try
                        Do
                            Dim partNoNew As String = preactor.ReadFieldString("Products", "Part No.", f)
                            If partNoNew = sku Then
                                Dim OpNo As Integer = preactor.ReadFieldInt("Products", "Op. No.", f)
                                Dim OpName As String = preactor.ReadFieldString("Products", "Operation Name", f)
                                Dim constraintsGroup1 As Integer = 0
                                If OpNo = 40 And OpName.Contains("CLAMP") Then
                                    'MsgBox(f)
                                    constraintsGroup1 = preactor.ReadFieldInt("Products", "Constraint Group 1", f)
                                    Dim test = preactor.MatrixFieldSize("Secondary Constraint Groups", "Secondary Constraints", constraintsGroup1)

                                    Dim constraint_Group_1_num As Integer = 0
                                    If Not constraintsGroup1 = -1 Then
                                        constraint_Group_1_num = preactor.FindMatchingRecord("Secondary Constraint Groups", "Number", constraint_Group_1_num, constraintsGroup1)
                                        Dim size = preactor.MatrixFieldSize("Secondary Constraint Groups", "Secondary Constraints", constraint_Group_1_num)
                                        Dim values = New Dictionary(Of Double, Double)()
                                        Dim constraint_1_val As Double
                                        For m As Integer = 1 To size.X
                                            Dim key = preactor.ReadFieldString("Secondary Constraint Groups", "Secondary Constraints", constraint_Group_1_num, i)
                                            'Dim qty = preactor.ReadFieldDouble("Secondary Constraint Groups", "Secondary Constraints", constraint_Group_1_num, i)
                                            Dim o As Integer = 1
                                            Dim recNo As Integer
                                            If secondaryConstraintsCount > 0 Then
                                                Do
                                                    Dim secondaryConstraintsName As String = preactor.ReadFieldString("Secondary Constraints", "Name", o)
                                                    If secondaryConstraintsName = key Then
                                                        recNo = o
                                                        ''  MsgBox(recNo)
                                                        ''  constraint_1_val = planningboard.GetSecondaryResourceCurrentState(recNo, currentDate).CurrentValue
                                                        Dim terminatorTime As DateTime = planningboard.TerminatorTime()
                                                        constraint_1_val = planningboard.GetSecondaryResourceCurrentState(recNo, terminatorTime).CurrentValue
                                                        'Dim recNoName As String
                                                        'recNoName = planningboard.GetSecondaryResourceName(recNo)

                                                        'values.Add(key, qty)
                                                        Dim availableQuantity As Integer = CInt(dt.Rows(c)("No Of CLM")) - CInt(constraint_1_val)
                                                        dt.Rows(c)("Available CLM") = availableQuantity
                                                    End If
                                                    o = o + 1
                                                Loop While o <= secondaryConstraintsCount
                                            End If

                                        Next
                                    End If
                                End If


                                If OpNo = 40 And OpName.Contains("COMPRESSION") Then
                                    constraintsGroup1 = preactor.ReadFieldInt("Products", "Constraint Group 1", f)
                                    Dim test = preactor.MatrixFieldSize("Secondary Constraint Groups", "Secondary Constraints", constraintsGroup1)
                                    Dim constraint_Group_1_num As Integer = 0
                                    If Not constraintsGroup1 = -1 Then
                                        constraint_Group_1_num = preactor.FindMatchingRecord("Secondary Constraint Groups", "Number", constraint_Group_1_num, constraintsGroup1)
                                        Dim size = preactor.MatrixFieldSize("Secondary Constraint Groups", "Secondary Constraints", constraint_Group_1_num)
                                        Dim values = New Dictionary(Of Double, Double)()
                                        Dim constraint_1_val As Double
                                        For m As Integer = 1 To size.X
                                            Dim key = preactor.ReadFieldString("Secondary Constraint Groups", "Secondary Constraints", constraint_Group_1_num, i)
                                            'Dim qty = preactor.ReadFieldDouble("Secondary Constraint Groups", "Secondary Constraints", constraint_Group_1_num, i)
                                            Dim o As Integer = 1
                                            Dim recNo As Integer
                                            If secondaryConstraintsCount > 0 Then
                                                Do
                                                    Dim secondaryConstraintsName As String = preactor.ReadFieldString("Secondary Constraints", "Name", o)
                                                    If secondaryConstraintsName = key Then
                                                        recNo = o
                                                        ''  MsgBox(recNo)
                                                        ''  constraint_1_val = planningboard.GetSecondaryResourceCurrentState(recNo, currentDate).CurrentValue
                                                        Dim terminatorTime As DateTime = planningboard.TerminatorTime()
                                                        constraint_1_val = planningboard.GetSecondaryResourceCurrentState(recNo, terminatorTime).CurrentValue

                                                        'Dim recNoName As String
                                                        'recNoName = planningboard.GetSecondaryResourceName(recNo)

                                                        'values.Add(key, qty)
                                                        Dim availableQuantity As Integer = CInt(dt.Rows(c)("No Of CMP")) - CInt(constraint_1_val)
                                                        dt.Rows(c)("Available CMP") = availableQuantity
                                                    End If
                                                    o = o + 1
                                                Loop While o <= secondaryConstraintsCount
                                            End If


                                        Next
                                    End If
                                End If

                            End If
                            f = f + 1
                        Loop While f <= productCount
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try

                End If
            End If
        Next


        For c As Integer = 0 To dt.Rows.Count - 1
            If c >= 1 Then
                Dim balanceQty As Integer = CInt(dt.Rows(c)("Monthly Demand")) - CInt(dt.Rows(c)("On Plan"))
                dt.Rows(c)("Balance Qty") = balanceQty
                If balanceQty > 0 Then
                    dt.Rows(c)("Urgency") = "U"
                End If
            End If
        Next

        dt.Columns("SKU").ReadOnly = True
        dt.Columns("Description").ReadOnly = True
        dt.Columns("Planning Qty").ReadOnly = True
        dt.Columns("No Of CLM").ReadOnly = True
        dt.Columns("No Of CMP").ReadOnly = True
        dt.Columns("On Hand").ReadOnly = True
        dt.Columns("On Plan").ReadOnly = True
        dt.Columns("Running Qty").ReadOnly = True

        Dim quantityPlanningInterface As QuantityPlanningInterface = New QuantityPlanningInterface()
        quantityPlanningInterface.dataTable = dt
        quantityPlanningInterface.ShowDialog()
        'Try
        If quantityPlanningInterface.saveBtnClicked = True Then
            Dim maxNumberOfOrderNo As Integer = 0
            Dim firstLetterOfOrderNo As String = "Q"
            Dim selectedOrderList As New List(Of String)
            If currentYear = 2022 Then
                firstLetterOfOrderNo = "Q"
            ElseIf currentYear = 2023 Then
                firstLetterOfOrderNo = "R"
            ElseIf currentYear = 2024 Then
                firstLetterOfOrderNo = "S"
            End If
            Dim firstPart As String = firstLetterOfOrderNo + currentMonthName.Substring(0, 1) + currentDateWithNumber.ToString

            For e As Integer = 0 To dt.Rows.Count - 1
                If e >= 1 Then
                    Dim sku As String = dt.Rows(e)("SKU").ToString
                    Dim f As Integer = 1
                    Dim newOrderNo As String
                    If ordersCount > 0 Then
                        Do
                            Dim partNo As String = preactor.ReadFieldString("Orders", "Part No.", f)
                            Dim orderNoOfF As String = preactor.ReadFieldString("Orders", "Order No.", f)
                            If Not orderNoOfF.Contains("/") Then
                                If orderNoOfF.Contains(firstLetterOfOrderNo + currentMonthName.Substring(0, 1) + currentDateWithNumber.ToString) Then
                                    Dim lastThreeChar As Integer = CInt(orderNoOfF.Substring(orderNoOfF.Length - 4, 4))
                                    If maxNumberOfOrderNo < lastThreeChar Then
                                        maxNumberOfOrderNo = lastThreeChar
                                    End If
                                End If
                            End If
                            f = f + 1
                        Loop While f <= ordersCount
                    End If

                    Dim decimalLength As Integer
                    decimalLength = 4

                    'If (dt.Rows(e)("Is Updated").ToString = "True") Or (Not dt.Rows(e)("Is Updated") Is DBNull.Value) Then

                    Dim orderQtyCLM As String = dt.Rows(e)("Order Qty CLM").ToString
                    If Not orderQtyCLM = "" Then
                        If Not CInt(orderQtyCLM) = 0 Then
                            Dim newBlock As Integer = preactor.CreateRecord("Orders")
                            Dim newRecordNum As Integer = preactor.ReadFieldInt("Orders", "Number", newBlock)

                            Dim maxStr As String = (maxNumberOfOrderNo + 1).ToString("D" + decimalLength.ToString())
                            newOrderNo = firstPart + maxStr
                            selectedOrderList.Add(newOrderNo)
                            'preactor.WriteField("Orders", "Number", newBlock, newRecordNum)
                            preactor.WriteField("Orders", "Order No.", newBlock, newOrderNo)
                            preactor.WriteField("Orders", "Quantity", newBlock, orderQtyCLM)
                            preactor.WriteField("Orders", "Part No.", newBlock, sku)
                            preactor.WriteField("Orders", "Op. No.", newBlock, 20)
                            preactor.WriteField("Orders", "Disable Operation", newBlock, True)
                            preactor.WriteField("Orders", "K202_APSParentNumber", newBlock, newOrderNo)
                            If dt.Rows(e)("Urgency").ToString = "U" Then
                                preactor.WriteField("Orders", "Priority", newBlock, 11)
                            End If
                            preactor.ExpandJob("Orders", newBlock)
                            maxNumberOfOrderNo = maxNumberOfOrderNo + 1
                            preactor.WriteField("Orders", "String Attribute 2", newBlock, "CLM")
                        End If
                    End If
                    Dim orderQtyCMP As String = dt.Rows(e)("Order Qty CMP").ToString
                    If Not orderQtyCMP = "" Then
                        If Not CInt(orderQtyCMP) = 0 Then
                            Dim newBlock As Integer = preactor.CreateRecord("Orders")
                            Dim newRecordNum As Integer = preactor.ReadFieldInt("Orders", "Number", newBlock)

                            Dim maxStr As String = (maxNumberOfOrderNo + 1).ToString("D" + decimalLength.ToString())
                            newOrderNo = firstPart + maxStr
                            selectedOrderList.Add(newOrderNo)
                            'preactor.WriteField("Orders", "Number", newBlock, newRecordNum)
                            preactor.WriteField("Orders", "Order No.", newBlock, newOrderNo)
                            preactor.WriteField("Orders", "Quantity", newBlock, orderQtyCMP)
                            preactor.WriteField("Orders", "Part No.", newBlock, sku)
                            preactor.WriteField("Orders", "Op. No.", newBlock, 20)
                            preactor.WriteField("Orders", "Disable Operation", newBlock, False)
                            preactor.WriteField("Orders", "K202_APSParentNumber", newBlock, newOrderNo)
                            If dt.Rows(e)("Urgency").ToString = "U" Then
                                preactor.WriteField("Orders", "Priority", newBlock, 11)
                            End If
                            preactor.ExpandJob("Orders", newBlock)
                            preactor.WriteField("Orders", "String Attribute 2", newBlock, "CMP")
                            maxNumberOfOrderNo = maxNumberOfOrderNo + 1
                        End If
                    End If
                    'End If
                End If
            Next
            preactor.Commit("Orders")
            Dim ordersCountNew = preactor.RecordCount("Orders")
            For Each element As String In selectedOrderList
                If ordersCountNew > 0 Then
                    Dim g As Integer = 1
                    Do
                        Dim orderNoOfG As String = preactor.ReadFieldString("Orders", "Order No.", g)
                        Dim belongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", g)
                        If orderNoOfG = element Then
                            Dim stringAttribute2 As String = preactor.ReadFieldString("Orders", "String Attribute 2", g)
                            If stringAttribute2 = "CLM" And belongsToOrderNo = "PARENT" Then
                                preactor.WriteField("Orders", "Disable Operation", g, False)
                                Dim h As Integer = 1
                                Do
                                    Dim orderNoOfH As String = preactor.ReadFieldString("Orders", "Order No.", h)
                                    Dim opNoOfH As Integer = preactor.ReadFieldInt("Orders", "Op. No.", h)
                                    Dim opName As String = preactor.ReadFieldString("Orders", "Operation Name", h)
                                    If orderNoOfH = orderNoOfG Then
                                        If opNoOfH = 40 And opName.Contains("COMPRESSION") Then
                                            preactor.WriteField("Orders", "Disable Operation", h, True)
                                        End If
                                        If opNoOfH = 50 Then
                                            preactor.WriteField("Orders", "Disable Operation", h, True)
                                        End If
                                    End If
                                    h = h + 1
                                Loop While h <= ordersCountNew
                            End If
                            If stringAttribute2 = "CMP" And belongsToOrderNo = "PARENT" Then
                                preactor.WriteField("Orders", "Disable Operation", g, False)
                                Dim h As Integer = 1
                                Do
                                    Dim orderNoOfH As String = preactor.ReadFieldString("Orders", "Order No.", h)
                                    Dim opNoOfH As Integer = preactor.ReadFieldInt("Orders", "Op. No.", h)
                                    Dim opName As String = preactor.ReadFieldString("Orders", "Operation Name", h)
                                    If orderNoOfH = orderNoOfG Then
                                        If opNoOfH = 40 And opName.Contains("CLAMP") Then
                                            preactor.WriteField("Orders", "Disable Operation", h, True)
                                        End If
                                    End If
                                    h = h + 1
                                Loop While h <= ordersCountNew
                            End If
                        End If
                        g = g + 1
                    Loop While g <= ordersCountNew
                End If
            Next
            preactor.Commit("Orders")
            quantityPlanningInterface.Close_Interface()
        End If
        'Catch ex As Exception
        '    MsgBox("Error")
        'End Try

        Return 0
    End Function
    Public Function GetMaxMolds(ByRef connetionString As String, ByRef intSecondaryConstraintRecNumber As Integer, ByRef dtStartTime As Date) As Integer
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetMaxMold_Sp"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@secondaryConstraintNumber", intSecondaryConstraintRecNumber)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int64
            command.Parameters.Add(param)
            param = New SqlParameter("@StartDate", dtStartTime)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Date
            command.Parameters.Add(param)
            Dim intTotalSpindle As Integer = 0
            param = New SqlParameter("@MaxValue", intTotalSpindle)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int64
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CInt(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            '' 18_02_2022
            '' MsgBox("Available former not define",, "error")
            MsgBox(ex.Message)
        Finally

        End Try
    End Function

    Public Function GetMaximumMoldNumber(ByRef connetionString As String, ByRef partNo As String, ByRef OpName As String) As Integer
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetMaximumNumberOfMolds"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@partNo", partNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@OpName", OpName)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim intTotalSpindle As Integer = 0
            param = New SqlParameter("@maxMold", intTotalSpindle)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int64
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            Return CInt(param.Value)
            connection.Close()
        Catch ex As Exception
            '' 18_02_2022
            '' MsgBox("Available former not define",, "error")
            'MsgBox(ex.Message)
        Finally

        End Try
    End Function

    Public Function GetShipmentDate(ByRef connetionString As String, ByRef orderNo As String) As DateTime
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetMaxDemandDate"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@orderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim intTotalSpindle As DateTime
            param = New SqlParameter("@shipmentDate", intTotalSpindle)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.DateTime
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            Dim shipmentDate As DateTime
            If param.Value Is DBNull.Value Then
                Dim currentYear As Integer = CInt(DateTime.Now.ToString("yyyy"))
                Dim endOfYear As String = currentYear & "/12" & "/31"
                shipmentDate = CDate(endOfYear)
            Else
                shipmentDate = CDate(param.Value)
            End If

            Return shipmentDate

            connection.Close()
        Catch ex As Exception
        Finally

        End Try
    End Function

    Public Function GetMaxOpTimePerItemInSelectedOrders(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim numOfOrders As Integer = preactor.RecordCount("Orders")
        Dim y As Integer = 1
        Dim SelectedOrderList As New List(Of Integer)
        Dim functionRunOnSecondTime As Boolean = False
        Do
            If (planningboard.GetOperationLocateState(y)) Then
                Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", y)
                Dim endDateOfY As Date = preactor.ReadFieldDateTime("Orders", "End Time", y)
                Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", y)
                'MsgBox("Record Number is " & y)
                SelectedOrderList.Add(y)
            End If
            y = y + 1
        Loop While y <= numOfOrders

        For Each index As Integer In SelectedOrderList
            Dim toggleAttribute2 As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 2", index)
            If toggleAttribute2 = 1 Then
                functionRunOnSecondTime = True
            End If
        Next

        If functionRunOnSecondTime = True Then
            Dim firstItem As Integer = SelectedOrderList(0)
            Dim mainResourceNo As Integer = preactor.ReadFieldInt("Orders", "Resource", firstItem)
            If Not SelectedOrderList.Count = numOfOrders Then
                For Each index As Integer In SelectedOrderList
                    Dim resourceNoOfEachIndex As Integer = preactor.ReadFieldInt("Orders", "Resource", index)
                    If Not mainResourceNo = resourceNoOfEachIndex Then
                        MsgBox("Need to select operations in same resource.")
                        Return 0
                    End If
                Next

                For Each index As Integer In SelectedOrderList
                    Dim toggleAttribute2 As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 2", index)
                    If toggleAttribute2 = 1 Then
                        Dim durationAttribute3 As Double = preactor.ReadFieldDouble("Orders", "Duration Attribute 3", index)
                        Dim operationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", index)
                        preactor.WriteField("Orders", "Op. Time per Item", index, (operationTimePerItem - durationAttribute3))
                        preactor.WriteField("Orders", "Duration Attribute 3", index, -1)
                        preactor.WriteField("Orders", "Toggle Attribute 2", index, 0)
                    End If
                Next
            Else
                MsgBox("Inactive for single operation.")
            End If
        Else
            If Not SelectedOrderList.Count = numOfOrders Then
                Dim mainResourceNo As Integer
                Dim mainTableAttribute5 As Integer
                Dim mainIndex As Integer
                Dim maxOperationTime As Double = 0
                For Each index As Integer In SelectedOrderList

                    Dim toggleAttribute2 As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 2", index)

                    Dim operationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", index)
                    If operationTimePerItem > maxOperationTime Then
                        maxOperationTime = operationTimePerItem
                        Dim resourceNoOfFirstIndex As Integer = preactor.ReadFieldInt("Orders", "Resource", index)
                        Dim tableAttribute5 As Integer = preactor.ReadFieldInt("Orders", "Table Attribute 5", index)
                        mainResourceNo = resourceNoOfFirstIndex
                        mainTableAttribute5 = tableAttribute5
                        mainIndex = index
                    End If

                Next

                For Each index As Integer In SelectedOrderList
                    Dim resourceNoOfEachIndex As Integer = preactor.ReadFieldInt("Orders", "Resource", index)
                    If Not mainResourceNo = resourceNoOfEachIndex Then
                        MsgBox("Need to select operations in same resource.")
                        Return 0
                    End If
                Next

                For Each index As Integer In SelectedOrderList
                    Dim toggleAttribute2 As Integer = preactor.ReadFieldInt("Orders", "Toggle Attribute 2", index)
                    Dim operationTimePerItem As Double = preactor.ReadFieldDouble("Orders", "Op. Time per Item", index)
                    Dim tableAttribute5 As Integer = preactor.ReadFieldInt("Orders", "Table Attribute 5", index)
                    If Not index = mainIndex Then
                        If Not operationTimePerItem = maxOperationTime And tableAttribute5 = mainTableAttribute5 Then
                            Dim differenceOfOperationTimes As Double = maxOperationTime - operationTimePerItem
                            preactor.WriteField("Orders", "Duration Attribute 3", index, differenceOfOperationTimes)
                            preactor.WriteField("Orders", "Op. Time per Item", index, (operationTimePerItem + differenceOfOperationTimes))
                            preactor.WriteField("Orders", "Toggle Attribute 2", index, 1)
                        End If
                    End If

                Next
            Else
                MsgBox("Inactive for single operation.")
            End If
        End If
        preactor.Commit("Orders")
        Return 0
    End Function
    Public Function UpstreamOrderGenarate(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim oForm As New UpstreamOrderGenarateForm()
        oForm.connetionString = connetionString
        oForm.ShowDialog()
    End Function

    Public Function AddCalendarPeriods(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef recordNumber As Integer) As Integer
        Try
            Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
            Dim resource As Integer = preactor.ReadFieldInt("Orders", "Resource", recordNumber)
            Dim StartDate As DateTime = preactor.ReadFieldDateTime("Orders", "Setup Start", recordNumber)
            Dim EndDate As DateTime = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)
            Dim planningboard As IPlanningBoard = preactor.PlanningBoard
            planningboard.CreatePrimaryCalendarException(resource, StartDate, EndDate, "Off Shift")
            preactor.Redraw()
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try

    End Function

    Public Function RemoveCalendarPeriods(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef recordNumber As Integer) As Integer
        Try
            Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
            Dim resource As Integer = preactor.ReadFieldInt("Orders", "Resource", recordNumber)
            Dim StartDate As DateTime = preactor.ReadFieldDateTime("Orders", "Setup Start", recordNumber)
            Dim EndDate As DateTime = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)
            Dim planningboard As IPlanningBoard = preactor.PlanningBoard
            planningboard.CreatePrimaryCalendarException(resource, StartDate, EndDate, "On Shift")
            preactor.Redraw()
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try

    End Function

    Public Function K202_GetSerialsForJobSplit(ByRef connetionString As String, ByRef Year As String, ByRef Month As String) As Integer
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GetSerialsForJobSplit_Sp"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@Year", Year)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@Month", Month)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim serialnumber As Integer
            param = New SqlParameter("@serialNumber", SqlDbType.Int)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int64
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            Return CInt(param.Value)

            connection.Close()
        Catch ex As Exception
        Finally

        End Try
    End Function

    Public Function K202_UpdateSerialsForJobSplit(ByRef connetionString As String, ByRef Year As String, ByRef Month As String) As Integer
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_UpdateSerialsForJobSplit_Sp"
            '
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@Year", Year)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@Month", Month)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
        Catch ex As Exception
        Finally

        End Try
    End Function

    Public Function RecordValidation(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim ordersCount As Integer = preactor.RecordCount("Orders")
        Dim i As Integer = 1
        Dim opNoList As New List(Of Integer)
        Do
            Dim opNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
            opNoList.Add(opNo)
            i = i + 1
        Loop While i <= ordersCount
        opNoList = opNoList.Distinct.ToList()

        Dim k As Integer = 1
        'Dim table As DataTable = New DataTable()
        'table.Columns.Add("OpNo", GetType(Integer))
        'table.Columns.Add("Imported Quantity", GetType(Integer))

        Dim oForm As New RecordValidation()
        For Each element In opNoList
            Dim j As Integer = 1
            Dim sumOfQuantity As Integer = 0
            Do
                Dim opNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", j)
                If element = opNo Then
                    Dim quantity As Integer = preactor.ReadFieldInt("Orders", "Quantity", j)
                    sumOfQuantity = sumOfQuantity + quantity
                End If
                j = j + 1
            Loop While j <= ordersCount
            'table.Rows.Add(element, sumOfQuantity)


            Dim lable As New Label
            lable.Name = "Lable" + k.ToString
            lable.Location = New System.Drawing.Point(10, 40 * k)
            lable.Text = "Op No. " + element.ToString + " : " + "Total Quantity is " + sumOfQuantity.ToString
            lable.Padding = New Padding(10, 0, 0, 10)
            lable.Size = New System.Drawing.Size(250, 30)
            lable.Font = New Drawing.Font(lable.Text, 10, Drawing.FontStyle.Regular)
            oForm.Controls.Add(lable)
            k = k + 1
        Next
        'oForm.DataGridView.DataSource = table
        oForm.ShowDialog()
    End Function

    ''Added by  Milana amarasooriya 20230512
    Public Function CheckUserIsInMemoryEditMode(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim strPathO As String
        Dim strPath As String
        Dim strLockUser As String
        Dim strCurrentUser As String


        strPathO = preactor.ParseShellString("{PATH}")
        strCurrentUser = preactor.ParseShellString("{USER NAME}")

        strPath = strPathO + "\locked.tmp"
        If File.Exists(strPath) Then

            Try
                preactor.SetShellVariable("K202_IsMemoryEditMode", 0)

                Using fs = New FileStream(strPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Using sr = New StreamReader(fs, Encoding.Default)
                        strLockUser = sr.ReadToEnd()
                        If (strLockUser <> strCurrentUser) Then
                            MsgBox("You can't Execute the process,because '" + strLockUser + "' User is accessing the system")
                            preactor.SetShellVariable("K202_IsMemoryEditMode", 1)
                            'MsgBox(preactor.GetShellVariable("K202_IsMemoryEditMode"))
                        End If
                    End Using
                End Using

            Catch ex As Exception
                ''MsgBox(preactor.GetShellVariable("K202_IsMemoryEditMode"))
                MsgBox("Error " + ex.Message)

            End Try

        Else
            MsgBox(strPath)
        End If
        Return 0
    End Function

    Public Function AdjustCuringStartAndEnd(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef recordNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim resourceCount As Integer = preactor.RecordCount("Resources")
        Dim maxEndDate As Date = #9/23/1899 01:00 AM#

        Dim strOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
        Dim strOrderOprName As String = preactor.ReadFieldString("Orders", "Operation Name", recordNumber)
        Dim endDate As Date = preactor.ReadFieldDateTime("Orders", "End Time", recordNumber)
        Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)

        'Dim newStartDate As Date = startDate.AddMinutes(-30)
        'Dim newEndDate As Date = endDate.AddMinutes(30)
        Dim newStartDate As Date = startDate
        Dim newEndDate As Date = endDate

        Dim resource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)
        'Read the attribute 1 from resources table.

        'Read the attribute 1 of the order.
        Dim attribute01 As String = ""
        Dim resourceCountStart = 1
        Do
            Dim resourceNew As String = preactor.ReadFieldString("Resources", "Name", resourceCountStart)
            If resourceNew = resource Then
                attribute01 = preactor.ReadFieldString("Resources", "Attribute 1", resourceCountStart)
            End If
            resourceCountStart = resourceCountStart + 1
        Loop While resourceCountStart <= resourceCount

        Dim i As Integer = 1
        Dim list As New List(Of Integer)
        Do
            Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)
            Dim operationNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
            Dim resource2 As String = preactor.ReadFieldString("Orders", "Resource", i)
            Dim resourceCountStart2 As Integer = 1
            Dim eachReseourceAttribute1 As String = ""
            Do
                Dim eachReseourceName As String = ""
                eachReseourceName = preactor.ReadFieldString("Resources", "Name", resourceCountStart2)
                If eachReseourceName = resource2 Then
                    eachReseourceAttribute1 = preactor.ReadFieldString("Resources", "Attribute 1", resourceCountStart2)
                End If
                resourceCountStart2 = resourceCountStart2 + 1
            Loop While resourceCountStart2 <= resourceCount

            If eachReseourceAttribute1 = attribute01 Then
                Dim endDate2 As Date = preactor.ReadFieldDateTime("Orders", "End Time", i)
                Dim startDate2 As Date = preactor.ReadFieldDateTime("Orders", "Start Time", i)
                If (endDate2 >= newStartDate And endDate2 <= newEndDate) Then
                    If operationName.Contains("COMPRESSION") Then
                        If operationNo = 40 Then
                            list.Add(i)
                        End If
                    End If
                End If
            End If
            i = i + 1
        Loop While i <= num


        For Each record As Integer In list
            Dim endDateOfRecord As Date = preactor.ReadFieldDateTime("Orders", "End Time", record)
            If maxEndDate < endDateOfRecord Then
                maxEndDate = endDateOfRecord
            End If
        Next

        For Each record As Integer In list
            Dim endDateOfRecordd As Date = preactor.ReadFieldDateTime("Orders", "End Time", record)
            Dim timeDifference As TimeSpan = maxEndDate.Subtract(endDateOfRecordd)
            Dim setupTime As Double = preactor.ReadFieldDouble("Orders", "Setup Time", record)
            Dim totalSetupTime As Double = setupTime + timeDifference.TotalDays
            preactor.WriteField("Orders", "Setup Time", record, totalSetupTime)
        Next
        ' Designed to be used without concurrent setup option for resources
        preactor.Commit("Orders")
        Return 0
    End Function

    Function RemoveSetupTime(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef recordNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim ordersCount As Integer = preactor.RecordCount("Orders")

        Dim strOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
        Dim strOrderOprName As String = preactor.ReadFieldString("Orders", "Operation Name", recordNumber)
        Dim endDate As Date = preactor.ReadFieldDateTime("Orders", "End Time", recordNumber)
        Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)
        Dim resource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)

        Dim i As Integer = 1
        Dim resourceCount As Integer = preactor.RecordCount("Resources")

        Dim attribute01 As String = ""
        Dim resourceCountStart = 1
        Do
            Dim resourceNew As String = preactor.ReadFieldString("Resources", "Name", resourceCountStart)
            If resourceNew = resource Then
                attribute01 = preactor.ReadFieldString("Resources", "Attribute 1", resourceCountStart)
            End If
            resourceCountStart = resourceCountStart + 1
        Loop While resourceCountStart <= resourceCount


        'MsgBox("Resource " & resource)
        Dim list As New List(Of Integer)
        Do
            Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)
            If operationName.Contains("COMPRESSION") Then
                Dim endDate2 As Date = preactor.ReadFieldDateTime("Orders", "End Time", i)
                Dim startDate2 As Date = preactor.ReadFieldDateTime("Orders", "Start Time", i)
                Dim resource2 As String = preactor.ReadFieldString("Orders", "Resource", i)
                Dim resourceCountStart2 As Integer = 1
                Dim eachReseourceAttribute1 As String = ""
                Do
                    Dim eachReseourceName As String = ""
                    eachReseourceName = preactor.ReadFieldString("Resources", "Name", resourceCountStart2)
                    If eachReseourceName = resource2 Then
                        eachReseourceAttribute1 = preactor.ReadFieldString("Resources", "Attribute 1", resourceCountStart2)
                    End If
                    resourceCountStart2 = resourceCountStart2 + 1
                Loop While resourceCountStart2 <= resourceCount

                If (endDate2 >= startDate And endDate2 <= endDate And eachReseourceAttribute1 = attribute01) Then
                    list.Add(i)
                End If
            End If
            i = i + 1
        Loop While i <= ordersCount

        For Each record As Integer In list
            preactor.WriteField("Orders", "Setup Time", record, 0)
        Next
        preactor.Commit("Orders")
        Return 0
    End Function


    'Code backup 2023/05/23

    'Public Function AdjustCuringStartAndEndOld(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard
    '    Dim num As Integer = preactor.RecordCount("Orders")
    '    Dim resourceCount As Integer = preactor.RecordCount("Resources")
    '    Dim y As Integer = 1
    '    Dim maxEndDate As Date = #9/23/1899 01:00 AM#
    '    Dim recordNumber As Integer
    '    Do
    '        If (planningboard.GetOperationLocateState(y)) Then
    '            Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", y)
    '            Dim operationNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", y)
    '            Dim endDateOfY As Date = preactor.ReadFieldDateTime("Orders", "End Time", y)
    '            'MsgBox("End date of y " & endDateOfY)
    '            'Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", y)
    '            If operationName.Contains("COMPRESSION") Then
    '                If operationNo = 40 Then
    '                    If endDateOfY > maxEndDate Then
    '                        maxEndDate = endDateOfY
    '                        recordNumber = y
    '                    End If
    '                End If
    '            End If
    '        End If
    '        y = y + 1
    '    Loop While y <= num



    '    Dim strOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
    '    Dim strOrderOprName As String = preactor.ReadFieldString("Orders", "Operation Name", recordNumber)
    '    Dim endDate As Date = preactor.ReadFieldDateTime("Orders", "End Time", recordNumber)
    '    Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)

    '    Dim newStartDate As Date = startDate.AddMinutes(-30)
    '    Dim newEndDate As Date = endDate.AddMinutes(30)
    '    'MsgBox("Start Date : " & startDate & "/ New Start Date : " & newStartDate)
    '    'MsgBox("End Date : " & endDate & "/ New End Date : " & newEndDate)
    '    Dim resource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)
    '    MsgBox("Resource is " & resource & " Order no " & strOrderNo)
    '    'Read the attribute 1 from resources table.
    '    Dim attribute01 As String = ""
    '    Dim resourceCountStart = 1
    '    Do
    '        Dim resourceNew As String = preactor.ReadFieldString("Resources", "Name", resourceCountStart)
    '        MsgBox(resourceNew)
    '        If resourceNew = resource Then
    '            MsgBox("resources are same")
    '            attribute01 = preactor.ReadFieldString("Resources", "Attribute 1", resourceCountStart)
    '        End If
    '        resourceCountStart = resourceCountStart + 1
    '    Loop While resourceCountStart <= resourceCount

    '    MsgBox("Attribute 01 " & attribute01)
    '    Dim i As Integer = 1
    '    'MsgBox("Resource " & resource)
    '    Dim list As New List(Of Integer)
    '    Do
    '        Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)
    '        Dim operationNo As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)

    '        If operationName.Contains("COMPRESSION") Then
    '            If operationNo = 40 Then
    '                Dim endDate2 As Date = preactor.ReadFieldDateTime("Orders", "End Time", i)
    '                Dim startDate2 As Date = preactor.ReadFieldDateTime("Orders", "Start Time", i)
    '                Dim resource2 As String = preactor.ReadFieldString("Orders", "Resource", i)
    '                If (endDate2 >= newStartDate And endDate2 <= newEndDate And resource2 = resource) Or (startDate2 >= newStartDate And startDate2 <= newEndDate And resource2 = resource) Then
    '                    list.Add(i)
    '                End If
    '            End If
    '        End If
    '        i = i + 1
    '    Loop While i <= num


    '    For Each record As Integer In list
    '        Dim endDateOfRecord As Date = preactor.ReadFieldDateTime("Orders", "End Time", record)
    '        If maxEndDate < endDateOfRecord Then
    '            maxEndDate = endDateOfRecord
    '        End If
    '    Next

    '    For Each record As Integer In list
    '        Dim endDateOfRecordd As Date = preactor.ReadFieldDateTime("Orders", "End Time", record)
    '        'MsgBox("End date" & endDateOfRecordd)
    '        Dim timeDifference As TimeSpan = maxEndDate.Subtract(endDateOfRecordd)
    '        'MsgBox("Time difference in Minutes" & timeDifference.TotalMinutes)
    '        Dim setupTime As Double = preactor.ReadFieldDouble("Orders", "Setup Time", record)
    '        Dim totalSetupTime As Double = setupTime + timeDifference.TotalDays
    '        preactor.WriteField("Orders", "Setup Time", record, totalSetupTime)
    '    Next
    '    ' Designed to be used without concurrent setup option for resources
    '    preactor.Commit("Orders")
    '    Return 0
    'End Function

    'Function RemoveSetupTimeOld(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
    '    Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

    '    Dim planningboard As IPlanningBoard = preactor.PlanningBoard
    '    Dim num As Integer = preactor.RecordCount("Orders")
    '    Dim y As Integer = 1
    '    Dim maxEndDate As Date = #9/23/1899 01:00 AM#
    '    Dim recordNumber As Integer
    '    Do
    '        If (planningboard.GetOperationLocateState(y)) Then
    '            Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", y)
    '            Dim endDateOfY As Date = preactor.ReadFieldDateTime("Orders", "End Time", y)
    '            'MsgBox("End date of y " & endDateOfY)
    '            'Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", y)
    '            If operationName.Contains("COMPRESSION") Then
    '                If endDateOfY > maxEndDate Then
    '                    maxEndDate = endDateOfY
    '                    recordNumber = y
    '                End If
    '            End If
    '        End If
    '        y = y + 1
    '    Loop While y <= num



    '    Dim strOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", recordNumber)
    '    Dim strOrderOprName As String = preactor.ReadFieldString("Orders", "Operation Name", recordNumber)
    '    Dim endDate As Date = preactor.ReadFieldDateTime("Orders", "End Time", recordNumber)
    '    Dim startDate As Date = preactor.ReadFieldDateTime("Orders", "Start Time", recordNumber)

    '    Dim newStartDate As Date = startDate.AddMinutes(-30)
    '    Dim newEndDate As Date = endDate.AddMinutes(30)
    '    'MsgBox("Start Date : " & startDate & "/ New Start Date : " & newStartDate)
    '    'MsgBox("End Date : " & endDate & "/ New End Date : " & newEndDate)
    '    Dim resource As String = preactor.ReadFieldString("Orders", "Resource", recordNumber)
    '    Dim i As Integer = 1


    '    'MsgBox("Resource " & resource)
    '    Dim list As New List(Of Integer)
    '    Do
    '        Dim operationName As String = preactor.ReadFieldString("Orders", "Operation Name", i)
    '        If operationName.Contains("COMPRESSION") Then
    '            Dim endDate2 As Date = preactor.ReadFieldDateTime("Orders", "End Time", i)
    '            Dim startDate2 As Date = preactor.ReadFieldDateTime("Orders", "Start Time", i)
    '            Dim resource2 As String = preactor.ReadFieldString("Orders", "Resource", i)
    '            If (endDate2 >= newStartDate And endDate2 <= newEndDate And resource2 = resource) Or (startDate2 >= newStartDate And startDate2 <= newEndDate And resource2 = resource) Then
    '                list.Add(i)
    '            End If
    '        End If
    '        i = i + 1
    '    Loop While i <= num


    '    For Each record As Integer In list
    '        preactor.WriteField("Orders", "Setup Time", record, 0)
    '    Next
    '    preactor.Commit("Orders")
    '    Return 0
    'End Function

    ''Sahan added K202_CustomSchedulingWindowFunction
    ''20230523
    Function K202_CustomSchedulingWindowFunction(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim K202_CustomSchedulingWindowForm As New K202_CustomSchedulingWindow()
        Dim K202_CustomSchedulingWindowSecondPartForm As New K202_CustomSchedulingWindowSecondPart()
        'Get orders count
        Dim ordersCount As Integer = preactor.RecordCount("Orders")
        'Get resource count
        Dim resourceCount As Integer = preactor.RecordCount("Resources")

        Dim resourcesLoopStart As Integer = 1
        Dim ordersLoopStart As Integer = 1
        Dim attribute01List As New List(Of String)
        attribute01List.Add("")
        Do
            Dim attribute01OfResource As String = preactor.ReadFieldString("Resources", "Attribute 1", resourcesLoopStart)
            If Not attribute01OfResource = "" Then
                attribute01List.Add(attribute01OfResource)
            End If
            resourcesLoopStart = resourcesLoopStart + 1
        Loop While resourcesLoopStart <= resourceCount

        attribute01List = attribute01List.Distinct.ToList()
        K202_CustomSchedulingWindowForm.ComboBox1.DataSource = attribute01List
        'Get attribute01 list in resource table
        K202_CustomSchedulingWindowForm.clickedShowCavity_01Btn = False
        K202_CustomSchedulingWindowForm.ShowDialog()
        AddHandler K202_CustomSchedulingWindowForm.Button1.Click, AddressOf K202_CustomSchedulingWindowForm.Button1_Click
        Try
            If K202_CustomSchedulingWindowForm.clickedShowCavity_01Btn = True Then
                Dim selectedAttribute01 As String = K202_CustomSchedulingWindowForm.selectedAttribute01

                Dim resourceList As New List(Of String)
                Dim resourceLoopStart2 As Integer = 1
                Do
                    Dim attribute01OfResource As String = preactor.ReadFieldString("Resources", "Attribute 1", resourceLoopStart2)
                    If attribute01OfResource = selectedAttribute01 Then
                        Dim resourceName As String = preactor.ReadFieldString("Resources", "Name", resourceLoopStart2)
                        If Not resourceName.Contains("TABLES") Then
                            resourceList.Add(resourceName)
                        End If
                    End If
                        resourceLoopStart2 = resourceLoopStart2 + 1
                Loop While resourceLoopStart2 <= resourceCount

                Dim dt As DataTable = New DataTable()
                dt.Columns.Add(New DataColumn("Resources", Type.[GetType]("System.String")))
                Dim dt_row As DataRow
                For Each element In resourceList
                    dt_row = dt.NewRow()
                    dt_row("Resources") = element
                    dt.Rows.Add(dt_row)
                Next
                dt.Columns.Add(New DataColumn("Mold Set 01", Type.[GetType]("System.String")))
                dt.Columns.Add(New DataColumn("Mold Set 02", Type.[GetType]("System.String")))

                K202_CustomSchedulingWindowSecondPartForm.dataTable = dt
                K202_CustomSchedulingWindowSecondPartForm.ShowDialog()

                Try
                    If K202_CustomSchedulingWindowSecondPartForm.okBtnClicked = True Then

                        For c As Integer = 0 To dt.Columns.Count - 1
                            If dt.Columns(c).ColumnName = "Mold Set 01" Then
                                For e As Integer = 0 To dt.Rows.Count - 1
                                    Dim userEnteredProductName As String = dt.Rows(e)("Mold Set 01").ToString()
                                    Dim selectedResourceName As String = dt.Rows(e)("Resources").ToString

                                    Dim ordersLoopStart2 As Integer = 1
                                    Dim priorityValue As Integer = 1
                                    Do
                                        Dim orderProductName As String = preactor.ReadFieldString("Orders", "Product", ordersLoopStart2)
                                        If orderProductName = userEnteredProductName Then
                                            Dim unscheduleOrder As String = preactor.ReadFieldString("Orders", "Resource", ordersLoopStart2)
                                            Dim opNoOfOrder As Integer = preactor.ReadFieldInt("Orders", "Op. No.", ordersLoopStart2)

                                            If unscheduleOrder = "Unspecified" And opNoOfOrder = 20 Then
                                                preactor.WriteField("Orders", "Priority", ordersLoopStart2, priorityValue)
                                                preactor.WriteField("Orders", "Required Resource", ordersLoopStart2, selectedResourceName)
                                                priorityValue = priorityValue + 2
                                            End If

                                            If unscheduleOrder = "Unspecified" And opNoOfOrder = 40 Then
                                                preactor.WriteField("Orders", "Required Resource", ordersLoopStart2, selectedResourceName)
                                                Dim priorityOfOpNo40 As Integer = preactor.ReadFieldInt("Orders", "Priority", ordersLoopStart2)
                                                preactor.WriteField("Orders", "String Attribute 5", ordersLoopStart2, "Mold_Set_01_" & priorityOfOpNo40)
                                            End If

                                        End If
                                        ordersLoopStart2 = ordersLoopStart2 + 1
                                    Loop While ordersLoopStart2 <= ordersCount
                                Next
                            End If

                            If dt.Columns(c).ColumnName = "Mold Set 02" Then
                                For e As Integer = 0 To dt.Rows.Count - 1
                                    Dim userEnteredProductName As String = dt.Rows(e)("Mold Set 02").ToString()
                                    Dim selectedResourceName As String = dt.Rows(e)("Resources").ToString

                                    Dim ordersLoopStart2 As Integer = 1
                                    Dim priorityValue As Integer = 2
                                    Do
                                        Dim orderProductName As String = preactor.ReadFieldString("Orders", "Product", ordersLoopStart2)
                                        If orderProductName = userEnteredProductName Then
                                            Dim unscheduleOrder As String = preactor.ReadFieldString("Orders", "Resource", ordersLoopStart2)
                                            Dim opNoOfOrder As Integer = preactor.ReadFieldInt("Orders", "Op. No.", ordersLoopStart2)
                                            If unscheduleOrder = "Unspecified" And opNoOfOrder = 20 Then
                                                preactor.WriteField("Orders", "Priority", ordersLoopStart2, priorityValue)
                                                priorityValue = priorityValue + 2
                                            End If

                                            If unscheduleOrder = "Unspecified" And opNoOfOrder = 40 Then
                                                preactor.WriteField("Orders", "Required Resource", ordersLoopStart2, selectedResourceName)
                                                Dim priorityOfOpNo40 As Integer = preactor.ReadFieldInt("Orders", "Priority", ordersLoopStart2)
                                                preactor.WriteField("Orders", "String Attribute 5", ordersLoopStart2, "Mold_Set_02_" & priorityOfOpNo40)
                                            End If

                                        End If
                                        ordersLoopStart2 = ordersLoopStart2 + 1
                                    Loop While ordersLoopStart2 <= ordersCount
                                Next
                            End If
                        Next


                    End If
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception

        End Try

    End Function

End Class


