Option Explicit On
Option Strict On
Imports System.Data.OleDb
Module DBController
    Public Const connectString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= ExamReview.accdb"
    Public br As String = Environment.NewLine
    ''' <summary>
    ''' Test the connection
    ''' </summary>
    ''' <param name="connectString">Connect String</param>
    ''' <returns>Result of test</returns>
    Public Function controlDB(ByVal connectString As String) As String
        ' Setup variable
        Dim connectResult As String = String.Empty
        ' Initial Connection Setup
        Using connectSetup As New OleDbConnection(connectString)
            ' Create Command Controller
            Dim commandSQL As New OleDbCommand
            Try
                ' Open Connection
                connectSetup.Open()
                ' Command to connect
                commandSQL.Connection = connectSetup
                ' Clost Connection
                connectSetup.Close()
                connectResult = "Connect Successful!"
            Catch ex As Exception
                connectResult = ex.Message
            End Try
        End Using
        Return connectResult
    End Function
    ''' <summary>
    ''' Use for Insert, Update, Delete Command SQL
    ''' </summary>
    ''' <param name="connectString">Connect String</param>
    ''' <param name="queryString">Query String</param>
    ''' <returns>Result of Command SQL</returns>
    Public Function controlDB(ByVal connectString As String, ByVal queryString As String) As Boolean
        ' Variable
        Dim querySuccess As Boolean = False
        ' Initial Connection Setup
        Using connectSetup As New OleDbConnection(connectString)
            ' Create Command Controller
            Dim commandDB As New OleDbCommand(queryString)
            Try
                ' Open Connection
                connectSetup.Open()
                ' Command to connect
                commandDB.Connection = connectSetup
                ' Command to Execute Query
                commandDB.ExecuteNonQuery()
                ' Close Connection
                connectSetup.Close()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Using
        Return False
    End Function
    ''' <summary>
    ''' Use for select command SQL
    ''' </summary>
    ''' <param name="connectString">Connect String</param>
    ''' <param name="queryString">Query String</param>
    ''' <param name="IsItASelectQuery">TRUE</param>
    ''' <returns>Virtual Table from select command</returns>
    Public Function controlDB(ByVal connectString As String, ByVal queryString As String, ByVal IsItASelectQuery As Boolean) As List(Of Hashtable)
        ' Variable
        Dim listOfHashTableResult As New List(Of Hashtable)
        ' Initial Connection Setup
        Using connectSetup As New OleDbConnection(connectString)
            ' Create Command Controller
            Dim commandDB As New OleDbCommand(queryString)
            Try
                ' Open Connection
                connectSetup.Open()
                ' Command to connect
                commandDB.Connection = connectSetup
                ' *******************************************
                ' Task: Get Header of Table
                ' *******************************************
                ' Create virtual table to store field name and its data type
                ' -------------------------------------------
                Dim VirtualTable As New DataTable
                Dim ColumnOfVirtualTable As DataColumn
                Dim RowOfVirtualTable As DataRow
                Dim VirtualTableReader As OleDbDataReader ' Like eye of that virtual table
                Dim GetHeaderNameOutOfVirtualTable As New List(Of String)
                Dim NumberOfHeaderNameThatExistInVirtualTable As Integer = 0
                Dim CellOfVirtualTableValue As String = String.Empty
                ' -------------------------------------------
                ' Command Virtual Table to read meta of the table, you will not receive any record.
                VirtualTableReader = commandDB.ExecuteReader(CommandBehavior.KeyInfo)
                ' Input them to the Virtual Table
                VirtualTable = VirtualTableReader.GetSchemaTable()
                ' Turn off the Eye of Virtual Table
                VirtualTableReader.Close()
                ' Read the virtual Table
                ' Loop to get only the header
                ' -------------------------------------------
                ' Check For Each Row in Virtual table
                For Each RowOfVirtualTable In VirtualTable.Rows
                    ' Check for Each column in that row
                    For Each ColumnOfVirtualTable In VirtualTable.Columns
                        ' If the column have name = ColumnName Retreive that value and put it into array HearderName
                        If ColumnOfVirtualTable.ColumnName = "ColumnName" Then
                            ' Count up header name by 1
                            NumberOfHeaderNameThatExistInVirtualTable = NumberOfHeaderNameThatExistInVirtualTable + 1
                            ' Get Cell Value According by Row And Column and convert it value to string
                            CellOfVirtualTableValue = RowOfVirtualTable(ColumnOfVirtualTable).ToString
#If ShowHowItWork = True Then
                            Debug.Write(CellOfVirtualTableValue + " | ")
#End If
                            ' Insert that Value to header String
                            GetHeaderNameOutOfVirtualTable.Add(CellOfVirtualTableValue)
                            ' Clean current cell value
                            CellOfVirtualTableValue = String.Empty
                        End If
                    Next
                Next
#If ShowHowItWork = True Then
                Debug.WriteLine(br + "--------------------------------")
#End If
                ' Now we already have the header
                ' -------------------------------------------
                ' *******************************************
                ' Task: Insert Record to the real table
                ' *******************************************
                ' -------------------------------------------
                Dim recordValues As New Hashtable
                Dim numberOfRowOfCurrentReadingTable As Integer = 0
                Dim currentCellValueOfRecord As String = String.Empty
                ' Current Column readed must equal to NumberOfHeaderNameThatExistInVirtualTable all the field must be filled
                Dim currentColumnReaded As Integer = 0
                ' -------------------------------------------
                ' Open Eye of virtual table again to read the full data
                VirtualTableReader = commandDB.ExecuteReader()
                ' Fetch All Array Out 
                Do While VirtualTableReader.Read()
                    ' Count up the numberOfRowOfCurrentReadingTable
                    numberOfRowOfCurrentReadingTable = numberOfRowOfCurrentReadingTable + 1
                    ' Clean all field of last record
                    currentCellValueOfRecord = String.Empty
                    recordValues = New Hashtable
                    currentColumnReaded = 0
                    ' Loop through all the column to input value
                    While currentColumnReaded < NumberOfHeaderNameThatExistInVirtualTable
                        ' Get Current Cell Value
                        currentCellValueOfRecord = VirtualTableReader(GetHeaderNameOutOfVirtualTable.Item(currentColumnReaded)).ToString
                        If currentCellValueOfRecord.Length = 0 Then
                            currentCellValueOfRecord = "Null"
                        End If
                        ' Insert Value to cell according to headers and rows
                        recordValues(GetHeaderNameOutOfVirtualTable.Item(currentColumnReaded)) = currentCellValueOfRecord
#If ShowHowItWork = True Then
                        Debug.Write(currentCellValueOfRecord + " | ")
#End If
                        ' Reading Count Up 1
                        currentColumnReaded = currentColumnReaded + 1
                    End While
                    ' Add that record to the list of record
                    listOfHashTableResult.Add(recordValues)
#If ShowHowItWork = True Then
                    Debug.Write(br)
#End If
                Loop
                ' Close Eye
                VirtualTableReader.Close()
                ' Close Connection
                connectSetup.Close()
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
        End Using
        Return listOfHashTableResult
    End Function
End Module
