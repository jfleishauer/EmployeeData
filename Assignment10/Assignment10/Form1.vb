Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports Assignment10.Assignment10
Imports System.ComponentModel

'Assignment 10 by Jaime Fleishauer

Public Class frmMain
    'Class variables
    Friend numOfEmployees As Integer 'Keeps track of current pk
    Dim db As New CompanyEntities 'DB class variable


    'Loading procedures
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        'CountQuery() '10 records
        InsertData("EmpData.txt") 'Read & insert data into db
        CountQuery() 'Reload view - 25 records total

    End Sub

    'Closing procedures
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'CountQuery() 'Reload view - 25 records total
        DeleteData("EmpData.txt") 'Read and delete data
        'CountQuery() '10 records
    End Sub

    'Read CSV txt file and import into existing DB
    Private Sub InsertData(ByVal file As String)
        'Declare table and class
        Dim db As New CompanyEntities
        'Dim employee As New Employees

        'Open text file as a csv file and split the record by delimiters
        Dim empDataFile As New TextFieldParser(file)
        empDataFile.TextFieldType = FieldType.Delimited
        empDataFile.SetDelimiters(",")

        Dim currentRow As String() 'Reads each row and saves into an array

        'Loop iteration to read until the end of the file
        While Not empDataFile.EndOfData
            Try
                currentRow = empDataFile.ReadFields() 'Read each line with delimiters

                With currentRow
                    Dim employee As New Employees
                    'Set employee object with CSV file
                    employee.Emp_Id = currentRow(0)
                    employee.First_Name = currentRow(1)
                    employee.Last_Name = currentRow(2)
                    employee.Dept_Id = CInt(currentRow(3))
                    'MsgBox(currentRow(0) & " " & currentRow(1) & " " & currentRow(2) & " " & currentRow(3) & vbCrLf & vbCrLf &
                    '       employee.Emp_Id & " " & employee.First_Name & " " & employee.Last_Name & " " & employee.Dept_Id) 'Test to see how the data appears

                    'Add & save new record into db
                    db.Employees1.Add(employee)


                    db.SaveChanges()

                End With

            Catch ex As FileIO.MalformedLineException
                MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
            End Try

        End While 'End readFile
    End Sub 'End InsertData(file)

    Private Sub CountQuery()
        'Declare variables
        Dim db As New CompanyEntities
        numOfEmployees = 0 'Refresh class variable counter

        'Use SQL query: SELECT COUNT(*) FROM employee
        Dim numOfEmp = (From employee In db.Employees1 Select employee.Emp_Id).Count

        numOfEmployees = numOfEmp 'Update class variable counter

        'Display number of employees 
        MsgBox(“Number of Employees: ” & numOfEmployees.ToString())
    End Sub

    Private Sub DeleteData(ByVal file As String)
        'Declare table and class
        Dim db As New CompanyEntities
        'Dim employee As New Employees

        'Open text file as a csv file and split the record by delimiters
        Dim empDataFile As New TextFieldParser(file)
        empDataFile.TextFieldType = FieldType.Delimited
        empDataFile.SetDelimiters(",")

        Dim currentRow As String() 'Reads each row and saves into an array
        Dim empID As String

        'Loop iteration to read until the end of the file
        While Not empDataFile.EndOfData
            Try
                currentRow = empDataFile.ReadFields() 'Read each line with delimiters

                With currentRow
                    Dim employee As New Employees
                    empID = currentRow(0)
                    employee = (From emp In db.Employees1 Where emp.Emp_Id = empID).FirstOrDefault

                    'Remove & save new record out of db
                    db.Employees1.Remove(employee)

                    db.SaveChanges()

                End With

            Catch ex As FileIO.MalformedLineException
                MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
            End Try

        End While 'End readFile
    End Sub 'End DeleteData(file)


End Class
