Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Windows.Forms
Imports System
Imports System.Web.UI.WebControls

Public Class ConnectDB
    Private csConn As New SqlConnection
    Public csConn2 As New SqlConnection
    Private csConnect As String = ConfigurationManager.ConnectionStrings("ConnectionString1").ConnectionString
    Dim objConn As New SqlConnection
    Dim objCmd As New SqlCommand
    Dim dtAdapter As New SqlDataAdapter
    Dim ds As New DataSet


    Public ReadOnly Property ConnectDB() As SqlConnection
        Get
            csConn = New SqlConnection(csConnect)
            Return csConn
        End Get
    End Property
    Public ReadOnly Property ConnectDBHr() As SqlConnection
        Get
            csConn = New SqlConnection(csConnect)
            Return csConn
        End Get
    End Property


    Public Function ReadDataSet(ByVal vSQL As String, ByVal vTBName As String) As DataSet
        Dim da As New SqlDataAdapter
        Dim ds As New DataSet
        Try
            csConn = ConnectDB
            csConn.Open()
            da = New SqlDataAdapter(vSQL, csConn)
            da.Fill(ds, vTBName)
            csConn.Close()
            Return ds
        Catch ex As Exception
            If csConn.State = ConnectionState.Open Then csConn.Close()
            Return ds
        Finally
            csConn.Close()
        End Try
    End Function




    Public Function GetData(ByVal vSql As String, ByVal vTbName As String, ByVal vColName As String, ByVal vDef As String) As String
        Dim strValue As String = vDef
        Dim ds As New DataSet
        Try

            ds = ReadDataSet(vSql, vTbName)
            If ds.Tables(vTbName).Rows.Count > 0 Then
                strValue = ds.Tables(vTbName).Rows(0).Item(vColName).ToString()
            End If
            ds.Clear()

            If strValue = "" Then strValue = vDef
            Return strValue
        Catch ex As Exception
            Return strValue
        Finally
            ds.Dispose()
        End Try
    End Function

    Public Function ExecuteSQL(ByVal vSql As String) As Integer
        Dim dc As New SqlCommand
        Dim strResult As Integer = 0
        Try
            csConn = ConnectDB
            csConn.Open()
            dc = New SqlCommand(vSql, csConn)
            strResult = dc.ExecuteNonQuery
            csConn.Close()
            Return strResult
        Catch ex As Exception
            If csConn.State = ConnectionState.Open Then csConn.Close()
            Return strResult
        End Try
    End Function

    Public Sub LoadDataToChk(ByVal vSQL As String, ByVal vTable As String, ByVal vFldShow As String, ByVal vFldValue As String, ByVal vCheck As String, ByVal vStartShow As String, ByVal vStartValue As String, ByVal vObj As CheckBoxList)
        Dim i As Long = 0
        Try
            Dim ds As New DataSet
            Dim ir As Long = 0
            Dim iCheck As Long = 0
            Dim strCheck As String
            ds = ReadDataSet(vSQL, vTable)
            vObj.Items.Clear()
            If vStartShow <> "" Or vStartValue <> "" Then
                vObj.Items.Add(vStartShow)
                vObj.Items(i).Value = vStartValue
                i = i + 1
            End If
            With ds.Tables(vTable)
                If .Rows.Count > 0 Then
                    For ir = 0 To .Rows.Count - 1
                        If Not IsDBNull(.Rows(ir).Item(vFldShow)) Then
                            strCheck = .Rows(ir).Item(vFldValue)
                            vObj.Items.Add(.Rows(ir).Item(vFldShow))
                            vObj.Items(i).Value = strCheck
                            If CStr(strCheck) = CStr(vCheck) Then iCheck = i
                            i = i + 1
                        End If
                    Next
                    If iCheck > 0 Then
                        vObj.SelectedIndex = iCheck
                    End If
                End If
            End With
            ds.Clear()
            ds.Dispose()
        Catch ex As Exception
            vObj.Items.Clear()
            If vStartShow <> "" Or vStartValue <> "" Then
                i = vObj.Items.Count
                vObj.Items.Add(vStartShow)
                vObj.Items(i).Value = vStartValue
            End If
        End Try

    End Sub

    Public Function ValueNull(ByVal sValue As String) As String
        If sValue Is DBNull.Value Then
            ValueNull = ""
        Else
            ValueNull = sValue
        End If
        Return ValueNull
    End Function

    Public Function ConvertStr(ByVal Text As String) As String
        ConvertStr = Replace(Text, "'", "''")
        Return ConvertStr
    End Function


    Public Sub CloseDBHr()
        Dim ds As New DataSet
        ds.Clear()
        ds.Dispose()

    End Sub

    Public Function MsgBoxAlert(ByVal strMsg As String) As String
        Dim strAlert As String = ""

        strAlert = strAlert & "<script language='javascript'>"
        strAlert = strAlert & "alert('" & strMsg & "');"
        strAlert = strAlert & "</script>"
        Return strAlert

    End Function

    Public Function WindowOpen(ByVal vURL As String, ByVal vtitle As String, ByVal vWidth As Long, ByVal vHeight As Long) As String

        Try
            Dim strStringAlert As String = ""
            strStringAlert = strStringAlert & "<script language='javascript'>"
            strStringAlert = strStringAlert & vbCrLf & "window.open('" & vURL & "', 'tinyWindow', 'width=" & vWidth & ",height=" & vHeight & ",toolbar=no')"
            strStringAlert = strStringAlert & "</script>"
            Return strStringAlert
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function QueryAdapter(ByVal strSQL1 As String) As SqlDataAdapter

        objConn.ConnectionString = csConnect
        With objCmd
            .Connection = objConn
            .CommandText = strSQL1
            .CommandType = CommandType.Text
        End With
        dtAdapter.SelectCommand = objCmd

        dtAdapter.Fill(ds)

        Return dtAdapter

    End Function

    Public Function QueryDataReader(ByVal strSQL As String) As SqlDataReader
        Dim dtReader As SqlDataReader
        objConn = New SqlConnection
        With objConn
            .ConnectionString = csConnect
            .Open()
        End With
        objCmd = New SqlCommand(strSQL, objConn)
        dtReader = objCmd.ExecuteReader()
        Return dtReader '*** Return DataReader ***'

        dtReader.Close()
        dtReader = Nothing
    End Function

    Public Function QueryDataSet(ByVal strSQL As String) As DataSet

        If objConn.State = ConnectionState.Open Then objConn.Close()
        Dim ds As New DataSet
        Dim dtAdapter As New SqlDataAdapter
        objConn = New SqlConnection
        With objConn
            .ConnectionString = csConnect
            .Open()
        End With
        objCmd = New SqlCommand
        With objCmd
            .Connection = objConn
            .CommandText = strSQL
            .CommandType = CommandType.Text
        End With
        dtAdapter.SelectCommand = objCmd
        dtAdapter.Fill(ds)

        Return ds '*** Return DataSet ***'

    End Function

    Public Function QueryDataTable(ByVal strSQL As String) As DataTable
        Dim dtAdapter As SqlDataAdapter
        Dim dt As New DataTable
        objConn = New SqlConnection
        'If objconn.State = ConnectionState.Open Then
        '    Close()
        'End If
        With objConn
            .ConnectionString = csConnect
            .Open()
        End With
        dtAdapter = New SqlDataAdapter(strSQL, objConn)
        dtAdapter.Fill(dt)
        Return dt '*** Return DataTable ***'
        ' dtAdapter.Dispose()

    End Function


End Class
