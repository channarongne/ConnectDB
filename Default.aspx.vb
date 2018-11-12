Imports System.Data

Partial Class _Default
    Inherits System.Web.UI.Page
    Dim condb As New ConnectDB
    Dim ds As New DataSet
    Dim sql1 As String = ""

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            'Code this page

        End If

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Dim empcode As String = ""
            empcode = TextBox1.Text.Trim.Replace("'", "")

            'sql = "sp_storeprocedure '1','2'"
            'sql = "sp_storeprocedure @Parameter1='Test',@Parameter2='Test2' "

            'sql1 = "   SELECT TOP 10 FLDEMP_CODE,FLDEMP_TNAME FROM dbo.TB_VW_EMP_YNEW WHERE FLDEMP_CODE = '" & empcode & "' "
            sql1 = "   SELECT TOP 10 FLDEMP_CODE,FLDEMP_TNAME FROM dbo.TB_VW_EMP_YNEW "
            ds = condb.QueryDataSet(sql1)
            If ds.Tables(0).Rows.Count > 0 Then
                Label1.Text = ""
                For i = 0 To ds.Tables(0).Rows.Count - 1

                    Label1.Text = Label1.Text + ds.Tables(0).Rows(i).Item("FLDEMP_TNAME").ToString + vbNewLine

                Next

            Else
                Response.Write(condb.MsgBoxAlert("ไม่พบข้อมูล"))
            End If

        Catch ex As Exception
            Response.Write(condb.MsgBoxAlert(ex.Message))
        End Try


    End Sub
End Class
