<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">



    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim sum_ha_total2, sum_abi_total2 As Integer
        sum_ha_total2 = 0
        sum_abi_total2 = 0

        Dim sum2_1 As Integer
        sum2_1 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList1.Items(i).Selected = True) Then
                sum2_1 = (sum2_1 + (i + 1)) * 2

            End If

        Next


        Dim sum2_2 As Integer
        sum2_2 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList2.Items(i).Selected = True) Then
                sum2_2 = (sum2_2 + (i + 1)) * 2

            End If

        Next

        Dim sum2_3 As Integer
        sum2_3 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList3.Items(i).Selected = True) Then
                sum2_3 = (sum2_3 + (i + 1)) * 2

            End If

        Next

        Dim sum2_4 As Integer
        sum2_4 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList4.Items(i).Selected = True) Then
                sum2_4 = (sum2_4 + (i + 1)) * 2

            End If

        Next

        Dim sum2_5 As Integer
        sum2_5 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList5.Items(i).Selected = True) Then
                sum2_5 = (sum2_5 + (i + 1)) * 2

            End If

        Next

        Dim sum2_6 As Integer
        sum2_6 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList6.Items(i).Selected = True) Then
                sum2_6 = (sum2_6 + (i + 1)) * 2

            End If

        Next

        Dim sum2_7 As Integer
        sum2_7 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList7.Items(i).Selected = True) Then
                sum2_7 = (sum2_7 + (i + 1)) * 2

            End If

        Next

        Dim sum2_8 As Integer
        sum2_8 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList8.Items(i).Selected = True) Then
                sum2_8 = (sum2_8 + (i + 1)) * 2

            End If

        Next

        Dim sum2_9 As Integer
        sum2_9 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList9.Items(i).Selected = True) Then
                sum2_9 = (sum2_9 + (i + 1)) * 2

            End If

        Next

        Dim sum2_10 As Integer
        sum2_10 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList10.Items(i).Selected = True) Then
                sum2_10 = (sum2_10 + (i + 1)) * 2

            End If

        Next




        sum_ha_total2 = sum2_1 + sum2_2 + sum2_3 + sum2_4 + sum2_5
        sum_abi_total2 = sum2_6 + sum2_7 + sum2_8 + sum2_9 + sum2_10

        TextBox1.Text = sum_ha_total2
        TextBox2.Text = sum_abi_total2
        SqlDataSource1.Update()
        Server.Transfer("ch4.aspx")

    End Sub


    '不可以刪
    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub


    Protected Sub RadioButtonList7_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs)
        TextBox1.Visible = False
        TextBox2.Visible = False
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server" style="background-image: url('背景.jpg')">
    <div style="background-image: url('背景.jpg'); height: 2122px; font-family: 微軟正黑體;">
    
        <p id="docs-internal-guid-b4f5a044-7fff-a80b-19e2-45eafdbccf45" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">11.<span id="docs-internal-guid-387ab147-7fff-b4c7-844e-b7c6442d8332" style="background-color: transparent; color: rgb(0, 0, 0); font-family: Arial; font-size: 11pt; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; vertical-align: baseline; -webkit-text-stroke-width: 0px; white-space: pre-wrap; word-spacing: 0px;">.是否喜歡在有清楚規範的環境下工作?</span></span></p>
        <asp:RadioButtonList ID="RadioButtonList1" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <p id="docs-internal-guid-b4f5a044-7fff-a80b-19e2-45eafdbccf46" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">12.是否做事講求規矩和精確?</span></p>
        <asp:RadioButtonList ID="RadioButtonList2" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <p id="docs-internal-guid-1b97d264-7fff-5b61-19cc-377466fbeefc" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">13.是否喜歡做事按部就班?</span></p>
        <asp:RadioButtonList ID="RadioButtonList3" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p id="docs-internal-guid-1b97d264-7fff-5b61-19cc-377466fbeefc0" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">14.<span id="docs-internal-guid-296f0e86-7fff-6e51-592b-747336f2cd17" style="background-color: transparent; color: rgb(0, 0, 0); font-family: Arial; font-size: 11pt; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; vertical-align: baseline; -webkit-text-stroke-width: 0px; white-space: pre-wrap; word-spacing: 0px;">是否不喜歡改變?</span></span></p>
        <asp:RadioButtonList ID="RadioButtonList4" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p id="docs-internal-guid-5b37a361-7fff-031b-5cc1-4b5354a97813" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">15.是否喜歡安穩的生活?</span></p>
        <asp:RadioButtonList ID="RadioButtonList5" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p id="docs-internal-guid-4af4ff9b-7fff-502e-8e0b-1021d6b04f0a" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">16.是否個性謹慎?</span></p>
        <asp:RadioButtonList ID="RadioButtonList6" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p id="docs-internal-guid-149785f9-7fff-18da-aa2c-8e9f3eee752f" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <span style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">17.是否常給人的感覺是有效率?</span></p>
        <asp:RadioButtonList ID="RadioButtonList7" runat="server" style="height: 122px; width: 106px">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            18<span style="background-color: transparent; color: rgb(0, 0, 0); font-family: Arial; font-size: 11pt; font-style: normal; font-variant: normal; font-weight: 400; text-decoration: none; vertical-align: baseline; white-space: pre-wrap;">.</span><span id="docs-internal-guid-38f53765-7fff-a658-eb54-e9d170a9491f0" style="background-color: transparent; color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12pt; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; vertical-align: baseline; -webkit-text-stroke-width: 0px; white-space: pre-wrap; word-spacing: 0px;">是否做事仔細?</span></p>
        <asp:RadioButtonList ID="RadioButtonList8" runat="server" Height="137px">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label1" runat="server" Font-Size="12pt" Text="19.是否專注力強?"></asp:Label>
        </p>
        <asp:RadioButtonList ID="RadioButtonList9" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label2" runat="server" Font-Size="12pt" Text="20.是否有耐心?" Font-Names="微軟正黑體"></asp:Label>
        </p>
        <asp:RadioButtonList ID="RadioButtonList10" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        <br />
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT * FROM tester.DSS" UpdateCommand="UPDATE tester.DSS SET page3_HABIT = @sum_ha_total3, page3_ABI = @sum_abi_total3  WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))">
            <UpdateParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="sum_ha_total3" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox2" Name="sum_abi_total3" PropertyName="Text" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <br />
    
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="下一頁"/>
    
        <br />
    
    </div>
    </form>
</body>
</html>
