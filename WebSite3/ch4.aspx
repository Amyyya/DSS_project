<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">

    Protected Sub Page_Load(sender As Object, e As EventArgs)
        TextBox1.Visible = False
        TextBox2.Visible = False
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim sum_ha_total3, sum_abi_total3 As Integer
        sum_ha_total3 = 0
        sum_abi_total3 = 0

        Dim sum3_1 As Integer
        sum3_1 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList1.Items(i).Selected = True) Then
                sum3_1 = (sum3_1 + (i + 1)) * 2

            End If

        Next


        Dim sum3_2 As Integer
        sum3_2 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList2.Items(i).Selected = True) Then
                sum3_2 = (sum3_2 + (i + 1)) * 2

            End If

        Next

        Dim sum3_3 As Integer
        sum3_3 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList3.Items(i).Selected = True) Then
                sum3_3 = (sum3_3 + (i + 1)) * 2

            End If

        Next

        Dim sum3_4 As Integer
        sum3_4 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList4.Items(i).Selected = True) Then
                sum3_4 = (sum3_4 + (i + 1)) * 2

            End If

        Next

        Dim sum3_5 As Integer
        sum3_5 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList5.Items(i).Selected = True) Then
                sum3_5 = (sum3_5 + (i + 1)) * 2

            End If

        Next

        Dim sum3_6 As Integer
        sum3_6 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList6.Items(i).Selected = True) Then
                sum3_6 = (sum3_6 + (i + 1)) * 2

            End If

        Next

        Dim sum3_7 As Integer
        sum3_7 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList7.Items(i).Selected = True) Then
                sum3_7 = (sum3_7 + (i + 1)) * 2

            End If

        Next

        Dim sum3_8 As Integer
        sum3_8 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList8.Items(i).Selected = True) Then
                sum3_8 = (sum3_8 + (i + 1)) * 2

            End If

        Next

        Dim sum3_9 As Integer
        sum3_9 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList9.Items(i).Selected = True) Then
                sum3_9 = (sum3_9 + (i + 1)) * 2

            End If

        Next

        Dim sum3_10 As Integer
        sum3_10 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList10.Items(i).Selected = True) Then
                sum3_10 = (sum3_10 + (i + 1)) * 2

            End If

        Next

        sum_ha_total3 = sum3_1 + sum3_2 + sum3_3 + sum3_4 + sum3_5
        sum_abi_total3 = sum3_6 + sum3_7 + sum3_8 + sum3_9 + sum3_10


        TextBox1.Text = sum_ha_total3
        TextBox2.Text = sum_abi_total3
        SqlDataSource1.Update()
        Server.Transfer("ch5.aspx")

    End Sub


    '不可以刪
    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server" style="background-image: url('背景.jpg')">
    <div style="background-image: url('背景.jpg'); height: 2052px; font-family: 微軟正黑體;">
    
        <asp:Label ID="Label1" runat="server" Text="21.是否喜歡是創造不平凡的事物?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList1" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label2" runat="server" Text="22.是否不喜歡管人和被別人管?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList2" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label3" runat="server" Text="23.心情是否受到音樂、色彩、寫作、和美麗事務的影響極大?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList3" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label4" runat="server" Text="24.是否崇尚獨創?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList4" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label5" runat="server" Text="25.是否喜歡設計?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList5" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label6" runat="server" Text="26.是否善於表達?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList6" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label7" runat="server" Text="27.是否善於創新?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList7" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label8" runat="server" Text="28.當從事創造事物時，我會忘掉一切舊經驗?"></asp:Label>
        <br />
        <asp:RadioButtonList ID="RadioButtonList8" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label9" runat="server" Text="29.是否缺乏文書事務的能力?"></asp:Label>
        <br />
        <asp:RadioButtonList ID="RadioButtonList9" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label10" runat="server" Text="30.是否有創造力?"></asp:Label>
        <br />
        <asp:RadioButtonList ID="RadioButtonList10" runat="server">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        <br />
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT id, habit, ability, page2_HABIT, page2_ABI, page3_HABIT, page3_ABI, page4_HABIT, page4_ABI, page5_HABIT, page5_ABI, page6_HABIT, page6_ABI, page7_HABIT, page7_ABI FROM tester.DSS" UpdateCommand="UPDATE tester.DSS SET page4_HABIT = @sum_ha_total4, page4_ABI = @sum_abi_total4  WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))">
            <UpdateParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="sum_ha_total4" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox2" Name="sum_abi_total4" PropertyName="Text" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <br />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="下一頁"  />
        <br />
    
    </div>
    </form>
</body>
</html>
