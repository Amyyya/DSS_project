<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">


    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim sum_ha_total1, sum_abi_total1 As Integer

        sum_ha_total1 = 0
        sum_abi_total1 = 0

        Dim sum1_1 As Integer
        sum1_1 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList2.Items(i).Selected = True) Then
                sum1_1 = (sum1_1 + (i + 1)) * 2

            End If

        Next


        Dim sum1_2 As Integer
        sum1_2 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList3.Items(i).Selected = True) Then
                sum1_2 = (sum1_2 + (i + 1)) * 2

            End If

        Next

        Dim sum1_3 As Integer
        sum1_3 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList4.Items(i).Selected = True) Then
                sum1_3 = (sum1_3 + (i + 1)) * 2

            End If

        Next

        Dim sum1_4 As Integer
        sum1_4 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList5.Items(i).Selected = True) Then
                sum1_4 = (sum1_4 + (i + 1)) * 2

            End If

        Next

        Dim sum1_5 As Integer
        sum1_5 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList6.Items(i).Selected = True) Then
                sum1_5 = (sum1_5 + (i + 1)) * 2

            End If

        Next

        Dim sum1_6 As Integer
        sum1_6 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList7.Items(i).Selected = True) Then
                sum1_6 = (sum1_6 + (i + 1)) * 2

            End If

        Next

        Dim sum1_7 As Integer
        sum1_7 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList8.Items(i).Selected = True) Then
                sum1_7 = (sum1_7 + (i + 1)) * 2

            End If

        Next

        Dim sum1_8 As Integer
        sum1_8 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList9.Items(i).Selected = True) Then
                sum1_8 = (sum1_8 + (i + 1)) * 2

            End If

        Next

        Dim sum1_9 As Integer
        sum1_9 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList10.Items(i).Selected = True) Then
                sum1_9 = (sum1_9 + (i + 1)) * 2

            End If

        Next

        Dim sum1_10 As Integer
        sum1_10 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList11.Items(i).Selected = True) Then
                sum1_10 = (sum1_10 + (i + 1)) * 2

            End If

        Next




        sum_ha_total1 = sum1_1 + sum1_2 + sum1_3 + sum1_4 + sum1_5
        sum_abi_total1 = sum1_6 + sum1_7 + sum1_8 + sum1_9 + sum1_10

        TextBox1.Text = sum_ha_total1
        TextBox2.Text = sum_abi_total1
        SqlDataSource1.Update()
        Server.Transfer("ch3.aspx")





    End Sub


    '不可以刪



    Protected Sub Page_Load(sender As Object, e As EventArgs)
        TextBox1.Visible = False
        TextBox2.Visible = False
    End Sub

    Protected Sub SqlDataSource1_Selecting(sender As Object, e As SqlDataSourceSelectingEventArgs)

    End Sub

    Protected Sub RadioButtonList2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Protected Sub RadioButtonList11_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub



    Protected Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        #form1 {
            height: 1993px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server" style="background-image: url('背景.jpg'); font-family: 微軟正黑體; font-size: 16px;">
    
        <asp:Label ID="Label9" runat="server" Text="1.是否常常覺得精力旺盛?"></asp:Label>
        <br />
    
   
        <asp:RadioButtonList ID="RadioButtonList2" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
        <asp:Label ID="Label10" runat="server" Text="2.是否喜歡做事有計畫?"></asp:Label>
        <br />
        <asp:RadioButtonList ID="RadioButtonList3" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <p id="docs-internal-guid-09901a45-7fff-85b7-625a-f4fc1cc146e4" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            &nbsp;</p>
        <p dir="ltr" style="color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label11" runat="server" Text="3.是否喜歡冒險競爭?"></asp:Label>
        </p>
        <asp:RadioButtonList ID="RadioButtonList4" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
    
   
        <asp:Label ID="Label12" runat="server" Text="4.是否做事喜歡立刻行動?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList5" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
    
   
        <asp:Label ID="Label13" runat="server" Text="5.是否有正義感?"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList6" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
        <asp:Label ID="Label14" runat="server" Text="6.是否擅長說服別人?"></asp:Label>
    
   
        <br />
    
   
        <asp:RadioButtonList ID="RadioButtonList7" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
        <p id="docs-internal-guid-5a728ac2-7fff-5980-82a3-cbece0a4cf70" dir="ltr" style="color: rgb(0, 0, 0); font-family: 微軟正黑體; font-size: 12px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label15" runat="server" Font-Size="12pt" Text="7.是否具有組織能力?"></asp:Label>
        </p>
        <asp:RadioButtonList ID="RadioButtonList8" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
        <p id="docs-internal-guid-a75f8dcc-7fff-03d3-d53d-24ff9f80ad5b" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label16" runat="server" Font-Size="12pt" Text="8.是否常常是團體的焦點人物?"></asp:Label>
        </p>
        <asp:RadioButtonList ID="RadioButtonList9" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
        <p id="docs-internal-guid-d8d8990e-7fff-6526-5c35-3713961ed084" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 4pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label17" runat="server" Text="9.是否擅長銷售?"></asp:Label>
        </p>
    
   
        <asp:RadioButtonList ID="RadioButtonList10" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
          <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
   
        <br />
        <p id="docs-internal-guid-ffc688e2-7fff-d245-f4fd-666edf478805" dir="ltr" style="color: rgb(0, 0, 0); font-family: Times New Roman; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; letter-spacing: normal; line-height: 1.38; margin-bottom: 0pt; margin-top: 0pt; orphans: 2; text-align: left; text-decoration: none; text-indent: 0px; text-transform: none; -webkit-text-stroke-width: 0px; white-space: normal; word-spacing: 0px;">
            <asp:Label ID="Label18" runat="server" Text="10.是否善於說話?" Font-Overline="False" Font-Names="微軟正黑體"></asp:Label>
        </p>
    
   
        <asp:RadioButtonList ID="RadioButtonList11" runat="server" OnSelectedIndexChanged="RadioButtonList2_SelectedIndexChanged">
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
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" OnSelecting="SqlDataSource1_Selecting" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT * FROM tester.DSS


" UpdateCommand="UPDATE tester.DSS SET page2_HABIT = @sum_ha_total2, page2_ABI = @sum_abi_total2 WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))">
            <UpdateParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="sum_ha_total2" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox2" Name="sum_abi_total2" PropertyName="Text" />
            </UpdateParameters>
        </asp:SqlDataSource>
     
        <br />
    
   
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="下一頁" />
    
   
    </form>
</body>
</html>
