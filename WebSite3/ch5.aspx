<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">



    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim sum_ha_total4, sum_abi_total4 As Integer
        sum_ha_total4 = 0
        sum_abi_total4 = 0

        Dim sum4_1 As Integer
        sum4_1 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList1.Items(i).Selected = True) Then
                sum4_1 = (sum4_1 + (i + 1)) * 2

            End If

        Next


        Dim sum4_2 As Integer
        sum4_2 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList2.Items(i).Selected = True) Then
                sum4_2 = (sum4_2 + (i + 1)) * 2

            End If

        Next

        Dim sum4_3 As Integer
        sum4_3 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList3.Items(i).Selected = True) Then
                sum4_3 = (sum4_3 + (i + 1)) * 2

            End If

        Next

        Dim sum4_4 As Integer
        sum4_4 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList4.Items(i).Selected = True) Then
                sum4_4 = (sum4_4 + (i + 1)) * 2

            End If

        Next

        Dim sum4_5 As Integer
        sum4_5 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList5.Items(i).Selected = True) Then
                sum4_5 = (sum4_5 + (i + 1)) * 2

            End If

        Next

        Dim sum4_6 As Integer
        sum4_6 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList6.Items(i).Selected = True) Then
                sum4_6 = (sum4_6 + (i + 1)) * 2

            End If

        Next

        Dim sum4_7 As Integer
        sum4_7 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList7.Items(i).Selected = True) Then
                sum4_7 = (sum4_7 + (i + 1)) * 2

            End If

        Next

        Dim sum4_8 As Integer
        sum4_8 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList8.Items(i).Selected = True) Then
                sum4_8 = (sum4_8 + (i + 1)) * 2

            End If

        Next

        Dim sum4_9 As Integer
        sum4_9 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList9.Items(i).Selected = True) Then
                sum4_9 = (sum4_9 + (i + 1)) * 2

            End If

        Next

        Dim sum4_10 As Integer
        sum4_10 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList10.Items(i).Selected = True) Then
                sum4_10 = (sum4_10 + (i + 1)) * 2

            End If

        Next




        sum_ha_total4 = sum4_1 + sum4_2 + sum4_3 + sum4_4 + sum4_5
        sum_abi_total4 = sum4_6 + sum4_7 + sum4_8 + sum4_9 + sum4_10

        TextBox1.Text = sum_ha_total4
        TextBox2.Text = sum_abi_total4

        SqlDataSource1.Update()
        Server.Transfer("ch6.aspx")


    End Sub


    '不可以刪
    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub


    Protected Sub Page_Load(sender As Object, e As EventArgs)
        TextBox1.Visible = False
        TextBox2.Visible = False
    End Sub

    Protected Sub SqlDataSource1_Selecting(sender As Object, e As SqlDataSourceSelectingEventArgs)

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
        <meta charset="utf-8" />
            <meta charset="utf-8" />
    <meta charset="utf-8" />
    <meta charset="utf-8" />
    <meta charset="utf-8" />
    <meta charset="utf-8" />
    <meta charset="utf-8" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body style="background-image: url('背景.jpg'); height: 2139px; width: 820px; font-family: 微軟正黑體; font-size: 12px; margin-top: 12px;">
    <form id="form1" runat="server" style="background-image: url('背景.jpg'); height: 2139px; width: 820px; font-family: 微軟正黑體; font-size: 12px; margin-top: 0px;" submitdisabledcontrols="True">
    <div>
    
        <asp:Label ID="Label2" runat="server" Text="31.是否喜歡傾聽別人?" Font-Size="12pt"></asp:Label >
    
        <meta charset="utf-8" />
        </div>
        <asp:RadioButtonList ID="RadioButtonList1" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <p>
            <asp:Label ID="Label3" runat="server" Text="32.是否喜歡了解別人?" Font-Size="12pt"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList2" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        </p>
        <b id="docs-internal-guid-6a6d74e1-7fff-b141-429f-3f9956b0b3f2" style="font-weight:normal;">
        <p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt;">
            <asp:Label ID="Label5" runat="server" Font-Size="12pt" Text="33.是否願意付出時間和精力解決他人的衝突?"></asp:Label>
        </p>
        </b>
        <asp:RadioButtonList ID="RadioButtonList3" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label6" runat="server" Font-Size="12pt" Text=" 34.是否喜歡教導別人?"></asp:Label>
        <meta charset="utf-8" />
        <br class="Apple-interchange-newline" />
        <asp:RadioButtonList ID="RadioButtonList4" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label7" runat="server" Font-Size="12pt" Text="35.是否喜歡團體一起做事勝於獨立作業?"></asp:Label>
        <meta charset="utf-8" />
        <br />
        <asp:RadioButtonList ID="RadioButtonList5" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        
        <br />
        
        <asp:Label ID="Label8" runat="server" Font-Size="12pt" Text="36.是否善於傾聽?"></asp:Label>
          <meta charset="utf-8" />
        <br />
        <asp:RadioButtonList ID="RadioButtonList6" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>

      
  
            <br />
            <asp:Label ID="Label9" runat="server" Font-Size="12pt" Text="37.是否能體會到某人想要和他人溝通的需要?"></asp:Label>
             <meta charset="utf-8" />
        <br/>       
        <asp:RadioButtonList ID="RadioButtonList7" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
         <br />
       
            <asp:Label ID="Label10" runat="server" Font-Size="12pt" Text="38.是否能幫助別人自我改進?"></asp:Label>
             <meta charset="utf-8" />
           <br />
     
        <asp:RadioButtonList ID="RadioButtonList8" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
          <br />
       
        
            <asp:Label ID="Label11" runat="server" Font-Size="12pt" Text="39.是否具有與他人溫和相處的能力?"></asp:Label>
         <meta charset="utf-8" />
         <br />
      
        <asp:RadioButtonList ID="RadioButtonList9" runat="server" Height="118px" Width="155px" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />

        <asp:Label ID="Label12" runat="server" Font-Size="12pt" Text="40.經常關注到社會上有許多人需要幫助?"></asp:Label>
        <meta charset="utf-8" />
         <br />
        <asp:RadioButtonList ID="RadioButtonList10" runat="server" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
       
        <br />
        <br />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="下一頁" style="width: 74px"/>
        </b>
        </b>
        </b>
        <br />
        </span>
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        </b>
        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        <br />
        </b>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT * FROM [DSS]" UpdateCommand="UPDATE tester.DSS SET page5_HABIT = @sum_ha_total5, page5_ABI = @sum_abi_total5  WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))">
            <UpdateParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="sum_ha_total5" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox2" Name="sum_abi_total5" PropertyName="Text" />
            </UpdateParameters>
        </asp:SqlDataSource>
        </b>
        </b>
        </span></b>
        <br class="Apple-interchange-newline" />
        <br class="Apple-interchange-newline" />
    </form>
</body>
</html>
