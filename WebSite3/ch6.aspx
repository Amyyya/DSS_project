<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">

    Protected Sub Page_Load(sender As Object, e As EventArgs)
        TextBox1.Visible = False
        TextBox2.Visible = False
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim sum_ha_total5, sum_abi_total5 As Integer
        sum_ha_total5 = 0
        sum_abi_total5 = 0

        Dim sum5_1 As Integer
        sum5_1 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList1.Items(i).Selected = True) Then
                sum5_1 = (sum5_1 + (i + 1)) * 2

            End If

        Next


        Dim sum5_2 As Integer
        sum5_2 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList2.Items(i).Selected = True) Then
                sum5_2 = (sum5_2 + (i + 1)) * 2

            End If

        Next

        Dim sum5_3 As Integer
        sum5_3 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList3.Items(i).Selected = True) Then
                sum5_3 = (sum5_3 + (i + 1)) * 2

            End If

        Next

        Dim sum5_4 As Integer
        sum5_4 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList4.Items(i).Selected = True) Then
                sum5_4 = (sum5_4 + (i + 1)) * 2

            End If

        Next

        Dim sum5_5 As Integer
        sum5_5 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList5.Items(i).Selected = True) Then
                sum5_5 = (sum5_5 + (i + 1)) * 2

            End If

        Next

        Dim sum5_6 As Integer
        sum5_6 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList6.Items(i).Selected = True) Then
                sum5_6 = (sum5_6 + (i + 1)) * 2

            End If

        Next

        Dim sum5_7 As Integer
        sum5_7 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList7.Items(i).Selected = True) Then
                sum5_7 = (sum5_7 + (i + 1)) * 2

            End If

        Next

        Dim sum5_8 As Integer
        sum5_8 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList8.Items(i).Selected = True) Then
                sum5_8 = (sum5_8 + (i + 1)) * 2

            End If

        Next

        Dim sum5_9 As Integer
        sum5_9 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList9.Items(i).Selected = True) Then
                sum5_9 = (sum5_9 + (i + 1)) * 2

            End If

        Next

        Dim sum5_10 As Integer
        sum5_10 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList10.Items(i).Selected = True) Then
                sum5_10 = (sum5_10 + (i + 1)) * 2

            End If

        Next

        sum_ha_total5 = sum5_1 + sum5_2 + sum5_3 + sum5_4 + sum5_5
        sum_abi_total5 = sum5_6 + sum5_7 + sum5_8 + sum5_9 + sum5_10

        TextBox1.Text = sum_ha_total5
        TextBox2.Text = sum_abi_total5

        SqlDataSource1.Update()
        Server.Transfer("ch7.aspx")


    End Sub


    '不可以刪
    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    
        <meta charset="utf-8" /><meta charset="utf-8" /><meta charset="utf-8" /><meta charset="utf-8" /><meta charset="utf-8" />
    
        <meta charset="utf-8" />
    <meta charset="utf-8" />
        <meta charset="utf-8" />
    <meta charset="utf-8" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body style="background-image: url('背景.jpg'); font-family: '微軟正黑體'; font-size: 12px;">
    <form id="form1" runat="server" style="background-image: url('背景.jpg'); font-family: '微軟正黑體 Light'; font-size: 12px;">
    <div style="font-family: 微軟正黑體; font-size: 12px">
    
        <asp:Label ID="Label3" runat="server" Text="41.從小喜歡動手組裝新家具?" Font-Size="12pt"></asp:Label></div>
       <meta charset="utf-8" />
          <asp:RadioButtonList ID="RadioButtonList1" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
   
       <br />
            <asp:Label ID="Label1" runat="server" Font-Size="12pt" Text="42.小時候對於機器人、遙控車特別有興趣?" Font-Names="微軟正黑體"></asp:Label>
        <meta charset="utf-8" />
           <br />
         <asp:RadioButtonList ID="RadioButtonList2" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
      
    
   
        <br />
        <asp:Label ID="Label4" runat="server" Font-Size="12pt" Text="43.與其天馬行空，更喜歡講求實際?" Font-Names="微軟正黑體"></asp:Label>
        <meta charset="utf-8" />
<asp:RadioButtonList ID="RadioButtonList3" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
    
        <br />
      
      
          
            <asp:Label ID="Label5" runat="server" Font-Size="12pt" Text="44.朋友常說自己是個務實、勤勞的人?" Font-Names="微軟正黑體"></asp:Label>
           <meta charset="utf-8" />
             <asp:RadioButtonList ID="RadioButtonList4" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
       
               <br />
     
       
         
      
            <asp:Label ID="Label2" runat="server" Font-Size="12pt" Text="45.認為做事必須踏實，不須有太多想像?" Font-Names="微軟正黑體"></asp:Label>
            <meta charset="utf-8" />
            <asp:RadioButtonList ID="RadioButtonList5" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
              <br />

    
        
            <asp:Label ID="Label6" runat="server" Font-Size="12pt" Text="46.是否能夠常常獨立完成某項任務?" Font-Names="微軟正黑體"></asp:Label>
             <meta charset="utf-8" />
            <asp:RadioButtonList ID="RadioButtonList6" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
            <br/>

  
            <asp:Label ID="Label7" runat="server" Font-Size="12pt" Text="47.喜歡獨自做事勝過於團體合作?" Font-Names="微軟正黑體"></asp:Label>
          <meta charset="utf-8" />
     <asp:RadioButtonList ID="RadioButtonList7" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
       
        <br />
        <asp:Label ID="Label8" runat="server" Font-Size="12pt" Text="48.從小喜歡動手算數學?" Font-Names="微軟正黑體"></asp:Label>
           <meta charset="utf-8" />
     <asp:RadioButtonList ID="RadioButtonList8" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>

        <br />

        <asp:Label ID="Label9" runat="server" Font-Size="12pt" Text="49.與其紙上談兵，不如動手實作才知道結果?" Font-Names="微軟正黑體"></asp:Label>
         <meta charset="utf-8" />
         <asp:RadioButtonList ID="RadioButtonList9" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>

        <br />
       
         <asp:Label ID="Label10" runat="server" Font-Size="12pt" Text="50.大家總認為我是個手腳靈活的人?" Font-Names="微軟正黑體"></asp:Label>
        <meta charset="utf-8" />
        <asp:RadioButtonList ID="RadioButtonList10" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br/>
        <b>
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        </br/>
        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        <br />
        </b>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT * FROM tester.DSS" UpdateCommand="UPDATE tester.DSS SET page6_HABIT = @sum_ha_total6, page6_ABI = @sum_abi_total6  WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))">
            <UpdateParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="sum_ha_total6" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox2" Name="sum_abi_total6" PropertyName="Text" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <b>
        <br/>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="下一頁"  />
        </b>
    </form>
    <br class="Apple-interchange-newline" />
    <br class="Apple-interchange-newline" />
</body>
</html>
