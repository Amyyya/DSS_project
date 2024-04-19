<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">

    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim sum_ha_total6, sum_abi_total6 As Integer
        sum_ha_total6 = 0
        sum_abi_total6 = 0

        Dim sum6_1 As Integer
        sum6_1 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList1.Items(i).Selected = True) Then
                sum6_1 = (sum6_1 + (i + 1)) * 2

            End If

        Next


        Dim sum6_2 As Integer
        sum6_2 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList2.Items(i).Selected = True) Then
                sum6_2 = (sum6_2 + (i + 1)) * 2

            End If

        Next

        Dim sum6_3 As Integer
        sum6_3 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList3.Items(i).Selected = True) Then
                sum6_3 = (sum6_3 + (i + 1)) * 2

            End If

        Next

        Dim sum6_4 As Integer
        sum6_4 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList4.Items(i).Selected = True) Then
                sum6_4 = (sum6_4 + (i + 1)) * 2

            End If

        Next

        Dim sum6_5 As Integer
        sum6_5 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList5.Items(i).Selected = True) Then
                sum6_5 = (sum6_5 + (i + 1)) * 2

            End If

        Next

        Dim sum6_6 As Integer
        sum6_6 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList6.Items(i).Selected = True) Then
                sum6_6 = (sum6_6 + (i + 1)) * 2

            End If

        Next

        Dim sum6_7 As Integer
        sum6_7 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList7.Items(i).Selected = True) Then
                sum6_7 = (sum6_7 + (i + 1)) * 2

            End If

        Next

        Dim sum6_8 As Integer
        sum6_8 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList8.Items(i).Selected = True) Then
                sum6_8 = (sum6_8 + (i + 1)) * 2

            End If

        Next

        Dim sum6_9 As Integer
        sum6_9 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList9.Items(i).Selected = True) Then
                sum6_9 = (sum6_9 + (i + 1)) * 2

            End If

        Next

        Dim sum6_10 As Integer
        sum6_10 = 0
        For i As Integer = 0 To 4
            If (RadioButtonList10.Items(i).Selected = True) Then
                sum6_10 = (sum6_10 + (i + 1)) * 2

            End If

        Next


        sum_ha_total6 = sum6_1 + sum6_2 + sum6_3 + sum6_4 + sum6_5
        sum_abi_total6 = sum6_6 + sum6_7 + sum6_8 + sum6_9 + sum6_10


        TextBox1.Text = sum_ha_total6
        TextBox2.Text = sum_abi_total6

        SqlDataSource1.Update()
        Server.Transfer("ch8.aspx")



    End Sub


    '不可以刪
    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs)
        TextBox1.Visible = False
        TextBox2.Visible = False
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    
        <meta charset="utf-8" /><meta charset="utf-8" /><meta charset="utf-8" />
    
        <meta charset="utf-8" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body style="font-size: large">
    <form id="form1" runat="server">
    <div style="background-image: url('背景.jpg')">
    
        <asp:Label ID="Label1" runat="server" Font-Size="12pt" Text="51.在團體中，喜歡偷偷觀察其他人?" Font-Names="微軟正黑體"></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList1" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
     
        <br />
        <asp:Label ID="Label2" runat="server" Font-Size="12pt" Text="52.喜歡不斷創新、變化，或是追求流行?" Font-Names="微軟正黑體"></asp:Label>
        <meta charset="utf-8" />
     
            <asp:RadioButtonList ID="RadioButtonList2" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
           <br />
        
        <asp:Label ID="Label3" runat="server" Font-Size="12pt" Text="53.常常天馬行空，對萬物保持開放態度?" Font-Names="微軟正黑體"></asp:Label>
        <meta charset="utf-8" />
        <asp:RadioButtonList ID="RadioButtonList3" runat="server" Font-Names="微軟正黑體" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
          <br />
      
        <asp:Label ID="Label4" runat="server" Font-Size="12pt" Text="54.喜歡觀察大自然的變化與各種生物?" Font-Names="微軟正黑體"></asp:Label>
         <meta charset="utf-8" />
         <asp:RadioButtonList ID="RadioButtonList4" runat="server" Font-Names="微軟正黑體" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />
 
            <asp:Label ID="Label5" runat="server" Font-Size="12pt" Text="55.課外時間，總是喜歡到圖書館閱讀、增進其他知識?" Font-Names="微軟正黑體"></asp:Label>
         <meta charset="utf-8" />   
         <asp:RadioButtonList ID="RadioButtonList5" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
              <br />
       
  
            <asp:Label ID="Label6" runat="server" Font-Size="12pt" Text="56.對於一件事情，喜歡不斷思考更好的執行步驟方法?" Font-Names="微軟正黑體"></asp:Label>
            <meta charset="utf-8" />
           
         <asp:RadioButtonList ID="RadioButtonList6" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
           
          <br />

            <asp:Label ID="Label7" runat="server" Font-Size="12pt" Text="57.對於實驗，總能細心觀察，並且了解其中原因?" Font-Names="微軟正黑體"></asp:Label>
     
        <meta charset="utf-8" />
        <asp:RadioButtonList ID="RadioButtonList7" runat="server" Font-Size="12pt" Font-Names="微軟正黑體">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
              <br />
 
            <asp:Label ID="Label8" runat="server" EnableViewState="False" Font-Size="12pt" Text="58.喜歡學習新的知識領域與創新?" Font-Names="微軟正黑體"></asp:Label>
                   <meta charset="utf-8" />
         <asp:RadioButtonList ID="RadioButtonList8" runat="server" Font-Names="微軟正黑體" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />

            <asp:Label ID="Label9" runat="server" Font-Size="12pt" Text="59.喜歡追求新鮮更勝於安定?" Font-Names="微軟正黑體"></asp:Label>
                       <meta charset="utf-8" />
         <asp:RadioButtonList ID="RadioButtonList9" runat="server" Font-Names="微軟正黑體" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
        <br />

            <asp:Label ID="Label10" runat="server" Font-Size="12pt" Text="60.喜歡打破常規，認為人必須不斷進步?" Font-Names="微軟正黑體"></asp:Label>
                 <meta charset="utf-8" />
        <asp:RadioButtonList ID="RadioButtonList10" runat="server" style="height: 132px" Font-Names="微軟正黑體" Font-Size="12pt">
            <asp:ListItem>非常不喜歡</asp:ListItem>
            <asp:ListItem>不喜歡</asp:ListItem>
            <asp:ListItem>普通</asp:ListItem>
            <asp:ListItem>喜歡</asp:ListItem>
            <asp:ListItem>非常喜歡</asp:ListItem>
        </asp:RadioButtonList>
          <br />
        <b>
        </b>
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <b>
        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        <br />
        </b>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT * FROM tester.DSS" UpdateCommand="UPDATE tester.DSS SET page7_HABIT = @sum_ha_total7, page7_ABI = @sum_abi_total7  WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))">
            <UpdateParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="sum_ha_total7" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox2" Name="sum_abi_total7" PropertyName="Text" />
            </UpdateParameters>
        </asp:SqlDataSource>
   
        <br class="Apple-interchange-newline" />
        <b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd" style="font-weight:normal;"><span style="font-size:13.999999999999998pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;"><b id="docs-internal-guid-0e76d61d-7fff-c37f-5650-154047f3d0c1" style="font-weight:normal;">
        <b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d" style="font-weight:normal;">
        <b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c" style="font-weight:normal;">
        <b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd59" style="font-weight:normal;">
        <b id="docs-internal-guid-fbeeea91-7fff-ab91-2eed-5dd6937c2e72" style="font-weight:normal;">
        <b id="docs-internal-guid-c25ec24e-7fff-1484-60ac-a2206a96a925" style="font-weight:normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd27" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d20" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c13" style="font-weight: normal;"><b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd66" style="font-weight: normal;"><b id="docs-internal-guid-fbeeea91-7fff-ab91-2eed-5dd6937c2e73" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd28" style="font-weight: normal;"><b id="docs-internal-guid-0e76d61d-7fff-c37f-5650-154047f3d0c8" style="font-weight:normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d21" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd29" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c14" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd30" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d22" style="font-weight: normal;"><b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd67" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd31" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d23" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c15" style="font-weight: normal;"><b id="docs-internal-guid-b7b9ca37-7fff-9d71-1ec1-3c4d0f8c94e13" style="font-weight: normal;"><b id="docs-internal-guid-75db62b4-7fff-4e3c-e4b4-608c8c955ed6" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd32" style="font-weight: normal;"><b id="docs-internal-guid-0e76d61d-7fff-c37f-5650-154047f3d0c9" style="font-weight:normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d24" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd33" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c16" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd34" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d25" style="font-weight: normal;"><b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd68" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd35" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d26" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c17" style="font-weight: normal;"><b id="docs-internal-guid-b7b9ca37-7fff-9d71-1ec1-3c4d0f8c94e14" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd36" style="font-weight: normal;"><b id="docs-internal-guid-0e76d61d-7fff-c37f-5650-154047f3d0c10" style="font-weight:normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d27" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd37" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c18" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd38" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d28" style="font-weight: normal;"><b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd69" style="font-weight: normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd39" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d29" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c19" style="font-weight: normal;"><b id="docs-internal-guid-f2c01782-7fff-93fd-9b47-0d2fe960c876" style="font-weight:normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd40" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d30" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c20" style="font-weight: normal;"><b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd70" style="font-weight: normal;"><b id="docs-internal-guid-fbeeea91-7fff-ab91-2eed-5dd6937c2e74" style="font-weight: normal;"><b id="docs-internal-guid-c25ec24e-7fff-1484-60ac-a2206a96a926" style="font-weight: normal;"><b id="docs-internal-guid-2e449f1d-7fff-e2b6-67fd-97d2c289ad03" style="font-weight:normal;"><b id="docs-internal-guid-e04d8868-7fff-4c73-ba97-d3966dc089bd41" style="font-weight: normal;"><b id="docs-internal-guid-6f9267a4-7fff-bc77-f288-789bc8d1dd8d31" style="font-weight: normal;"><b id="docs-internal-guid-d5aefc74-7fff-ff6c-6bcf-e990ef59b75c21" style="font-weight: normal;"><b id="docs-internal-guid-a2cda041-7fff-5b01-b0a3-a829e3a6fd71" style="font-weight: normal;"><b id="docs-internal-guid-fbeeea91-7fff-ab91-2eed-5dd6937c2e75" style="font-weight: normal;"><b id="docs-internal-guid-c25ec24e-7fff-1484-60ac-a2206a96a927" style="font-weight: normal;">
        <asp:Button ID="Button1"  runat="server" OnClick="Button1_Click" Text="下一頁"  />
        </b></span></b>
        <br class="Apple-interchange-newline" />
        
      
    
    </div>
    </form>
</body>
</html>
