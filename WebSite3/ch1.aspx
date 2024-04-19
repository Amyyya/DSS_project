

<!DOCTYPE html>
    
    
<%@ Page Language="VB" %>
<script runat="server">


    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        SqlDataSource1.Insert()
        Server.Transfer("ch2.aspx")
    End Sub




    Protected Sub Page_Load(sender As Object, e As EventArgs)
        GridView1.Visible = False

    End Sub



</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">
        #fram1 {
            height: 762px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div id="fram1" style="background-image: url('背景.jpg'); font-family: 微軟正黑體; font-size: 20px;">
    
        <asp:Label ID="Label1" runat="server" Text="歡迎使用性向測驗" Font-Size="Large"></asp:Label>
    
        <br />
    
        <br />
        <asp:Label ID="Label2" runat="server" Text="姓名:"></asp:Label>
    
    
       
    
        <asp:TextBox ID="TextBox4" runat="server"></asp:TextBox>
    
    
       
    
        <br />
        <asp:Label ID="Label3" runat="server" Text="興趣取向:"></asp:Label>
        <asp:TextBox ID="TextBox2" runat="server" Width="96px"></asp:TextBox>
        <asp:Label ID="Label6" runat="server" Text="%"></asp:Label>
        <br />
        <asp:Label ID="Label4" runat="server" Text="能力取向:"></asp:Label>
        <asp:TextBox ID="TextBox3" runat="server" Width="96px"></asp:TextBox>
        <asp:Label ID="Label5" runat="server" Text="%"></asp:Label>
        <br />
        <br />
        <br />
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" InsertCommand="INSERT INTO tester.DSS(habit, ability) VALUES (@habit, @ability)" SelectCommand="SELECT habit, ability FROM tester.DSS">
            <InsertParameters>
                <asp:ControlParameter ControlID="TextBox2" Name="habit" PropertyName="Text" />
                <asp:ControlParameter ControlID="TextBox3" Name="ability" PropertyName="Text" />
            </InsertParameters>
        </asp:SqlDataSource>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1">
            <Columns>
                <asp:BoundField DataField="habit" HeaderText="habit" SortExpression="habit" />
                <asp:BoundField DataField="ability" HeaderText="ability" SortExpression="ability" />
            </Columns>
        </asp:GridView>
        <br />
        <br />
        
        <br />
    
        <p>
        
        <asp:Button ID="Button1" runat="server" Text="開始測試"  OnClick="Button1_Click" />
        
        </p>
    
    </div>
    </form>
</body>
</html>
