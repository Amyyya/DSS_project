<%@ Page Language="VB" %>

<!DOCTYPE html>

<script runat="server">

    Protected Sub Page_Load(sender As Object, e As EventArgs)
        GridView2.Visible = False
        GridView3.Visible = False
        GridView4.Visible = False
        GridView5.Visible = False
        TextBox1.Visible = False
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        Dim habit As Single
        Dim ability As Single
        Dim habbit_value(6) As Single
        Dim ability_value(6) As Single
        Dim habit_percent As Integer
        Dim ability_percent As Integer
        Dim habbit_name() As String = {"Eh", "Ch", "Ah", "Sh", "Rh", "Ih"}
        Dim habbit_answer As String
        Dim ability_name() As String = {"Ea", "Ca", "Aa", "Sa", "Ra", "Ia"}
        Dim ability_answer As String

        Dim i As Integer
        Dim j As Integer
        Dim x As Integer
        Dim y As Integer
        Dim max As String

        i = 0
        j = 0
        x = 0
        y = 0
        SqlDataSource2.Insert()
        GridView2.Visible = True
        SqlDataSource3.Insert()
        GridView5.Visible = True

        habit = Me.GridView2.Rows(0).Cells(6).Text.Trim()
        ability = Me.GridView5.Rows(0).Cells(6).Text.Trim()
        habit_percent = Me.GridView1.Rows(0).Cells(0).Text.Trim()
        ability_percent = Me.GridView1.Rows(0).Cells(1).Text.Trim()

        For i = 0 To 5
            habbit_value(i) = Me.GridView2.Rows(0).Cells(i).Text.Trim()
        Next
        For j = 0 To 5
            ability_value(j) = Me.GridView5.Rows(0).Cells(j).Text.Trim()
        Next
        For x = 0 To 5
            If (habit = habbit_value(x)) Then
                habbit_answer = habbit_name(x)
            End If
        Next
        For y = 0 To 5
            If (ability = ability_value(y)) Then
                ability_answer = ability_name(y)
            End If
        Next
        If habit > ability Then
            max = habbit_answer

        Else
            max = ability_answer
        End If

        If (habit_percent > ability_percent) Then
            Select Case max
                Case "Eh"
                    Label4.Text = "你選擇了更加注重興趣，而測出來的結果最高分落在企業型的興趣取向，因此非常建議你往企業型發展。"
                Case "Ch"
                    Label4.Text = "你選擇了更加注重興趣，而測出來的結果最高分落在事務型的興趣取向，因此非常建議你往事務型發展。"
                Case "Ah"
                    Label4.Text = "你選擇了更加注重興趣，而測出來的結果最高分落在藝術型的興趣取向，因此非常建議你往藝術型發展。"
                Case "Sh"
                    Label4.Text = "你選擇了更加注重興趣，而測出來的結果最高分落在社會型的興趣取向，因此非常建議你往社會型發展。"
                Case "Rh"
                    Label4.Text = "你選擇了更加注重興趣，而測出來的結果最高分落在實用型的興趣取向，因此非常建議你往實用型發展。"
                Case "Ih"
                    Label4.Text = "你選擇了更加注重興趣，而測出來的結果最高分落在研究型的興趣取向，因此非常建議你往研究型發展。"
                Case "Ea"
                    Label4.Text = "你選擇了更加注重興趣的比例，但測出來結果最高分仍然為能力取向的企業型。事實上以你的能力非常適合企業型。因此你可以多考慮參考自己的興趣往企業型發展。"
                Case "Ca"
                    Label4.Text = "你選擇了更加注重興趣的比例，但測出來結果最高分仍然為能力取向的事務型。事實上以你的能力非常適合事務型。因此你可以多考慮參考自己的能力往事務型發展。"
                Case "Aa"
                    Label4.Text = "你選擇了更加注重興趣的比例，但測出來結果最高分仍然為能力取向的藝術型。事實上以你的能力非常適合藝術型。因此你可以多考慮參考自己的能力往藝術型發展。"
                Case "Sa"
                    Label4.Text = "你選擇了更加注重興趣的比例，但測出來結果最高分仍然為能力取向的社會型。事實上以你的能力非常適合社會型。因此你可以多考慮參考自己的能力往社會型發展。"
                Case "Ra"
                    Label4.Text = "你選擇了更加注重興趣的比例，但測出來結果最高分仍然為能力取向的實用型。事實上以你的能力非常適合實用型。因此你可以多考慮參考自己的能力往實用型發展。"
                Case "Ia"
                    Label4.Text = "你選擇了更加注重興趣的比例，但測出來結果最高分仍然為能力取向的研究型。事實上以你的能力非常適合研究型。因此你可以多考慮參考自己的能力往研究型發展。"

            End Select
        ElseIf (habit_percent = ability_percent) Then
            Select Case max
                Case "Eh"
                    Label4.Text = "測出來的結果為，以你的興趣非常適合企業型，因此建議你可以往企業型發展。"
                Case "Ch"
                    Label4.Text = "測出來的結果為，以你的興趣非常適合事務型，因此建議你可以往事務型發展。"
                Case "Ah"
                    Label4.Text = "測出來的結果為，以你的興趣非常適合藝術型，因此建議你可以往藝術型發展。"
                Case "Sh"
                    Label4.Text = "測出來的結果為，以你的興趣非常適合社會型，因此建議你可以往社會型發展。"
                Case "Rh"
                    Label4.Text = "測出來的結果為，以你的興趣非常適合實用型，因此建議你可以往實用型發展。"
                Case "Ih"
                    Label4.Text = "測出來的結果為，以你的興趣非常適合研究型，因此建議你可以往研究型發展。"
                Case "Ea"
                    Label4.Text = "測出來的結果為，以你的能力非常適合企業型，因此建議你可以往企業型發展。"
                Case "Ca"
                    Label4.Text = "測出來的結果為，以你的能力非常適合事務型，因此建議你可以往事務型發展。"
                Case "Aa"
                    Label4.Text = "測出來的結果為，以你的能力非常適合藝術型，因此建議你可以往藝術型發展。"
                Case "Sa"
                    Label4.Text = "測出來的結果為，以你的能力非常適合社會型，因此建議你可以往社會型發展。"
                Case "Ra"
                    Label4.Text = "測出來的結果為，以你的能力非常適合實用型，因此建議你可以往實用型發展。"
                Case "Ia"
                    Label4.Text = "測出來的結果為，以你的能力非常適合研究型，因此建議你可以往研究型發展。"
            End Select


        Else
            Select Case max
                Case "Eh"
                    Label4.Text = "你選擇了更加注重能力的比例，但測出來結果最高分仍然為興趣取向的企業型。事實上以你的興趣非常適合企業型。因此你可以多考慮參考自己的興趣往企業型發展。"
                Case "Ch"
                    Label4.Text = "你選擇了更加注重能力的比例，但測出來結果最高分仍然為興趣取向的事務型。事實上以你的興趣非常適合事務型。因此你可以多考慮參考自己的興趣往事務型發展。"
                Case "Ah"
                    Label4.Text = "你選擇了更加注重能力的比例，但測出來結果最高分仍然為興趣取向的藝術型。事實上以你的興趣非常適合藝術型。因此你可以多考慮參考自己的興趣往藝術型發展。"
                Case "Sh"
                    Label4.Text = "你選擇了更加注重能力的比例，但測出來結果最高分仍然為興趣取向的社會型。事實上以你的興趣非常適合社會型。因此你可以多考慮參考自己的興趣往社會型發展。"
                Case "Rh"
                    Label4.Text = "你選擇了更加注重能力的比例，但測出來結果最高分仍然為興趣取向的實用型。事實上以你的興趣非常適合實用型。因此你可以多考慮參考自己的興趣往實用型發展。"
                Case "Ih"
                    Label4.Text = "你選擇了更加注重能力的比例，但測出來結果最高分仍然為興趣取向的研究型。事實上以你的興趣非常適合研究型。因此你可以多考慮參考自己的興趣往研究型發展。"
                Case "Ea"
                    Label4.Text = "你選擇了更加注重能力，而測出來的結果最高分落在企業型的能力取向，因此非常建議你往企業型發展。"
                Case "Ca"
                    Label4.Text = "你選擇了更加注重能力，而測出來的結果最高分落在事務型的能力取向，因此非常建議你往事務型發展。"
                Case "Aa"
                    Label4.Text = "你選擇了更加注重能力，而測出來的結果最高分落在藝術型的能力取向，因此非常建議你往藝術型發展。"
                Case "Sa"
                    Label4.Text = "{你選擇了更加注重能力，而測出來的結果最高分落在社會型的能力取向，因此非常建議你往社會型發展。"
                Case "Ra"
                    Label4.Text = "你選擇了更加注重能力，而測出來的結果最高分落在實用型的能力取向，因此非常建議你往實用型發展。"
                Case "Ia"
                    Label4.Text = "你選擇了更加注重能力，而測出來的結果最高分落在研究型的能力取向，因此非常建議你往研究型發展。"
            End Select
        End If



        If max = "Ia" Or max = "Ih" Then
            TextBox1.Text = "I"
        ElseIf max = "Ea" Or max = "Eh" Then
            TextBox1.Text = "E"

        ElseIf max = "Ra" Or max = "Rh" Then
            TextBox1.Text = "R"

        ElseIf max = "Aa" Or max = "Ah" Then
            TextBox1.Text = "A"

        ElseIf max = "Ca" Or max = "Ch" Then
            TextBox1.Text = "C"
        ElseIf max = "Sa" Or max = "Sh" Then
            TextBox1.Text = "S"
        End If
        SqlDataSource6.Select(DataSourceSelectArguments.Empty)

    End Sub

    Protected Sub SqlDataSource2_Selecting(sender As Object, e As SqlDataSourceSelectingEventArgs)

    End Sub

    Protected Sub GridView4_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Protected Sub SqlDataSource3_Selecting(sender As Object, e As SqlDataSourceSelectingEventArgs)

    End Sub

    Protected Sub SqlDataSource6_Selecting(sender As Object, e As SqlDataSourceSelectingEventArgs)

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server" style="background-image: url('背景.jpg')">
       同學您的測試結果顯示如下:<br />
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT habit as N'興趣權重', ability as N'能力權重', page2_HABIT as N'企業型興趣得分', page2_ABI as N'企業型能力得分', page3_HABIT as N'事務型興趣得分', page3_ABI as N'事務型能力得分', page4_HABIT as N'藝術型興趣得分', page4_ABI as N'藝術型能力得分', page5_HABIT as N'社會型興趣得分', page5_ABI as N'社會型能力得分', page6_HABIT as N'實用型興趣得分', page6_ABI as N'實用型能力得分', page7_HABIT as N'研究型興趣得分', page7_ABI as N'研究型能力得分' FROM tester.DSS WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))"></asp:SqlDataSource>
        <br />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="興趣權重" HeaderText="興趣權重" SortExpression="興趣權重" />
                <asp:BoundField DataField="能力權重" HeaderText="能力權重" SortExpression="能力權重" />
                <asp:BoundField DataField="企業型興趣得分" HeaderText="企業型興趣得分" SortExpression="企業型興趣得分" />
                <asp:BoundField DataField="企業型能力得分" HeaderText="企業型能力得分" SortExpression="企業型能力得分" />
                <asp:BoundField DataField="事務型興趣得分" HeaderText="事務型興趣得分" SortExpression="事務型興趣得分" />
                <asp:BoundField DataField="事務型能力得分" HeaderText="事務型能力得分" SortExpression="事務型能力得分" />
                <asp:BoundField DataField="藝術型興趣得分" HeaderText="藝術型興趣得分" SortExpression="藝術型興趣得分" />
                <asp:BoundField DataField="藝術型能力得分" HeaderText="藝術型能力得分" SortExpression="藝術型能力得分" />
                <asp:BoundField DataField="社會型興趣得分" HeaderText="社會型興趣得分" SortExpression="社會型興趣得分" />
                <asp:BoundField DataField="社會型能力得分" HeaderText="社會型能力得分" SortExpression="社會型能力得分" />
                <asp:BoundField DataField="實用型興趣得分" HeaderText="實用型興趣得分" SortExpression="實用型興趣得分" />
                <asp:BoundField DataField="實用型能力得分" HeaderText="實用型能力得分" SortExpression="實用型能力得分" />
                <asp:BoundField DataField="研究型興趣得分" HeaderText="研究型興趣得分" SortExpression="研究型興趣得分" />
                <asp:BoundField DataField="研究型能力得分" HeaderText="研究型能力得分" SortExpression="研究型能力得分" />
            </Columns>
        </asp:GridView>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label5" runat="server" Text="%"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label6" runat="server" Text="%"></asp:Label>
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="查看更多" />
        <br />
        <br />
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT habit * 0.01 * page2_HABIT AS '企業型興趣', habit * 0.01 * page3_HABIT AS '事務型興趣', habit * 0.01 * page4_HABIT AS '藝術型興趣', habit * 0.01 * page5_HABIT AS '社會型興趣', habit * 0.01 * page6_HABIT AS '實用型興趣', habit * 0.01 * page7_HABIT AS '研究型興趣' FROM tester.DSS WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))" InsertCommand="INSERT INTO tester.finally_habit(E_habit, C_habit, A_habit, S_habit, R_habit, I_habit) SELECT habit * 0.01 * page2_HABIT AS Expr1, habit * 0.01 * page3_HABIT AS Expr2, habit * 0.01 * page4_HABIT AS Expr3,habit * 0.01 * page5_HABIT AS Expr4, habit * 0.01 * page6_HABIT AS Expr5,habit * 0.01 * page7_HABIT AS Expr6 FROM tester.DSS WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))"></asp:SqlDataSource>
        <br />
        <asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource2" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="企業型興趣" HeaderText="企業型興趣" ReadOnly="True" SortExpression="企業型興趣" />
                <asp:BoundField DataField="事務型興趣" HeaderText="事務型興趣" ReadOnly="True" SortExpression="事務型興趣" />
                <asp:BoundField DataField="藝術型興趣" HeaderText="藝術型興趣" ReadOnly="True" SortExpression="藝術型興趣" />
                <asp:BoundField DataField="社會型興趣" HeaderText="社會型興趣" ReadOnly="True" SortExpression="社會型興趣" />
                <asp:BoundField DataField="實用型興趣" HeaderText="實用型興趣" ReadOnly="True" SortExpression="實用型興趣" />
                <asp:BoundField DataField="研究型興趣" HeaderText="研究型興趣" ReadOnly="True" SortExpression="研究型興趣" />
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT  E_habit  ,C_habit , A_habit, S_habit , R_habit , I_habit , (SELECT max(c) FROM (VALUES (E_habit) , (C_habit) , (A_habit), (S_habit) ,(R_habit) ,(I_habit)) AS cValues(c)) AS N'分數高的興趣指標' FROM tester.finally_habit WHERE id= (SELECT max(id) FROM tester.finally_habit)"></asp:SqlDataSource>
        <br />
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource4" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="E_habit" HeaderText="E_habit" SortExpression="E_habit" />
                <asp:BoundField DataField="C_habit" HeaderText="C_habit" SortExpression="C_habit" />
                <asp:BoundField DataField="A_habit" HeaderText="A_habit" SortExpression="A_habit" />
                <asp:BoundField DataField="S_habit" HeaderText="S_habit" SortExpression="S_habit" />
                <asp:BoundField DataField="R_habit" HeaderText="R_habit" SortExpression="R_habit" />
                <asp:BoundField DataField="I_habit" HeaderText="I_habit" SortExpression="I_habit" />
                <asp:BoundField DataField="分數高的興趣指標" HeaderText="分數高的興趣指標" SortExpression="分數高的興趣指標" ReadOnly="True" />
            </Columns>
        </asp:GridView>
        <br />
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT ability * 0.01 * page2_ABI AS '企業型能力', ability * 0.01 * page3_ABI AS '事務型能力', ability * 0.01 * page4_ABI AS '藝術型能力', ability * 0.01 * page5_ABI AS '社會型能力', ability * 0.01 * page6_ABI AS '實用型能力', ability * 0.01 * page7_ABI AS '研究型能力' FROM tester.DSS WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))" OnSelecting="SqlDataSource3_Selecting" InsertCommand="INSERT INTO tester.finally_ability(E_ability, C_ability, A_ability, S_ability, R_ability, I_ability) SELECT ability * 0.01 * page2_ABI AS Expr1, ability * 0.01 * page3_ABI AS Expr2, ability * 0.01 * page4_ABI AS Expr3, ability * 0.01 * page5_ABI AS Expr4, ability * 0.01 * page6_ABI AS Expr5, ability * 0.01 * page7_ABI AS Expr6 FROM tester.DSS WHERE (id = (SELECT MAX(id) AS Expr1 FROM tester.DSS AS DSS_1))"></asp:SqlDataSource>
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource3" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="企業型能力" HeaderText="企業型能力" ReadOnly="True" SortExpression="企業型能力" />
                <asp:BoundField DataField="事務型能力" HeaderText="事務型能力" ReadOnly="True" SortExpression="事務型能力" />
                <asp:BoundField DataField="藝術型能力" HeaderText="藝術型能力" ReadOnly="True" SortExpression="藝術型能力" />
                <asp:BoundField DataField="社會型能力" HeaderText="社會型能力" ReadOnly="True" SortExpression="社會型能力" />
                <asp:BoundField DataField="實用型能力" HeaderText="實用型能力" ReadOnly="True" SortExpression="實用型能力" />
                <asp:BoundField DataField="研究型能力" HeaderText="研究型能力" ReadOnly="True" SortExpression="研究型能力" />
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT  E_ability ,C_ability , A_ability, S_ability , R_ability , I_ability , (SELECT max(c) FROM (VALUES (E_ability) , (C_ability) , (A_ability), (S_ability) ,(R_ability) ,(I_ability)) AS cValues(c)) AS N'分數高的能力指標' FROM tester.finally_ability WHERE id= (SELECT max(id) FROM tester.finally_ability)"></asp:SqlDataSource>
        <br />
        <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource5" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="E_ability" HeaderText="E_ability" SortExpression="E_ability" />
                <asp:BoundField DataField="C_ability" HeaderText="C_ability" SortExpression="C_ability" />
                <asp:BoundField DataField="A_ability" HeaderText="A_ability" SortExpression="A_ability" />
                <asp:BoundField DataField="S_ability" HeaderText="S_ability" SortExpression="S_ability" />
                <asp:BoundField DataField="R_ability" HeaderText="R_ability" SortExpression="R_ability" />
                <asp:BoundField DataField="I_ability" HeaderText="I_ability" SortExpression="I_ability" />
                <asp:BoundField DataField="分數高的能力指標" HeaderText="分數高的能力指標" SortExpression="分數高的能力指標" ReadOnly="True" />
            </Columns>
        </asp:GridView>
        <br />
        <asp:Label ID="Label4" runat="server"></asp:Label>
        <br />
        <br />
        <asp:SqlDataSource ID="SqlDataSource6" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT department as N'適合科系'    FROM tester.type_department WHERE (test_type = @max)" OnSelecting="SqlDataSource6_Selecting">
            <SelectParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="max" PropertyName="Text" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:GridView ID="GridView6" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource6" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="適合科系" HeaderText="適合科系" SortExpression="適合科系" />
            </Columns>
        </asp:GridView>
        <br />
        <asp:SqlDataSource ID="SqlDataSource7" runat="server" ConnectionString="<%$ ConnectionStrings:DSS_projectConnectionString2 %>" SelectCommand="SELECT  job as N'適合工作' FROM tester.type_job WHERE (test_type = @max2)">
            <SelectParameters>
                <asp:ControlParameter ControlID="TextBox1" Name="max2" PropertyName="Text" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:GridView ID="GridView7" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource7" Font-Names="微軟正黑體">
            <Columns>
                <asp:BoundField DataField="適合工作" HeaderText="適合工作" SortExpression="適合工作" />
            </Columns>
        </asp:GridView>
        <br />
        <asp:TextBox ID="TextBox1" runat="server" Width="86px"></asp:TextBox>
        <br />
    </form>
</body>
</html>
