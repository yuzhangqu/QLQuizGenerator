<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ExportXlsToDownload._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>武汉市庆龄幼儿园珠心算试卷制作</title>
</head>
<body>
    <form id="form1" runat="server">
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox1" runat="server" ReadOnly="True" Text="9级*1 + 8级*2 + 7级*2" Width="50%"></asp:TextBox>
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="下载"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox2" runat="server" ReadOnly="True" Text="8级*5" Width="50%"></asp:TextBox>
            <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="下载" style="height: 21px"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox3" runat="server" ReadOnly="True" Text="10级*1 + 9级*1 + 8级*2" Width="50%"></asp:TextBox>
            <asp:Button ID="Button3" runat="server" OnClick="Button3_Click" Text="下载" style="height: 21px"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox4" runat="server" ReadOnly="True" Text="3——5位数 10笔" Width="50%"></asp:TextBox>
            <asp:Button ID="Button4" runat="server" OnClick="Button4_Click" Text="下载" style="height: 21px"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox5" runat="server" ReadOnly="True" Text="9级*5" Width="50%"></asp:TextBox>
            <asp:Button ID="Button5" runat="server" OnClick="Button5_Click" Text="下载" style="height: 21px"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox6" runat="server" ReadOnly="True" Text="纯10级" Width="50%"></asp:TextBox>
            <asp:Button ID="Button6" runat="server" OnClick="Button6_Click" Text="下载" style="height: 21px"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox7" runat="server" ReadOnly="True" Text="2018年武汉市比赛" Width="50%"></asp:TextBox>
            <asp:Button ID="Button7" runat="server" OnClick="Button7_Click" Text="下载" style="height: 21px"/>
        </div>
        <div style="margin:20px auto; width:100%" align="center">
            <asp:TextBox ID="TextBox8" runat="server" ReadOnly="True" Text="庆龄珠心算小学组训练题" Width="50%"></asp:TextBox>
            <asp:Button ID="Button8" runat="server" OnClick="Button8_Click" Text="下载" style="height: 21px"/>
        </div>
    </form>
</body>
</html>
