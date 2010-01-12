<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApp._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:LinkButton ID="linkNewOpinionDocument" Text="New Opinion Document" runat="server" />
    </div>
    <div>
        <asp:Repeater ID="repeaterDocuments" runat="server">
            <HeaderTemplate>
                <table>
                    <thead>
                        <tr>
                            <th>Id</th>
                            <th>Type</th>
                            <th>State</th>
                            <th>Edit</th>
                        </tr>
                    </thead>
                    <tbody>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td><%# DataBinder.Eval(Container.DataItem, "Id") %></td>
                    <td><%# DataBinder.Eval(Container.DataItem, "DocumentType")%></td>
                    <td><%# DataBinder.Eval(Container.DataItem, "State")%></td>
                    <td><asp:HyperLink runat="server" Text="Edit" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/Edit.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' /></td>
                </tr>
            </ItemTemplate>
            <FooterTemplate>
                    </tbody>
                </table>
            </FooterTemplate>
        </asp:Repeater>
    </div>
    </form>
</body>
</html>
