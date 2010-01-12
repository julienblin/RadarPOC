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
                            <th colspan="5">Actions</th>
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
                    <td><asp:LinkButton runat="server" Text="Publish" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Id")%>' OnClick="Publish_Click" /></td>
                    <td><asp:HyperLink runat="server" Text="Print-HTML" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintHTML.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' /></td>
                    <td><asp:HyperLink ID="HyperLink1" runat="server" Text="Print-PDF" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintPDF.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' /></td>
                    <td><asp:LinkButton runat="server" Text="Delete" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Id")%>' OnClick="Delete_Click" /></td>
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
