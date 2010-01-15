<%@ Page Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true"
    CodeBehind="EditDocument.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.OpinionDocuments.EditDocument"
    Title="RadarPOC" ValidateRequest="false" %>
<%@ Register Src="~/OpinionDocuments/Components/SectionEditor.ascx" TagPrefix="radar" TagName="SectionEditor" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContentPlaceHolder" runat="server">
    <div id="ProductNav">
        <table style="width: 100%" cellspacing="2" cellpadding="2">
            <tbody>
                <tr id="buttonRow">
                    <td valign="top">
                        <div id="PageTemplate_ucContentResearch_panTitle">
                            <h1>
                                Opinion</h1>
                        </div>
                    </td>
                    <td>
                        <div align="right">
                            <asp:ImageButton ID="linkOutput" runat="server" ImageUrl="~/images/btn_Output.gif" />
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    <div id="contentData">
        <h3>
            Overall Evaluation&nbsp;&nbsp;<asp:DropDownList ID="ddlOverallRank" runat="server" /></h3>
        <asp:TextBox ID="textBoxOverall" TextMode="MultiLine" Rows="10" Width="100%" runat="server" />
        
        <h3>Discussion</h3>
        <asp:TextBox ID="textBoxDiscussion" TextMode="MultiLine" Rows="10" Columns="120" CssClass="makeckeditor" runat="server" />
        
        <radar:SectionEditor id="SectionInvestementStaff" Title="Investment Staff" runat="server" />
        <radar:SectionEditor id="SectionOrganizationalStability" Title="Organizational Stability" runat="server" />
        <radar:SectionEditor id="SectionAssetAllocation" Title="Asset Allocation" runat="server" />
        <radar:SectionEditor id="SectionResearch" Title="Research" runat="server" />
        <radar:SectionEditor id="SectionCountrySelection" Title="Country Selection" runat="server" />
        <radar:SectionEditor id="SectionPortfolioConstruction" Title="Portfolio Construction" runat="server" />
        <radar:SectionEditor id="SectionCurrencyManagement" Title="Currency Management" runat="server" />
        <radar:SectionEditor id="SectionImplementation" Title="Implementation" runat="server" />
        <radar:SectionEditor id="SectionSecuritySelection" Title="Security Selection" runat="server" />
        <radar:SectionEditor id="SectionSellDiscipline" Title="Sell Discipline" runat="server" />
        <br />
        <div>
            <asp:Button ID="buttonOk" Text="Save" runat="server" />
            <asp:Button ID="buttonCancel" Text="Cancel" runat="server" />
        </div>
    </div>
</asp:Content>
<asp:Content ContentPlaceHolderID="ScriptsPlaceHolder" runat="server">
    <script type="text/javascript">
        $(".makeckeditor").ckeditor();
    </script>

</asp:Content>
