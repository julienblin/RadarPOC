<%@ Page Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true"
    CodeBehind="EditDocument.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.OpinionDocuments.EditDocument"
    Title="RadarPOC" %>

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
            Overall Evaluation</h3>
        <table class="current" cellspacing="0" cellpadding="0">
            <tbody>
                <tr>
                    <td width="5%">
                        <img src="/images/2(2).gif" border="0">
                    </td>
                    <td valign="top">
                        <span id="PageTemplate_ucOpinionContent_lblRankDescription">We recommend that mutual
                            clients actively evaluate replacement managers.</span>
                    </td>
                </tr>
            </tbody>
        </table>
        <h3>
            Discussion</h3>
        <p class="Normal-P">
            We recommend that mutual clients actively evaluate replacement managers. Morgan
            Stanley Investment Management’s (MSIM’s) Asia ex Japan equity product is based on
            fundamental bottom-up research within a top-down country and sector/thematic overlay.
            MSIM’s approach emphasises bottom-up stock selection as the primary source of added
            value with the team seeking to identify stocks with attractive growth prospects.
            This approach can lead to significant bets against the index at sector and stock
            level. Under the leadership of Ashutosh Sinha we have observed a bias towards smaller
            companies in portfolios. We expect the product to perform in line with its benchmark
            over 3-5 years within a tracking error of 5%-6% relative to the MSCI Far East Free
            ex Japan index.
        </p>
        <h3>
            Investment Staff
            <img src="/images/3(1).gif" /></h3>
        <p class="Normal-P">
            We recommend that mutual clients actively evaluate replacement managers. Morgan
            Stanley Investment Management’s (MSIM’s) Asia ex Japan equity product is based on
            fundamental bottom-up research within a top-down country and sector/thematic overlay.
            MSIM’s approach emphasises bottom-up stock selection as the primary source of added
            value with the team seeking to identify stocks with attractive growth prospects.
            This approach can lead to significant bets against the index at sector and stock
            level. Under the leadership of Ashutosh Sinha we have observed a bias towards smaller
            companies in portfolios. We expect the product to perform in line with its benchmark
            over 3-5 years within a tracking error of 5%-6% relative to the MSCI Far East Free
            ex Japan index.
        </p>
    </div>
</asp:Content>
