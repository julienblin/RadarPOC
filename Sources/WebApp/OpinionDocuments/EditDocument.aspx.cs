﻿using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;
using Russell.RADAR.POC.Entities;
using Russell.RADAR.POC.Infrastructure.NH;
using Russell.RADAR.POC.AuthoringServices;

namespace Russell.RADAR.POC.WebApp.OpinionDocuments
{
    public partial class EditDocument : BasePage
    {
        OpinionDocument doc;

        protected override void OnInitComplete(EventArgs e)
        {
            base.OnInitComplete(e);
            buttonOk.Click += new EventHandler(buttonOk_Click);
            buttonCancel.Click += new EventHandler(buttonCancel_Click);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            using (var uow = UnitOfWork.Start())
            {
                var authoringService = Resolve<IAuthoringService>();
                doc = (OpinionDocument)authoringService.Retrieve(Convert.ToInt32(Request.Params["id"]));
                uow.Commit();
            }

            if (!IsPostBack)
            {
                ddlOverallRank.DataSource = new int[] { 1, 2, 3, 4 };
                ddlOverallRank.DataBind();

                textBoxOverall.Text = doc.OverallEvaluation.Content;
                ddlOverallRank.SelectedIndex = doc.OverallEvaluation.Rank - 1;

                textBoxDiscussion.Text = doc.Discussion;

                SectionInvestementStaff.Section = doc.InvestmentStaff;
                SectionOrganizationalStability.Section = doc.OrganizationalStability;
                SectionAssetAllocation.Section = doc.AssetAllocation;
                SectionResearch.Section = doc.Research;
                SectionCountrySelection.Section = doc.CountrySelection;
                SectionPortfolioConstruction.Section = doc.PortfolioConstruction;
                SectionCurrencyManagement.Section = doc.CurrencyManagement;
                SectionImplementation.Section = doc.Implementation;
                SectionSecuritySelection.Section = doc.SecuritySelection;
                SectionSellDiscipline.Section = doc.SellDiscipline;
            }
        }

        void buttonOk_Click(object sender, EventArgs e)
        {
            using (var uow = UnitOfWork.Start())
            {
                doc.OverallEvaluation.Content = textBoxOverall.Text;
                doc.OverallEvaluation.Rank = ddlOverallRank.SelectedIndex + 1;

                doc.Discussion = textBoxDiscussion.Text;

                doc.InvestmentStaff.Rank = SectionInvestementStaff.GetRank();
                doc.InvestmentStaff.Content = SectionInvestementStaff.GetContent();
                doc.OrganizationalStability.Rank = SectionOrganizationalStability.GetRank();
                doc.OrganizationalStability.Content = SectionOrganizationalStability.GetContent();
                doc.AssetAllocation.Rank = SectionAssetAllocation.GetRank();
                doc.AssetAllocation.Content = SectionAssetAllocation.GetContent();
                doc.Research.Rank = SectionResearch.GetRank();
                doc.Research.Content = SectionResearch.GetContent();
                doc.CountrySelection.Rank = SectionCountrySelection.GetRank();
                doc.CountrySelection.Content = SectionCountrySelection.GetContent();
                doc.PortfolioConstruction.Rank = SectionPortfolioConstruction.GetRank();
                doc.PortfolioConstruction.Content = SectionPortfolioConstruction.GetContent();
                doc.CurrencyManagement.Rank = SectionCurrencyManagement.GetRank();
                doc.CurrencyManagement.Content = SectionCurrencyManagement.GetContent();
                doc.Implementation.Rank = SectionImplementation.GetRank();
                doc.Implementation.Content = SectionImplementation.GetContent();
                doc.SecuritySelection.Rank = SectionSecuritySelection.GetRank();
                doc.SecuritySelection.Content = SectionSecuritySelection.GetContent();
                doc.SellDiscipline.Rank = SectionSellDiscipline.GetRank();
                doc.SellDiscipline.Content = SectionSellDiscipline.GetContent();
                uow.Commit();
            }
            Response.Redirect("~/Default.aspx");
        }

        void buttonCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Default.aspx");
        }
    }
}