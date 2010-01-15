using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentNHibernate.Mapping;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocumentMap : SubclassMap<OpinionDocument>
    {
        public OpinionDocumentMap()
        {
            Component(x => x.OverallEvaluation, m =>
            {
                m.Map(x => x.Rank, "OverallEvaluationRank");
                m.Map(x => x.Content, "OverallEvaluationContent");
            }
            );

            Map(x => x.Discussion).Length(10000);

            Component(x => x.InvestmentStaff, m =>
                {
                    m.Map(x => x.Rank, "InvestmentStaffRank");
                    m.Map(x => x.Content, "InvestmentStaffContent");
                }
            );

            Component(x => x.OrganizationalStability, m =>
            {
                m.Map(x => x.Rank, "OrganizationalStabilityRank");
                m.Map(x => x.Content, "OrganizationalStabilityContent");
            }
            );

            Component(x => x.AssetAllocation, m =>
            {
                m.Map(x => x.Rank, "AssetAllocationRank");
                m.Map(x => x.Content, "AssetAllocationContent");
            }
            );

            Component(x => x.Research, m =>
            {
                m.Map(x => x.Rank, "ResearchRank");
                m.Map(x => x.Content, "ResearchContent");
            }
            );

            Component(x => x.CountrySelection, m =>
            {
                m.Map(x => x.Rank, "CountrySelectionRank");
                m.Map(x => x.Content, "CountrySelectionContent");
            }
            );

            Component(x => x.PortfolioConstruction, m =>
            {
                m.Map(x => x.Rank, "PortfolioConstructionRank");
                m.Map(x => x.Content, "PortfolioConstructionContent");
            }
            );

            Component(x => x.CurrencyManagement, m =>
            {
                m.Map(x => x.Rank, "CurrencyManagementRank");
                m.Map(x => x.Content, "CurrencyManagementContent");
            }
            );

            Component(x => x.Implementation, m =>
            {
                m.Map(x => x.Rank, "ImplementationRank");
                m.Map(x => x.Content, "ImplementationContent");
            }
            );

            Component(x => x.SecuritySelection, m =>
            {
                m.Map(x => x.Rank, "SecuritySelectionRank");
                m.Map(x => x.Content, "SecuritySelectionContent");
            }
            );

            Component(x => x.SellDiscipline, m =>
            {
                m.Map(x => x.Rank, "SellDisciplineRank");
                m.Map(x => x.Content, "SellDisciplineContent");
            }
            );

        }
    }
}
