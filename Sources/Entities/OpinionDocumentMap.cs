using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentNHibernate.Mapping;
using Russell.RADAR.POC.Entities.Content;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocumentMap : SubclassMap<OpinionDocument>
    {
        public OpinionDocumentMap()
        {
            Map(x => x.OverallEvaluationRank);
            Map(x => x.OverallEvaluationContent);

            Map(x => x.Discussion).CustomType<FormattedContentUserType>();

            Component(x => x.InvestmentStaff, m =>
                {
                    m.Map(x => x.Rank, "InvestmentStaffRank");
                    m.Map(x => x.Content, "InvestmentStaffContent").CustomType<FormattedContentUserType>();
                }
            );

            Component(x => x.OrganizationalStability, m =>
            {
                m.Map(x => x.Rank, "OrganizationalStabilityRank");
                m.Map(x => x.Content, "OrganizationalStabilityContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.AssetAllocation, m =>
            {
                m.Map(x => x.Rank, "AssetAllocationRank");
                m.Map(x => x.Content, "AssetAllocationContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.Research, m =>
            {
                m.Map(x => x.Rank, "ResearchRank");
                m.Map(x => x.Content, "ResearchContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.CountrySelection, m =>
            {
                m.Map(x => x.Rank, "CountrySelectionRank");
                m.Map(x => x.Content, "CountrySelectionContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.PortfolioConstruction, m =>
            {
                m.Map(x => x.Rank, "PortfolioConstructionRank");
                m.Map(x => x.Content, "PortfolioConstructionContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.CurrencyManagement, m =>
            {
                m.Map(x => x.Rank, "CurrencyManagementRank");
                m.Map(x => x.Content, "CurrencyManagementContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.Implementation, m =>
            {
                m.Map(x => x.Rank, "ImplementationRank");
                m.Map(x => x.Content, "ImplementationContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.SecuritySelection, m =>
            {
                m.Map(x => x.Rank, "SecuritySelectionRank");
                m.Map(x => x.Content, "SecuritySelectionContent").CustomType<FormattedContentUserType>();
            }
            );

            Component(x => x.SellDiscipline, m =>
            {
                m.Map(x => x.Rank, "SellDisciplineRank");
                m.Map(x => x.Content, "SellDisciplineContent").CustomType<FormattedContentUserType>();
            }
            );

        }
    }
}
