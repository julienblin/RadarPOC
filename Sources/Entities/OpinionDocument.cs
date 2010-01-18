using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities.Content;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocument : Document
    {
        public OpinionDocument()
        {
            Discussion = new FormattedContent();
            OverallEvaluation = new OpinionDocumentSection();
            InvestmentStaff = new OpinionDocumentSection();
            OrganizationalStability = new OpinionDocumentSection();
            AssetAllocation = new OpinionDocumentSection();
            Research = new OpinionDocumentSection();
            CountrySelection = new OpinionDocumentSection();
            PortfolioConstruction = new OpinionDocumentSection();
            CurrencyManagement = new OpinionDocumentSection();
            Implementation = new OpinionDocumentSection();
            SecuritySelection = new OpinionDocumentSection();
            SellDiscipline = new OpinionDocumentSection();
        }

        public virtual OpinionDocumentSection OverallEvaluation { get; set; }
        public virtual FormattedContent Discussion { get; set; }
        public virtual OpinionDocumentSection InvestmentStaff { get; set; }
        public virtual OpinionDocumentSection OrganizationalStability { get; set; }
        public virtual OpinionDocumentSection AssetAllocation { get; set; }
        public virtual OpinionDocumentSection Research { get; set; }
        public virtual OpinionDocumentSection CountrySelection { get; set; }
        public virtual OpinionDocumentSection PortfolioConstruction { get; set; }
        public virtual OpinionDocumentSection CurrencyManagement { get; set; }
        public virtual OpinionDocumentSection Implementation { get; set; }
        public virtual OpinionDocumentSection SecuritySelection { get; set; }
        public virtual OpinionDocumentSection SellDiscipline { get; set; }

        public override DocumentType DocumentType { get { return DocumentType.Opinion; } }
    }
}
