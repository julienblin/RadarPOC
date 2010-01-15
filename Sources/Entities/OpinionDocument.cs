﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocument : Document
    {
        public virtual OpinionDocumentSection OverallEvaluation { get; set; }
        public virtual string Discussion { get; set; }
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
