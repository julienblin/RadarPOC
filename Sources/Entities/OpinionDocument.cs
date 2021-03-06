﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities.Content;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Russell.RADAR.POC.Entities
{
    public partial class OpinionDocument : Document
    {
        public OpinionDocument()
        {
            OverallEvaluationRank = 1;
            OverallEvaluationContent = string.Empty;
            Discussion = new FormattedContent();
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

        public virtual int OverallEvaluationRank { get; set; }
        public virtual string OverallEvaluationContent { get; set; }
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

        public override void StreamOpenXMLDocument(Stream stream)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                /*var mainDocumentPart = package.AddMainDocumentPart();
                var document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                mainDocumentPart.Document = document;

                var body = new Body();
                document.Append(body);

                foreach (var para in InvestmentStaff.Content.GetParagraphs())
                {
                    body.Append(para);
                }*/

                CreateParts(package);
            }
        }
    }
}
