namespace Russell.RADAR.POC.PublishingServices.TelerikTemplates
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;
    using Telerik.Reporting;
    using Telerik.Reporting.Drawing;

    /// <summary>
    /// Summary description for OpinionDocument.
    /// </summary>
    public partial class OpinionDocumentTemplate : Telerik.Reporting.Report
    {
        public OpinionDocumentTemplate()
        {
            /// <summary>
            /// Required for telerik Reporting designer support
            /// </summary>
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        public string Discussion
        {
            get { return textBoxDiscussions.Value; }
            set { textBoxDiscussions.Value = value; }
        }

        public string InvestementStaff
        {
            get { return textBoxInvestementStaff.Value; }
            set { textBoxInvestementStaff.Value = value; }
        }
    }
}