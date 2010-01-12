namespace Russell.RADAR.POC.PublishingServices.TelerikTemplates
{
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;
    using Telerik.Reporting;
    using Telerik.Reporting.Drawing;

    partial class OpinionDocumentTemplate
    {
        #region Component Designer generated code
        /// <summary>
        /// Required method for telerik Reporting designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Telerik.Reporting.HtmlTextBox htmlTextBox1;
            Telerik.Reporting.HtmlTextBox htmlTextBox2;
            this.detail = new Telerik.Reporting.DetailSection();
            this.textBoxDiscussions = new Telerik.Reporting.HtmlTextBox();
            this.textBoxInvestementStaff = new Telerik.Reporting.HtmlTextBox();
            htmlTextBox1 = new Telerik.Reporting.HtmlTextBox();
            htmlTextBox2 = new Telerik.Reporting.HtmlTextBox();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            // 
            // htmlTextBox1
            // 
            htmlTextBox1.Name = "htmlTextBox1";
            htmlTextBox1.Size = new Telerik.Reporting.Drawing.SizeU(new Telerik.Reporting.Drawing.Unit(14.999899864196777, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(1.0000001192092896, Telerik.Reporting.Drawing.UnitType.Cm));
            htmlTextBox1.Value = "<span style=\"font-size: 18pt\"><strong>Discussion</strong></span>";
            // 
            // htmlTextBox2
            // 
            htmlTextBox2.Location = new Telerik.Reporting.Drawing.PointU(new Telerik.Reporting.Drawing.Unit(0, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(4.4001002311706543, Telerik.Reporting.Drawing.UnitType.Cm));
            htmlTextBox2.Name = "htmlTextBox2";
            htmlTextBox2.Size = new Telerik.Reporting.Drawing.SizeU(new Telerik.Reporting.Drawing.Unit(14.999899864196777, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(1.0000001192092896, Telerik.Reporting.Drawing.UnitType.Cm));
            htmlTextBox2.Value = "<span style=\"font-size: 18pt\"><strong>Investement Staff</strong></span>";
            // 
            // detail
            // 
            this.detail.Height = new Telerik.Reporting.Drawing.Unit(10.899999618530273, Telerik.Reporting.Drawing.UnitType.Cm);
            this.detail.Items.AddRange(new Telerik.Reporting.ReportItemBase[] {
            this.textBoxDiscussions,
            htmlTextBox1,
            htmlTextBox2,
            this.textBoxInvestementStaff});
            this.detail.Name = "detail";
            // 
            // textBoxDiscussions
            // 
            this.textBoxDiscussions.Location = new Telerik.Reporting.Drawing.PointU(new Telerik.Reporting.Drawing.Unit(0, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(1.0002001523971558, Telerik.Reporting.Drawing.UnitType.Cm));
            this.textBoxDiscussions.Name = "textBoxDiscussions";
            this.textBoxDiscussions.Size = new Telerik.Reporting.Drawing.SizeU(new Telerik.Reporting.Drawing.Unit(14.999899864196777, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(3.3996999263763428, Telerik.Reporting.Drawing.UnitType.Cm));
            // 
            // textBoxInvestementStaff
            // 
            this.textBoxInvestementStaff.Location = new Telerik.Reporting.Drawing.PointU(new Telerik.Reporting.Drawing.Unit(0.00010012308484874666, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(5.4003009796142578, Telerik.Reporting.Drawing.UnitType.Cm));
            this.textBoxInvestementStaff.Name = "textBoxInvestementStaff";
            this.textBoxInvestementStaff.Size = new Telerik.Reporting.Drawing.SizeU(new Telerik.Reporting.Drawing.Unit(14.999799728393555, Telerik.Reporting.Drawing.UnitType.Cm), new Telerik.Reporting.Drawing.Unit(4.1996989250183105, Telerik.Reporting.Drawing.UnitType.Cm));
            // 
            // OpinionDocumentTemplate
            // 
            this.Items.AddRange(new Telerik.Reporting.ReportItemBase[] {
            this.detail});
            this.PageSettings.Landscape = false;
            this.PageSettings.Margins.Bottom = new Telerik.Reporting.Drawing.Unit(2.5399999618530273, Telerik.Reporting.Drawing.UnitType.Cm);
            this.PageSettings.Margins.Left = new Telerik.Reporting.Drawing.Unit(2.5399999618530273, Telerik.Reporting.Drawing.UnitType.Cm);
            this.PageSettings.Margins.Right = new Telerik.Reporting.Drawing.Unit(2.5399999618530273, Telerik.Reporting.Drawing.UnitType.Cm);
            this.PageSettings.Margins.Top = new Telerik.Reporting.Drawing.Unit(2.5399999618530273, Telerik.Reporting.Drawing.UnitType.Cm);
            this.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter;
            this.Style.BackgroundColor = System.Drawing.Color.White;
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();

        }
        #endregion

        private Telerik.Reporting.DetailSection detail;
        private HtmlTextBox textBoxDiscussions;
        private HtmlTextBox textBoxInvestementStaff;
    }
}