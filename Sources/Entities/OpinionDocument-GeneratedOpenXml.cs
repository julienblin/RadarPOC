using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using Russell.RADAR.POC.Entities.Content;

namespace Russell.RADAR.POC.Entities
{
    public partial class OpinionDocument
    {
        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            ImagePart imagePart1 = mainDocumentPart1.AddNewPart<ImagePart>("image/gif", "rId8");
            GenerateImagePart1Content(imagePart1);

            FooterPart footerPart1 = mainDocumentPart1.AddNewPart<FooterPart>("rId13");
            GenerateFooterPart1Content(footerPart1);

            ImagePart imagePart2 = footerPart1.AddNewPart<ImagePart>("image/gif", "rId2");
            GenerateImagePart2Content(imagePart2);

            ImagePart imagePart3 = footerPart1.AddNewPart<ImagePart>("image/gif", "rId1");
            GenerateImagePart3Content(imagePart3);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            documentSettingsPart1.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", new System.Uri("file:///C:\\Documents%20and%20Settings\\ppelletier.RUSSELL\\Application%20Data\\Microsoft\\Templates\\RADAR%20Template.dot", System.UriKind.Absolute), "rId1");
            ImagePart imagePart4 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rId7");
            GenerateImagePart4Content(imagePart4);

            HeaderPart headerPart1 = mainDocumentPart1.AddNewPart<HeaderPart>("rId12");
            GenerateHeaderPart1Content(headerPart1);

            ImagePart imagePart5 = headerPart1.AddNewPart<ImagePart>("image/png", "rId1");
            GenerateImagePart5Content(imagePart5);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId1");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId6");
            GenerateEndnotesPart1Content(endnotesPart1);

            FooterPart footerPart2 = mainDocumentPart1.AddNewPart<FooterPart>("rId11");
            GenerateFooterPart2Content(footerPart2);

            footerPart2.AddPart(imagePart2, "rId2");

            footerPart2.AddPart(imagePart3, "rId1");

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId5");
            GenerateFootnotesPart1Content(footnotesPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId15");
            GenerateThemePart1Content(themePart1);

            HeaderPart headerPart2 = mainDocumentPart1.AddNewPart<HeaderPart>("rId10");
            GenerateHeaderPart2Content(headerPart2);

            ImagePart imagePart6 = headerPart2.AddNewPart<ImagePart>("image/jpeg", "rId1");
            GenerateImagePart6Content(imagePart6);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId4");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            ImagePart imagePartTopicRank1 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rIdImgPartTopicRank1");
            GenerateImagePartTopicRankContent(imagePartTopicRank1, imagePartTopicRank1Data);

            ImagePart imagePartTopicRank2 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rIdImgPartTopicRank2");
            GenerateImagePartTopicRankContent(imagePartTopicRank2, imagePartTopicRank2Data);

            ImagePart imagePartTopicRank3 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rIdImgPartTopicRank3");
            GenerateImagePartTopicRankContent(imagePartTopicRank3, imagePartTopicRank3Data);

            ImagePart imagePartTopicRank4 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rIdImgPartTopicRank4");
            GenerateImagePartTopicRankContent(imagePartTopicRank4, imagePartTopicRank4Data);

            ImagePart imagePartTopicRank5 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rIdImgPartTopicRank5");
            GenerateImagePartTopicRankContent(imagePartTopicRank5, imagePartTopicRank5Data);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId14");
            GenerateFontTablePart1Content(fontTablePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            Ap.Template template1 = new Ap.Template();
            template1.Text = "RADAR Template.dot";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "7";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "151";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "831";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "6";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "CGI";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "981";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            DocumentFormat.OpenXml.Wordprocessing.Document document1 = new DocumentFormat.OpenXml.Wordprocessing.Document();

            Body body1 = new Body();

            CustomXmlBlock customXmlBlock1 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "reportdoc" };

            CustomXmlBlock customXmlBlock2 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "body" };

            CustomXmlBlock customXmlBlock3 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "opinion" };

            CustomXmlBlock customXmlBlock4 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "product" };

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "006A0B1E", RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00397A9E", RsidRunAdditionDefault = "005546F4" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "StyleProductNameBefore0ptAfter8pt" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "PRODUCT: ASIA EX JAPAN EQUITIES";

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "Grilledutableau" };
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 141, RightFromText = 141, VerticalAnchor = VerticalAnchorValues.Text, TablePositionY = 74 };
            TableWidth tableWidth1 = new TableWidth() { Width = "10485", Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 45, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 45, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook1 = new TableLook() { Val = "01E0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tablePositionProperties1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLayout1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2124" };
            GridColumn gridColumn2 = new GridColumn() { Width = "2892" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2576" };
            GridColumn gridColumn4 = new GridColumn() { Width = "2893" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "004427CC", RsidTableRowProperties = "00443CD0" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)182U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2124", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "000C4F36", RsidRunAdditionDefault = "004427CC" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "TableHeading" };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(spacingBetweenLines1);

            Run run2 = new Run();
            Text text2 = new Text();
            text2.Text = "ASSET CLASS";

            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2892", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "000C4F36", RsidRunAdditionDefault = "004427CC" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "TableHeading" };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(spacingBetweenLines2);

            Run run3 = new Run();
            Text text3 = new Text();
            text3.Text = "GEOGRAPHIC EMPHASIS";

            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2576", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "000C4F36", RsidRunAdditionDefault = "004427CC" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "TableHeading" };
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties4.Append(paragraphStyleId4);
            paragraphProperties4.Append(spacingBetweenLines3);

            Run run4 = new Run();
            Text text4 = new Text();
            text4.Text = "STYLE";

            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "2893", Type = TableWidthUnitValues.Dxa };

            tableCellProperties4.Append(tableCellWidth4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "000C4F36", RsidRunAdditionDefault = "004427CC" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "TableHeading" };
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties5.Append(paragraphStyleId5);
            paragraphProperties5.Append(spacingBetweenLines4);

            Run run5 = new Run();
            Text text5 = new Text();
            text5.Text = "SUBSTYLE";

            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "004427CC", RsidTableRowProperties = "00443CD0" };

            CustomXmlCell customXmlCell1 = new CustomXmlCell() { Uri = "http://hubblereports.com/namespace", Element = "AssetClass" };

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "2124", Type = TableWidthUnitValues.Dxa };

            tableCellProperties5.Append(tableCellWidth5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "005546F4", RsidRunAdditionDefault = "005546F4" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties6.Append(paragraphStyleId6);
            paragraphProperties6.Append(spacingBetweenLines5);
            paragraphProperties6.Append(paragraphMarkRunProperties1);

            Run run6 = new Run();

            RunProperties runProperties1 = new RunProperties();
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "18" };

            runProperties1.Append(fontSizeComplexScript2);
            Text text6 = new Text();
            text6.Text = "Equity";

            run6.Append(runProperties1);
            run6.Append(text6);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run6);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph6);

            customXmlCell1.Append(tableCell5);

            CustomXmlCell customXmlCell2 = new CustomXmlCell() { Uri = "http://hubblereports.com/namespace", Element = "GeoEmphasis" };

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "2892", Type = TableWidthUnitValues.Dxa };

            tableCellProperties6.Append(tableCellWidth6);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "005546F4", RsidRunAdditionDefault = "005546F4" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

            paragraphProperties7.Append(paragraphStyleId7);
            paragraphProperties7.Append(spacingBetweenLines6);
            paragraphProperties7.Append(paragraphMarkRunProperties2);

            Run run7 = new Run();

            RunProperties runProperties2 = new RunProperties();
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "18" };

            runProperties2.Append(fontSizeComplexScript4);
            Text text7 = new Text();
            text7.Text = "Asia ex Japan";

            run7.Append(runProperties2);
            run7.Append(text7);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run7);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph7);

            customXmlCell2.Append(tableCell6);

            CustomXmlCell customXmlCell3 = new CustomXmlCell() { Uri = "http://hubblereports.com/namespace", Element = "Style" };

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2576", Type = TableWidthUnitValues.Dxa };

            tableCellProperties7.Append(tableCellWidth7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "005546F4", RsidRunAdditionDefault = "005546F4" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

            paragraphProperties8.Append(paragraphStyleId8);
            paragraphProperties8.Append(spacingBetweenLines7);
            paragraphProperties8.Append(paragraphMarkRunProperties3);

            Run run8 = new Run();

            RunProperties runProperties3 = new RunProperties();
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "18" };

            runProperties3.Append(fontSizeComplexScript6);
            Text text8 = new Text();
            text8.Text = "-";

            run8.Append(runProperties3);
            run8.Append(text8);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run8);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph8);

            customXmlCell3.Append(tableCell7);

            CustomXmlCell customXmlCell4 = new CustomXmlCell() { Uri = "http://hubblereports.com/namespace", Element = "Substyle" };

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "2893", Type = TableWidthUnitValues.Dxa };

            tableCellProperties8.Append(tableCellWidth8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "004427CC", RsidParagraphProperties = "005546F4", RsidRunAdditionDefault = "005546F4" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties4.Append(fontSizeComplexScript7);

            paragraphProperties9.Append(paragraphStyleId9);
            paragraphProperties9.Append(spacingBetweenLines8);
            paragraphProperties9.Append(paragraphMarkRunProperties4);

            Run run9 = new Run();
            Text text9 = new Text();
            text9.Text = "-";

            run9.Append(text9);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run9);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph9);

            customXmlCell4.Append(tableCell8);

            tableRow2.Append(customXmlCell1);
            tableRow2.Append(customXmlCell2);
            tableRow2.Append(customXmlCell3);
            tableRow2.Append(customXmlCell4);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);

            customXmlBlock4.Append(paragraph1);
            customXmlBlock4.Append(table1);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00E12034", RsidParagraphAddition = "009F7E7F", RsidParagraphProperties = "00397A9E", RsidRunAdditionDefault = "009F7E7F" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "StyleProductsReviewedHeading6ptBefore15ptAfter0pt" };

            paragraphProperties10.Append(paragraphStyleId10);

            paragraph10.Append(paragraphProperties10);

            CustomXmlBlock customXmlBlock5 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "opiniondata" };

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "009F7E7F", RsidParagraphAddition = "00EE7B69", RsidParagraphProperties = "009F7E7F", RsidRunAdditionDefault = "00FC0F0D" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId11 = new ParagraphStyleId() { Val = "RankHeading" };
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties11.Append(paragraphStyleId11);
            paragraphProperties11.Append(spacingBetweenLines9);

            Run run10 = new Run() { RsidRunProperties = "009F7E7F" };
            Text text10 = new Text();
            text10.Text = "OVERALL EVaLUATION";

            run10.Append(text10);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run10);

            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableStyle tableStyle2 = new TableStyle() { Val = "Grilledutableau" };
            TableWidth tableWidth2 = new TableWidth() { Width = "10485", Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder2);
            tableBorders2.Append(leftBorder2);
            tableBorders2.Append(bottomBorder2);
            tableBorders2.Append(rightBorder2);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);
            TableLayout tableLayout2 = new TableLayout() { Type = TableLayoutValues.Fixed };

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 45, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 45, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);
            TableLook tableLook2 = new TableLook() { Val = "01E0" };

            tableProperties2.Append(tableStyle2);
            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableBorders2);
            tableProperties2.Append(tableLayout2);
            tableProperties2.Append(tableCellMarginDefault2);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn5 = new GridColumn() { Width = "3175" };
            GridColumn gridColumn6 = new GridColumn() { Width = "4070" };
            GridColumn gridColumn7 = new GridColumn() { Width = "3240" };

            tableGrid2.Append(gridColumn5);
            tableGrid2.Append(gridColumn6);
            tableGrid2.Append(gridColumn7);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "002E7D22", RsidTableRowProperties = "00837232" };

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "3175", Type = TableWidthUnitValues.Dxa };

            tableCellProperties9.Append(tableCellWidth9);

            CustomXmlBlock customXmlBlock6 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "RankValueImage" };

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00BA7E3F", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId12 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties5.Append(fontSizeComplexScript8);

            paragraphProperties12.Append(paragraphStyleId12);
            paragraphProperties12.Append(spacingBetweenLines10);
            paragraphProperties12.Append(paragraphMarkRunProperties5);

            Run run11 = new Run();

            RunProperties runProperties4 = new RunProperties();
            NoProof noProof1 = new NoProof();
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "18" };
            Languages languages1 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties4.Append(noProof1);
            runProperties4.Append(fontSizeComplexScript9);
            runProperties4.Append(languages1);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent1 = new Wp.Extent() { Cx = 1485900L, Cy = 428625L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)3U, Name = "Image 3", Description = "rank_1" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 3", Description = "rank_1" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();
            A.Blip blip1 = new A.Blip() { Embed = "rId7", CompressionState = A.BlipCompressionValues.Print };
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle();

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 1485900L, Cy = 428625L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline() { Width = 9525 };
            A.NoFill noFill2 = new A.NoFill();
            A.Miter miter1 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

            outline1.Append(noFill2);
            outline1.Append(miter1);
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(outline1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run11.Append(runProperties4);
            run11.Append(drawing1);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run11);

            customXmlBlock6.Append(paragraph12);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(customXmlBlock6);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "4070", Type = TableWidthUnitValues.Dxa };

            tableCellProperties10.Append(tableCellWidth10);

            CustomXmlBlock customXmlBlock7 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "StatementForOverall" };

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00E340CC", RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "003F1967", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId13 = new ParagraphStyleId() { Val = "RankStatement" };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { EastAsia = "Arial Unicode MS" };
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties6.Append(runFonts1);
            paragraphMarkRunProperties6.Append(fontSize1);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript10);

            paragraphProperties13.Append(paragraphStyleId13);
            paragraphProperties13.Append(paragraphMarkRunProperties6);

            Run run12 = new Run() { RsidRunProperties = "003F1967" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { EastAsia = "Arial Unicode MS" };

            runProperties5.Append(runFonts2);
            Text text11 = new Text();
            text11.Text = "Our preliminary view of this product is positive, and we therefore intend to gather and review additional information in the near future, prior to assigning a formal rank.";

            run12.Append(runProperties5);
            run12.Append(text11);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run12);

            customXmlBlock7.Append(paragraph13);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(customXmlBlock7);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "3240", Type = TableWidthUnitValues.Dxa };
            NoWrap noWrap1 = new NoWrap();

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(noWrap1);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00256073", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId14 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            FontSize fontSize2 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties7.Append(fontSize2);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript11);

            paragraphProperties14.Append(paragraphStyleId14);
            paragraphProperties14.Append(spacingBetweenLines11);
            paragraphProperties14.Append(paragraphMarkRunProperties7);

            Run run13 = new Run() { RsidRunProperties = "00DC3ED5" };

            RunProperties runProperties6 = new RunProperties();
            Bold bold1 = new Bold();

            runProperties6.Append(bold1);
            Text text12 = new Text();
            text12.Text = "Updated By:";

            run13.Append(runProperties6);
            run13.Append(text12);

            Run run14 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties7 = new RunProperties();
            Bold bold2 = new Bold();
            NoProof noProof2 = new NoProof();
            Languages languages2 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties7.Append(bold2);
            runProperties7.Append(noProof2);
            runProperties7.Append(languages2);

            Drawing drawing2 = new Drawing();

            Wp.Inline inline2 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent2 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent2 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties2 = new Wp.DocProperties() { Id = (UInt32Value)4U, Name = "Image 4", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties2 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks2 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties2.Append(graphicFrameLocks2);

            A.Graphic graphic2 = new A.Graphic();

            A.GraphicData graphicData2 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture2 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties2 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 4", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties2 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks2 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties2.Append(pictureLocks2);

            nonVisualPictureProperties2.Append(nonVisualDrawingProperties2);
            nonVisualPictureProperties2.Append(nonVisualPictureDrawingProperties2);

            Pic.BlipFill blipFill2 = new Pic.BlipFill();
            A.Blip blip2 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle2 = new A.SourceRectangle();

            A.Stretch stretch2 = new A.Stretch();
            A.FillRectangle fillRectangle2 = new A.FillRectangle();

            stretch2.Append(fillRectangle2);

            blipFill2.Append(blip2);
            blipFill2.Append(sourceRectangle2);
            blipFill2.Append(stretch2);

            Pic.ShapeProperties shapeProperties2 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D2.Append(offset2);
            transform2D2.Append(extents2);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);
            A.NoFill noFill3 = new A.NoFill();

            A.Outline outline2 = new A.Outline() { Width = 9525 };
            A.NoFill noFill4 = new A.NoFill();
            A.Miter miter2 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd2 = new A.HeadEnd();
            A.TailEnd tailEnd2 = new A.TailEnd();

            outline2.Append(noFill4);
            outline2.Append(miter2);
            outline2.Append(headEnd2);
            outline2.Append(tailEnd2);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(noFill3);
            shapeProperties2.Append(outline2);

            picture2.Append(nonVisualPictureProperties2);
            picture2.Append(blipFill2);
            picture2.Append(shapeProperties2);

            graphicData2.Append(picture2);

            graphic2.Append(graphicData2);

            inline2.Append(extent2);
            inline2.Append(effectExtent2);
            inline2.Append(docProperties2);
            inline2.Append(nonVisualGraphicFrameDrawingProperties2);
            inline2.Append(graphic2);

            drawing2.Append(inline2);

            run14.Append(runProperties7);
            run14.Append(drawing2);

            Run run15 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties8 = new RunProperties();
            Bold bold3 = new Bold();
            NoProof noProof3 = new NoProof();
            Languages languages3 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties8.Append(bold3);
            runProperties8.Append(noProof3);
            runProperties8.Append(languages3);

            Drawing drawing3 = new Drawing();

            Wp.Inline inline3 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent3 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent3 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties3 = new Wp.DocProperties() { Id = (UInt32Value)5U, Name = "Image 5", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties3 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks3 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties3.Append(graphicFrameLocks3);

            A.Graphic graphic3 = new A.Graphic();

            A.GraphicData graphicData3 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture3 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties3 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 5", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties3 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks3 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties3.Append(pictureLocks3);

            nonVisualPictureProperties3.Append(nonVisualDrawingProperties3);
            nonVisualPictureProperties3.Append(nonVisualPictureDrawingProperties3);

            Pic.BlipFill blipFill3 = new Pic.BlipFill();
            A.Blip blip3 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle3 = new A.SourceRectangle();

            A.Stretch stretch3 = new A.Stretch();
            A.FillRectangle fillRectangle3 = new A.FillRectangle();

            stretch3.Append(fillRectangle3);

            blipFill3.Append(blip3);
            blipFill3.Append(sourceRectangle3);
            blipFill3.Append(stretch3);

            Pic.ShapeProperties shapeProperties3 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset3 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents3 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D3.Append(offset3);
            transform2D3.Append(extents3);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);
            A.NoFill noFill5 = new A.NoFill();

            A.Outline outline3 = new A.Outline() { Width = 9525 };
            A.NoFill noFill6 = new A.NoFill();
            A.Miter miter3 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd3 = new A.HeadEnd();
            A.TailEnd tailEnd3 = new A.TailEnd();

            outline3.Append(noFill6);
            outline3.Append(miter3);
            outline3.Append(headEnd3);
            outline3.Append(tailEnd3);

            shapeProperties3.Append(transform2D3);
            shapeProperties3.Append(presetGeometry3);
            shapeProperties3.Append(noFill5);
            shapeProperties3.Append(outline3);

            picture3.Append(nonVisualPictureProperties3);
            picture3.Append(blipFill3);
            picture3.Append(shapeProperties3);

            graphicData3.Append(picture3);

            graphic3.Append(graphicData3);

            inline3.Append(extent3);
            inline3.Append(effectExtent3);
            inline3.Append(docProperties3);
            inline3.Append(nonVisualGraphicFrameDrawingProperties3);
            inline3.Append(graphic3);

            drawing3.Append(inline3);

            run15.Append(runProperties8);
            run15.Append(drawing3);

            CustomXmlRun customXmlRun1 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "UpdatedBy" };

            Run run16 = new Run() { RsidRunAddition = "005546F4" };
            Text text13 = new Text();
            text13.Text = "Julien Blin";

            run16.Append(text13);

            customXmlRun1.Append(run16);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run13);
            paragraph14.Append(run14);
            paragraph14.Append(run15);
            paragraph14.Append(customXmlRun1);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00256073", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId15 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            FontSize fontSize3 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties8.Append(fontSize3);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript12);

            paragraphProperties15.Append(paragraphStyleId15);
            paragraphProperties15.Append(spacingBetweenLines12);
            paragraphProperties15.Append(paragraphMarkRunProperties8);

            Run run17 = new Run() { RsidRunProperties = "00DC3ED5" };

            RunProperties runProperties9 = new RunProperties();
            Bold bold4 = new Bold();

            runProperties9.Append(bold4);
            Text text14 = new Text();
            text14.Text = "Target Excess Return:";

            run17.Append(runProperties9);
            run17.Append(text14);

            Run run18 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties10 = new RunProperties();
            Bold bold5 = new Bold();
            NoProof noProof4 = new NoProof();
            Languages languages4 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties10.Append(bold5);
            runProperties10.Append(noProof4);
            runProperties10.Append(languages4);

            Drawing drawing4 = new Drawing();

            Wp.Inline inline4 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent4 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent4 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties4 = new Wp.DocProperties() { Id = (UInt32Value)6U, Name = "Image 6", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties4 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks4 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties4.Append(graphicFrameLocks4);

            A.Graphic graphic4 = new A.Graphic();

            A.GraphicData graphicData4 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture4 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties4 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 6", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties4 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks4 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties4.Append(pictureLocks4);

            nonVisualPictureProperties4.Append(nonVisualDrawingProperties4);
            nonVisualPictureProperties4.Append(nonVisualPictureDrawingProperties4);

            Pic.BlipFill blipFill4 = new Pic.BlipFill();
            A.Blip blip4 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle4 = new A.SourceRectangle();

            A.Stretch stretch4 = new A.Stretch();
            A.FillRectangle fillRectangle4 = new A.FillRectangle();

            stretch4.Append(fillRectangle4);

            blipFill4.Append(blip4);
            blipFill4.Append(sourceRectangle4);
            blipFill4.Append(stretch4);

            Pic.ShapeProperties shapeProperties4 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents4 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D4.Append(offset4);
            transform2D4.Append(extents4);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);
            A.NoFill noFill7 = new A.NoFill();

            A.Outline outline4 = new A.Outline() { Width = 9525 };
            A.NoFill noFill8 = new A.NoFill();
            A.Miter miter4 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd4 = new A.HeadEnd();
            A.TailEnd tailEnd4 = new A.TailEnd();

            outline4.Append(noFill8);
            outline4.Append(miter4);
            outline4.Append(headEnd4);
            outline4.Append(tailEnd4);

            shapeProperties4.Append(transform2D4);
            shapeProperties4.Append(presetGeometry4);
            shapeProperties4.Append(noFill7);
            shapeProperties4.Append(outline4);

            picture4.Append(nonVisualPictureProperties4);
            picture4.Append(blipFill4);
            picture4.Append(shapeProperties4);

            graphicData4.Append(picture4);

            graphic4.Append(graphicData4);

            inline4.Append(extent4);
            inline4.Append(effectExtent4);
            inline4.Append(docProperties4);
            inline4.Append(nonVisualGraphicFrameDrawingProperties4);
            inline4.Append(graphic4);

            drawing4.Append(inline4);

            run18.Append(runProperties10);
            run18.Append(drawing4);

            Run run19 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties11 = new RunProperties();
            Bold bold6 = new Bold();
            NoProof noProof5 = new NoProof();
            FontSize fontSize4 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "20" };
            Languages languages5 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties11.Append(bold6);
            runProperties11.Append(noProof5);
            runProperties11.Append(fontSize4);
            runProperties11.Append(fontSizeComplexScript13);
            runProperties11.Append(languages5);

            Drawing drawing5 = new Drawing();

            Wp.Inline inline5 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent5 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent5 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties5 = new Wp.DocProperties() { Id = (UInt32Value)7U, Name = "Image 7", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties5 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks5 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties5.Append(graphicFrameLocks5);

            A.Graphic graphic5 = new A.Graphic();

            A.GraphicData graphicData5 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture5 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties5 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties5 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 7", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties5 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks5 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties5.Append(pictureLocks5);

            nonVisualPictureProperties5.Append(nonVisualDrawingProperties5);
            nonVisualPictureProperties5.Append(nonVisualPictureDrawingProperties5);

            Pic.BlipFill blipFill5 = new Pic.BlipFill();
            A.Blip blip5 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle5 = new A.SourceRectangle();

            A.Stretch stretch5 = new A.Stretch();
            A.FillRectangle fillRectangle5 = new A.FillRectangle();

            stretch5.Append(fillRectangle5);

            blipFill5.Append(blip5);
            blipFill5.Append(sourceRectangle5);
            blipFill5.Append(stretch5);

            Pic.ShapeProperties shapeProperties5 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset5 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents5 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D5.Append(offset5);
            transform2D5.Append(extents5);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);
            A.NoFill noFill9 = new A.NoFill();

            A.Outline outline5 = new A.Outline() { Width = 9525 };
            A.NoFill noFill10 = new A.NoFill();
            A.Miter miter5 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd5 = new A.HeadEnd();
            A.TailEnd tailEnd5 = new A.TailEnd();

            outline5.Append(noFill10);
            outline5.Append(miter5);
            outline5.Append(headEnd5);
            outline5.Append(tailEnd5);

            shapeProperties5.Append(transform2D5);
            shapeProperties5.Append(presetGeometry5);
            shapeProperties5.Append(noFill9);
            shapeProperties5.Append(outline5);

            picture5.Append(nonVisualPictureProperties5);
            picture5.Append(blipFill5);
            picture5.Append(shapeProperties5);

            graphicData5.Append(picture5);

            graphic5.Append(graphicData5);

            inline5.Append(extent5);
            inline5.Append(effectExtent5);
            inline5.Append(docProperties5);
            inline5.Append(nonVisualGraphicFrameDrawingProperties5);
            inline5.Append(graphic5);

            drawing5.Append(inline5);

            run19.Append(runProperties11);
            run19.Append(drawing5);

            Run run20 = new Run() { RsidRunAddition = "005546F4" };
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = "0 to 0 bp ";

            run20.Append(text15);
            CustomXmlRun customXmlRun2 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "TargetExcessReturnMaxCurrent" };

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run17);
            paragraph15.Append(run18);
            paragraph15.Append(run19);
            paragraph15.Append(run20);
            paragraph15.Append(customXmlRun2);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00256073", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            FontSize fontSize5 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties9.Append(fontSize5);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript14);

            paragraphProperties16.Append(paragraphStyleId16);
            paragraphProperties16.Append(spacingBetweenLines13);
            paragraphProperties16.Append(paragraphMarkRunProperties9);

            Run run21 = new Run() { RsidRunProperties = "00DC3ED5" };

            RunProperties runProperties12 = new RunProperties();
            Bold bold7 = new Bold();

            runProperties12.Append(bold7);
            Text text16 = new Text();
            text16.Text = "Target Tracking Error:";

            run21.Append(runProperties12);
            run21.Append(text16);

            Run run22 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties13 = new RunProperties();
            Bold bold8 = new Bold();
            NoProof noProof6 = new NoProof();
            FontSize fontSize6 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "20" };
            Languages languages6 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties13.Append(bold8);
            runProperties13.Append(noProof6);
            runProperties13.Append(fontSize6);
            runProperties13.Append(fontSizeComplexScript15);
            runProperties13.Append(languages6);

            Drawing drawing6 = new Drawing();

            Wp.Inline inline6 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent6 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent6 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties6 = new Wp.DocProperties() { Id = (UInt32Value)8U, Name = "Image 8", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties6 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks6 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties6.Append(graphicFrameLocks6);

            A.Graphic graphic6 = new A.Graphic();

            A.GraphicData graphicData6 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture6 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties6 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties6 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 8", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties6 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks6 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties6.Append(pictureLocks6);

            nonVisualPictureProperties6.Append(nonVisualDrawingProperties6);
            nonVisualPictureProperties6.Append(nonVisualPictureDrawingProperties6);

            Pic.BlipFill blipFill6 = new Pic.BlipFill();
            A.Blip blip6 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle6 = new A.SourceRectangle();

            A.Stretch stretch6 = new A.Stretch();
            A.FillRectangle fillRectangle6 = new A.FillRectangle();

            stretch6.Append(fillRectangle6);

            blipFill6.Append(blip6);
            blipFill6.Append(sourceRectangle6);
            blipFill6.Append(stretch6);

            Pic.ShapeProperties shapeProperties6 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset6 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents6 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D6.Append(offset6);
            transform2D6.Append(extents6);

            A.PresetGeometry presetGeometry6 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetGeometry6.Append(adjustValueList6);
            A.NoFill noFill11 = new A.NoFill();

            A.Outline outline6 = new A.Outline() { Width = 9525 };
            A.NoFill noFill12 = new A.NoFill();
            A.Miter miter6 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd6 = new A.HeadEnd();
            A.TailEnd tailEnd6 = new A.TailEnd();

            outline6.Append(noFill12);
            outline6.Append(miter6);
            outline6.Append(headEnd6);
            outline6.Append(tailEnd6);

            shapeProperties6.Append(transform2D6);
            shapeProperties6.Append(presetGeometry6);
            shapeProperties6.Append(noFill11);
            shapeProperties6.Append(outline6);

            picture6.Append(nonVisualPictureProperties6);
            picture6.Append(blipFill6);
            picture6.Append(shapeProperties6);

            graphicData6.Append(picture6);

            graphic6.Append(graphicData6);

            inline6.Append(extent6);
            inline6.Append(effectExtent6);
            inline6.Append(docProperties6);
            inline6.Append(nonVisualGraphicFrameDrawingProperties6);
            inline6.Append(graphic6);

            drawing6.Append(inline6);

            run22.Append(runProperties13);
            run22.Append(drawing6);

            Run run23 = new Run() { RsidRunAddition = "005546F4" };
            Text text17 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text17.Text = " 0 to 0 bp";

            run23.Append(text17);
            CustomXmlRun customXmlRun3 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "TargetTrackingErrorMaxCurrent" };

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run21);
            paragraph16.Append(run22);
            paragraph16.Append(run23);
            paragraph16.Append(customXmlRun3);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00256073", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId17 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            FontSize fontSize7 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties10.Append(fontSize7);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript16);

            paragraphProperties17.Append(paragraphStyleId17);
            paragraphProperties17.Append(spacingBetweenLines14);
            paragraphProperties17.Append(paragraphMarkRunProperties10);

            Run run24 = new Run() { RsidRunProperties = "00DC3ED5" };

            RunProperties runProperties14 = new RunProperties();
            Bold bold9 = new Bold();

            runProperties14.Append(bold9);
            Text text18 = new Text();
            text18.Text = "Time Period:";

            run24.Append(runProperties14);
            run24.Append(text18);

            Run run25 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties15 = new RunProperties();
            Bold bold10 = new Bold();
            NoProof noProof7 = new NoProof();
            FontSize fontSize8 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "20" };
            Languages languages7 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties15.Append(bold10);
            runProperties15.Append(noProof7);
            runProperties15.Append(fontSize8);
            runProperties15.Append(fontSizeComplexScript17);
            runProperties15.Append(languages7);

            Drawing drawing7 = new Drawing();

            Wp.Inline inline7 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent7 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent7 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties7 = new Wp.DocProperties() { Id = (UInt32Value)9U, Name = "Image 9", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties7 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks7 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties7.Append(graphicFrameLocks7);

            A.Graphic graphic7 = new A.Graphic();

            A.GraphicData graphicData7 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture7 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties7 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties7 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 9", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties7 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks7 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties7.Append(pictureLocks7);

            nonVisualPictureProperties7.Append(nonVisualDrawingProperties7);
            nonVisualPictureProperties7.Append(nonVisualPictureDrawingProperties7);

            Pic.BlipFill blipFill7 = new Pic.BlipFill();
            A.Blip blip7 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle7 = new A.SourceRectangle();

            A.Stretch stretch7 = new A.Stretch();
            A.FillRectangle fillRectangle7 = new A.FillRectangle();

            stretch7.Append(fillRectangle7);

            blipFill7.Append(blip7);
            blipFill7.Append(sourceRectangle7);
            blipFill7.Append(stretch7);

            Pic.ShapeProperties shapeProperties7 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset7 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents7 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D7.Append(offset7);
            transform2D7.Append(extents7);

            A.PresetGeometry presetGeometry7 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList7 = new A.AdjustValueList();

            presetGeometry7.Append(adjustValueList7);
            A.NoFill noFill13 = new A.NoFill();

            A.Outline outline7 = new A.Outline() { Width = 9525 };
            A.NoFill noFill14 = new A.NoFill();
            A.Miter miter7 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd7 = new A.HeadEnd();
            A.TailEnd tailEnd7 = new A.TailEnd();

            outline7.Append(noFill14);
            outline7.Append(miter7);
            outline7.Append(headEnd7);
            outline7.Append(tailEnd7);

            shapeProperties7.Append(transform2D7);
            shapeProperties7.Append(presetGeometry7);
            shapeProperties7.Append(noFill13);
            shapeProperties7.Append(outline7);

            picture7.Append(nonVisualPictureProperties7);
            picture7.Append(blipFill7);
            picture7.Append(shapeProperties7);

            graphicData7.Append(picture7);

            graphic7.Append(graphicData7);

            inline7.Append(extent7);
            inline7.Append(effectExtent7);
            inline7.Append(docProperties7);
            inline7.Append(nonVisualGraphicFrameDrawingProperties7);
            inline7.Append(graphic7);

            drawing7.Append(inline7);

            run25.Append(runProperties15);
            run25.Append(drawing7);

            CustomXmlRun customXmlRun4 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "TimePeriodCurrent" };

            Run run26 = new Run() { RsidRunAddition = "005546F4" };
            Text text19 = new Text();
            text19.Text = "-";

            run26.Append(text19);

            customXmlRun4.Append(run26);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run24);
            paragraph17.Append(run25);
            paragraph17.Append(customXmlRun4);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "005546F4", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId18 = new ParagraphStyleId() { Val = "TableText" };
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties11.Append(fontSizeComplexScript18);

            paragraphProperties18.Append(paragraphStyleId18);
            paragraphProperties18.Append(spacingBetweenLines15);
            paragraphProperties18.Append(paragraphMarkRunProperties11);

            Run run27 = new Run() { RsidRunProperties = "00DC3ED5" };

            RunProperties runProperties16 = new RunProperties();
            Bold bold11 = new Bold();

            runProperties16.Append(bold11);
            Text text20 = new Text();
            text20.Text = "Russell-Assigned Benchmark:";

            run27.Append(runProperties16);
            run27.Append(text20);

            Run run28 = new Run() { RsidRunAddition = "00740A1C" };

            RunProperties runProperties17 = new RunProperties();
            NoProof noProof8 = new NoProof();
            Languages languages8 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties17.Append(noProof8);
            runProperties17.Append(languages8);

            Drawing drawing8 = new Drawing();

            Wp.Inline inline8 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent8 = new Wp.Extent() { Cx = 9525L, Cy = 9525L };
            Wp.EffectExtent effectExtent8 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties8 = new Wp.DocProperties() { Id = (UInt32Value)10U, Name = "Image 10", Description = "spacer" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties8 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks8 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties8.Append(graphicFrameLocks8);

            A.Graphic graphic8 = new A.Graphic();

            A.GraphicData graphicData8 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture8 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties8 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties8 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 10", Description = "spacer" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties8 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks8 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties8.Append(pictureLocks8);

            nonVisualPictureProperties8.Append(nonVisualDrawingProperties8);
            nonVisualPictureProperties8.Append(nonVisualPictureDrawingProperties8);

            Pic.BlipFill blipFill8 = new Pic.BlipFill();
            A.Blip blip8 = new A.Blip() { Embed = "rId8" };
            A.SourceRectangle sourceRectangle8 = new A.SourceRectangle();

            A.Stretch stretch8 = new A.Stretch();
            A.FillRectangle fillRectangle8 = new A.FillRectangle();

            stretch8.Append(fillRectangle8);

            blipFill8.Append(blip8);
            blipFill8.Append(sourceRectangle8);
            blipFill8.Append(stretch8);

            Pic.ShapeProperties shapeProperties8 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset8 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents8 = new A.Extents() { Cx = 9525L, Cy = 9525L };

            transform2D8.Append(offset8);
            transform2D8.Append(extents8);

            A.PresetGeometry presetGeometry8 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList8 = new A.AdjustValueList();

            presetGeometry8.Append(adjustValueList8);
            A.NoFill noFill15 = new A.NoFill();

            A.Outline outline8 = new A.Outline() { Width = 9525 };
            A.NoFill noFill16 = new A.NoFill();
            A.Miter miter8 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd8 = new A.HeadEnd();
            A.TailEnd tailEnd8 = new A.TailEnd();

            outline8.Append(noFill16);
            outline8.Append(miter8);
            outline8.Append(headEnd8);
            outline8.Append(tailEnd8);

            shapeProperties8.Append(transform2D8);
            shapeProperties8.Append(presetGeometry8);
            shapeProperties8.Append(noFill15);
            shapeProperties8.Append(outline8);

            picture8.Append(nonVisualPictureProperties8);
            picture8.Append(blipFill8);
            picture8.Append(shapeProperties8);

            graphicData8.Append(picture8);

            graphic8.Append(graphicData8);

            inline8.Append(extent8);
            inline8.Append(effectExtent8);
            inline8.Append(docProperties8);
            inline8.Append(nonVisualGraphicFrameDrawingProperties8);
            inline8.Append(graphic8);

            drawing8.Append(inline8);

            run28.Append(runProperties17);
            run28.Append(drawing8);

            CustomXmlRun customXmlRun5 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "RussellBenchmark" };

            Run run29 = new Run() { RsidRunAddition = "005546F4" };
            Text text21 = new Text();
            text21.Text = "-";

            run29.Append(text21);

            customXmlRun5.Append(run29);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run27);
            paragraph18.Append(run28);
            paragraph18.Append(customXmlRun5);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph14);
            tableCell11.Append(paragraph15);
            tableCell11.Append(paragraph16);
            tableCell11.Append(paragraph17);
            tableCell11.Append(paragraph18);

            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow3);

            customXmlBlock5.Append(paragraph11);
            customXmlBlock5.Append(table2);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00A6171D", RsidParagraphAddition = "00F342A0", RsidParagraphProperties = "00397A9E", RsidRunAdditionDefault = "00F342A0" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId19 = new ParagraphStyleId() { Val = "StyleProductsReviewedHeading4ptBefore15ptAfter0pt" };

            paragraphProperties19.Append(paragraphStyleId19);

            paragraph19.Append(paragraphProperties19);

            Paragraph paragraphDiscussionTitle = CreateTopicTitle("DISCUSSION", null);
            var paragraphsDiscussionContent = CreateTopicText(Discussion);

            CustomXmlBlock customXmlBlock8 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "category" };

            Paragraph paragraphInvestmentStaffTitle = CreateTopicTitle("INVESTMENT STAFF", InvestmentStaff.Rank.ToString());
            var paragraphsInvestmentStaffContent = CreateTopicText(InvestmentStaff.Content);

            Paragraph paragraphOrganizationalStabilityTitle = CreateTopicTitle("ORGANIZATIONAL STABILITY", OrganizationalStability.Rank.ToString());
            var paragraphsOrganizationalStabilityContent = CreateTopicText(OrganizationalStability.Content);

            Paragraph paragraphAssetAllocationTitle = CreateTopicTitle("ASSET ALLOCATION", AssetAllocation.Rank.ToString());
            var paragraphsAssetAllocationContent = CreateTopicText(AssetAllocation.Content);

            Paragraph paragraphResearchTitle = CreateTopicTitle("RESEARCH", Research.Rank.ToString());
            var paragraphsResearchContent = CreateTopicText(Research.Content);

            Paragraph paragraphCountrySelectionTitle = CreateTopicTitle("COUNTRY SELECTION", CountrySelection.Rank.ToString());
            var paragraphsCountrySelectionContent = CreateTopicText(CountrySelection.Content);

            Paragraph paragraphPortfolioConstructionTitle = CreateTopicTitle("PORTFOLIO CONSTRUCTION", PortfolioConstruction.Rank.ToString());
            var paragraphsPortfolioConstructionContent = CreateTopicText(PortfolioConstruction.Content);

            Paragraph paragraphCurrencyManagementTitle = CreateTopicTitle("CURRENCY MANAGEMENT", CurrencyManagement.Rank.ToString());
            var paragraphsCurrencyManagementContent = CreateTopicText(CurrencyManagement.Content);

            Paragraph paragraphImplementationTitle = CreateTopicTitle("IMPLEMENTATION", Implementation.Rank.ToString());
            var paragraphsImplementationContent = CreateTopicText(Implementation.Content);

            Paragraph paragraphSecuritySelectionTitle = CreateTopicTitle("SECURITY SELECTION", SecuritySelection.Rank.ToString());
            var paragraphsSecuritySelectionContent = CreateTopicText(SecuritySelection.Content);

            Paragraph paragraphSellDisciplineTitle = CreateTopicTitle("SELL DISCIPLINE", SellDiscipline.Rank.ToString());
            var paragraphsSellDisciplineContent = CreateTopicText(SellDiscipline.Content);


            Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "00EE7B69", RsidParagraphProperties = "00C32704", RsidRunAdditionDefault = "00957E57" };

            customXmlBlock3.Append(customXmlBlock4);
            customXmlBlock3.Append(paragraph10);
            customXmlBlock3.Append(customXmlBlock5);
            customXmlBlock3.Append(paragraph19);
            customXmlBlock3.Append(paragraphDiscussionTitle);
            customXmlBlock3.Append(paragraphsDiscussionContent.ToArray());
            customXmlBlock3.Append(paragraphInvestmentStaffTitle);
            customXmlBlock3.Append(paragraphsInvestmentStaffContent.ToArray());
            customXmlBlock3.Append(paragraphOrganizationalStabilityTitle);
            customXmlBlock3.Append(paragraphsOrganizationalStabilityContent.ToArray());
            customXmlBlock3.Append(paragraphAssetAllocationTitle);
            customXmlBlock3.Append(paragraphsAssetAllocationContent.ToArray());
            customXmlBlock3.Append(paragraphResearchTitle);
            customXmlBlock3.Append(paragraphsResearchContent.ToArray());
            customXmlBlock3.Append(paragraphCountrySelectionTitle);
            customXmlBlock3.Append(paragraphsCountrySelectionContent.ToArray());
            customXmlBlock3.Append(paragraphPortfolioConstructionTitle);
            customXmlBlock3.Append(paragraphsPortfolioConstructionContent.ToArray());
            customXmlBlock3.Append(paragraphCurrencyManagementTitle);
            customXmlBlock3.Append(paragraphsCurrencyManagementContent.ToArray());
            customXmlBlock3.Append(paragraphImplementationTitle);
            customXmlBlock3.Append(paragraphsImplementationContent.ToArray());
            customXmlBlock3.Append(paragraphSecuritySelectionTitle);
            customXmlBlock3.Append(paragraphsSecuritySelectionContent.ToArray());
            customXmlBlock3.Append(paragraphSellDisciplineTitle);
            customXmlBlock3.Append(paragraphsSellDisciplineContent.ToArray());
            customXmlBlock3.Append(paragraph23);

            customXmlBlock2.Append(customXmlBlock3);

            CustomXmlBlock customXmlBlock10 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "footer" };

            CustomXmlBlock customXmlBlock11 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "Headline" };

            Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "000A24FF", RsidParagraphProperties = "00E560D4", RsidRunAdditionDefault = "00707FFA" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId23 = new ParagraphStyleId() { Val = "DislaimerHeading" };
            WidowControl widowControl1 = new WidowControl() { Val = false };

            paragraphProperties23.Append(paragraphStyleId23);
            paragraphProperties23.Append(widowControl1);

            Run run37 = new Run() { RsidRunProperties = "00C62467" };
            Text text27 = new Text();
            text27.Text = "Healine";

            run37.Append(text27);

            paragraph24.Append(paragraphProperties23);
            paragraph24.Append(run37);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "00E70E3A", RsidParagraphProperties = "00E560D4", RsidRunAdditionDefault = "00957E57" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId24 = new ParagraphStyleId() { Val = "Disclaimer" };
            KeepNext keepNext1 = new KeepNext();
            WidowControl widowControl2 = new WidowControl() { Val = false };
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { Before = "80" };

            paragraphProperties24.Append(paragraphStyleId24);
            paragraphProperties24.Append(keepNext1);
            paragraphProperties24.Append(widowControl2);
            paragraphProperties24.Append(spacingBetweenLines17);

            CustomXmlRun customXmlRun8 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "LongDisclaimer" };

            Run run38 = new Run() { RsidRunProperties = "00C62467", RsidRunAddition = "00707FFA" };
            Text text28 = new Text();
            text28.Text = "Long Disclaimer";

            run38.Append(text28);

            customXmlRun8.Append(run38);

            paragraph25.Append(paragraphProperties24);
            paragraph25.Append(customXmlRun8);

            customXmlBlock11.Append(paragraph24);
            customXmlBlock11.Append(paragraph25);

            customXmlBlock10.Append(customXmlBlock11);

            CustomXmlBlock customXmlBlock12 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "PageBreak" };

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00233025", RsidParagraphAddition = "00E70E3A", RsidParagraphProperties = "00090FFC", RsidRunAdditionDefault = "00957E57" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunStyle runStyle6 = new RunStyle() { Val = "StyleBodoniMT" };

            paragraphMarkRunProperties14.Append(runStyle6);

            paragraphProperties25.Append(keepNext2);
            paragraphProperties25.Append(keepLines1);
            paragraphProperties25.Append(paragraphMarkRunProperties14);

            paragraph26.Append(paragraphProperties25);

            customXmlBlock12.Append(paragraph26);

            customXmlBlock1.Append(customXmlBlock2);
            customXmlBlock1.Append(customXmlBlock10);
            customXmlBlock1.Append(customXmlBlock12);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "00233025", RsidR = "00E70E3A", RsidSect = "00C913B8" };
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId10" };
            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId11" };
            HeaderReference headerReference2 = new HeaderReference() { Type = HeaderFooterValues.First, Id = "rId12" };
            FooterReference footerReference2 = new FooterReference() { Type = HeaderFooterValues.First, Id = "rId13" };
            SectionType sectionType1 = new SectionType() { Val = SectionMarkValues.Continuous };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12242U, Height = (UInt32Value)15842U, Code = (UInt16Value)1U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)720U, Bottom = 1440, Left = (UInt32Value)720U, Header = (UInt32Value)187U, Footer = (UInt32Value)115U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            TitlePage titlePage1 = new TitlePage();
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(headerReference1);
            sectionProperties1.Append(footerReference1);
            sectionProperties1.Append(headerReference2);
            sectionProperties1.Append(footerReference2);
            sectionProperties1.Append(sectionType1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(titlePage1);
            sectionProperties1.Append(docGrid1);

            body1.Append(customXmlBlock1);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        private static IEnumerable<OpenXmlElement> CreateTopicText(FormattedContent formattedContent)
        {
            var formattedContentParagraphs = formattedContent.ToOpenXmlElements();

            foreach (var para in formattedContentParagraphs)
            {
                if (para is Paragraph)
                {
                    if (((Paragraph)para).ParagraphProperties == null)
                    {
                        ParagraphProperties paragraphProperties22 = new ParagraphProperties();
                        ParagraphStyleId paragraphStyleId22 = new ParagraphStyleId() { Val = "StyleAfter0pt" };
                        SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { After = "360" };

                        paragraphProperties22.Append(paragraphStyleId22);
                        paragraphProperties22.Append(spacingBetweenLines16);
                        para.InsertAt(paragraphProperties22, 0);
                    }
                }
            }

            return formattedContentParagraphs;
        }

        private static Paragraph CreateTopicTitle(string title, string notation)
        {
            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00233025", RsidParagraphAddition = "00F8047A", RsidParagraphProperties = "00DD5BAE", RsidRunAdditionDefault = "00957E57" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "StyleBefore9ptAfter0pt" };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "StyleCategoryRankGraphic10pt" };

            paragraphMarkRunProperties13.Append(runStyle4);

            paragraphProperties21.Append(paragraphStyleId21);
            paragraphProperties21.Append(paragraphMarkRunProperties13);

            CustomXmlRun customXmlRun6 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "CatName" };

            Run run33 = new Run() { RsidRunProperties = "00233025", RsidRunAddition = "00F8047A" };

            RunProperties runProperties20 = new RunProperties();
            RunStyle runStyle5 = new RunStyle() { Val = "Style10ptBold" };

            runProperties20.Append(runStyle5);
            Text text24 = new Text();
            text24.Text = title;

            run33.Append(runProperties20);
            run33.Append(text24);

            customXmlRun6.Append(run33);

            CustomXmlRun customXmlRun7 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "CatRankValueImage" };

            if (!string.IsNullOrEmpty(notation))
            {
                Run run34 = new Run() { RsidRunAddition = "00740A1C" };

                RunProperties runProperties21 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { ComplexScript = "Arial" };
                Bold bold12 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                Caps caps1 = new Caps();
                NoProof noProof9 = new NoProof();
                Kern kern1 = new Kern() { Val = (UInt32Value)20U };
                Position position1 = new Position() { Val = "-4" };
                FontSize fontSize9 = new FontSize() { Val = "20" };
                Languages languages9 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

                runProperties21.Append(runFonts3);
                runProperties21.Append(bold12);
                runProperties21.Append(boldComplexScript1);
                runProperties21.Append(caps1);
                runProperties21.Append(noProof9);
                runProperties21.Append(kern1);
                runProperties21.Append(position1);
                runProperties21.Append(fontSize9);
                runProperties21.Append(languages9);

                Drawing drawing9 = new Drawing();

                Wp.Inline inline9 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
                Wp.Extent extent9 = new Wp.Extent() { Cx = 838200L, Cy = 152400L };
                Wp.EffectExtent effectExtent9 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
                Wp.DocProperties docProperties9 = new Wp.DocProperties() { Id = (UInt32Value)20U, Name = "Image 20", Description = "rank_category_5" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties9 = new Wp.NonVisualGraphicFrameDrawingProperties();
                A.GraphicFrameLocks graphicFrameLocks9 = new A.GraphicFrameLocks() { NoChangeAspect = true };

                nonVisualGraphicFrameDrawingProperties9.Append(graphicFrameLocks9);

                A.Graphic graphic9 = new A.Graphic();

                A.GraphicData graphicData9 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

                Pic.Picture picture9 = new Pic.Picture();

                Pic.NonVisualPictureProperties nonVisualPictureProperties9 = new Pic.NonVisualPictureProperties();
                Pic.NonVisualDrawingProperties nonVisualDrawingProperties9 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 20", Description = "rank_category_5" };

                Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties9 = new Pic.NonVisualPictureDrawingProperties();
                A.PictureLocks pictureLocks9 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

                nonVisualPictureDrawingProperties9.Append(pictureLocks9);

                nonVisualPictureProperties9.Append(nonVisualDrawingProperties9);
                nonVisualPictureProperties9.Append(nonVisualPictureDrawingProperties9);

                Pic.BlipFill blipFill9 = new Pic.BlipFill();
                A.Blip blip9 = new A.Blip() { Embed = "rIdImgPartTopicRank" + notation, CompressionState = A.BlipCompressionValues.Print };
                A.SourceRectangle sourceRectangle9 = new A.SourceRectangle();

                A.Stretch stretch9 = new A.Stretch();
                A.FillRectangle fillRectangle9 = new A.FillRectangle();

                stretch9.Append(fillRectangle9);

                blipFill9.Append(blip9);
                blipFill9.Append(sourceRectangle9);
                blipFill9.Append(stretch9);

                Pic.ShapeProperties shapeProperties9 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D9 = new A.Transform2D();
                A.Offset offset9 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents9 = new A.Extents() { Cx = 838200L, Cy = 152400L };

                transform2D9.Append(offset9);
                transform2D9.Append(extents9);

                A.PresetGeometry presetGeometry9 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList9 = new A.AdjustValueList();

                presetGeometry9.Append(adjustValueList9);
                A.NoFill noFill17 = new A.NoFill();

                A.Outline outline9 = new A.Outline() { Width = 9525 };
                A.NoFill noFill18 = new A.NoFill();
                A.Miter miter9 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd9 = new A.HeadEnd();
                A.TailEnd tailEnd9 = new A.TailEnd();

                outline9.Append(noFill18);
                outline9.Append(miter9);
                outline9.Append(headEnd9);
                outline9.Append(tailEnd9);

                shapeProperties9.Append(transform2D9);
                shapeProperties9.Append(presetGeometry9);
                shapeProperties9.Append(noFill17);
                shapeProperties9.Append(outline9);

                picture9.Append(nonVisualPictureProperties9);
                picture9.Append(blipFill9);
                picture9.Append(shapeProperties9);

                graphicData9.Append(picture9);

                graphic9.Append(graphicData9);

                inline9.Append(extent9);
                inline9.Append(effectExtent9);
                inline9.Append(docProperties9);
                inline9.Append(nonVisualGraphicFrameDrawingProperties9);
                inline9.Append(graphic9);

                drawing9.Append(inline9);

                run34.Append(runProperties21);
                run34.Append(drawing9);

                customXmlRun7.Append(run34);
            }

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(customXmlRun6);
            paragraph21.Append(customXmlRun7);
            return paragraph21;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // Generates content of footerPart1.
        private void GenerateFooterPart1Content(FooterPart footerPart1)
        {
            Footer footer1 = new Footer();

            Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "00C91FAF", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId25 = new ParagraphStyleId() { Val = "FooterRankLegend" };

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Left, Position = 3315 };

            tabs1.Append(tabStop1);
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines() { After = "240" };

            paragraphProperties26.Append(paragraphStyleId25);
            paragraphProperties26.Append(tabs1);
            paragraphProperties26.Append(spacingBetweenLines18);

            Run run39 = new Run();

            RunProperties runProperties22 = new RunProperties();
            NoProof noProof10 = new NoProof();
            Languages languages10 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties22.Append(noProof10);
            runProperties22.Append(languages10);

            Drawing drawing10 = new Drawing();

            Wp.Inline inline10 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent10 = new Wp.Extent() { Cx = 1447800L, Cy = 314325L };
            Wp.EffectExtent effectExtent10 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties10 = new Wp.DocProperties() { Id = (UInt32Value)23U, Name = "Image 23", Description = "RADAR_RankLegend" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties10 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks10 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties10.Append(graphicFrameLocks10);

            A.Graphic graphic10 = new A.Graphic();

            A.GraphicData graphicData10 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture10 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties10 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties10 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 23", Description = "RADAR_RankLegend" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties10 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks10 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties10.Append(pictureLocks10);

            nonVisualPictureProperties10.Append(nonVisualDrawingProperties10);
            nonVisualPictureProperties10.Append(nonVisualPictureDrawingProperties10);

            Pic.BlipFill blipFill10 = new Pic.BlipFill();
            A.Blip blip10 = new A.Blip() { Embed = "rId1" };
            A.SourceRectangle sourceRectangle10 = new A.SourceRectangle();

            A.Stretch stretch10 = new A.Stretch();
            A.FillRectangle fillRectangle10 = new A.FillRectangle();

            stretch10.Append(fillRectangle10);

            blipFill10.Append(blip10);
            blipFill10.Append(sourceRectangle10);
            blipFill10.Append(stretch10);

            Pic.ShapeProperties shapeProperties10 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset10 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents10 = new A.Extents() { Cx = 1447800L, Cy = 314325L };

            transform2D10.Append(offset10);
            transform2D10.Append(extents10);

            A.PresetGeometry presetGeometry10 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList10 = new A.AdjustValueList();

            presetGeometry10.Append(adjustValueList10);
            A.NoFill noFill19 = new A.NoFill();

            A.Outline outline10 = new A.Outline() { Width = 9525 };
            A.NoFill noFill20 = new A.NoFill();
            A.Miter miter10 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd10 = new A.HeadEnd();
            A.TailEnd tailEnd10 = new A.TailEnd();

            outline10.Append(noFill20);
            outline10.Append(miter10);
            outline10.Append(headEnd10);
            outline10.Append(tailEnd10);

            shapeProperties10.Append(transform2D10);
            shapeProperties10.Append(presetGeometry10);
            shapeProperties10.Append(noFill19);
            shapeProperties10.Append(outline10);

            picture10.Append(nonVisualPictureProperties10);
            picture10.Append(blipFill10);
            picture10.Append(shapeProperties10);

            graphicData10.Append(picture10);

            graphic10.Append(graphicData10);

            inline10.Append(extent10);
            inline10.Append(effectExtent10);
            inline10.Append(docProperties10);
            inline10.Append(nonVisualGraphicFrameDrawingProperties10);
            inline10.Append(graphic10);

            drawing10.Append(inline10);

            run39.Append(runProperties22);
            run39.Append(drawing10);

            Run run40 = new Run() { RsidRunAddition = "00C91FAF" };
            TabChar tabChar1 = new TabChar();

            run40.Append(tabChar1);

            paragraph27.Append(paragraphProperties26);
            paragraph27.Append(run39);
            paragraph27.Append(run40);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "00C91FAF", RsidRunAdditionDefault = "00C91FAF" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId26 = new ParagraphStyleId() { Val = "Disclaimer" };

            ParagraphBorders paragraphBorders1 = new ParagraphBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "66AADD", Size = (UInt32Value)48U, Space = (UInt32Value)1U };

            paragraphBorders1.Append(topBorder3);

            paragraphProperties27.Append(paragraphStyleId26);
            paragraphProperties27.Append(paragraphBorders1);

            paragraph28.Append(paragraphProperties27);

            Table table3 = new Table();

            TableProperties tableProperties3 = new TableProperties();
            TableStyle tableStyle3 = new TableStyle() { Val = "Grilledutableau" };
            TableWidth tableWidth3 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders3 = new TableBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders3.Append(topBorder4);
            tableBorders3.Append(leftBorder3);
            tableBorders3.Append(bottomBorder3);
            tableBorders3.Append(rightBorder3);
            tableBorders3.Append(insideHorizontalBorder3);
            tableBorders3.Append(insideVerticalBorder3);
            TableLayout tableLayout3 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook3 = new TableLook() { Val = "01E0" };

            tableProperties3.Append(tableStyle3);
            tableProperties3.Append(tableWidth3);
            tableProperties3.Append(tableBorders3);
            tableProperties3.Append(tableLayout3);
            tableProperties3.Append(tableLook3);

            TableGrid tableGrid3 = new TableGrid();
            GridColumn gridColumn8 = new GridColumn() { Width = "8388" };
            GridColumn gridColumn9 = new GridColumn() { Width = "540" };

            tableGrid3.Append(gridColumn8);
            tableGrid3.Append(gridColumn9);

            TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "00C91FAF", RsidTableRowProperties = "008F7383" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)618U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "8388", Type = TableWidthUnitValues.Dxa };

            tableCellProperties12.Append(tableCellWidth12);

            CustomXmlBlock customXmlBlock13 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "reportdoc" };

            CustomXmlBlock customXmlBlock14 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "footer" };

            CustomXmlBlock customXmlBlock15 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "ShortDisclaimer" };

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "003F1967", RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "008F7383", RsidRunAdditionDefault = "00C91FAF" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId27 = new ParagraphStyleId() { Val = "Disclaimer" };

            ParagraphBorders paragraphBorders2 = new ParagraphBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders2.Append(topBorder5);

            paragraphProperties28.Append(paragraphStyleId27);
            paragraphProperties28.Append(paragraphBorders2);

            Run run41 = new Run() { RsidRunProperties = "004B56C1" };
            Text text29 = new Text();
            text29.Text = "Confidential Proprietary Information of Russell Investments not to be distributed to third party without the express written consent of Russell Investments. Please see Important Legal Information for further information on this material.";

            run41.Append(text29);

            paragraph29.Append(paragraphProperties28);
            paragraph29.Append(run41);

            customXmlBlock15.Append(paragraph29);

            customXmlBlock14.Append(customXmlBlock15);

            customXmlBlock13.Append(customXmlBlock14);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "008F7383", RsidRunAdditionDefault = "00C91FAF" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId28 = new ParagraphStyleId() { Val = "FooterLogo" };

            paragraphProperties29.Append(paragraphStyleId28);

            paragraph30.Append(paragraphProperties29);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(customXmlBlock13);
            tableCell12.Append(paragraph30);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellVerticalAlignment1);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00FB4EAB", RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "008F7383", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId29 = new ParagraphStyleId() { Val = "FooterLogo" };
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties30.Append(paragraphStyleId29);
            paragraphProperties30.Append(justification1);

            Run run42 = new Run();

            RunProperties runProperties23 = new RunProperties();
            NoProof noProof11 = new NoProof();
            Languages languages11 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties23.Append(noProof11);
            runProperties23.Append(languages11);

            Drawing drawing11 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251657216U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "445770";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "102235";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent11 = new Wp.Extent() { Cx = 1085850L, Cy = 323850L };
            Wp.EffectExtent effectExtent11 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties11 = new Wp.DocProperties() { Id = (UInt32Value)71U, Name = "Image 71", Description = "RADAR_RLogo" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties11 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks11 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties11.Append(graphicFrameLocks11);

            A.Graphic graphic11 = new A.Graphic();

            A.GraphicData graphicData11 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture11 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties11 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties11 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 71", Description = "RADAR_RLogo" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties11 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks11 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties11.Append(pictureLocks11);

            nonVisualPictureProperties11.Append(nonVisualDrawingProperties11);
            nonVisualPictureProperties11.Append(nonVisualPictureDrawingProperties11);

            Pic.BlipFill blipFill11 = new Pic.BlipFill();
            A.Blip blip11 = new A.Blip() { Embed = "rId2" };
            A.SourceRectangle sourceRectangle11 = new A.SourceRectangle();

            A.Stretch stretch11 = new A.Stretch();
            A.FillRectangle fillRectangle11 = new A.FillRectangle();

            stretch11.Append(fillRectangle11);

            blipFill11.Append(blip11);
            blipFill11.Append(sourceRectangle11);
            blipFill11.Append(stretch11);

            Pic.ShapeProperties shapeProperties11 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D11 = new A.Transform2D();
            A.Offset offset11 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents11 = new A.Extents() { Cx = 1085850L, Cy = 323850L };

            transform2D11.Append(offset11);
            transform2D11.Append(extents11);

            A.PresetGeometry presetGeometry11 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList11 = new A.AdjustValueList();

            presetGeometry11.Append(adjustValueList11);
            A.NoFill noFill21 = new A.NoFill();

            shapeProperties11.Append(transform2D11);
            shapeProperties11.Append(presetGeometry11);
            shapeProperties11.Append(noFill21);

            picture11.Append(nonVisualPictureProperties11);
            picture11.Append(blipFill11);
            picture11.Append(shapeProperties11);

            graphicData11.Append(picture11);

            graphic11.Append(graphicData11);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent11);
            anchor1.Append(effectExtent11);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties11);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties11);
            anchor1.Append(graphic11);

            drawing11.Append(anchor1);

            run42.Append(runProperties23);
            run42.Append(drawing11);

            paragraph31.Append(paragraphProperties30);
            paragraph31.Append(run42);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph31);

            tableRow4.Append(tableRowProperties2);
            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);

            table3.Append(tableProperties3);
            table3.Append(tableGrid3);
            table3.Append(tableRow4);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "00C91FAF", RsidRunAdditionDefault = "00C91FAF" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId30 = new ParagraphStyleId() { Val = "FooterLogo" };

            paragraphProperties31.Append(paragraphStyleId30);

            paragraph32.Append(paragraphProperties31);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00F34666", RsidParagraphAddition = "00C91FAF", RsidParagraphProperties = "00C91FAF", RsidRunAdditionDefault = "00C91FAF" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId31 = new ParagraphStyleId() { Val = "FooterPageNumber" };
            SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines() { After = "320" };

            paragraphProperties32.Append(paragraphStyleId31);
            paragraphProperties32.Append(spacingBetweenLines19);

            paragraph33.Append(paragraphProperties32);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "00C91FAF", RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00C91FAF", RsidRunAdditionDefault = "002E7D22" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId32 = new ParagraphStyleId() { Val = "Pieddepage" };

            paragraphProperties33.Append(paragraphStyleId32);

            paragraph34.Append(paragraphProperties33);

            footer1.Append(paragraph27);
            footer1.Append(paragraph28);
            footer1.Append(table3);
            footer1.Append(paragraph32);
            footer1.Append(paragraph33);
            footer1.Append(paragraph34);

            footerPart1.Footer = footer1;
        }

        // Generates content of imagePart2.
        private void GenerateImagePart2Content(ImagePart imagePart2)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart2Data);
            imagePart2.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart3.
        private void GenerateImagePart3Content(ImagePart imagePart3)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart3Data);
            imagePart3.FeedData(data);
            data.Close();
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings();
            Zoom zoom1 = new Zoom() { Percent = "100" };
            AttachedTemplate attachedTemplate1 = new AttachedTemplate() { Id = "rId1" };
            LinkStyles linkStyles1 = new LinkStyles();
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            HyphenationZone hyphenationZone1 = new HyphenationZone() { Val = "425" };
            NoPunctuationKerning noPunctuationKerning1 = new NoPunctuationKerning();
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };
            DoNotValidateAgainstSchema doNotValidateAgainstSchema1 = new DoNotValidateAgainstSchema();
            SaveInvalidXml saveInvalidXml1 = new SaveInvalidXml();
            IgnoreMixedContent ignoreMixedContent1 = new IgnoreMixedContent();

            HeaderShapeDefaults headerShapeDefaults1 = new HeaderShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults1 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 300034 };

            headerShapeDefaults1.Append(shapeDefaults1);

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference() { Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference() { Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference() { Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference() { Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);

            Compatibility compatibility1 = new Compatibility();
            UseFarEastLayout useFarEastLayout1 = new UseFarEastLayout();

            compatibility1.Append(useFarEastLayout1);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "002E7D22" };
            Rsid rsid1 = new Rsid() { Val = "00000151" };
            Rsid rsid2 = new Rsid() { Val = "0000036F" };
            Rsid rsid3 = new Rsid() { Val = "000107CA" };
            Rsid rsid4 = new Rsid() { Val = "00011A83" };
            Rsid rsid5 = new Rsid() { Val = "00026250" };
            Rsid rsid6 = new Rsid() { Val = "0002709C" };
            Rsid rsid7 = new Rsid() { Val = "00033A61" };
            Rsid rsid8 = new Rsid() { Val = "000369F7" };
            Rsid rsid9 = new Rsid() { Val = "00037233" };
            Rsid rsid10 = new Rsid() { Val = "000432BA" };
            Rsid rsid11 = new Rsid() { Val = "00046BAE" };
            Rsid rsid12 = new Rsid() { Val = "00050CB1" };
            Rsid rsid13 = new Rsid() { Val = "00055488" };
            Rsid rsid14 = new Rsid() { Val = "00056308" };
            Rsid rsid15 = new Rsid() { Val = "00056347" };
            Rsid rsid16 = new Rsid() { Val = "00067CFE" };
            Rsid rsid17 = new Rsid() { Val = "00072CF0" };
            Rsid rsid18 = new Rsid() { Val = "00090FFC" };
            Rsid rsid19 = new Rsid() { Val = "000A24FF" };
            Rsid rsid20 = new Rsid() { Val = "000A318E" };
            Rsid rsid21 = new Rsid() { Val = "000A5DCE" };
            Rsid rsid22 = new Rsid() { Val = "000A6BF0" };
            Rsid rsid23 = new Rsid() { Val = "000A778A" };
            Rsid rsid24 = new Rsid() { Val = "000A7D4C" };
            Rsid rsid25 = new Rsid() { Val = "000B0A8A" };
            Rsid rsid26 = new Rsid() { Val = "000B25F2" };
            Rsid rsid27 = new Rsid() { Val = "000B5DD7" };
            Rsid rsid28 = new Rsid() { Val = "000C0418" };
            Rsid rsid29 = new Rsid() { Val = "000C29E0" };
            Rsid rsid30 = new Rsid() { Val = "000C42AA" };
            Rsid rsid31 = new Rsid() { Val = "000C4F36" };
            Rsid rsid32 = new Rsid() { Val = "000D1AC7" };
            Rsid rsid33 = new Rsid() { Val = "000D6093" };
            Rsid rsid34 = new Rsid() { Val = "000E1B96" };
            Rsid rsid35 = new Rsid() { Val = "000E2EEF" };
            Rsid rsid36 = new Rsid() { Val = "000E365D" };
            Rsid rsid37 = new Rsid() { Val = "000E5A65" };
            Rsid rsid38 = new Rsid() { Val = "000E62B4" };
            Rsid rsid39 = new Rsid() { Val = "000F60EB" };
            Rsid rsid40 = new Rsid() { Val = "000F68FD" };
            Rsid rsid41 = new Rsid() { Val = "00100B85" };
            Rsid rsid42 = new Rsid() { Val = "00103C4D" };
            Rsid rsid43 = new Rsid() { Val = "0010766E" };
            Rsid rsid44 = new Rsid() { Val = "00113178" };
            Rsid rsid45 = new Rsid() { Val = "0012032A" };
            Rsid rsid46 = new Rsid() { Val = "00121411" };
            Rsid rsid47 = new Rsid() { Val = "001255B1" };
            Rsid rsid48 = new Rsid() { Val = "0013124D" };
            Rsid rsid49 = new Rsid() { Val = "001357A3" };
            Rsid rsid50 = new Rsid() { Val = "00136DF8" };
            Rsid rsid51 = new Rsid() { Val = "00140622" };
            Rsid rsid52 = new Rsid() { Val = "00141D32" };
            Rsid rsid53 = new Rsid() { Val = "00150E17" };
            Rsid rsid54 = new Rsid() { Val = "00151C60" };
            Rsid rsid55 = new Rsid() { Val = "00156A9E" };
            Rsid rsid56 = new Rsid() { Val = "00157D8A" };
            Rsid rsid57 = new Rsid() { Val = "001717CD" };
            Rsid rsid58 = new Rsid() { Val = "001777F8" };
            Rsid rsid59 = new Rsid() { Val = "00185AFD" };
            Rsid rsid60 = new Rsid() { Val = "00187ACA" };
            Rsid rsid61 = new Rsid() { Val = "0019393C" };
            Rsid rsid62 = new Rsid() { Val = "00194B6D" };
            Rsid rsid63 = new Rsid() { Val = "00195CE8" };
            Rsid rsid64 = new Rsid() { Val = "00196E39" };
            Rsid rsid65 = new Rsid() { Val = "001974FE" };
            Rsid rsid66 = new Rsid() { Val = "001A0A63" };
            Rsid rsid67 = new Rsid() { Val = "001A4082" };
            Rsid rsid68 = new Rsid() { Val = "001A67D6" };
            Rsid rsid69 = new Rsid() { Val = "001B011D" };
            Rsid rsid70 = new Rsid() { Val = "001B225F" };
            Rsid rsid71 = new Rsid() { Val = "001B6F6C" };
            Rsid rsid72 = new Rsid() { Val = "001B746B" };
            Rsid rsid73 = new Rsid() { Val = "001C4579" };
            Rsid rsid74 = new Rsid() { Val = "001D5E74" };
            Rsid rsid75 = new Rsid() { Val = "001D60A9" };
            Rsid rsid76 = new Rsid() { Val = "001D6F62" };
            Rsid rsid77 = new Rsid() { Val = "001E724B" };
            Rsid rsid78 = new Rsid() { Val = "001F7499" };
            Rsid rsid79 = new Rsid() { Val = "002058A3" };
            Rsid rsid80 = new Rsid() { Val = "00206CCC" };
            Rsid rsid81 = new Rsid() { Val = "00216E92" };
            Rsid rsid82 = new Rsid() { Val = "00222633" };
            Rsid rsid83 = new Rsid() { Val = "0022367F" };
            Rsid rsid84 = new Rsid() { Val = "00227F18" };
            Rsid rsid85 = new Rsid() { Val = "00233025" };
            Rsid rsid86 = new Rsid() { Val = "00233DD2" };
            Rsid rsid87 = new Rsid() { Val = "00240AED" };
            Rsid rsid88 = new Rsid() { Val = "00243071" };
            Rsid rsid89 = new Rsid() { Val = "00244BC4" };
            Rsid rsid90 = new Rsid() { Val = "002462E2" };
            Rsid rsid91 = new Rsid() { Val = "002509D3" };
            Rsid rsid92 = new Rsid() { Val = "00256073" };
            Rsid rsid93 = new Rsid() { Val = "002614AB" };
            Rsid rsid94 = new Rsid() { Val = "00261E4D" };
            Rsid rsid95 = new Rsid() { Val = "0026354B" };
            Rsid rsid96 = new Rsid() { Val = "00264CFB" };
            Rsid rsid97 = new Rsid() { Val = "00267B3B" };
            Rsid rsid98 = new Rsid() { Val = "00272AD9" };
            Rsid rsid99 = new Rsid() { Val = "00273990" };
            Rsid rsid100 = new Rsid() { Val = "00275427" };
            Rsid rsid101 = new Rsid() { Val = "00280514" };
            Rsid rsid102 = new Rsid() { Val = "002825BD" };
            Rsid rsid103 = new Rsid() { Val = "00282F5E" };
            Rsid rsid104 = new Rsid() { Val = "002836D3" };
            Rsid rsid105 = new Rsid() { Val = "00283CC4" };
            Rsid rsid106 = new Rsid() { Val = "00292192" };
            Rsid rsid107 = new Rsid() { Val = "00295BC7" };
            Rsid rsid108 = new Rsid() { Val = "002A248B" };
            Rsid rsid109 = new Rsid() { Val = "002A33EC" };
            Rsid rsid110 = new Rsid() { Val = "002A4E0E" };
            Rsid rsid111 = new Rsid() { Val = "002A5362" };
            Rsid rsid112 = new Rsid() { Val = "002A7539" };
            Rsid rsid113 = new Rsid() { Val = "002B0C9D" };
            Rsid rsid114 = new Rsid() { Val = "002B7BE9" };
            Rsid rsid115 = new Rsid() { Val = "002C2313" };
            Rsid rsid116 = new Rsid() { Val = "002C57EF" };
            Rsid rsid117 = new Rsid() { Val = "002D3F1B" };
            Rsid rsid118 = new Rsid() { Val = "002D5864" };
            Rsid rsid119 = new Rsid() { Val = "002D5BB8" };
            Rsid rsid120 = new Rsid() { Val = "002E1DC5" };
            Rsid rsid121 = new Rsid() { Val = "002E32B1" };
            Rsid rsid122 = new Rsid() { Val = "002E4819" };
            Rsid rsid123 = new Rsid() { Val = "002E5FC9" };
            Rsid rsid124 = new Rsid() { Val = "002E707E" };
            Rsid rsid125 = new Rsid() { Val = "002E7D22" };
            Rsid rsid126 = new Rsid() { Val = "003000C4" };
            Rsid rsid127 = new Rsid() { Val = "0031558F" };
            Rsid rsid128 = new Rsid() { Val = "00317CC3" };
            Rsid rsid129 = new Rsid() { Val = "00336490" };
            Rsid rsid130 = new Rsid() { Val = "00336875" };
            Rsid rsid131 = new Rsid() { Val = "0034101A" };
            Rsid rsid132 = new Rsid() { Val = "00344459" };
            Rsid rsid133 = new Rsid() { Val = "00354726" };
            Rsid rsid134 = new Rsid() { Val = "00356266" };
            Rsid rsid135 = new Rsid() { Val = "00374165" };
            Rsid rsid136 = new Rsid() { Val = "00375BA9" };
            Rsid rsid137 = new Rsid() { Val = "00381D4A" };
            Rsid rsid138 = new Rsid() { Val = "00385C36" };
            Rsid rsid139 = new Rsid() { Val = "00387921" };
            Rsid rsid140 = new Rsid() { Val = "003907B3" };
            Rsid rsid141 = new Rsid() { Val = "003924C4" };
            Rsid rsid142 = new Rsid() { Val = "00397A9E" };
            Rsid rsid143 = new Rsid() { Val = "003C2666" };
            Rsid rsid144 = new Rsid() { Val = "003C5F5D" };
            Rsid rsid145 = new Rsid() { Val = "003C6DE9" };
            Rsid rsid146 = new Rsid() { Val = "003D7741" };
            Rsid rsid147 = new Rsid() { Val = "003D786B" };
            Rsid rsid148 = new Rsid() { Val = "003E0EAC" };
            Rsid rsid149 = new Rsid() { Val = "003E4D99" };
            Rsid rsid150 = new Rsid() { Val = "003F0DD8" };
            Rsid rsid151 = new Rsid() { Val = "003F1967" };
            Rsid rsid152 = new Rsid() { Val = "003F1E87" };
            Rsid rsid153 = new Rsid() { Val = "003F2779" };
            Rsid rsid154 = new Rsid() { Val = "00401509" };
            Rsid rsid155 = new Rsid() { Val = "00413F24" };
            Rsid rsid156 = new Rsid() { Val = "00417B92" };
            Rsid rsid157 = new Rsid() { Val = "00417CB6" };
            Rsid rsid158 = new Rsid() { Val = "00423094" };
            Rsid rsid159 = new Rsid() { Val = "00430E1B" };
            Rsid rsid160 = new Rsid() { Val = "00432185" };
            Rsid rsid161 = new Rsid() { Val = "004427CC" };
            Rsid rsid162 = new Rsid() { Val = "004431D7" };
            Rsid rsid163 = new Rsid() { Val = "004438D9" };
            Rsid rsid164 = new Rsid() { Val = "00443A55" };
            Rsid rsid165 = new Rsid() { Val = "00443CD0" };
            Rsid rsid166 = new Rsid() { Val = "00444A32" };
            Rsid rsid167 = new Rsid() { Val = "00447118" };
            Rsid rsid168 = new Rsid() { Val = "00451171" };
            Rsid rsid169 = new Rsid() { Val = "004522C9" };
            Rsid rsid170 = new Rsid() { Val = "00453E48" };
            Rsid rsid171 = new Rsid() { Val = "00456F84" };
            Rsid rsid172 = new Rsid() { Val = "00461D37" };
            Rsid rsid173 = new Rsid() { Val = "00464492" };
            Rsid rsid174 = new Rsid() { Val = "004657C1" };
            Rsid rsid175 = new Rsid() { Val = "00466898" };
            Rsid rsid176 = new Rsid() { Val = "00472DEA" };
            Rsid rsid177 = new Rsid() { Val = "004805C1" };
            Rsid rsid178 = new Rsid() { Val = "004826CB" };
            Rsid rsid179 = new Rsid() { Val = "0049162E" };
            Rsid rsid180 = new Rsid() { Val = "00495D69" };
            Rsid rsid181 = new Rsid() { Val = "004A5AE6" };
            Rsid rsid182 = new Rsid() { Val = "004A6E9F" };
            Rsid rsid183 = new Rsid() { Val = "004A7F93" };
            Rsid rsid184 = new Rsid() { Val = "004B2B7F" };
            Rsid rsid185 = new Rsid() { Val = "004C0DA7" };
            Rsid rsid186 = new Rsid() { Val = "004C4687" };
            Rsid rsid187 = new Rsid() { Val = "004C7C60" };
            Rsid rsid188 = new Rsid() { Val = "004D12BE" };
            Rsid rsid189 = new Rsid() { Val = "004D5ECC" };
            Rsid rsid190 = new Rsid() { Val = "004E16FD" };
            Rsid rsid191 = new Rsid() { Val = "004E195A" };
            Rsid rsid192 = new Rsid() { Val = "004E54D9" };
            Rsid rsid193 = new Rsid() { Val = "004E7907" };
            Rsid rsid194 = new Rsid() { Val = "004F2494" };
            Rsid rsid195 = new Rsid() { Val = "004F2A92" };
            Rsid rsid196 = new Rsid() { Val = "00506462" };
            Rsid rsid197 = new Rsid() { Val = "00514769" };
            Rsid rsid198 = new Rsid() { Val = "00517D57" };
            Rsid rsid199 = new Rsid() { Val = "00523FC2" };
            Rsid rsid200 = new Rsid() { Val = "00524AA7" };
            Rsid rsid201 = new Rsid() { Val = "00532951" };
            Rsid rsid202 = new Rsid() { Val = "00535256" };
            Rsid rsid203 = new Rsid() { Val = "00544156" };
            Rsid rsid204 = new Rsid() { Val = "00545261" };
            Rsid rsid205 = new Rsid() { Val = "00547DDD" };
            Rsid rsid206 = new Rsid() { Val = "005536AE" };
            Rsid rsid207 = new Rsid() { Val = "00554657" };
            Rsid rsid208 = new Rsid() { Val = "005546F4" };
            Rsid rsid209 = new Rsid() { Val = "00560517" };
            Rsid rsid210 = new Rsid() { Val = "00561A98" };
            Rsid rsid211 = new Rsid() { Val = "005623FA" };
            Rsid rsid212 = new Rsid() { Val = "00562F9B" };
            Rsid rsid213 = new Rsid() { Val = "00566334" };
            Rsid rsid214 = new Rsid() { Val = "00572029" };
            Rsid rsid215 = new Rsid() { Val = "00574E3B" };
            Rsid rsid216 = new Rsid() { Val = "005754DB" };
            Rsid rsid217 = new Rsid() { Val = "00583853" };
            Rsid rsid218 = new Rsid() { Val = "00583E34" };
            Rsid rsid219 = new Rsid() { Val = "00584020" };
            Rsid rsid220 = new Rsid() { Val = "00592E66" };
            Rsid rsid221 = new Rsid() { Val = "00594B09" };
            Rsid rsid222 = new Rsid() { Val = "005A4976" };
            Rsid rsid223 = new Rsid() { Val = "005A4CCC" };
            Rsid rsid224 = new Rsid() { Val = "005A4E70" };
            Rsid rsid225 = new Rsid() { Val = "005A6F62" };
            Rsid rsid226 = new Rsid() { Val = "005B5C7C" };
            Rsid rsid227 = new Rsid() { Val = "005B7379" };
            Rsid rsid228 = new Rsid() { Val = "005C5D0D" };
            Rsid rsid229 = new Rsid() { Val = "005C5E18" };
            Rsid rsid230 = new Rsid() { Val = "005C722F" };
            Rsid rsid231 = new Rsid() { Val = "005D27DE" };
            Rsid rsid232 = new Rsid() { Val = "005D5D40" };
            Rsid rsid233 = new Rsid() { Val = "005D7E7A" };
            Rsid rsid234 = new Rsid() { Val = "005E40AC" };
            Rsid rsid235 = new Rsid() { Val = "005F2848" };
            Rsid rsid236 = new Rsid() { Val = "005F4DB9" };
            Rsid rsid237 = new Rsid() { Val = "005F6B60" };
            Rsid rsid238 = new Rsid() { Val = "006209F6" };
            Rsid rsid239 = new Rsid() { Val = "006248C1" };
            Rsid rsid240 = new Rsid() { Val = "0063624F" };
            Rsid rsid241 = new Rsid() { Val = "0065191A" };
            Rsid rsid242 = new Rsid() { Val = "00660821" };
            Rsid rsid243 = new Rsid() { Val = "00675ED0" };
            Rsid rsid244 = new Rsid() { Val = "00681EB3" };
            Rsid rsid245 = new Rsid() { Val = "006862EE" };
            Rsid rsid246 = new Rsid() { Val = "0069278B" };
            Rsid rsid247 = new Rsid() { Val = "006A0B1E" };
            Rsid rsid248 = new Rsid() { Val = "006A58D3" };
            Rsid rsid249 = new Rsid() { Val = "006B1D99" };
            Rsid rsid250 = new Rsid() { Val = "006D1BD8" };
            Rsid rsid251 = new Rsid() { Val = "006D2E84" };
            Rsid rsid252 = new Rsid() { Val = "006D39A0" };
            Rsid rsid253 = new Rsid() { Val = "006D6972" };
            Rsid rsid254 = new Rsid() { Val = "006E1177" };
            Rsid rsid255 = new Rsid() { Val = "006E384B" };
            Rsid rsid256 = new Rsid() { Val = "006E5BF6" };
            Rsid rsid257 = new Rsid() { Val = "006E6F86" };
            Rsid rsid258 = new Rsid() { Val = "006F02A4" };
            Rsid rsid259 = new Rsid() { Val = "006F57DE" };
            Rsid rsid260 = new Rsid() { Val = "006F58EB" };
            Rsid rsid261 = new Rsid() { Val = "006F6E20" };
            Rsid rsid262 = new Rsid() { Val = "006F74B4" };
            Rsid rsid263 = new Rsid() { Val = "00700510" };
            Rsid rsid264 = new Rsid() { Val = "00703676" };
            Rsid rsid265 = new Rsid() { Val = "00707FFA" };
            Rsid rsid266 = new Rsid() { Val = "00713208" };
            Rsid rsid267 = new Rsid() { Val = "007171C9" };
            Rsid rsid268 = new Rsid() { Val = "007247F0" };
            Rsid rsid269 = new Rsid() { Val = "00731EBE" };
            Rsid rsid270 = new Rsid() { Val = "0073257C" };
            Rsid rsid271 = new Rsid() { Val = "00740A1C" };
            Rsid rsid272 = new Rsid() { Val = "00751786" };
            Rsid rsid273 = new Rsid() { Val = "00751A3A" };
            Rsid rsid274 = new Rsid() { Val = "00751D5C" };
            Rsid rsid275 = new Rsid() { Val = "00753A74" };
            Rsid rsid276 = new Rsid() { Val = "00754A89" };
            Rsid rsid277 = new Rsid() { Val = "00761A8E" };
            Rsid rsid278 = new Rsid() { Val = "007643CF" };
            Rsid rsid279 = new Rsid() { Val = "00782598" };
            Rsid rsid280 = new Rsid() { Val = "0078481C" };
            Rsid rsid281 = new Rsid() { Val = "0079064C" };
            Rsid rsid282 = new Rsid() { Val = "007A0670" };
            Rsid rsid283 = new Rsid() { Val = "007A234D" };
            Rsid rsid284 = new Rsid() { Val = "007A2948" };
            Rsid rsid285 = new Rsid() { Val = "007A41D7" };
            Rsid rsid286 = new Rsid() { Val = "007B0BC0" };
            Rsid rsid287 = new Rsid() { Val = "007B2876" };
            Rsid rsid288 = new Rsid() { Val = "007B6346" };
            Rsid rsid289 = new Rsid() { Val = "007B661A" };
            Rsid rsid290 = new Rsid() { Val = "007B66F2" };
            Rsid rsid291 = new Rsid() { Val = "007B7BF0" };
            Rsid rsid292 = new Rsid() { Val = "007C1300" };
            Rsid rsid293 = new Rsid() { Val = "007C34CD" };
            Rsid rsid294 = new Rsid() { Val = "007C3DB4" };
            Rsid rsid295 = new Rsid() { Val = "007C4997" };
            Rsid rsid296 = new Rsid() { Val = "007C6A01" };
            Rsid rsid297 = new Rsid() { Val = "007C6AF0" };
            Rsid rsid298 = new Rsid() { Val = "007D0EFC" };
            Rsid rsid299 = new Rsid() { Val = "007D64BE" };
            Rsid rsid300 = new Rsid() { Val = "007E1E2F" };
            Rsid rsid301 = new Rsid() { Val = "007E6AA4" };
            Rsid rsid302 = new Rsid() { Val = "007E7586" };
            Rsid rsid303 = new Rsid() { Val = "008154D4" };
            Rsid rsid304 = new Rsid() { Val = "0082549B" };
            Rsid rsid305 = new Rsid() { Val = "008278CF" };
            Rsid rsid306 = new Rsid() { Val = "0083140B" };
            Rsid rsid307 = new Rsid() { Val = "00835077" };
            Rsid rsid308 = new Rsid() { Val = "00837232" };
            Rsid rsid309 = new Rsid() { Val = "00837AE4" };
            Rsid rsid310 = new Rsid() { Val = "008439F9" };
            Rsid rsid311 = new Rsid() { Val = "00850F31" };
            Rsid rsid312 = new Rsid() { Val = "00851D16" };
            Rsid rsid313 = new Rsid() { Val = "00852F72" };
            Rsid rsid314 = new Rsid() { Val = "00855B1B" };
            Rsid rsid315 = new Rsid() { Val = "008602C0" };
            Rsid rsid316 = new Rsid() { Val = "00860BAA" };
            Rsid rsid317 = new Rsid() { Val = "00862EA1" };
            Rsid rsid318 = new Rsid() { Val = "0086518C" };
            Rsid rsid319 = new Rsid() { Val = "00871C48" };
            Rsid rsid320 = new Rsid() { Val = "00872DEA" };
            Rsid rsid321 = new Rsid() { Val = "008737BB" };
            Rsid rsid322 = new Rsid() { Val = "00880A12" };
            Rsid rsid323 = new Rsid() { Val = "0088564B" };
            Rsid rsid324 = new Rsid() { Val = "00890BFC" };
            Rsid rsid325 = new Rsid() { Val = "00894D97" };
            Rsid rsid326 = new Rsid() { Val = "00897ECC" };
            Rsid rsid327 = new Rsid() { Val = "008A25E5" };
            Rsid rsid328 = new Rsid() { Val = "008B2B2F" };
            Rsid rsid329 = new Rsid() { Val = "008B419B" };
            Rsid rsid330 = new Rsid() { Val = "008B561B" };
            Rsid rsid331 = new Rsid() { Val = "008C067B" };
            Rsid rsid332 = new Rsid() { Val = "008C2D52" };
            Rsid rsid333 = new Rsid() { Val = "008C4316" };
            Rsid rsid334 = new Rsid() { Val = "008D2C48" };
            Rsid rsid335 = new Rsid() { Val = "008D69D4" };
            Rsid rsid336 = new Rsid() { Val = "008D6C0E" };
            Rsid rsid337 = new Rsid() { Val = "008E00D5" };
            Rsid rsid338 = new Rsid() { Val = "008F2CC8" };
            Rsid rsid339 = new Rsid() { Val = "00902E88" };
            Rsid rsid340 = new Rsid() { Val = "00913955" };
            Rsid rsid341 = new Rsid() { Val = "00915758" };
            Rsid rsid342 = new Rsid() { Val = "009166B9" };
            Rsid rsid343 = new Rsid() { Val = "00934E6E" };
            Rsid rsid344 = new Rsid() { Val = "00937F8E" };
            Rsid rsid345 = new Rsid() { Val = "00940CCD" };
            Rsid rsid346 = new Rsid() { Val = "0094125B" };
            Rsid rsid347 = new Rsid() { Val = "00944624" };
            Rsid rsid348 = new Rsid() { Val = "009505D2" };
            Rsid rsid349 = new Rsid() { Val = "009527BD" };
            Rsid rsid350 = new Rsid() { Val = "00957E57" };
            Rsid rsid351 = new Rsid() { Val = "00965C1D" };
            Rsid rsid352 = new Rsid() { Val = "00983F27" };
            Rsid rsid353 = new Rsid() { Val = "009A4C13" };
            Rsid rsid354 = new Rsid() { Val = "009A55D5" };
            Rsid rsid355 = new Rsid() { Val = "009A6239" };
            Rsid rsid356 = new Rsid() { Val = "009A6370" };
            Rsid rsid357 = new Rsid() { Val = "009B6613" };
            Rsid rsid358 = new Rsid() { Val = "009C6FC7" };
            Rsid rsid359 = new Rsid() { Val = "009E2BFE" };
            Rsid rsid360 = new Rsid() { Val = "009E5742" };
            Rsid rsid361 = new Rsid() { Val = "009E5D63" };
            Rsid rsid362 = new Rsid() { Val = "009E5E64" };
            Rsid rsid363 = new Rsid() { Val = "009F02AD" };
            Rsid rsid364 = new Rsid() { Val = "009F0608" };
            Rsid rsid365 = new Rsid() { Val = "009F0B75" };
            Rsid rsid366 = new Rsid() { Val = "009F3BBF" };
            Rsid rsid367 = new Rsid() { Val = "009F453D" };
            Rsid rsid368 = new Rsid() { Val = "009F47A3" };
            Rsid rsid369 = new Rsid() { Val = "009F7E7F" };
            Rsid rsid370 = new Rsid() { Val = "00A129B7" };
            Rsid rsid371 = new Rsid() { Val = "00A15F1D" };
            Rsid rsid372 = new Rsid() { Val = "00A20A5F" };
            Rsid rsid373 = new Rsid() { Val = "00A21AB2" };
            Rsid rsid374 = new Rsid() { Val = "00A22E79" };
            Rsid rsid375 = new Rsid() { Val = "00A241C0" };
            Rsid rsid376 = new Rsid() { Val = "00A27025" };
            Rsid rsid377 = new Rsid() { Val = "00A3564C" };
            Rsid rsid378 = new Rsid() { Val = "00A41F12" };
            Rsid rsid379 = new Rsid() { Val = "00A53072" };
            Rsid rsid380 = new Rsid() { Val = "00A6171D" };
            Rsid rsid381 = new Rsid() { Val = "00A62399" };
            Rsid rsid382 = new Rsid() { Val = "00A65073" };
            Rsid rsid383 = new Rsid() { Val = "00A747BB" };
            Rsid rsid384 = new Rsid() { Val = "00A75BDE" };
            Rsid rsid385 = new Rsid() { Val = "00A7674D" };
            Rsid rsid386 = new Rsid() { Val = "00A77710" };
            Rsid rsid387 = new Rsid() { Val = "00A82025" };
            Rsid rsid388 = new Rsid() { Val = "00A8638E" };
            Rsid rsid389 = new Rsid() { Val = "00A918FB" };
            Rsid rsid390 = new Rsid() { Val = "00A921AB" };
            Rsid rsid391 = new Rsid() { Val = "00A935F6" };
            Rsid rsid392 = new Rsid() { Val = "00AA0A1A" };
            Rsid rsid393 = new Rsid() { Val = "00AA6279" };
            Rsid rsid394 = new Rsid() { Val = "00AB1206" };
            Rsid rsid395 = new Rsid() { Val = "00AB318F" };
            Rsid rsid396 = new Rsid() { Val = "00AB3D48" };
            Rsid rsid397 = new Rsid() { Val = "00AB4921" };
            Rsid rsid398 = new Rsid() { Val = "00AB5753" };
            Rsid rsid399 = new Rsid() { Val = "00AC0771" };
            Rsid rsid400 = new Rsid() { Val = "00AC1437" };
            Rsid rsid401 = new Rsid() { Val = "00AC1B75" };
            Rsid rsid402 = new Rsid() { Val = "00AD0D68" };
            Rsid rsid403 = new Rsid() { Val = "00AD1EF0" };
            Rsid rsid404 = new Rsid() { Val = "00AD5D8C" };
            Rsid rsid405 = new Rsid() { Val = "00AD5E16" };
            Rsid rsid406 = new Rsid() { Val = "00AD61F0" };
            Rsid rsid407 = new Rsid() { Val = "00AD6EBD" };
            Rsid rsid408 = new Rsid() { Val = "00AE3C1A" };
            Rsid rsid409 = new Rsid() { Val = "00AF0136" };
            Rsid rsid410 = new Rsid() { Val = "00AF598C" };
            Rsid rsid411 = new Rsid() { Val = "00AF701E" };
            Rsid rsid412 = new Rsid() { Val = "00AF7795" };
            Rsid rsid413 = new Rsid() { Val = "00B02C12" };
            Rsid rsid414 = new Rsid() { Val = "00B03AFD" };
            Rsid rsid415 = new Rsid() { Val = "00B0579B" };
            Rsid rsid416 = new Rsid() { Val = "00B062C8" };
            Rsid rsid417 = new Rsid() { Val = "00B105DC" };
            Rsid rsid418 = new Rsid() { Val = "00B14BDB" };
            Rsid rsid419 = new Rsid() { Val = "00B21366" };
            Rsid rsid420 = new Rsid() { Val = "00B33F8A" };
            Rsid rsid421 = new Rsid() { Val = "00B354C8" };
            Rsid rsid422 = new Rsid() { Val = "00B401C3" };
            Rsid rsid423 = new Rsid() { Val = "00B41D00" };
            Rsid rsid424 = new Rsid() { Val = "00B62341" };
            Rsid rsid425 = new Rsid() { Val = "00B64DAD" };
            Rsid rsid426 = new Rsid() { Val = "00B67E72" };
            Rsid rsid427 = new Rsid() { Val = "00B70A98" };
            Rsid rsid428 = new Rsid() { Val = "00B72CF9" };
            Rsid rsid429 = new Rsid() { Val = "00B72E61" };
            Rsid rsid430 = new Rsid() { Val = "00B75F5F" };
            Rsid rsid431 = new Rsid() { Val = "00B80285" };
            Rsid rsid432 = new Rsid() { Val = "00B900F7" };
            Rsid rsid433 = new Rsid() { Val = "00B95F6C" };
            Rsid rsid434 = new Rsid() { Val = "00B97B60" };
            Rsid rsid435 = new Rsid() { Val = "00BA224B" };
            Rsid rsid436 = new Rsid() { Val = "00BA33DE" };
            Rsid rsid437 = new Rsid() { Val = "00BA36E7" };
            Rsid rsid438 = new Rsid() { Val = "00BA543D" };
            Rsid rsid439 = new Rsid() { Val = "00BA5E83" };
            Rsid rsid440 = new Rsid() { Val = "00BA679D" };
            Rsid rsid441 = new Rsid() { Val = "00BA6DC3" };
            Rsid rsid442 = new Rsid() { Val = "00BA7E3F" };
            Rsid rsid443 = new Rsid() { Val = "00BB0D74" };
            Rsid rsid444 = new Rsid() { Val = "00BB40B9" };
            Rsid rsid445 = new Rsid() { Val = "00BB5522" };
            Rsid rsid446 = new Rsid() { Val = "00BB5D40" };
            Rsid rsid447 = new Rsid() { Val = "00BC106F" };
            Rsid rsid448 = new Rsid() { Val = "00BC4CAC" };
            Rsid rsid449 = new Rsid() { Val = "00BE25CC" };
            Rsid rsid450 = new Rsid() { Val = "00BE574F" };
            Rsid rsid451 = new Rsid() { Val = "00BF66F3" };
            Rsid rsid452 = new Rsid() { Val = "00BF6F89" };
            Rsid rsid453 = new Rsid() { Val = "00C01F31" };
            Rsid rsid454 = new Rsid() { Val = "00C03444" };
            Rsid rsid455 = new Rsid() { Val = "00C04CE0" };
            Rsid rsid456 = new Rsid() { Val = "00C062F9" };
            Rsid rsid457 = new Rsid() { Val = "00C1094B" };
            Rsid rsid458 = new Rsid() { Val = "00C166F8" };
            Rsid rsid459 = new Rsid() { Val = "00C24508" };
            Rsid rsid460 = new Rsid() { Val = "00C27D0F" };
            Rsid rsid461 = new Rsid() { Val = "00C31288" };
            Rsid rsid462 = new Rsid() { Val = "00C32704" };
            Rsid rsid463 = new Rsid() { Val = "00C33CF0" };
            Rsid rsid464 = new Rsid() { Val = "00C41ABA" };
            Rsid rsid465 = new Rsid() { Val = "00C44584" };
            Rsid rsid466 = new Rsid() { Val = "00C4492A" };
            Rsid rsid467 = new Rsid() { Val = "00C47206" };
            Rsid rsid468 = new Rsid() { Val = "00C6049A" };
            Rsid rsid469 = new Rsid() { Val = "00C62467" };
            Rsid rsid470 = new Rsid() { Val = "00C704CB" };
            Rsid rsid471 = new Rsid() { Val = "00C71CC7" };
            Rsid rsid472 = new Rsid() { Val = "00C7570A" };
            Rsid rsid473 = new Rsid() { Val = "00C803F9" };
            Rsid rsid474 = new Rsid() { Val = "00C82A8F" };
            Rsid rsid475 = new Rsid() { Val = "00C8492B" };
            Rsid rsid476 = new Rsid() { Val = "00C86FD9" };
            Rsid rsid477 = new Rsid() { Val = "00C913B8" };
            Rsid rsid478 = new Rsid() { Val = "00C91FAF" };
            Rsid rsid479 = new Rsid() { Val = "00CA34E5" };
            Rsid rsid480 = new Rsid() { Val = "00CA68C1" };
            Rsid rsid481 = new Rsid() { Val = "00CA7AED" };
            Rsid rsid482 = new Rsid() { Val = "00CB005B" };
            Rsid rsid483 = new Rsid() { Val = "00CB3330" };
            Rsid rsid484 = new Rsid() { Val = "00CB77EF" };
            Rsid rsid485 = new Rsid() { Val = "00CB7FBF" };
            Rsid rsid486 = new Rsid() { Val = "00CC0E6F" };
            Rsid rsid487 = new Rsid() { Val = "00CC2476" };
            Rsid rsid488 = new Rsid() { Val = "00CC46B9" };
            Rsid rsid489 = new Rsid() { Val = "00CD3A9B" };
            Rsid rsid490 = new Rsid() { Val = "00CD57B8" };
            Rsid rsid491 = new Rsid() { Val = "00CE0591" };
            Rsid rsid492 = new Rsid() { Val = "00CE06F6" };
            Rsid rsid493 = new Rsid() { Val = "00CE60E8" };
            Rsid rsid494 = new Rsid() { Val = "00CE7054" };
            Rsid rsid495 = new Rsid() { Val = "00CF0CEB" };
            Rsid rsid496 = new Rsid() { Val = "00CF67E9" };
            Rsid rsid497 = new Rsid() { Val = "00D03B21" };
            Rsid rsid498 = new Rsid() { Val = "00D11E55" };
            Rsid rsid499 = new Rsid() { Val = "00D2246F" };
            Rsid rsid500 = new Rsid() { Val = "00D232C4" };
            Rsid rsid501 = new Rsid() { Val = "00D23A7D" };
            Rsid rsid502 = new Rsid() { Val = "00D412D1" };
            Rsid rsid503 = new Rsid() { Val = "00D42739" };
            Rsid rsid504 = new Rsid() { Val = "00D45D2A" };
            Rsid rsid505 = new Rsid() { Val = "00D525A8" };
            Rsid rsid506 = new Rsid() { Val = "00D52961" };
            Rsid rsid507 = new Rsid() { Val = "00D53395" };
            Rsid rsid508 = new Rsid() { Val = "00D56C12" };
            Rsid rsid509 = new Rsid() { Val = "00D60895" };
            Rsid rsid510 = new Rsid() { Val = "00D62CD3" };
            Rsid rsid511 = new Rsid() { Val = "00D6522A" };
            Rsid rsid512 = new Rsid() { Val = "00D65F6F" };
            Rsid rsid513 = new Rsid() { Val = "00D70F0E" };
            Rsid rsid514 = new Rsid() { Val = "00D80B9E" };
            Rsid rsid515 = new Rsid() { Val = "00D83EDC" };
            Rsid rsid516 = new Rsid() { Val = "00D928EF" };
            Rsid rsid517 = new Rsid() { Val = "00D944B2" };
            Rsid rsid518 = new Rsid() { Val = "00D948F5" };
            Rsid rsid519 = new Rsid() { Val = "00D96806" };
            Rsid rsid520 = new Rsid() { Val = "00DB4ADC" };
            Rsid rsid521 = new Rsid() { Val = "00DC3ED5" };
            Rsid rsid522 = new Rsid() { Val = "00DD1EE3" };
            Rsid rsid523 = new Rsid() { Val = "00DD5BAE" };
            Rsid rsid524 = new Rsid() { Val = "00DD718E" };
            Rsid rsid525 = new Rsid() { Val = "00DE0ED8" };
            Rsid rsid526 = new Rsid() { Val = "00DE1401" };
            Rsid rsid527 = new Rsid() { Val = "00DE345E" };
            Rsid rsid528 = new Rsid() { Val = "00DE549B" };
            Rsid rsid529 = new Rsid() { Val = "00DE5857" };
            Rsid rsid530 = new Rsid() { Val = "00DE78CB" };
            Rsid rsid531 = new Rsid() { Val = "00DF14E1" };
            Rsid rsid532 = new Rsid() { Val = "00DF50CA" };
            Rsid rsid533 = new Rsid() { Val = "00DF71F8" };
            Rsid rsid534 = new Rsid() { Val = "00DF7A1D" };
            Rsid rsid535 = new Rsid() { Val = "00E01429" };
            Rsid rsid536 = new Rsid() { Val = "00E0169F" };
            Rsid rsid537 = new Rsid() { Val = "00E0607E" };
            Rsid rsid538 = new Rsid() { Val = "00E12034" };
            Rsid rsid539 = new Rsid() { Val = "00E1250A" };
            Rsid rsid540 = new Rsid() { Val = "00E23707" };
            Rsid rsid541 = new Rsid() { Val = "00E23AC5" };
            Rsid rsid542 = new Rsid() { Val = "00E23B0A" };
            Rsid rsid543 = new Rsid() { Val = "00E24E56" };
            Rsid rsid544 = new Rsid() { Val = "00E27210" };
            Rsid rsid545 = new Rsid() { Val = "00E31290" };
            Rsid rsid546 = new Rsid() { Val = "00E3130B" };
            Rsid rsid547 = new Rsid() { Val = "00E31C6F" };
            Rsid rsid548 = new Rsid() { Val = "00E32064" };
            Rsid rsid549 = new Rsid() { Val = "00E340CC" };
            Rsid rsid550 = new Rsid() { Val = "00E461C6" };
            Rsid rsid551 = new Rsid() { Val = "00E51784" };
            Rsid rsid552 = new Rsid() { Val = "00E55B54" };
            Rsid rsid553 = new Rsid() { Val = "00E560D4" };
            Rsid rsid554 = new Rsid() { Val = "00E61216" };
            Rsid rsid555 = new Rsid() { Val = "00E61BC8" };
            Rsid rsid556 = new Rsid() { Val = "00E6525C" };
            Rsid rsid557 = new Rsid() { Val = "00E70E3A" };
            Rsid rsid558 = new Rsid() { Val = "00E714A6" };
            Rsid rsid559 = new Rsid() { Val = "00E745A0" };
            Rsid rsid560 = new Rsid() { Val = "00E75E67" };
            Rsid rsid561 = new Rsid() { Val = "00E76728" };
            Rsid rsid562 = new Rsid() { Val = "00E85D8A" };
            Rsid rsid563 = new Rsid() { Val = "00EB02A3" };
            Rsid rsid564 = new Rsid() { Val = "00EB135B" };
            Rsid rsid565 = new Rsid() { Val = "00EB4A0C" };
            Rsid rsid566 = new Rsid() { Val = "00EB61FB" };
            Rsid rsid567 = new Rsid() { Val = "00ED3794" };
            Rsid rsid568 = new Rsid() { Val = "00ED7CA9" };
            Rsid rsid569 = new Rsid() { Val = "00EE2379" };
            Rsid rsid570 = new Rsid() { Val = "00EE7B69" };
            Rsid rsid571 = new Rsid() { Val = "00F02BCF" };
            Rsid rsid572 = new Rsid() { Val = "00F04EF5" };
            Rsid rsid573 = new Rsid() { Val = "00F1393E" };
            Rsid rsid574 = new Rsid() { Val = "00F22E15" };
            Rsid rsid575 = new Rsid() { Val = "00F23038" };
            Rsid rsid576 = new Rsid() { Val = "00F235A9" };
            Rsid rsid577 = new Rsid() { Val = "00F31361" };
            Rsid rsid578 = new Rsid() { Val = "00F31DFE" };
            Rsid rsid579 = new Rsid() { Val = "00F32438" };
            Rsid rsid580 = new Rsid() { Val = "00F342A0" };
            Rsid rsid581 = new Rsid() { Val = "00F34666" };
            Rsid rsid582 = new Rsid() { Val = "00F378DE" };
            Rsid rsid583 = new Rsid() { Val = "00F40AED" };
            Rsid rsid584 = new Rsid() { Val = "00F43E8A" };
            Rsid rsid585 = new Rsid() { Val = "00F561B1" };
            Rsid rsid586 = new Rsid() { Val = "00F617B8" };
            Rsid rsid587 = new Rsid() { Val = "00F61A04" };
            Rsid rsid588 = new Rsid() { Val = "00F66B7D" };
            Rsid rsid589 = new Rsid() { Val = "00F7153C" };
            Rsid rsid590 = new Rsid() { Val = "00F723D6" };
            Rsid rsid591 = new Rsid() { Val = "00F77926" };
            Rsid rsid592 = new Rsid() { Val = "00F8047A" };
            Rsid rsid593 = new Rsid() { Val = "00F811E2" };
            Rsid rsid594 = new Rsid() { Val = "00F83E0C" };
            Rsid rsid595 = new Rsid() { Val = "00F85916" };
            Rsid rsid596 = new Rsid() { Val = "00F90086" };
            Rsid rsid597 = new Rsid() { Val = "00F97D85" };
            Rsid rsid598 = new Rsid() { Val = "00FA196D" };
            Rsid rsid599 = new Rsid() { Val = "00FA3E05" };
            Rsid rsid600 = new Rsid() { Val = "00FA773F" };
            Rsid rsid601 = new Rsid() { Val = "00FA79D4" };
            Rsid rsid602 = new Rsid() { Val = "00FB4196" };
            Rsid rsid603 = new Rsid() { Val = "00FB4EAB" };
            Rsid rsid604 = new Rsid() { Val = "00FB6AB0" };
            Rsid rsid605 = new Rsid() { Val = "00FB6E9A" };
            Rsid rsid606 = new Rsid() { Val = "00FC0F0D" };
            Rsid rsid607 = new Rsid() { Val = "00FC430A" };
            Rsid rsid608 = new Rsid() { Val = "00FC4A2D" };
            Rsid rsid609 = new Rsid() { Val = "00FC5AC6" };
            Rsid rsid610 = new Rsid() { Val = "00FD3FAA" };
            Rsid rsid611 = new Rsid() { Val = "00FD741A" };
            Rsid rsid612 = new Rsid() { Val = "00FE5962" };
            Rsid rsid613 = new Rsid() { Val = "00FF1CD4" };
            Rsid rsid614 = new Rsid() { Val = "00FF61CB" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);
            rsids1.Append(rsid13);
            rsids1.Append(rsid14);
            rsids1.Append(rsid15);
            rsids1.Append(rsid16);
            rsids1.Append(rsid17);
            rsids1.Append(rsid18);
            rsids1.Append(rsid19);
            rsids1.Append(rsid20);
            rsids1.Append(rsid21);
            rsids1.Append(rsid22);
            rsids1.Append(rsid23);
            rsids1.Append(rsid24);
            rsids1.Append(rsid25);
            rsids1.Append(rsid26);
            rsids1.Append(rsid27);
            rsids1.Append(rsid28);
            rsids1.Append(rsid29);
            rsids1.Append(rsid30);
            rsids1.Append(rsid31);
            rsids1.Append(rsid32);
            rsids1.Append(rsid33);
            rsids1.Append(rsid34);
            rsids1.Append(rsid35);
            rsids1.Append(rsid36);
            rsids1.Append(rsid37);
            rsids1.Append(rsid38);
            rsids1.Append(rsid39);
            rsids1.Append(rsid40);
            rsids1.Append(rsid41);
            rsids1.Append(rsid42);
            rsids1.Append(rsid43);
            rsids1.Append(rsid44);
            rsids1.Append(rsid45);
            rsids1.Append(rsid46);
            rsids1.Append(rsid47);
            rsids1.Append(rsid48);
            rsids1.Append(rsid49);
            rsids1.Append(rsid50);
            rsids1.Append(rsid51);
            rsids1.Append(rsid52);
            rsids1.Append(rsid53);
            rsids1.Append(rsid54);
            rsids1.Append(rsid55);
            rsids1.Append(rsid56);
            rsids1.Append(rsid57);
            rsids1.Append(rsid58);
            rsids1.Append(rsid59);
            rsids1.Append(rsid60);
            rsids1.Append(rsid61);
            rsids1.Append(rsid62);
            rsids1.Append(rsid63);
            rsids1.Append(rsid64);
            rsids1.Append(rsid65);
            rsids1.Append(rsid66);
            rsids1.Append(rsid67);
            rsids1.Append(rsid68);
            rsids1.Append(rsid69);
            rsids1.Append(rsid70);
            rsids1.Append(rsid71);
            rsids1.Append(rsid72);
            rsids1.Append(rsid73);
            rsids1.Append(rsid74);
            rsids1.Append(rsid75);
            rsids1.Append(rsid76);
            rsids1.Append(rsid77);
            rsids1.Append(rsid78);
            rsids1.Append(rsid79);
            rsids1.Append(rsid80);
            rsids1.Append(rsid81);
            rsids1.Append(rsid82);
            rsids1.Append(rsid83);
            rsids1.Append(rsid84);
            rsids1.Append(rsid85);
            rsids1.Append(rsid86);
            rsids1.Append(rsid87);
            rsids1.Append(rsid88);
            rsids1.Append(rsid89);
            rsids1.Append(rsid90);
            rsids1.Append(rsid91);
            rsids1.Append(rsid92);
            rsids1.Append(rsid93);
            rsids1.Append(rsid94);
            rsids1.Append(rsid95);
            rsids1.Append(rsid96);
            rsids1.Append(rsid97);
            rsids1.Append(rsid98);
            rsids1.Append(rsid99);
            rsids1.Append(rsid100);
            rsids1.Append(rsid101);
            rsids1.Append(rsid102);
            rsids1.Append(rsid103);
            rsids1.Append(rsid104);
            rsids1.Append(rsid105);
            rsids1.Append(rsid106);
            rsids1.Append(rsid107);
            rsids1.Append(rsid108);
            rsids1.Append(rsid109);
            rsids1.Append(rsid110);
            rsids1.Append(rsid111);
            rsids1.Append(rsid112);
            rsids1.Append(rsid113);
            rsids1.Append(rsid114);
            rsids1.Append(rsid115);
            rsids1.Append(rsid116);
            rsids1.Append(rsid117);
            rsids1.Append(rsid118);
            rsids1.Append(rsid119);
            rsids1.Append(rsid120);
            rsids1.Append(rsid121);
            rsids1.Append(rsid122);
            rsids1.Append(rsid123);
            rsids1.Append(rsid124);
            rsids1.Append(rsid125);
            rsids1.Append(rsid126);
            rsids1.Append(rsid127);
            rsids1.Append(rsid128);
            rsids1.Append(rsid129);
            rsids1.Append(rsid130);
            rsids1.Append(rsid131);
            rsids1.Append(rsid132);
            rsids1.Append(rsid133);
            rsids1.Append(rsid134);
            rsids1.Append(rsid135);
            rsids1.Append(rsid136);
            rsids1.Append(rsid137);
            rsids1.Append(rsid138);
            rsids1.Append(rsid139);
            rsids1.Append(rsid140);
            rsids1.Append(rsid141);
            rsids1.Append(rsid142);
            rsids1.Append(rsid143);
            rsids1.Append(rsid144);
            rsids1.Append(rsid145);
            rsids1.Append(rsid146);
            rsids1.Append(rsid147);
            rsids1.Append(rsid148);
            rsids1.Append(rsid149);
            rsids1.Append(rsid150);
            rsids1.Append(rsid151);
            rsids1.Append(rsid152);
            rsids1.Append(rsid153);
            rsids1.Append(rsid154);
            rsids1.Append(rsid155);
            rsids1.Append(rsid156);
            rsids1.Append(rsid157);
            rsids1.Append(rsid158);
            rsids1.Append(rsid159);
            rsids1.Append(rsid160);
            rsids1.Append(rsid161);
            rsids1.Append(rsid162);
            rsids1.Append(rsid163);
            rsids1.Append(rsid164);
            rsids1.Append(rsid165);
            rsids1.Append(rsid166);
            rsids1.Append(rsid167);
            rsids1.Append(rsid168);
            rsids1.Append(rsid169);
            rsids1.Append(rsid170);
            rsids1.Append(rsid171);
            rsids1.Append(rsid172);
            rsids1.Append(rsid173);
            rsids1.Append(rsid174);
            rsids1.Append(rsid175);
            rsids1.Append(rsid176);
            rsids1.Append(rsid177);
            rsids1.Append(rsid178);
            rsids1.Append(rsid179);
            rsids1.Append(rsid180);
            rsids1.Append(rsid181);
            rsids1.Append(rsid182);
            rsids1.Append(rsid183);
            rsids1.Append(rsid184);
            rsids1.Append(rsid185);
            rsids1.Append(rsid186);
            rsids1.Append(rsid187);
            rsids1.Append(rsid188);
            rsids1.Append(rsid189);
            rsids1.Append(rsid190);
            rsids1.Append(rsid191);
            rsids1.Append(rsid192);
            rsids1.Append(rsid193);
            rsids1.Append(rsid194);
            rsids1.Append(rsid195);
            rsids1.Append(rsid196);
            rsids1.Append(rsid197);
            rsids1.Append(rsid198);
            rsids1.Append(rsid199);
            rsids1.Append(rsid200);
            rsids1.Append(rsid201);
            rsids1.Append(rsid202);
            rsids1.Append(rsid203);
            rsids1.Append(rsid204);
            rsids1.Append(rsid205);
            rsids1.Append(rsid206);
            rsids1.Append(rsid207);
            rsids1.Append(rsid208);
            rsids1.Append(rsid209);
            rsids1.Append(rsid210);
            rsids1.Append(rsid211);
            rsids1.Append(rsid212);
            rsids1.Append(rsid213);
            rsids1.Append(rsid214);
            rsids1.Append(rsid215);
            rsids1.Append(rsid216);
            rsids1.Append(rsid217);
            rsids1.Append(rsid218);
            rsids1.Append(rsid219);
            rsids1.Append(rsid220);
            rsids1.Append(rsid221);
            rsids1.Append(rsid222);
            rsids1.Append(rsid223);
            rsids1.Append(rsid224);
            rsids1.Append(rsid225);
            rsids1.Append(rsid226);
            rsids1.Append(rsid227);
            rsids1.Append(rsid228);
            rsids1.Append(rsid229);
            rsids1.Append(rsid230);
            rsids1.Append(rsid231);
            rsids1.Append(rsid232);
            rsids1.Append(rsid233);
            rsids1.Append(rsid234);
            rsids1.Append(rsid235);
            rsids1.Append(rsid236);
            rsids1.Append(rsid237);
            rsids1.Append(rsid238);
            rsids1.Append(rsid239);
            rsids1.Append(rsid240);
            rsids1.Append(rsid241);
            rsids1.Append(rsid242);
            rsids1.Append(rsid243);
            rsids1.Append(rsid244);
            rsids1.Append(rsid245);
            rsids1.Append(rsid246);
            rsids1.Append(rsid247);
            rsids1.Append(rsid248);
            rsids1.Append(rsid249);
            rsids1.Append(rsid250);
            rsids1.Append(rsid251);
            rsids1.Append(rsid252);
            rsids1.Append(rsid253);
            rsids1.Append(rsid254);
            rsids1.Append(rsid255);
            rsids1.Append(rsid256);
            rsids1.Append(rsid257);
            rsids1.Append(rsid258);
            rsids1.Append(rsid259);
            rsids1.Append(rsid260);
            rsids1.Append(rsid261);
            rsids1.Append(rsid262);
            rsids1.Append(rsid263);
            rsids1.Append(rsid264);
            rsids1.Append(rsid265);
            rsids1.Append(rsid266);
            rsids1.Append(rsid267);
            rsids1.Append(rsid268);
            rsids1.Append(rsid269);
            rsids1.Append(rsid270);
            rsids1.Append(rsid271);
            rsids1.Append(rsid272);
            rsids1.Append(rsid273);
            rsids1.Append(rsid274);
            rsids1.Append(rsid275);
            rsids1.Append(rsid276);
            rsids1.Append(rsid277);
            rsids1.Append(rsid278);
            rsids1.Append(rsid279);
            rsids1.Append(rsid280);
            rsids1.Append(rsid281);
            rsids1.Append(rsid282);
            rsids1.Append(rsid283);
            rsids1.Append(rsid284);
            rsids1.Append(rsid285);
            rsids1.Append(rsid286);
            rsids1.Append(rsid287);
            rsids1.Append(rsid288);
            rsids1.Append(rsid289);
            rsids1.Append(rsid290);
            rsids1.Append(rsid291);
            rsids1.Append(rsid292);
            rsids1.Append(rsid293);
            rsids1.Append(rsid294);
            rsids1.Append(rsid295);
            rsids1.Append(rsid296);
            rsids1.Append(rsid297);
            rsids1.Append(rsid298);
            rsids1.Append(rsid299);
            rsids1.Append(rsid300);
            rsids1.Append(rsid301);
            rsids1.Append(rsid302);
            rsids1.Append(rsid303);
            rsids1.Append(rsid304);
            rsids1.Append(rsid305);
            rsids1.Append(rsid306);
            rsids1.Append(rsid307);
            rsids1.Append(rsid308);
            rsids1.Append(rsid309);
            rsids1.Append(rsid310);
            rsids1.Append(rsid311);
            rsids1.Append(rsid312);
            rsids1.Append(rsid313);
            rsids1.Append(rsid314);
            rsids1.Append(rsid315);
            rsids1.Append(rsid316);
            rsids1.Append(rsid317);
            rsids1.Append(rsid318);
            rsids1.Append(rsid319);
            rsids1.Append(rsid320);
            rsids1.Append(rsid321);
            rsids1.Append(rsid322);
            rsids1.Append(rsid323);
            rsids1.Append(rsid324);
            rsids1.Append(rsid325);
            rsids1.Append(rsid326);
            rsids1.Append(rsid327);
            rsids1.Append(rsid328);
            rsids1.Append(rsid329);
            rsids1.Append(rsid330);
            rsids1.Append(rsid331);
            rsids1.Append(rsid332);
            rsids1.Append(rsid333);
            rsids1.Append(rsid334);
            rsids1.Append(rsid335);
            rsids1.Append(rsid336);
            rsids1.Append(rsid337);
            rsids1.Append(rsid338);
            rsids1.Append(rsid339);
            rsids1.Append(rsid340);
            rsids1.Append(rsid341);
            rsids1.Append(rsid342);
            rsids1.Append(rsid343);
            rsids1.Append(rsid344);
            rsids1.Append(rsid345);
            rsids1.Append(rsid346);
            rsids1.Append(rsid347);
            rsids1.Append(rsid348);
            rsids1.Append(rsid349);
            rsids1.Append(rsid350);
            rsids1.Append(rsid351);
            rsids1.Append(rsid352);
            rsids1.Append(rsid353);
            rsids1.Append(rsid354);
            rsids1.Append(rsid355);
            rsids1.Append(rsid356);
            rsids1.Append(rsid357);
            rsids1.Append(rsid358);
            rsids1.Append(rsid359);
            rsids1.Append(rsid360);
            rsids1.Append(rsid361);
            rsids1.Append(rsid362);
            rsids1.Append(rsid363);
            rsids1.Append(rsid364);
            rsids1.Append(rsid365);
            rsids1.Append(rsid366);
            rsids1.Append(rsid367);
            rsids1.Append(rsid368);
            rsids1.Append(rsid369);
            rsids1.Append(rsid370);
            rsids1.Append(rsid371);
            rsids1.Append(rsid372);
            rsids1.Append(rsid373);
            rsids1.Append(rsid374);
            rsids1.Append(rsid375);
            rsids1.Append(rsid376);
            rsids1.Append(rsid377);
            rsids1.Append(rsid378);
            rsids1.Append(rsid379);
            rsids1.Append(rsid380);
            rsids1.Append(rsid381);
            rsids1.Append(rsid382);
            rsids1.Append(rsid383);
            rsids1.Append(rsid384);
            rsids1.Append(rsid385);
            rsids1.Append(rsid386);
            rsids1.Append(rsid387);
            rsids1.Append(rsid388);
            rsids1.Append(rsid389);
            rsids1.Append(rsid390);
            rsids1.Append(rsid391);
            rsids1.Append(rsid392);
            rsids1.Append(rsid393);
            rsids1.Append(rsid394);
            rsids1.Append(rsid395);
            rsids1.Append(rsid396);
            rsids1.Append(rsid397);
            rsids1.Append(rsid398);
            rsids1.Append(rsid399);
            rsids1.Append(rsid400);
            rsids1.Append(rsid401);
            rsids1.Append(rsid402);
            rsids1.Append(rsid403);
            rsids1.Append(rsid404);
            rsids1.Append(rsid405);
            rsids1.Append(rsid406);
            rsids1.Append(rsid407);
            rsids1.Append(rsid408);
            rsids1.Append(rsid409);
            rsids1.Append(rsid410);
            rsids1.Append(rsid411);
            rsids1.Append(rsid412);
            rsids1.Append(rsid413);
            rsids1.Append(rsid414);
            rsids1.Append(rsid415);
            rsids1.Append(rsid416);
            rsids1.Append(rsid417);
            rsids1.Append(rsid418);
            rsids1.Append(rsid419);
            rsids1.Append(rsid420);
            rsids1.Append(rsid421);
            rsids1.Append(rsid422);
            rsids1.Append(rsid423);
            rsids1.Append(rsid424);
            rsids1.Append(rsid425);
            rsids1.Append(rsid426);
            rsids1.Append(rsid427);
            rsids1.Append(rsid428);
            rsids1.Append(rsid429);
            rsids1.Append(rsid430);
            rsids1.Append(rsid431);
            rsids1.Append(rsid432);
            rsids1.Append(rsid433);
            rsids1.Append(rsid434);
            rsids1.Append(rsid435);
            rsids1.Append(rsid436);
            rsids1.Append(rsid437);
            rsids1.Append(rsid438);
            rsids1.Append(rsid439);
            rsids1.Append(rsid440);
            rsids1.Append(rsid441);
            rsids1.Append(rsid442);
            rsids1.Append(rsid443);
            rsids1.Append(rsid444);
            rsids1.Append(rsid445);
            rsids1.Append(rsid446);
            rsids1.Append(rsid447);
            rsids1.Append(rsid448);
            rsids1.Append(rsid449);
            rsids1.Append(rsid450);
            rsids1.Append(rsid451);
            rsids1.Append(rsid452);
            rsids1.Append(rsid453);
            rsids1.Append(rsid454);
            rsids1.Append(rsid455);
            rsids1.Append(rsid456);
            rsids1.Append(rsid457);
            rsids1.Append(rsid458);
            rsids1.Append(rsid459);
            rsids1.Append(rsid460);
            rsids1.Append(rsid461);
            rsids1.Append(rsid462);
            rsids1.Append(rsid463);
            rsids1.Append(rsid464);
            rsids1.Append(rsid465);
            rsids1.Append(rsid466);
            rsids1.Append(rsid467);
            rsids1.Append(rsid468);
            rsids1.Append(rsid469);
            rsids1.Append(rsid470);
            rsids1.Append(rsid471);
            rsids1.Append(rsid472);
            rsids1.Append(rsid473);
            rsids1.Append(rsid474);
            rsids1.Append(rsid475);
            rsids1.Append(rsid476);
            rsids1.Append(rsid477);
            rsids1.Append(rsid478);
            rsids1.Append(rsid479);
            rsids1.Append(rsid480);
            rsids1.Append(rsid481);
            rsids1.Append(rsid482);
            rsids1.Append(rsid483);
            rsids1.Append(rsid484);
            rsids1.Append(rsid485);
            rsids1.Append(rsid486);
            rsids1.Append(rsid487);
            rsids1.Append(rsid488);
            rsids1.Append(rsid489);
            rsids1.Append(rsid490);
            rsids1.Append(rsid491);
            rsids1.Append(rsid492);
            rsids1.Append(rsid493);
            rsids1.Append(rsid494);
            rsids1.Append(rsid495);
            rsids1.Append(rsid496);
            rsids1.Append(rsid497);
            rsids1.Append(rsid498);
            rsids1.Append(rsid499);
            rsids1.Append(rsid500);
            rsids1.Append(rsid501);
            rsids1.Append(rsid502);
            rsids1.Append(rsid503);
            rsids1.Append(rsid504);
            rsids1.Append(rsid505);
            rsids1.Append(rsid506);
            rsids1.Append(rsid507);
            rsids1.Append(rsid508);
            rsids1.Append(rsid509);
            rsids1.Append(rsid510);
            rsids1.Append(rsid511);
            rsids1.Append(rsid512);
            rsids1.Append(rsid513);
            rsids1.Append(rsid514);
            rsids1.Append(rsid515);
            rsids1.Append(rsid516);
            rsids1.Append(rsid517);
            rsids1.Append(rsid518);
            rsids1.Append(rsid519);
            rsids1.Append(rsid520);
            rsids1.Append(rsid521);
            rsids1.Append(rsid522);
            rsids1.Append(rsid523);
            rsids1.Append(rsid524);
            rsids1.Append(rsid525);
            rsids1.Append(rsid526);
            rsids1.Append(rsid527);
            rsids1.Append(rsid528);
            rsids1.Append(rsid529);
            rsids1.Append(rsid530);
            rsids1.Append(rsid531);
            rsids1.Append(rsid532);
            rsids1.Append(rsid533);
            rsids1.Append(rsid534);
            rsids1.Append(rsid535);
            rsids1.Append(rsid536);
            rsids1.Append(rsid537);
            rsids1.Append(rsid538);
            rsids1.Append(rsid539);
            rsids1.Append(rsid540);
            rsids1.Append(rsid541);
            rsids1.Append(rsid542);
            rsids1.Append(rsid543);
            rsids1.Append(rsid544);
            rsids1.Append(rsid545);
            rsids1.Append(rsid546);
            rsids1.Append(rsid547);
            rsids1.Append(rsid548);
            rsids1.Append(rsid549);
            rsids1.Append(rsid550);
            rsids1.Append(rsid551);
            rsids1.Append(rsid552);
            rsids1.Append(rsid553);
            rsids1.Append(rsid554);
            rsids1.Append(rsid555);
            rsids1.Append(rsid556);
            rsids1.Append(rsid557);
            rsids1.Append(rsid558);
            rsids1.Append(rsid559);
            rsids1.Append(rsid560);
            rsids1.Append(rsid561);
            rsids1.Append(rsid562);
            rsids1.Append(rsid563);
            rsids1.Append(rsid564);
            rsids1.Append(rsid565);
            rsids1.Append(rsid566);
            rsids1.Append(rsid567);
            rsids1.Append(rsid568);
            rsids1.Append(rsid569);
            rsids1.Append(rsid570);
            rsids1.Append(rsid571);
            rsids1.Append(rsid572);
            rsids1.Append(rsid573);
            rsids1.Append(rsid574);
            rsids1.Append(rsid575);
            rsids1.Append(rsid576);
            rsids1.Append(rsid577);
            rsids1.Append(rsid578);
            rsids1.Append(rsid579);
            rsids1.Append(rsid580);
            rsids1.Append(rsid581);
            rsids1.Append(rsid582);
            rsids1.Append(rsid583);
            rsids1.Append(rsid584);
            rsids1.Append(rsid585);
            rsids1.Append(rsid586);
            rsids1.Append(rsid587);
            rsids1.Append(rsid588);
            rsids1.Append(rsid589);
            rsids1.Append(rsid590);
            rsids1.Append(rsid591);
            rsids1.Append(rsid592);
            rsids1.Append(rsid593);
            rsids1.Append(rsid594);
            rsids1.Append(rsid595);
            rsids1.Append(rsid596);
            rsids1.Append(rsid597);
            rsids1.Append(rsid598);
            rsids1.Append(rsid599);
            rsids1.Append(rsid600);
            rsids1.Append(rsid601);
            rsids1.Append(rsid602);
            rsids1.Append(rsid603);
            rsids1.Append(rsid604);
            rsids1.Append(rsid605);
            rsids1.Append(rsid606);
            rsids1.Append(rsid607);
            rsids1.Append(rsid608);
            rsids1.Append(rsid609);
            rsids1.Append(rsid610);
            rsids1.Append(rsid611);
            rsids1.Append(rsid612);
            rsids1.Append(rsid613);
            rsids1.Append(rsid614);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Off };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            AttachedSchema attachedSchema1 = new AttachedSchema() { Val = "http://hubblereports.com/namespace" };
            AttachedSchema attachedSchema2 = new AttachedSchema() { Val = "errors@http://hubblereports.com/namespace" };
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "fr-CA" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults2 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 300034 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults2.Append(shapeDefaults3);
            shapeDefaults2.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "," };
            ListSeparator listSeparator1 = new ListSeparator() { Val = ";" };

            settings1.Append(zoom1);
            settings1.Append(attachedTemplate1);
            settings1.Append(linkStyles1);
            settings1.Append(defaultTabStop1);
            settings1.Append(hyphenationZone1);
            settings1.Append(noPunctuationKerning1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(doNotValidateAgainstSchema1);
            settings1.Append(saveInvalidXml1);
            settings1.Append(ignoreMixedContent1);
            settings1.Append(headerShapeDefaults1);
            settings1.Append(footnoteDocumentWideProperties1);
            settings1.Append(endnoteDocumentWideProperties1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(attachedSchema1);
            settings1.Append(attachedSchema2);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults2);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of imagePart4.
        private void GenerateImagePart4Content(ImagePart imagePart4)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart4Data);
            imagePart4.FeedData(data);
            data.Close();
        }

        // Generates content of headerPart1.
        private void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header();

            Paragraph paragraph35 = new Paragraph() { RsidParagraphAddition = "00AB4921", RsidParagraphProperties = "002A7539", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId33 = new ParagraphStyleId() { Val = "En-tte" };

            ParagraphBorders paragraphBorders3 = new ParagraphBorders();
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders3.Append(bottomBorder4);
            SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines() { After = "60" };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunStyle runStyle7 = new RunStyle() { Val = "DateCar" };
            FontSize fontSize10 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties15.Append(runStyle7);
            paragraphMarkRunProperties15.Append(fontSize10);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript19);

            paragraphProperties34.Append(paragraphStyleId33);
            paragraphProperties34.Append(paragraphBorders3);
            paragraphProperties34.Append(spacingBetweenLines20);
            paragraphProperties34.Append(paragraphMarkRunProperties15);

            Run run43 = new Run();

            RunProperties runProperties24 = new RunProperties();
            NoProof noProof12 = new NoProof();
            Languages languages12 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties24.Append(noProof12);
            runProperties24.Append(languages12);

            Drawing drawing12 = new Drawing();

            Wp.Anchor anchor2 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251658240U, BehindDoc = true, Locked = false, LayoutInCell = true, AllowOverlap = true };
            Wp.SimplePosition simplePosition2 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition2 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset3 = new Wp.PositionOffset();
            positionOffset3.Text = "0";

            horizontalPosition2.Append(positionOffset3);

            Wp.VerticalPosition verticalPosition2 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset4 = new Wp.PositionOffset();
            positionOffset4.Text = "0";

            verticalPosition2.Append(positionOffset4);
            Wp.Extent extent12 = new Wp.Extent() { Cx = 6858000L, Cy = 714375L };
            Wp.EffectExtent effectExtent12 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone2 = new Wp.WrapNone();
            Wp.DocProperties docProperties12 = new Wp.DocProperties() { Id = (UInt32Value)54U, Name = "Image 54", Description = "RADAR_Opinion_BNR" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties12 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks12 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties12.Append(graphicFrameLocks12);

            A.Graphic graphic12 = new A.Graphic();

            A.GraphicData graphicData12 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture12 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties12 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties12 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 54", Description = "RADAR_Opinion_BNR" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties12 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks12 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties12.Append(pictureLocks12);

            nonVisualPictureProperties12.Append(nonVisualDrawingProperties12);
            nonVisualPictureProperties12.Append(nonVisualPictureDrawingProperties12);

            Pic.BlipFill blipFill12 = new Pic.BlipFill();
            A.Blip blip12 = new A.Blip() { Embed = "rId1" };
            A.SourceRectangle sourceRectangle12 = new A.SourceRectangle();

            A.Stretch stretch12 = new A.Stretch();
            A.FillRectangle fillRectangle12 = new A.FillRectangle();

            stretch12.Append(fillRectangle12);

            blipFill12.Append(blip12);
            blipFill12.Append(sourceRectangle12);
            blipFill12.Append(stretch12);

            Pic.ShapeProperties shapeProperties12 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset12 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents12 = new A.Extents() { Cx = 6858000L, Cy = 714375L };

            transform2D12.Append(offset12);
            transform2D12.Append(extents12);

            A.PresetGeometry presetGeometry12 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList12 = new A.AdjustValueList();

            presetGeometry12.Append(adjustValueList12);
            A.NoFill noFill22 = new A.NoFill();

            shapeProperties12.Append(transform2D12);
            shapeProperties12.Append(presetGeometry12);
            shapeProperties12.Append(noFill22);

            picture12.Append(nonVisualPictureProperties12);
            picture12.Append(blipFill12);
            picture12.Append(shapeProperties12);

            graphicData12.Append(picture12);

            graphic12.Append(graphicData12);

            anchor2.Append(simplePosition2);
            anchor2.Append(horizontalPosition2);
            anchor2.Append(verticalPosition2);
            anchor2.Append(extent12);
            anchor2.Append(effectExtent12);
            anchor2.Append(wrapNone2);
            anchor2.Append(docProperties12);
            anchor2.Append(nonVisualGraphicFrameDrawingProperties12);
            anchor2.Append(graphic12);

            drawing12.Append(anchor2);

            run43.Append(runProperties24);
            run43.Append(drawing12);

            paragraph35.Append(paragraphProperties34);
            paragraph35.Append(run43);

            CustomXmlBlock customXmlBlock16 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "reportdoc" };

            CustomXmlBlock customXmlBlock17 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "header" };

            CustomXmlBlock customXmlBlock18 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "ReportDate" };

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "006B1D99", RsidParagraphAddition = "00AB4921", RsidParagraphProperties = "002A7539", RsidRunAdditionDefault = "00AB4921" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId34 = new ParagraphStyleId() { Val = "En-tte" };

            ParagraphBorders paragraphBorders4 = new ParagraphBorders();
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders4.Append(bottomBorder5);
            SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines() { Before = "240", After = "60" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunStyle runStyle8 = new RunStyle() { Val = "DateCar" };
            RunFonts runFonts4 = new RunFonts() { EastAsia = "MS Mincho" };
            FontSize fontSize11 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties16.Append(runStyle8);
            paragraphMarkRunProperties16.Append(runFonts4);
            paragraphMarkRunProperties16.Append(fontSize11);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript20);

            paragraphProperties35.Append(paragraphStyleId34);
            paragraphProperties35.Append(paragraphBorders4);
            paragraphProperties35.Append(spacingBetweenLines21);
            paragraphProperties35.Append(paragraphMarkRunProperties16);

            Run run44 = new Run() { RsidRunProperties = "006B1D99" };

            RunProperties runProperties25 = new RunProperties();
            RunStyle runStyle9 = new RunStyle() { Val = "DateCar" };
            RunFonts runFonts5 = new RunFonts() { EastAsia = "MS Mincho" };
            FontSize fontSize12 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "20" };

            runProperties25.Append(runStyle9);
            runProperties25.Append(runFonts5);
            runProperties25.Append(fontSize12);
            runProperties25.Append(fontSizeComplexScript21);
            Text text30 = new Text();
            text30.Text = "NOVEMBER 30, 2005";

            run44.Append(runProperties25);
            run44.Append(text30);

            paragraph36.Append(paragraphProperties35);
            paragraph36.Append(run44);

            customXmlBlock18.Append(paragraph36);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00A445CA", RsidParagraphAddition = "00AB4921", RsidParagraphProperties = "00E340CC", RsidRunAdditionDefault = "00957E57" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId35 = new ParagraphStyleId() { Val = "Titre" };
            Justification justification2 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties36.Append(paragraphStyleId35);
            paragraphProperties36.Append(justification2);
            CustomXmlRun customXmlRun9 = new CustomXmlRun() { Uri = "errors@http://hubblereports.com/namespace", Element = "ContentTypeDesc" };

            paragraph37.Append(paragraphProperties36);
            paragraph37.Append(customXmlRun9);

            CustomXmlBlock customXmlBlock19 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "ManagerName" };

            Paragraph paragraph38 = new Paragraph() { RsidParagraphAddition = "00AB4921", RsidParagraphProperties = "000C4F36", RsidRunAdditionDefault = "005546F4" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId36 = new ParagraphStyleId() { Val = "ManagerName" };
            SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines() { After = "0" };

            paragraphProperties37.Append(paragraphStyleId36);
            paragraphProperties37.Append(spacingBetweenLines22);

            Run run45 = new Run();
            Text text31 = new Text();
            text31.Text = "Morgan Stanley Investement Management, Inc.";

            run45.Append(text31);

            paragraph38.Append(paragraphProperties37);
            paragraph38.Append(run45);

            customXmlBlock19.Append(paragraph38);

            customXmlBlock17.Append(customXmlBlock18);
            customXmlBlock17.Append(paragraph37);
            customXmlBlock17.Append(customXmlBlock19);

            customXmlBlock16.Append(customXmlBlock17);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00545261", RsidRunAdditionDefault = "002E7D22" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines() { After = "0", Line = "40", LineRule = LineSpacingRuleValues.Exact };

            paragraphProperties38.Append(spacingBetweenLines23);

            paragraph39.Append(paragraphProperties38);

            header1.Append(paragraph35);
            header1.Append(customXmlBlock16);
            header1.Append(paragraph39);

            headerPart1.Header = header1;
        }

        // Generates content of imagePart5.
        private void GenerateImagePart5Content(ImagePart imagePart5)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart5Data);
            imagePart5.FeedData(data);
            data.Close();
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "MS Mincho", ComplexScript = "Times New Roman" };
            Languages languages13 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts6);
            runPropertiesBaseStyle1.Append(languages13);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 0, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleException latentStyleException1 = new LatentStyleException() { Name = "Normal", PrimaryStyle = true };
            LatentStyleException latentStyleException2 = new LatentStyleException() { Name = "heading 1", PrimaryStyle = true };
            LatentStyleException latentStyleException3 = new LatentStyleException() { Name = "heading 2", PrimaryStyle = true };
            LatentStyleException latentStyleException4 = new LatentStyleException() { Name = "heading 3", PrimaryStyle = true };
            LatentStyleException latentStyleException5 = new LatentStyleException() { Name = "heading 4", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException6 = new LatentStyleException() { Name = "heading 5", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException7 = new LatentStyleException() { Name = "heading 6", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException8 = new LatentStyleException() { Name = "heading 7", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException9 = new LatentStyleException() { Name = "heading 8", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException10 = new LatentStyleException() { Name = "heading 9", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException11 = new LatentStyleException() { Name = "caption", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleException latentStyleException12 = new LatentStyleException() { Name = "Title", PrimaryStyle = true };
            LatentStyleException latentStyleException13 = new LatentStyleException() { Name = "Subtitle", PrimaryStyle = true };
            LatentStyleException latentStyleException14 = new LatentStyleException() { Name = "Strong", PrimaryStyle = true };
            LatentStyleException latentStyleException15 = new LatentStyleException() { Name = "Emphasis", PrimaryStyle = true };
            LatentStyleException latentStyleException16 = new LatentStyleException() { Name = "Placeholder Text", UiPriority = 99, SemiHidden = true };
            LatentStyleException latentStyleException17 = new LatentStyleException() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleException latentStyleException18 = new LatentStyleException() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleException latentStyleException19 = new LatentStyleException() { Name = "Light List", UiPriority = 61 };
            LatentStyleException latentStyleException20 = new LatentStyleException() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleException latentStyleException21 = new LatentStyleException() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleException latentStyleException22 = new LatentStyleException() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleException latentStyleException23 = new LatentStyleException() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleException latentStyleException24 = new LatentStyleException() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleException latentStyleException25 = new LatentStyleException() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleException latentStyleException26 = new LatentStyleException() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleException latentStyleException27 = new LatentStyleException() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleException latentStyleException28 = new LatentStyleException() { Name = "Dark List", UiPriority = 70 };
            LatentStyleException latentStyleException29 = new LatentStyleException() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleException latentStyleException30 = new LatentStyleException() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleException latentStyleException31 = new LatentStyleException() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleException latentStyleException32 = new LatentStyleException() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleException latentStyleException33 = new LatentStyleException() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleException latentStyleException34 = new LatentStyleException() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleException latentStyleException35 = new LatentStyleException() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleException latentStyleException36 = new LatentStyleException() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleException latentStyleException37 = new LatentStyleException() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleException latentStyleException38 = new LatentStyleException() { Name = "Revision", UiPriority = 99, SemiHidden = true };
            LatentStyleException latentStyleException39 = new LatentStyleException() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleException latentStyleException40 = new LatentStyleException() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleException latentStyleException41 = new LatentStyleException() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleException latentStyleException42 = new LatentStyleException() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleException latentStyleException43 = new LatentStyleException() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleException latentStyleException44 = new LatentStyleException() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleException latentStyleException45 = new LatentStyleException() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleException latentStyleException46 = new LatentStyleException() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleException latentStyleException47 = new LatentStyleException() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleException latentStyleException48 = new LatentStyleException() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleException latentStyleException49 = new LatentStyleException() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleException latentStyleException50 = new LatentStyleException() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleException latentStyleException51 = new LatentStyleException() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleException latentStyleException52 = new LatentStyleException() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleException latentStyleException53 = new LatentStyleException() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleException latentStyleException54 = new LatentStyleException() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleException latentStyleException55 = new LatentStyleException() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleException latentStyleException56 = new LatentStyleException() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleException latentStyleException57 = new LatentStyleException() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleException latentStyleException58 = new LatentStyleException() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleException latentStyleException59 = new LatentStyleException() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleException latentStyleException60 = new LatentStyleException() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleException latentStyleException61 = new LatentStyleException() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleException latentStyleException62 = new LatentStyleException() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleException latentStyleException63 = new LatentStyleException() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleException latentStyleException64 = new LatentStyleException() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleException latentStyleException65 = new LatentStyleException() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleException latentStyleException66 = new LatentStyleException() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleException latentStyleException67 = new LatentStyleException() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleException latentStyleException68 = new LatentStyleException() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleException latentStyleException69 = new LatentStyleException() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleException latentStyleException70 = new LatentStyleException() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleException latentStyleException71 = new LatentStyleException() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleException latentStyleException72 = new LatentStyleException() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleException latentStyleException73 = new LatentStyleException() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleException latentStyleException74 = new LatentStyleException() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleException latentStyleException75 = new LatentStyleException() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleException latentStyleException76 = new LatentStyleException() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleException latentStyleException77 = new LatentStyleException() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleException latentStyleException78 = new LatentStyleException() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleException latentStyleException79 = new LatentStyleException() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleException latentStyleException80 = new LatentStyleException() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleException latentStyleException81 = new LatentStyleException() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleException latentStyleException82 = new LatentStyleException() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleException latentStyleException83 = new LatentStyleException() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleException latentStyleException84 = new LatentStyleException() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleException latentStyleException85 = new LatentStyleException() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleException latentStyleException86 = new LatentStyleException() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleException latentStyleException87 = new LatentStyleException() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleException latentStyleException88 = new LatentStyleException() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleException latentStyleException89 = new LatentStyleException() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleException latentStyleException90 = new LatentStyleException() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleException latentStyleException91 = new LatentStyleException() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleException latentStyleException92 = new LatentStyleException() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleException latentStyleException93 = new LatentStyleException() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleException latentStyleException94 = new LatentStyleException() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleException latentStyleException95 = new LatentStyleException() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleException latentStyleException96 = new LatentStyleException() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleException latentStyleException97 = new LatentStyleException() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleException latentStyleException98 = new LatentStyleException() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleException latentStyleException99 = new LatentStyleException() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleException latentStyleException100 = new LatentStyleException() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleException latentStyleException101 = new LatentStyleException() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleException latentStyleException102 = new LatentStyleException() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleException latentStyleException103 = new LatentStyleException() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleException latentStyleException104 = new LatentStyleException() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleException latentStyleException105 = new LatentStyleException() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleException latentStyleException106 = new LatentStyleException() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleException latentStyleException107 = new LatentStyleException() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleException latentStyleException108 = new LatentStyleException() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleException latentStyleException109 = new LatentStyleException() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleException latentStyleException110 = new LatentStyleException() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleException latentStyleException111 = new LatentStyleException() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleException latentStyleException112 = new LatentStyleException() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleException latentStyleException113 = new LatentStyleException() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleException latentStyleException114 = new LatentStyleException() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleException latentStyleException115 = new LatentStyleException() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleException latentStyleException116 = new LatentStyleException() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleException latentStyleException117 = new LatentStyleException() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleException latentStyleException118 = new LatentStyleException() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleException latentStyleException119 = new LatentStyleException() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleException latentStyleException120 = new LatentStyleException() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleException latentStyleException121 = new LatentStyleException() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleException latentStyleException122 = new LatentStyleException() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleException latentStyleException123 = new LatentStyleException() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleException latentStyleException124 = new LatentStyleException() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleException latentStyleException125 = new LatentStyleException() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleException latentStyleException126 = new LatentStyleException() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };

            latentStyles1.Append(latentStyleException1);
            latentStyles1.Append(latentStyleException2);
            latentStyles1.Append(latentStyleException3);
            latentStyles1.Append(latentStyleException4);
            latentStyles1.Append(latentStyleException5);
            latentStyles1.Append(latentStyleException6);
            latentStyles1.Append(latentStyleException7);
            latentStyles1.Append(latentStyleException8);
            latentStyles1.Append(latentStyleException9);
            latentStyles1.Append(latentStyleException10);
            latentStyles1.Append(latentStyleException11);
            latentStyles1.Append(latentStyleException12);
            latentStyles1.Append(latentStyleException13);
            latentStyles1.Append(latentStyleException14);
            latentStyles1.Append(latentStyleException15);
            latentStyles1.Append(latentStyleException16);
            latentStyles1.Append(latentStyleException17);
            latentStyles1.Append(latentStyleException18);
            latentStyles1.Append(latentStyleException19);
            latentStyles1.Append(latentStyleException20);
            latentStyles1.Append(latentStyleException21);
            latentStyles1.Append(latentStyleException22);
            latentStyles1.Append(latentStyleException23);
            latentStyles1.Append(latentStyleException24);
            latentStyles1.Append(latentStyleException25);
            latentStyles1.Append(latentStyleException26);
            latentStyles1.Append(latentStyleException27);
            latentStyles1.Append(latentStyleException28);
            latentStyles1.Append(latentStyleException29);
            latentStyles1.Append(latentStyleException30);
            latentStyles1.Append(latentStyleException31);
            latentStyles1.Append(latentStyleException32);
            latentStyles1.Append(latentStyleException33);
            latentStyles1.Append(latentStyleException34);
            latentStyles1.Append(latentStyleException35);
            latentStyles1.Append(latentStyleException36);
            latentStyles1.Append(latentStyleException37);
            latentStyles1.Append(latentStyleException38);
            latentStyles1.Append(latentStyleException39);
            latentStyles1.Append(latentStyleException40);
            latentStyles1.Append(latentStyleException41);
            latentStyles1.Append(latentStyleException42);
            latentStyles1.Append(latentStyleException43);
            latentStyles1.Append(latentStyleException44);
            latentStyles1.Append(latentStyleException45);
            latentStyles1.Append(latentStyleException46);
            latentStyles1.Append(latentStyleException47);
            latentStyles1.Append(latentStyleException48);
            latentStyles1.Append(latentStyleException49);
            latentStyles1.Append(latentStyleException50);
            latentStyles1.Append(latentStyleException51);
            latentStyles1.Append(latentStyleException52);
            latentStyles1.Append(latentStyleException53);
            latentStyles1.Append(latentStyleException54);
            latentStyles1.Append(latentStyleException55);
            latentStyles1.Append(latentStyleException56);
            latentStyles1.Append(latentStyleException57);
            latentStyles1.Append(latentStyleException58);
            latentStyles1.Append(latentStyleException59);
            latentStyles1.Append(latentStyleException60);
            latentStyles1.Append(latentStyleException61);
            latentStyles1.Append(latentStyleException62);
            latentStyles1.Append(latentStyleException63);
            latentStyles1.Append(latentStyleException64);
            latentStyles1.Append(latentStyleException65);
            latentStyles1.Append(latentStyleException66);
            latentStyles1.Append(latentStyleException67);
            latentStyles1.Append(latentStyleException68);
            latentStyles1.Append(latentStyleException69);
            latentStyles1.Append(latentStyleException70);
            latentStyles1.Append(latentStyleException71);
            latentStyles1.Append(latentStyleException72);
            latentStyles1.Append(latentStyleException73);
            latentStyles1.Append(latentStyleException74);
            latentStyles1.Append(latentStyleException75);
            latentStyles1.Append(latentStyleException76);
            latentStyles1.Append(latentStyleException77);
            latentStyles1.Append(latentStyleException78);
            latentStyles1.Append(latentStyleException79);
            latentStyles1.Append(latentStyleException80);
            latentStyles1.Append(latentStyleException81);
            latentStyles1.Append(latentStyleException82);
            latentStyles1.Append(latentStyleException83);
            latentStyles1.Append(latentStyleException84);
            latentStyles1.Append(latentStyleException85);
            latentStyles1.Append(latentStyleException86);
            latentStyles1.Append(latentStyleException87);
            latentStyles1.Append(latentStyleException88);
            latentStyles1.Append(latentStyleException89);
            latentStyles1.Append(latentStyleException90);
            latentStyles1.Append(latentStyleException91);
            latentStyles1.Append(latentStyleException92);
            latentStyles1.Append(latentStyleException93);
            latentStyles1.Append(latentStyleException94);
            latentStyles1.Append(latentStyleException95);
            latentStyles1.Append(latentStyleException96);
            latentStyles1.Append(latentStyleException97);
            latentStyles1.Append(latentStyleException98);
            latentStyles1.Append(latentStyleException99);
            latentStyles1.Append(latentStyleException100);
            latentStyles1.Append(latentStyleException101);
            latentStyles1.Append(latentStyleException102);
            latentStyles1.Append(latentStyleException103);
            latentStyles1.Append(latentStyleException104);
            latentStyles1.Append(latentStyleException105);
            latentStyles1.Append(latentStyleException106);
            latentStyles1.Append(latentStyleException107);
            latentStyles1.Append(latentStyleException108);
            latentStyles1.Append(latentStyleException109);
            latentStyles1.Append(latentStyleException110);
            latentStyles1.Append(latentStyleException111);
            latentStyles1.Append(latentStyleException112);
            latentStyles1.Append(latentStyleException113);
            latentStyles1.Append(latentStyleException114);
            latentStyles1.Append(latentStyleException115);
            latentStyles1.Append(latentStyleException116);
            latentStyles1.Append(latentStyleException117);
            latentStyles1.Append(latentStyleException118);
            latentStyles1.Append(latentStyleException119);
            latentStyles1.Append(latentStyleException120);
            latentStyles1.Append(latentStyleException121);
            latentStyles1.Append(latentStyleException122);
            latentStyles1.Append(latentStyleException123);
            latentStyles1.Append(latentStyleException124);
            latentStyles1.Append(latentStyleException125);
            latentStyles1.Append(latentStyleException126);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid615 = new Rsid() { Val = "006F57DE" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines24 = new SpacingBetweenLines() { After = "120" };

            styleParagraphProperties1.Append(spacingBetweenLines24);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman" };
            FontSize fontSize13 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "22" };
            Languages languages14 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties1.Append(runFonts7);
            styleRunProperties1.Append(fontSize13);
            styleRunProperties1.Append(fontSizeComplexScript22);
            styleRunProperties1.Append(languages14);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid615);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Titre1" };
            StyleName styleName2 = new StyleName() { Val = "heading 1" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Titre1Car" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid616 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext3 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines() { Before = "120", Line = "280", LineRule = LineSpacingRuleValues.Exact };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties2.Append(keepNext3);
            styleParagraphProperties2.Append(spacingBetweenLines25);
            styleParagraphProperties2.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman", ComplexScript = "Arial" };
            Bold bold13 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Caps caps2 = new Caps();
            Kern kern2 = new Kern() { Val = (UInt32Value)20U };
            Languages languages15 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties2.Append(runFonts8);
            styleRunProperties2.Append(bold13);
            styleRunProperties2.Append(boldComplexScript2);
            styleRunProperties2.Append(caps2);
            styleRunProperties2.Append(kern2);
            styleRunProperties2.Append(languages15);

            style2.Append(styleName2);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(primaryStyle2);
            style2.Append(rsid616);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "Titre2" };
            StyleName styleName3 = new StyleName() { Val = "heading 2" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "Normal" };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid617 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            KeepNext keepNext4 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines() { Before = "120", Line = "280", LineRule = LineSpacingRuleValues.Exact };
            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties3.Append(keepNext4);
            styleParagraphProperties3.Append(spacingBetweenLines26);
            styleParagraphProperties3.Append(outlineLevel2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman", ComplexScript = "Arial" };
            Bold bold14 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize14 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "22" };
            Languages languages16 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties3.Append(runFonts9);
            styleRunProperties3.Append(bold14);
            styleRunProperties3.Append(boldComplexScript3);
            styleRunProperties3.Append(italicComplexScript1);
            styleRunProperties3.Append(fontSize14);
            styleRunProperties3.Append(fontSizeComplexScript23);
            styleRunProperties3.Append(languages16);

            style3.Append(styleName3);
            style3.Append(nextParagraphStyle2);
            style3.Append(primaryStyle3);
            style3.Append(rsid617);
            style3.Append(styleParagraphProperties3);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() { Type = StyleValues.Paragraph, StyleId = "Titre3" };
            StyleName styleName4 = new StyleName() { Val = "heading 3" };
            BasedOn basedOn1 = new BasedOn() { Val = "Titre2" };
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "Normal" };
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid618 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            OutlineLevel outlineLevel3 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties4.Append(outlineLevel3);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            Bold bold15 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript4 = new BoldComplexScript() { Val = false };
            Italic italic1 = new Italic();
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties4.Append(bold15);
            styleRunProperties4.Append(boldComplexScript4);
            styleRunProperties4.Append(italic1);
            styleRunProperties4.Append(fontSizeComplexScript24);

            style4.Append(styleName4);
            style4.Append(basedOn1);
            style4.Append(nextParagraphStyle3);
            style4.Append(primaryStyle4);
            style4.Append(rsid618);
            style4.Append(styleParagraphProperties4);
            style4.Append(styleRunProperties4);

            Style style5 = new Style() { Type = StyleValues.Character, StyleId = "Policepardfaut", Default = true };
            StyleName styleName5 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style5.Append(styleName5);
            style5.Append(uIPriority1);
            style5.Append(unhideWhenUsed1);

            Style style6 = new Style() { Type = StyleValues.Table, StyleId = "TableauNormal", Default = true };
            StyleName styleName6 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle5 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin3 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin3 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault3.Append(topMargin1);
            tableCellMarginDefault3.Append(tableCellLeftMargin3);
            tableCellMarginDefault3.Append(bottomMargin1);
            tableCellMarginDefault3.Append(tableCellRightMargin3);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault3);

            style6.Append(styleName6);
            style6.Append(uIPriority2);
            style6.Append(semiHidden1);
            style6.Append(unhideWhenUsed2);
            style6.Append(primaryStyle5);
            style6.Append(styleTableProperties1);

            Style style7 = new Style() { Type = StyleValues.Numbering, StyleId = "Aucuneliste", Default = true };
            StyleName styleName7 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style7.Append(styleName7);
            style7.Append(uIPriority3);
            style7.Append(semiHidden2);
            style7.Append(unhideWhenUsed3);

            Style style8 = new Style() { Type = StyleValues.Paragraph, StyleId = "En-tte" };
            StyleName styleName8 = new StyleName() { Val = "header" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "En-tteCar" };
            Rsid rsid619 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders5 = new ParagraphBorders();
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "264C73", Size = (UInt32Value)48U, Space = (UInt32Value)1U };

            paragraphBorders5.Append(bottomBorder6);

            Tabs tabs2 = new Tabs();
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 10800 };

            tabs2.Append(tabStop2);
            SpacingBetweenLines spacingBetweenLines27 = new SpacingBetweenLines() { After = "0" };
            Justification justification3 = new Justification() { Val = JustificationValues.Right };

            styleParagraphProperties5.Append(paragraphBorders5);
            styleParagraphProperties5.Append(tabs2);
            styleParagraphProperties5.Append(spacingBetweenLines27);
            styleParagraphProperties5.Append(justification3);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            Caps caps3 = new Caps();
            Kern kern3 = new Kern() { Val = (UInt32Value)16U };
            FontSize fontSize15 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties5.Append(caps3);
            styleRunProperties5.Append(kern3);
            styleRunProperties5.Append(fontSize15);
            styleRunProperties5.Append(fontSizeComplexScript25);

            style8.Append(styleName8);
            style8.Append(basedOn2);
            style8.Append(linkedStyle2);
            style8.Append(rsid619);
            style8.Append(styleParagraphProperties5);
            style8.Append(styleRunProperties5);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "Pieddepage" };
            StyleName styleName9 = new StyleName() { Val = "footer" };
            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            Rsid rsid620 = new Rsid() { Val = "00233025" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines28 = new SpacingBetweenLines() { After = "0" };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            styleParagraphProperties6.Append(spacingBetweenLines28);
            styleParagraphProperties6.Append(justification4);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            FontSize fontSize16 = new FontSize() { Val = "16" };

            styleRunProperties6.Append(fontSize16);

            style9.Append(styleName9);
            style9.Append(basedOn3);
            style9.Append(rsid620);
            style9.Append(styleParagraphProperties6);
            style9.Append(styleRunProperties6);

            Style style10 = new Style() { Type = StyleValues.Table, StyleId = "Grilledutableau" };
            StyleName styleName10 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn4 = new BasedOn() { Val = "TableauNormal" };
            Rsid rsid621 = new Rsid() { Val = "00233025" };

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines29 = new SpacingBetweenLines() { After = "120", Line = "280", LineRule = LineSpacingRuleValues.Exact };

            styleParagraphProperties7.Append(spacingBetweenLines29);

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman" };

            styleRunProperties7.Append(runFonts10);

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders4 = new TableBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder4 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder4 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders4.Append(topBorder6);
            tableBorders4.Append(leftBorder4);
            tableBorders4.Append(bottomBorder7);
            tableBorders4.Append(rightBorder4);
            tableBorders4.Append(insideHorizontalBorder4);
            tableBorders4.Append(insideVerticalBorder4);

            TableCellMarginDefault tableCellMarginDefault4 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin4 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin4 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault4.Append(topMargin2);
            tableCellMarginDefault4.Append(tableCellLeftMargin4);
            tableCellMarginDefault4.Append(bottomMargin2);
            tableCellMarginDefault4.Append(tableCellRightMargin4);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableBorders4);
            styleTableProperties2.Append(tableCellMarginDefault4);

            style10.Append(styleName10);
            style10.Append(basedOn4);
            style10.Append(rsid621);
            style10.Append(styleParagraphProperties7);
            style10.Append(styleRunProperties7);
            style10.Append(styleTableProperties2);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "Numrodepage" };
            StyleName styleName11 = new StyleName() { Val = "page number" };
            BasedOn basedOn5 = new BasedOn() { Val = "Policepardfaut" };
            Rsid rsid622 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };

            styleRunProperties8.Append(runFonts11);

            style11.Append(styleName11);
            style11.Append(basedOn5);
            style11.Append(rsid622);
            style11.Append(styleRunProperties8);

            Style style12 = new Style() { Type = StyleValues.Paragraph, StyleId = "Listepuces" };
            StyleName styleName12 = new StyleName() { Val = "List Bullet" };
            BasedOn basedOn6 = new BasedOn() { Val = "Normal" };
            Rsid rsid623 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 360 };

            tabs3.Append(tabStop3);
            Indentation indentation1 = new Indentation() { Left = "360", Hanging = "360" };

            styleParagraphProperties8.Append(tabs3);
            styleParagraphProperties8.Append(indentation1);

            style12.Append(styleName12);
            style12.Append(basedOn6);
            style12.Append(rsid623);
            style12.Append(styleParagraphProperties8);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "Titre" };
            StyleName styleName13 = new StyleName() { Val = "Title" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "TitreCar" };
            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            Rsid rsid624 = new Rsid() { Val = "00233025" };

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();
            Justification justification5 = new Justification() { Val = JustificationValues.Right };
            OutlineLevel outlineLevel4 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties9.Append(justification5);
            styleParagraphProperties9.Append(outlineLevel4);

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman", ComplexScript = "Arial" };
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Color color1 = new Color() { Val = "264C73" };
            Kern kern4 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize17 = new FontSize() { Val = "46" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "36" };
            Languages languages17 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties9.Append(runFonts12);
            styleRunProperties9.Append(boldComplexScript5);
            styleRunProperties9.Append(color1);
            styleRunProperties9.Append(kern4);
            styleRunProperties9.Append(fontSize17);
            styleRunProperties9.Append(fontSizeComplexScript26);
            styleRunProperties9.Append(languages17);

            style13.Append(styleName13);
            style13.Append(linkedStyle3);
            style13.Append(primaryStyle6);
            style13.Append(rsid624);
            style13.Append(styleParagraphProperties9);
            style13.Append(styleRunProperties9);

            Style style14 = new Style() { Type = StyleValues.Paragraph, StyleId = "ManagerName", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "Manager Name" };
            Rsid rsid625 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines30 = new SpacingBetweenLines() { After = "40" };

            styleParagraphProperties10.Append(spacingBetweenLines30);

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman" };
            Bold bold16 = new Bold();
            FontSize fontSize18 = new FontSize() { Val = "36" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "36" };
            Languages languages18 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties10.Append(runFonts13);
            styleRunProperties10.Append(bold16);
            styleRunProperties10.Append(fontSize18);
            styleRunProperties10.Append(fontSizeComplexScript27);
            styleRunProperties10.Append(languages18);

            style14.Append(styleName14);
            style14.Append(rsid625);
            style14.Append(styleParagraphProperties10);
            style14.Append(styleRunProperties10);

            Style style15 = new Style() { Type = StyleValues.Paragraph, StyleId = "TableText", CustomStyle = true };
            StyleName styleName15 = new StyleName() { Val = "Table Text" };
            BasedOn basedOn7 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "TableTextChar" };
            Rsid rsid626 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines31 = new SpacingBetweenLines() { After = "0" };

            styleParagraphProperties11.Append(spacingBetweenLines31);

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            FontSize fontSize19 = new FontSize() { Val = "18" };

            styleRunProperties11.Append(fontSize19);

            style15.Append(styleName15);
            style15.Append(basedOn7);
            style15.Append(linkedStyle4);
            style15.Append(rsid626);
            style15.Append(styleParagraphProperties11);
            style15.Append(styleRunProperties11);

            Style style16 = new Style() { Type = StyleValues.Paragraph, StyleId = "ProductsReviewedHeading", CustomStyle = true };
            StyleName styleName16 = new StyleName() { Val = "Products Reviewed Heading" };
            BasedOn basedOn8 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "ProductsReviewedHeadingChar" };
            Rsid rsid627 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties12 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders6 = new ParagraphBorders();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "808080", Size = (UInt32Value)4U, Space = (UInt32Value)7U };

            paragraphBorders6.Append(topBorder7);
            SpacingBetweenLines spacingBetweenLines32 = new SpacingBetweenLines() { After = "140" };

            styleParagraphProperties12.Append(paragraphBorders6);
            styleParagraphProperties12.Append(spacingBetweenLines32);

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            Bold bold17 = new Bold();
            Caps caps4 = new Caps();

            styleRunProperties12.Append(bold17);
            styleRunProperties12.Append(caps4);

            style16.Append(styleName16);
            style16.Append(basedOn8);
            style16.Append(nextParagraphStyle4);
            style16.Append(linkedStyle5);
            style16.Append(rsid627);
            style16.Append(styleParagraphProperties12);
            style16.Append(styleRunProperties12);

            Style style17 = new Style() { Type = StyleValues.Paragraph, StyleId = "Date" };
            StyleName styleName17 = new StyleName() { Val = "Date" };
            BasedOn basedOn9 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "DateCar" };
            Rsid rsid628 = new Rsid() { Val = "00233025" };

            StyleParagraphProperties styleParagraphProperties13 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines33 = new SpacingBetweenLines() { After = "0" };
            Justification justification6 = new Justification() { Val = JustificationValues.Right };

            styleParagraphProperties13.Append(spacingBetweenLines33);
            styleParagraphProperties13.Append(justification6);

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            Caps caps5 = new Caps();
            Color color2 = new Color() { Val = "5C5C5C" };
            Kern kern5 = new Kern() { Val = (UInt32Value)22U };

            styleRunProperties13.Append(caps5);
            styleRunProperties13.Append(color2);
            styleRunProperties13.Append(kern5);

            style17.Append(styleName17);
            style17.Append(basedOn9);
            style17.Append(nextParagraphStyle5);
            style17.Append(linkedStyle6);
            style17.Append(rsid628);
            style17.Append(styleParagraphProperties13);
            style17.Append(styleRunProperties13);

            Style style18 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header2", CustomStyle = true };
            StyleName styleName18 = new StyleName() { Val = "Header 2" };
            BasedOn basedOn10 = new BasedOn() { Val = "En-tte" };
            Rsid rsid629 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties14 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders7 = new ParagraphBorders();
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders7.Append(bottomBorder8);

            styleParagraphProperties14.Append(paragraphBorders7);

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            FontSize fontSize20 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties14.Append(fontSize20);
            styleRunProperties14.Append(fontSizeComplexScript28);

            style18.Append(styleName18);
            style18.Append(basedOn10);
            style18.Append(rsid629);
            style18.Append(styleParagraphProperties14);
            style18.Append(styleRunProperties14);

            Style style19 = new Style() { Type = StyleValues.Paragraph, StyleId = "FooterPageNumber", CustomStyle = true };
            StyleName styleName19 = new StyleName() { Val = "Footer Page Number" };
            BasedOn basedOn11 = new BasedOn() { Val = "Pieddepage" };
            Rsid rsid630 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            FontSize fontSize21 = new FontSize() { Val = "20" };

            styleRunProperties15.Append(fontSize21);

            style19.Append(styleName19);
            style19.Append(basedOn11);
            style19.Append(rsid630);
            style19.Append(styleRunProperties15);

            Style style20 = new Style() { Type = StyleValues.Paragraph, StyleId = "Textedebulles" };
            StyleName styleName20 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn12 = new BasedOn() { Val = "Normal" };
            SemiHidden semiHidden3 = new SemiHidden();
            Rsid rsid631 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize22 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties16.Append(runFonts14);
            styleRunProperties16.Append(fontSize22);
            styleRunProperties16.Append(fontSizeComplexScript29);

            style20.Append(styleName20);
            style20.Append(basedOn12);
            style20.Append(semiHidden3);
            style20.Append(rsid631);
            style20.Append(styleRunProperties16);

            Style style21 = new Style() { Type = StyleValues.Paragraph, StyleId = "Disclaimer", CustomStyle = true };
            StyleName styleName21 = new StyleName() { Val = "Disclaimer" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "DisclaimerChar" };
            AutoRedefine autoRedefine1 = new AutoRedefine();
            Rsid rsid632 = new Rsid() { Val = "00233025" };

            StyleParagraphProperties styleParagraphProperties15 = new StyleParagraphProperties();
            KeepLines keepLines2 = new KeepLines();

            ParagraphBorders paragraphBorders8 = new ParagraphBorders();
            TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "264C73", Size = (UInt32Value)18U, Space = (UInt32Value)6U };

            paragraphBorders8.Append(topBorder8);
            SpacingBetweenLines spacingBetweenLines34 = new SpacingBetweenLines() { Before = "120", Line = "200", LineRule = LineSpacingRuleValues.Exact };

            styleParagraphProperties15.Append(keepLines2);
            styleParagraphProperties15.Append(paragraphBorders8);
            styleParagraphProperties15.Append(spacingBetweenLines34);

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman" };
            Color color3 = new Color() { Val = "808080" };
            FontSize fontSize23 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "22" };
            Languages languages19 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties17.Append(runFonts15);
            styleRunProperties17.Append(color3);
            styleRunProperties17.Append(fontSize23);
            styleRunProperties17.Append(fontSizeComplexScript30);
            styleRunProperties17.Append(languages19);

            style21.Append(styleName21);
            style21.Append(linkedStyle7);
            style21.Append(autoRedefine1);
            style21.Append(rsid632);
            style21.Append(styleParagraphProperties15);
            style21.Append(styleRunProperties17);

            Style style22 = new Style() { Type = StyleValues.Paragraph, StyleId = "TableHeading", CustomStyle = true };
            StyleName styleName22 = new StyleName() { Val = "Table Heading" };
            BasedOn basedOn13 = new BasedOn() { Val = "TableText" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "TableHeadingChar" };
            Rsid rsid633 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            Bold bold18 = new Bold();
            Caps caps6 = new Caps();
            Kern kern6 = new Kern() { Val = (UInt32Value)16U };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties18.Append(bold18);
            styleRunProperties18.Append(caps6);
            styleRunProperties18.Append(kern6);
            styleRunProperties18.Append(fontSizeComplexScript31);

            style22.Append(styleName22);
            style22.Append(basedOn13);
            style22.Append(linkedStyle8);
            style22.Append(rsid633);
            style22.Append(styleRunProperties18);

            Style style23 = new Style() { Type = StyleValues.Paragraph, StyleId = "HorizontalLine", CustomStyle = true };
            StyleName styleName23 = new StyleName() { Val = "Horizontal Line" };
            BasedOn basedOn14 = new BasedOn() { Val = "Normal" };
            Rsid rsid634 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties16 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders9 = new ParagraphBorders();
            BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "808080", Size = (UInt32Value)4U, Space = (UInt32Value)1U };

            paragraphBorders9.Append(bottomBorder9);
            SpacingBetweenLines spacingBetweenLines35 = new SpacingBetweenLines() { After = "240" };

            styleParagraphProperties16.Append(paragraphBorders9);
            styleParagraphProperties16.Append(spacingBetweenLines35);

            style23.Append(styleName23);
            style23.Append(basedOn14);
            style23.Append(rsid634);
            style23.Append(styleParagraphProperties16);

            Style style24 = new Style() { Type = StyleValues.Paragraph, StyleId = "FooterLogo", CustomStyle = true };
            StyleName styleName24 = new StyleName() { Val = "Footer Logo" };
            BasedOn basedOn15 = new BasedOn() { Val = "Pieddepage" };
            Rsid rsid635 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties17 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines36 = new SpacingBetweenLines() { Before = "120" };
            Justification justification7 = new Justification() { Val = JustificationValues.Right };

            styleParagraphProperties17.Append(spacingBetweenLines36);
            styleParagraphProperties17.Append(justification7);

            style24.Append(styleName24);
            style24.Append(basedOn15);
            style24.Append(rsid635);
            style24.Append(styleParagraphProperties17);

            Style style25 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header2ManagerName", CustomStyle = true };
            StyleName styleName25 = new StyleName() { Val = "Header 2 Manager Name" };
            BasedOn basedOn16 = new BasedOn() { Val = "Header2" };
            Rsid rsid636 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties18 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders10 = new ParagraphBorders();
            TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "264C73", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "264C73", Size = (UInt32Value)4U, Space = (UInt32Value)1U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "264C73", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "264C73", Size = (UInt32Value)4U, Space = (UInt32Value)1U };

            paragraphBorders10.Append(topBorder9);
            paragraphBorders10.Append(leftBorder5);
            paragraphBorders10.Append(bottomBorder10);
            paragraphBorders10.Append(rightBorder5);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "264C73" };
            SpacingBetweenLines spacingBetweenLines37 = new SpacingBetweenLines() { After = "60" };
            Justification justification8 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties18.Append(paragraphBorders10);
            styleParagraphProperties18.Append(shading1);
            styleParagraphProperties18.Append(spacingBetweenLines37);
            styleParagraphProperties18.Append(justification8);

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            Bold bold19 = new Bold();
            Color color4 = new Color() { Val = "FFFFFF" };

            styleRunProperties19.Append(bold19);
            styleRunProperties19.Append(color4);

            style25.Append(styleName25);
            style25.Append(basedOn16);
            style25.Append(rsid636);
            style25.Append(styleParagraphProperties18);
            style25.Append(styleRunProperties19);

            Style style26 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header2Title", CustomStyle = true };
            StyleName styleName26 = new StyleName() { Val = "Header 2 Title" };
            BasedOn basedOn17 = new BasedOn() { Val = "Titre" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "Header2TitleChar" };
            Rsid rsid637 = new Rsid() { Val = "00233025" };

            StyleParagraphProperties styleParagraphProperties19 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines38 = new SpacingBetweenLines() { After = "60" };

            styleParagraphProperties19.Append(spacingBetweenLines38);

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            FontSize fontSize24 = new FontSize() { Val = "26" };

            styleRunProperties20.Append(fontSize24);

            style26.Append(styleName26);
            style26.Append(basedOn17);
            style26.Append(linkedStyle9);
            style26.Append(rsid637);
            style26.Append(styleParagraphProperties19);
            style26.Append(styleRunProperties20);

            Style style27 = new Style() { Type = StyleValues.Paragraph, StyleId = "ProductName", CustomStyle = true };
            StyleName styleName27 = new StyleName() { Val = "Product Name" };
            Rsid rsid638 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties20 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders11 = new ParagraphBorders();
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "808080", Size = (UInt32Value)4U, Space = (UInt32Value)7U };

            paragraphBorders11.Append(bottomBorder11);
            SpacingBetweenLines spacingBetweenLines39 = new SpacingBetweenLines() { Before = "60", After = "240" };

            styleParagraphProperties20.Append(paragraphBorders11);
            styleParagraphProperties20.Append(spacingBetweenLines39);

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman" };
            Bold bold20 = new Bold();
            Caps caps7 = new Caps();
            Kern kern7 = new Kern() { Val = (UInt32Value)16U };
            FontSize fontSize25 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "22" };
            Languages languages20 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties21.Append(runFonts16);
            styleRunProperties21.Append(bold20);
            styleRunProperties21.Append(caps7);
            styleRunProperties21.Append(kern7);
            styleRunProperties21.Append(fontSize25);
            styleRunProperties21.Append(fontSizeComplexScript32);
            styleRunProperties21.Append(languages20);

            style27.Append(styleName27);
            style27.Append(rsid638);
            style27.Append(styleParagraphProperties20);
            style27.Append(styleRunProperties21);

            Style style28 = new Style() { Type = StyleValues.Paragraph, StyleId = "DislaimerHeading", CustomStyle = true };
            StyleName styleName28 = new StyleName() { Val = "Dislaimer Heading" };
            BasedOn basedOn18 = new BasedOn() { Val = "Disclaimer" };
            NextParagraphStyle nextParagraphStyle6 = new NextParagraphStyle() { Val = "Disclaimer" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "DislaimerHeadingChar" };
            AutoRedefine autoRedefine2 = new AutoRedefine();
            Rsid rsid639 = new Rsid() { Val = "00782598" };

            StyleParagraphProperties styleParagraphProperties21 = new StyleParagraphProperties();
            KeepNext keepNext5 = new KeepNext();

            styleParagraphProperties21.Append(keepNext5);

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            Bold bold21 = new Bold();

            styleRunProperties22.Append(bold21);

            style28.Append(styleName28);
            style28.Append(basedOn18);
            style28.Append(nextParagraphStyle6);
            style28.Append(linkedStyle10);
            style28.Append(autoRedefine2);
            style28.Append(rsid639);
            style28.Append(styleParagraphProperties21);
            style28.Append(styleRunProperties22);

            Style style29 = new Style() { Type = StyleValues.Character, StyleId = "DateCar", CustomStyle = true };
            StyleName styleName29 = new StyleName() { Val = "Date Car" };
            BasedOn basedOn19 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle11 = new LinkedStyle() { Val = "Date" };
            Rsid rsid640 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Caps caps8 = new Caps();
            Color color5 = new Color() { Val = "5C5C5C" };
            Kern kern8 = new Kern() { Val = (UInt32Value)22U };
            FontSize fontSize26 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "22" };
            Languages languages21 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties23.Append(runFonts17);
            styleRunProperties23.Append(caps8);
            styleRunProperties23.Append(color5);
            styleRunProperties23.Append(kern8);
            styleRunProperties23.Append(fontSize26);
            styleRunProperties23.Append(fontSizeComplexScript33);
            styleRunProperties23.Append(languages21);

            style29.Append(styleName29);
            style29.Append(basedOn19);
            style29.Append(linkedStyle11);
            style29.Append(rsid640);
            style29.Append(styleRunProperties23);

            Style style30 = new Style() { Type = StyleValues.Character, StyleId = "TitreCar", CustomStyle = true };
            StyleName styleName30 = new StyleName() { Val = "Titre Car" };
            BasedOn basedOn20 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle12 = new LinkedStyle() { Val = "Titre" };
            Rsid rsid641 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            Color color6 = new Color() { Val = "264C73" };
            Kern kern9 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize27 = new FontSize() { Val = "46" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "36" };
            Languages languages22 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties24.Append(runFonts18);
            styleRunProperties24.Append(boldComplexScript6);
            styleRunProperties24.Append(color6);
            styleRunProperties24.Append(kern9);
            styleRunProperties24.Append(fontSize27);
            styleRunProperties24.Append(fontSizeComplexScript34);
            styleRunProperties24.Append(languages22);

            style30.Append(styleName30);
            style30.Append(basedOn20);
            style30.Append(linkedStyle12);
            style30.Append(rsid641);
            style30.Append(styleRunProperties24);

            Style style31 = new Style() { Type = StyleValues.Character, StyleId = "Header2TitleChar", CustomStyle = true };
            StyleName styleName31 = new StyleName() { Val = "Header 2 Title Char" };
            BasedOn basedOn21 = new BasedOn() { Val = "TitreCar" };
            LinkedStyle linkedStyle13 = new LinkedStyle() { Val = "Header2Title" };
            Rsid rsid642 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            Color color7 = new Color() { Val = "264C73" };
            Kern kern10 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize28 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "36" };
            Languages languages23 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties25.Append(runFonts19);
            styleRunProperties25.Append(boldComplexScript7);
            styleRunProperties25.Append(color7);
            styleRunProperties25.Append(kern10);
            styleRunProperties25.Append(fontSize28);
            styleRunProperties25.Append(fontSizeComplexScript35);
            styleRunProperties25.Append(languages23);

            style31.Append(styleName31);
            style31.Append(basedOn21);
            style31.Append(linkedStyle13);
            style31.Append(rsid642);
            style31.Append(styleRunProperties25);

            Style style32 = new Style() { Type = StyleValues.Character, StyleId = "Titre1Car", CustomStyle = true };
            StyleName styleName32 = new StyleName() { Val = "Titre 1 Car" };
            BasedOn basedOn22 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle14 = new LinkedStyle() { Val = "Titre1" };
            Rsid rsid643 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold22 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            Caps caps9 = new Caps();
            Kern kern11 = new Kern() { Val = (UInt32Value)20U };
            Languages languages24 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties26.Append(runFonts20);
            styleRunProperties26.Append(bold22);
            styleRunProperties26.Append(boldComplexScript8);
            styleRunProperties26.Append(caps9);
            styleRunProperties26.Append(kern11);
            styleRunProperties26.Append(languages24);

            style32.Append(styleName32);
            style32.Append(basedOn22);
            style32.Append(linkedStyle14);
            style32.Append(rsid643);
            style32.Append(styleRunProperties26);

            Style style33 = new Style() { Type = StyleValues.Character, StyleId = "TableTextChar", CustomStyle = true };
            StyleName styleName33 = new StyleName() { Val = "Table Text Char" };
            BasedOn basedOn23 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle15 = new LinkedStyle() { Val = "TableText" };
            Rsid rsid644 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            FontSize fontSize29 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "22" };
            Languages languages25 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties27.Append(runFonts21);
            styleRunProperties27.Append(fontSize29);
            styleRunProperties27.Append(fontSizeComplexScript36);
            styleRunProperties27.Append(languages25);

            style33.Append(styleName33);
            style33.Append(basedOn23);
            style33.Append(linkedStyle15);
            style33.Append(rsid644);
            style33.Append(styleRunProperties27);

            Style style34 = new Style() { Type = StyleValues.Character, StyleId = "TableHeadingChar", CustomStyle = true };
            StyleName styleName34 = new StyleName() { Val = "Table Heading Char" };
            BasedOn basedOn24 = new BasedOn() { Val = "TableTextChar" };
            LinkedStyle linkedStyle16 = new LinkedStyle() { Val = "TableHeading" };
            Rsid rsid645 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Bold bold23 = new Bold();
            Caps caps10 = new Caps();
            Kern kern12 = new Kern() { Val = (UInt32Value)16U };
            FontSize fontSize30 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "18" };
            Languages languages26 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties28.Append(runFonts22);
            styleRunProperties28.Append(bold23);
            styleRunProperties28.Append(caps10);
            styleRunProperties28.Append(kern12);
            styleRunProperties28.Append(fontSize30);
            styleRunProperties28.Append(fontSizeComplexScript37);
            styleRunProperties28.Append(languages26);

            style34.Append(styleName34);
            style34.Append(basedOn24);
            style34.Append(linkedStyle16);
            style34.Append(rsid645);
            style34.Append(styleRunProperties28);

            Style style35 = new Style() { Type = StyleValues.Paragraph, StyleId = "Liste" };
            StyleName styleName35 = new StyleName() { Val = "List" };
            BasedOn basedOn25 = new BasedOn() { Val = "Normal" };
            Rsid rsid646 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties22 = new StyleParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 360 };

            tabs4.Append(tabStop4);

            styleParagraphProperties22.Append(tabs4);

            style35.Append(styleName35);
            style35.Append(basedOn25);
            style35.Append(rsid646);
            style35.Append(styleParagraphProperties22);

            Style style36 = new Style() { Type = StyleValues.Paragraph, StyleId = "Listepuces2" };
            StyleName styleName36 = new StyleName() { Val = "List Bullet 2" };
            BasedOn basedOn26 = new BasedOn() { Val = "Normal" };
            Rsid rsid647 = new Rsid() { Val = "00314DD5" };

            StyleParagraphProperties styleParagraphProperties23 = new StyleParagraphProperties();

            Tabs tabs5 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 360 };

            tabs5.Append(tabStop5);

            styleParagraphProperties23.Append(tabs5);

            style36.Append(styleName36);
            style36.Append(basedOn26);
            style36.Append(rsid647);
            style36.Append(styleParagraphProperties23);

            Style style37 = new Style() { Type = StyleValues.Character, StyleId = "En-tteCar", CustomStyle = true };
            StyleName styleName37 = new StyleName() { Val = "En-tête Car" };
            BasedOn basedOn27 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle17 = new LinkedStyle() { Val = "En-tte" };
            Rsid rsid648 = new Rsid() { Val = "00314DD5" };

            StyleRunProperties styleRunProperties29 = new StyleRunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Caps caps11 = new Caps();
            Kern kern13 = new Kern() { Val = (UInt32Value)16U };
            Languages languages27 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties29.Append(runFonts23);
            styleRunProperties29.Append(caps11);
            styleRunProperties29.Append(kern13);
            styleRunProperties29.Append(languages27);

            style37.Append(styleName37);
            style37.Append(basedOn27);
            style37.Append(linkedStyle17);
            style37.Append(rsid648);
            style37.Append(styleRunProperties29);

            Style style38 = new Style() { Type = StyleValues.Paragraph, StyleId = "RankStatement", CustomStyle = true };
            StyleName styleName38 = new StyleName() { Val = "Rank Statement" };
            BasedOn basedOn28 = new BasedOn() { Val = "TableText" };
            LinkedStyle linkedStyle18 = new LinkedStyle() { Val = "RankStatementChar" };
            AutoRedefine autoRedefine3 = new AutoRedefine();
            Rsid rsid649 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties30 = new StyleRunProperties();
            Bold bold24 = new Bold();
            Color color8 = new Color() { Val = "DD6600" };

            styleRunProperties30.Append(bold24);
            styleRunProperties30.Append(color8);

            style38.Append(styleName38);
            style38.Append(basedOn28);
            style38.Append(linkedStyle18);
            style38.Append(autoRedefine3);
            style38.Append(rsid649);
            style38.Append(styleRunProperties30);

            Style style39 = new Style() { Type = StyleValues.Character, StyleId = "RankStatementChar", CustomStyle = true };
            StyleName styleName39 = new StyleName() { Val = "Rank Statement Char" };
            BasedOn basedOn29 = new BasedOn() { Val = "TableTextChar" };
            LinkedStyle linkedStyle19 = new LinkedStyle() { Val = "RankStatement" };
            Rsid rsid650 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties31 = new StyleRunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Bold bold25 = new Bold();
            Color color9 = new Color() { Val = "DD6600" };
            FontSize fontSize31 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "22" };
            Languages languages28 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties31.Append(runFonts24);
            styleRunProperties31.Append(bold25);
            styleRunProperties31.Append(color9);
            styleRunProperties31.Append(fontSize31);
            styleRunProperties31.Append(fontSizeComplexScript38);
            styleRunProperties31.Append(languages28);

            style39.Append(styleName39);
            style39.Append(basedOn29);
            style39.Append(linkedStyle19);
            style39.Append(rsid650);
            style39.Append(styleRunProperties31);

            Style style40 = new Style() { Type = StyleValues.Paragraph, StyleId = "RankHeading", CustomStyle = true };
            StyleName styleName40 = new StyleName() { Val = "Rank Heading" };
            BasedOn basedOn30 = new BasedOn() { Val = "Titre1" };
            NextParagraphStyle nextParagraphStyle7 = new NextParagraphStyle() { Val = "Normal" };
            Rsid rsid651 = new Rsid() { Val = "00EE7B69" };

            StyleParagraphProperties styleParagraphProperties24 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines40 = new SpacingBetweenLines() { Before = "0", After = "120" };

            styleParagraphProperties24.Append(spacingBetweenLines40);

            style40.Append(styleName40);
            style40.Append(basedOn30);
            style40.Append(nextParagraphStyle7);
            style40.Append(rsid651);
            style40.Append(styleParagraphProperties24);

            Style style41 = new Style() { Type = StyleValues.Character, StyleId = "CategoryRankGraphic", CustomStyle = true };
            StyleName styleName41 = new StyleName() { Val = "Category Rank Graphic" };
            BasedOn basedOn31 = new BasedOn() { Val = "Policepardfaut" };
            Rsid rsid652 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties32 = new StyleRunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Position position2 = new Position() { Val = "-4" };

            styleRunProperties32.Append(runFonts25);
            styleRunProperties32.Append(position2);

            style41.Append(styleName41);
            style41.Append(basedOn31);
            style41.Append(rsid652);
            style41.Append(styleRunProperties32);

            Style style42 = new Style() { Type = StyleValues.Paragraph, StyleId = "FooterRankLegend", CustomStyle = true };
            StyleName styleName42 = new StyleName() { Val = "Footer Rank Legend" };
            BasedOn basedOn32 = new BasedOn() { Val = "Normal" };
            Rsid rsid653 = new Rsid() { Val = "003F2779" };

            StyleParagraphProperties styleParagraphProperties25 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines41 = new SpacingBetweenLines() { Before = "360" };

            styleParagraphProperties25.Append(spacingBetweenLines41);

            style42.Append(styleName42);
            style42.Append(basedOn32);
            style42.Append(rsid653);
            style42.Append(styleParagraphProperties25);

            Style style43 = new Style() { Type = StyleValues.Character, StyleId = "DisclaimerChar", CustomStyle = true };
            StyleName styleName43 = new StyleName() { Val = "Disclaimer Char" };
            BasedOn basedOn33 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle20 = new LinkedStyle() { Val = "Disclaimer" };
            Rsid rsid654 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties33 = new StyleRunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Color color10 = new Color() { Val = "808080" };
            FontSize fontSize32 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "22" };
            Languages languages29 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties33.Append(runFonts26);
            styleRunProperties33.Append(color10);
            styleRunProperties33.Append(fontSize32);
            styleRunProperties33.Append(fontSizeComplexScript39);
            styleRunProperties33.Append(languages29);

            style43.Append(styleName43);
            style43.Append(basedOn33);
            style43.Append(linkedStyle20);
            style43.Append(rsid654);
            style43.Append(styleRunProperties33);

            Style style44 = new Style() { Type = StyleValues.Character, StyleId = "DislaimerHeadingChar", CustomStyle = true };
            StyleName styleName44 = new StyleName() { Val = "Dislaimer Heading Char" };
            BasedOn basedOn34 = new BasedOn() { Val = "DisclaimerChar" };
            LinkedStyle linkedStyle21 = new LinkedStyle() { Val = "DislaimerHeading" };
            Rsid rsid655 = new Rsid() { Val = "00782598" };

            StyleRunProperties styleRunProperties34 = new StyleRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Bold bold26 = new Bold();
            Color color11 = new Color() { Val = "808080" };
            FontSize fontSize33 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "22" };
            Languages languages30 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties34.Append(runFonts27);
            styleRunProperties34.Append(bold26);
            styleRunProperties34.Append(color11);
            styleRunProperties34.Append(fontSize33);
            styleRunProperties34.Append(fontSizeComplexScript40);
            styleRunProperties34.Append(languages30);

            style44.Append(styleName44);
            style44.Append(basedOn34);
            style44.Append(linkedStyle21);
            style44.Append(rsid655);
            style44.Append(styleRunProperties34);

            Style style45 = new Style() { Type = StyleValues.Character, StyleId = "ProductsReviewedHeadingChar", CustomStyle = true };
            StyleName styleName45 = new StyleName() { Val = "Products Reviewed Heading Char" };
            BasedOn basedOn35 = new BasedOn() { Val = "Policepardfaut" };
            LinkedStyle linkedStyle22 = new LinkedStyle() { Val = "ProductsReviewedHeading" };
            Rsid rsid656 = new Rsid() { Val = "00443CD0" };

            StyleRunProperties styleRunProperties35 = new StyleRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Bold bold27 = new Bold();
            Caps caps12 = new Caps();
            FontSize fontSize34 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "22" };
            Languages languages31 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            styleRunProperties35.Append(runFonts28);
            styleRunProperties35.Append(bold27);
            styleRunProperties35.Append(caps12);
            styleRunProperties35.Append(fontSize34);
            styleRunProperties35.Append(fontSizeComplexScript41);
            styleRunProperties35.Append(languages31);

            style45.Append(styleName45);
            style45.Append(basedOn35);
            style45.Append(linkedStyle22);
            style45.Append(rsid656);
            style45.Append(styleRunProperties35);

            Style style46 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleAfter0pt", CustomStyle = true };
            StyleName styleName46 = new StyleName() { Val = "Style After:  0 pt" };
            BasedOn basedOn36 = new BasedOn() { Val = "Normal" };
            Rsid rsid657 = new Rsid() { Val = "00983F27" };

            StyleParagraphProperties styleParagraphProperties26 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines42 = new SpacingBetweenLines() { After = "0" };

            styleParagraphProperties26.Append(spacingBetweenLines42);

            StyleRunProperties styleRunProperties36 = new StyleRunProperties();
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties36.Append(fontSizeComplexScript42);

            style46.Append(styleName46);
            style46.Append(basedOn36);
            style46.Append(rsid657);
            style46.Append(styleParagraphProperties26);
            style46.Append(styleRunProperties36);

            Style style47 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleHeading2Before0ptAfter2pt", CustomStyle = true };
            StyleName styleName47 = new StyleName() { Val = "Style Heading 2 + Before:  0 pt After:  2 pt" };
            BasedOn basedOn37 = new BasedOn() { Val = "Titre2" };
            Rsid rsid658 = new Rsid() { Val = "00AC1437" };

            StyleParagraphProperties styleParagraphProperties27 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines43 = new SpacingBetweenLines() { Before = "0", After = "40", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties27.Append(spacingBetweenLines43);

            StyleRunProperties styleRunProperties37 = new StyleRunProperties();
            RunFonts runFonts29 = new RunFonts() { ComplexScript = "Times New Roman" };
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript() { Val = false };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties37.Append(runFonts29);
            styleRunProperties37.Append(italicComplexScript2);
            styleRunProperties37.Append(fontSizeComplexScript43);

            style47.Append(styleName47);
            style47.Append(basedOn37);
            style47.Append(rsid658);
            style47.Append(styleParagraphProperties27);
            style47.Append(styleRunProperties37);

            Style style48 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleProductsReviewedHeadingBefore12pt", CustomStyle = true };
            StyleName styleName48 = new StyleName() { Val = "Style Products Reviewed Heading + Before:  12 pt" };
            BasedOn basedOn38 = new BasedOn() { Val = "ProductsReviewedHeading" };
            Rsid rsid659 = new Rsid() { Val = "009F7E7F" };

            StyleParagraphProperties styleParagraphProperties28 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines44 = new SpacingBetweenLines() { Before = "360" };

            styleParagraphProperties28.Append(spacingBetweenLines44);

            StyleRunProperties styleRunProperties38 = new StyleRunProperties();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties38.Append(boldComplexScript9);
            styleRunProperties38.Append(fontSizeComplexScript44);

            style48.Append(styleName48);
            style48.Append(basedOn38);
            style48.Append(rsid659);
            style48.Append(styleParagraphProperties28);
            style48.Append(styleRunProperties38);

            Style style49 = new Style() { Type = StyleValues.Paragraph, StyleId = "NumberedList", CustomStyle = true };
            StyleName styleName49 = new StyleName() { Val = "Numbered List" };
            BasedOn basedOn39 = new BasedOn() { Val = "Normal" };
            Rsid rsid660 = new Rsid() { Val = "00BB632E" };

            StyleParagraphProperties styleParagraphProperties29 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingId numberingId1 = new NumberingId() { Val = 2 };

            numberingProperties1.Append(numberingId1);

            styleParagraphProperties29.Append(numberingProperties1);

            style49.Append(styleName49);
            style49.Append(basedOn39);
            style49.Append(rsid660);
            style49.Append(styleParagraphProperties29);

            Style style50 = new Style() { Type = StyleValues.Paragraph, StyleId = "Explorateurdedocuments" };
            StyleName styleName50 = new StyleName() { Val = "Document Map" };
            BasedOn basedOn40 = new BasedOn() { Val = "Normal" };
            SemiHidden semiHidden4 = new SemiHidden();
            Rsid rsid661 = new Rsid() { Val = "002A7539" };

            StyleParagraphProperties styleParagraphProperties30 = new StyleParagraphProperties();
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "000080" };

            styleParagraphProperties30.Append(shading2);

            StyleRunProperties styleRunProperties39 = new StyleRunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize35 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties39.Append(runFonts30);
            styleRunProperties39.Append(fontSize35);
            styleRunProperties39.Append(fontSizeComplexScript45);

            style50.Append(styleName50);
            style50.Append(basedOn40);
            style50.Append(semiHidden4);
            style50.Append(rsid661);
            style50.Append(styleParagraphProperties30);
            style50.Append(styleRunProperties39);

            Style style51 = new Style() { Type = StyleValues.Character, StyleId = "Style10ptBold", CustomStyle = true };
            StyleName styleName51 = new StyleName() { Val = "Style 10 pt Bold" };
            BasedOn basedOn41 = new BasedOn() { Val = "Policepardfaut" };
            Rsid rsid662 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties40 = new StyleRunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Bold bold28 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            FontSize fontSize36 = new FontSize() { Val = "20" };

            styleRunProperties40.Append(runFonts31);
            styleRunProperties40.Append(bold28);
            styleRunProperties40.Append(boldComplexScript10);
            styleRunProperties40.Append(fontSize36);

            style51.Append(styleName51);
            style51.Append(basedOn41);
            style51.Append(rsid662);
            style51.Append(styleRunProperties40);

            Style style52 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleBefore9ptAfter0pt", CustomStyle = true };
            StyleName styleName52 = new StyleName() { Val = "Style Before:  9 pt After:  0 pt" };
            BasedOn basedOn42 = new BasedOn() { Val = "Normal" };
            AutoRedefine autoRedefine4 = new AutoRedefine();
            Rsid rsid663 = new Rsid() { Val = "00DD5BAE" };

            StyleParagraphProperties styleParagraphProperties31 = new StyleParagraphProperties();
            KeepNext keepNext6 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines45 = new SpacingBetweenLines() { Before = "180", After = "0" };

            styleParagraphProperties31.Append(keepNext6);
            styleParagraphProperties31.Append(spacingBetweenLines45);

            StyleRunProperties styleRunProperties41 = new StyleRunProperties();
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties41.Append(fontSizeComplexScript46);

            style52.Append(styleName52);
            style52.Append(basedOn42);
            style52.Append(autoRedefine4);
            style52.Append(rsid663);
            style52.Append(styleParagraphProperties31);
            style52.Append(styleRunProperties41);

            Style style53 = new Style() { Type = StyleValues.Character, StyleId = "StyleBodoniMT", CustomStyle = true };
            StyleName styleName53 = new StyleName() { Val = "Style Bodoni MT" };
            BasedOn basedOn43 = new BasedOn() { Val = "Policepardfaut" };
            Rsid rsid664 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties42 = new StyleRunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };

            styleRunProperties42.Append(runFonts32);

            style53.Append(styleName53);
            style53.Append(basedOn43);
            style53.Append(rsid664);
            style53.Append(styleRunProperties42);

            Style style54 = new Style() { Type = StyleValues.Character, StyleId = "StyleCategoryRankGraphic10pt", CustomStyle = true };
            StyleName styleName54 = new StyleName() { Val = "Style Category Rank Graphic + 10 pt" };
            BasedOn basedOn44 = new BasedOn() { Val = "CategoryRankGraphic" };
            Rsid rsid665 = new Rsid() { Val = "00233025" };

            StyleRunProperties styleRunProperties43 = new StyleRunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" };
            Position position3 = new Position() { Val = "0" };
            FontSize fontSize37 = new FontSize() { Val = "20" };

            styleRunProperties43.Append(runFonts33);
            styleRunProperties43.Append(position3);
            styleRunProperties43.Append(fontSize37);

            style54.Append(styleName54);
            style54.Append(basedOn44);
            style54.Append(rsid665);
            style54.Append(styleRunProperties43);

            Style style55 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleProductNameBefore0ptAfter8pt", CustomStyle = true };
            StyleName styleName55 = new StyleName() { Val = "Style Product Name + Before:  0 pt After:  8 pt" };
            BasedOn basedOn45 = new BasedOn() { Val = "ProductName" };
            AutoRedefine autoRedefine5 = new AutoRedefine();
            Rsid rsid666 = new Rsid() { Val = "00397A9E" };

            StyleParagraphProperties styleParagraphProperties32 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines46 = new SpacingBetweenLines() { Before = "0", After = "160" };

            styleParagraphProperties32.Append(spacingBetweenLines46);

            StyleRunProperties styleRunProperties44 = new StyleRunProperties();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties44.Append(boldComplexScript11);
            styleRunProperties44.Append(fontSizeComplexScript47);

            style55.Append(styleName55);
            style55.Append(basedOn45);
            style55.Append(autoRedefine5);
            style55.Append(rsid666);
            style55.Append(styleParagraphProperties32);
            style55.Append(styleRunProperties44);

            Style style56 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleProductsReviewedHeading6ptBefore15ptAfter0pt", CustomStyle = true };
            StyleName styleName56 = new StyleName() { Val = "Style Products Reviewed Heading + 6 pt Before:  15 pt After:  0 pt" };
            BasedOn basedOn46 = new BasedOn() { Val = "ProductsReviewedHeading" };
            AutoRedefine autoRedefine6 = new AutoRedefine();
            Rsid rsid667 = new Rsid() { Val = "00397A9E" };

            StyleParagraphProperties styleParagraphProperties33 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines47 = new SpacingBetweenLines() { Before = "300", After = "0" };

            styleParagraphProperties33.Append(spacingBetweenLines47);

            StyleRunProperties styleRunProperties45 = new StyleRunProperties();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            FontSize fontSize38 = new FontSize() { Val = "12" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties45.Append(boldComplexScript12);
            styleRunProperties45.Append(fontSize38);
            styleRunProperties45.Append(fontSizeComplexScript48);

            style56.Append(styleName56);
            style56.Append(basedOn46);
            style56.Append(autoRedefine6);
            style56.Append(rsid667);
            style56.Append(styleParagraphProperties33);
            style56.Append(styleRunProperties45);

            Style style57 = new Style() { Type = StyleValues.Paragraph, StyleId = "StyleProductsReviewedHeading4ptBefore15ptAfter0pt", CustomStyle = true };
            StyleName styleName57 = new StyleName() { Val = "Style Products Reviewed Heading + 4 pt Before:  15 pt After:  0 pt" };
            BasedOn basedOn47 = new BasedOn() { Val = "ProductsReviewedHeading" };
            AutoRedefine autoRedefine7 = new AutoRedefine();
            Rsid rsid668 = new Rsid() { Val = "00397A9E" };

            StyleParagraphProperties styleParagraphProperties34 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines48 = new SpacingBetweenLines() { Before = "300", After = "0" };

            styleParagraphProperties34.Append(spacingBetweenLines48);

            StyleRunProperties styleRunProperties46 = new StyleRunProperties();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            FontSize fontSize39 = new FontSize() { Val = "8" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties46.Append(boldComplexScript13);
            styleRunProperties46.Append(fontSize39);
            styleRunProperties46.Append(fontSizeComplexScript49);

            style57.Append(styleName57);
            style57.Append(basedOn47);
            style57.Append(autoRedefine7);
            style57.Append(rsid668);
            style57.Append(styleParagraphProperties34);
            style57.Append(styleRunProperties46);

            Style styleUnorderedList = new Style() { Type = StyleValues.Paragraph, StyleId = "UnorderedListStyle", CustomStyle = true };
            StyleName styleNameUnorderedList = new StyleName() { Val = "UnorderedList Style" };
            BasedOn basedOnUnorderedList = new BasedOn() { Val = "Normal" };

            StyleParagraphProperties styleParagraphPropertiesUnorderedList = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLinesUnorderedList = new SpacingBetweenLines() { After = "0" };

            styleParagraphPropertiesUnorderedList.Append(spacingBetweenLinesUnorderedList);

            StyleRunProperties styleRunPropertiesUnorderedList = new StyleRunProperties();
            FontSizeComplexScript fontSizeComplexScriptUnorderedList = new FontSizeComplexScript() { Val = "20" };

            styleRunPropertiesUnorderedList.Append(fontSizeComplexScriptUnorderedList);

            styleUnorderedList.Append(styleNameUnorderedList);
            styleUnorderedList.Append(basedOnUnorderedList);
            styleUnorderedList.Append(styleParagraphPropertiesUnorderedList);
            styleUnorderedList.Append(styleRunPropertiesUnorderedList);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);
            styles1.Append(style25);
            styles1.Append(style26);
            styles1.Append(style27);
            styles1.Append(style28);
            styles1.Append(style29);
            styles1.Append(style30);
            styles1.Append(style31);
            styles1.Append(style32);
            styles1.Append(style33);
            styles1.Append(style34);
            styles1.Append(style35);
            styles1.Append(style36);
            styles1.Append(style37);
            styles1.Append(style38);
            styles1.Append(style39);
            styles1.Append(style40);
            styles1.Append(style41);
            styles1.Append(style42);
            styles1.Append(style43);
            styles1.Append(style44);
            styles1.Append(style45);
            styles1.Append(style46);
            styles1.Append(style47);
            styles1.Append(style48);
            styles1.Append(style49);
            styles1.Append(style50);
            styles1.Append(style51);
            styles1.Append(style52);
            styles1.Append(style53);
            styles1.Append(style54);
            styles1.Append(style55);
            styles1.Append(style56);
            styles1.Append(style57);
            styles1.Append(styleUnorderedList);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering();

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            Nsid nsid1 = new Nsid() { Val = "2BE110B5" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "263EA1D2" };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            ParagraphStyleIdInLevel paragraphStyleIdInLevel1 = new ParagraphStyleIdInLevel() { Val = "NumberedList" };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

            tabs1.Append(tabStop1);
            Indentation indentation1 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties1.Append(tabs1);
            previousParagraphProperties1.Append(indentation1);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties1.Append(runFonts1);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(paragraphStyleIdInLevel1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText2 = new LevelText() { Val = "%1.%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 792 };

            tabs2.Append(tabStop2);
            Indentation indentation2 = new Indentation() { Left = "792", Hanging = "432" };

            previousParagraphProperties2.Append(tabs2);
            previousParagraphProperties2.Append(indentation2);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties2.Append(runFonts2);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText3 = new LevelText() { Val = "%1.%2.%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 1224 };

            tabs3.Append(tabStop3);
            Indentation indentation3 = new Indentation() { Left = "1224", Hanging = "504" };

            previousParagraphProperties3.Append(tabs3);
            previousParagraphProperties3.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties3.Append(runFonts3);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);

            Level level4 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 1728 };

            tabs4.Append(tabStop4);
            Indentation indentation4 = new Indentation() { Left = "1728", Hanging = "648" };

            previousParagraphProperties4.Append(tabs4);
            previousParagraphProperties4.Append(indentation4);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties4.Append(runFonts4);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Level level5 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

            Tabs tabs5 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 2232 };

            tabs5.Append(tabStop5);
            Indentation indentation5 = new Indentation() { Left = "2232", Hanging = "792" };

            previousParagraphProperties5.Append(tabs5);
            previousParagraphProperties5.Append(indentation5);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties5.Append(runFonts5);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);
            level5.Append(numberingSymbolRunProperties5);

            Level level6 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

            Tabs tabs6 = new Tabs();
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 2736 };

            tabs6.Append(tabStop6);
            Indentation indentation6 = new Indentation() { Left = "2736", Hanging = "936" };

            previousParagraphProperties6.Append(tabs6);
            previousParagraphProperties6.Append(indentation6);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties6.Append(runFonts6);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);
            level6.Append(numberingSymbolRunProperties6);

            Level level7 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

            Tabs tabs7 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 3240 };

            tabs7.Append(tabStop7);
            Indentation indentation7 = new Indentation() { Left = "3240", Hanging = "1080" };

            previousParagraphProperties7.Append(tabs7);
            previousParagraphProperties7.Append(indentation7);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties7.Append(runFonts7);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);
            level7.Append(numberingSymbolRunProperties7);

            Level level8 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

            Tabs tabs8 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 3744 };

            tabs8.Append(tabStop8);
            Indentation indentation8 = new Indentation() { Left = "3744", Hanging = "1224" };

            previousParagraphProperties8.Append(tabs8);
            previousParagraphProperties8.Append(indentation8);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties8.Append(runFonts8);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);
            level8.Append(numberingSymbolRunProperties8);

            Level level9 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

            Tabs tabs9 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 4320 };

            tabs9.Append(tabStop9);
            Indentation indentation9 = new Indentation() { Left = "4320", Hanging = "1440" };

            previousParagraphProperties9.Append(tabs9);
            previousParagraphProperties9.Append(indentation9);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties9.Append(runFonts9);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);
            level9.Append(numberingSymbolRunProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            AbstractNum abstractNum2 = new AbstractNum() { AbstractNumberId = 1 };
            Nsid nsid2 = new Nsid() { Val = "460E3EC6" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "DD28C6DC" };

            Level level10 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText10 = new LevelText() { Val = "·" };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();

            Tabs tabs10 = new Tabs();
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

            tabs10.Append(tabStop10);
            Indentation indentation10 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties10.Append(tabs10);
            previousParagraphProperties10.Append(indentation10);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };
            FontSize fontSize1 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties10.Append(runFonts10);
            numberingSymbolRunProperties10.Append(fontSize1);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties10);

            Level level11 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText11 = new LevelText() { Val = "n" };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();

            Tabs tabs11 = new Tabs();
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Number, Position = 1080 };

            tabs11.Append(tabStop11);
            Indentation indentation11 = new Indentation() { Left = "1080", Hanging = "360" };

            previousParagraphProperties11.Append(tabs11);
            previousParagraphProperties11.Append(indentation11);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize2 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties11.Append(runFonts11);
            numberingSymbolRunProperties11.Append(fontSize2);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);
            level11.Append(numberingSymbolRunProperties11);

            Level level12 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText12 = new LevelText() { Val = "n" };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();

            Tabs tabs12 = new Tabs();
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

            tabs12.Append(tabStop12);
            Indentation indentation12 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties12.Append(tabs12);
            previousParagraphProperties12.Append(indentation12);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize3 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties12.Append(runFonts12);
            numberingSymbolRunProperties12.Append(fontSize3);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);
            level12.Append(numberingSymbolRunProperties12);

            Level level13 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText13 = new LevelText() { Val = "n" };
            LevelJustification levelJustification13 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();

            Tabs tabs13 = new Tabs();
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Number, Position = 1800 };

            tabs13.Append(tabStop13);
            Indentation indentation13 = new Indentation() { Left = "1800", Hanging = "360" };

            previousParagraphProperties13.Append(tabs13);
            previousParagraphProperties13.Append(indentation13);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize4 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties13.Append(runFonts13);
            numberingSymbolRunProperties13.Append(fontSize4);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);
            level13.Append(numberingSymbolRunProperties13);

            Level level14 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText14 = new LevelText() { Val = "n" };
            LevelJustification levelJustification14 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();

            Tabs tabs14 = new Tabs();
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

            tabs14.Append(tabStop14);
            Indentation indentation14 = new Indentation() { Left = "2160", Hanging = "360" };

            previousParagraphProperties14.Append(tabs14);
            previousParagraphProperties14.Append(indentation14);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize5 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties14.Append(runFonts14);
            numberingSymbolRunProperties14.Append(fontSize5);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);
            level14.Append(numberingSymbolRunProperties14);

            Level level15 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText15 = new LevelText() { Val = "n" };
            LevelJustification levelJustification15 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();

            Tabs tabs15 = new Tabs();
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Number, Position = 2520 };

            tabs15.Append(tabStop15);
            Indentation indentation15 = new Indentation() { Left = "2520", Hanging = "360" };

            previousParagraphProperties15.Append(tabs15);
            previousParagraphProperties15.Append(indentation15);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize6 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties15.Append(runFonts15);
            numberingSymbolRunProperties15.Append(fontSize6);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);
            level15.Append(numberingSymbolRunProperties15);

            Level level16 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText16 = new LevelText() { Val = "n" };
            LevelJustification levelJustification16 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();

            Tabs tabs16 = new Tabs();
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs16.Append(tabStop16);
            Indentation indentation16 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties16.Append(tabs16);
            previousParagraphProperties16.Append(indentation16);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize7 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties16.Append(runFonts16);
            numberingSymbolRunProperties16.Append(fontSize7);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);
            level16.Append(numberingSymbolRunProperties16);

            Level level17 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText17 = new LevelText() { Val = "n" };
            LevelJustification levelJustification17 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();

            Tabs tabs17 = new Tabs();
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Number, Position = 3240 };

            tabs17.Append(tabStop17);
            Indentation indentation17 = new Indentation() { Left = "3240", Hanging = "360" };

            previousParagraphProperties17.Append(tabs17);
            previousParagraphProperties17.Append(indentation17);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts17 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize8 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties17.Append(runFonts17);
            numberingSymbolRunProperties17.Append(fontSize8);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);
            level17.Append(numberingSymbolRunProperties17);

            Level level18 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText18 = new LevelText() { Val = "n" };
            LevelJustification levelJustification18 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();

            Tabs tabs18 = new Tabs();
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

            tabs18.Append(tabStop18);
            Indentation indentation18 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties18.Append(tabs18);
            previousParagraphProperties18.Append(indentation18);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize9 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties18.Append(runFonts18);
            numberingSymbolRunProperties18.Append(fontSize9);

            level18.Append(startNumberingValue18);
            level18.Append(numberingFormat18);
            level18.Append(levelText18);
            level18.Append(levelJustification18);
            level18.Append(previousParagraphProperties18);
            level18.Append(numberingSymbolRunProperties18);

            abstractNum2.Append(nsid2);
            abstractNum2.Append(multiLevelType2);
            abstractNum2.Append(templateCode2);
            abstractNum2.Append(level10);
            abstractNum2.Append(level11);
            abstractNum2.Append(level12);
            abstractNum2.Append(level13);
            abstractNum2.Append(level14);
            abstractNum2.Append(level15);
            abstractNum2.Append(level16);
            abstractNum2.Append(level17);
            abstractNum2.Append(level18);

            AbstractNum abstractNum3 = new AbstractNum() { AbstractNumberId = 2 };
            Nsid nsid3 = new Nsid() { Val = "70913756" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "624EA66A" };
            AbstractNumDefinitionName abstractNumDefinitionName1 = new AbstractNumDefinitionName() { Val = "RussellSubbullet" };

            Level level19 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue19 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat19 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText19 = new LevelText() { Val = "n" };
            LevelJustification levelJustification19 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties19 = new PreviousParagraphProperties();

            Tabs tabs19 = new Tabs();
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Number, Position = 360 };

            tabs19.Append(tabStop19);
            Indentation indentation19 = new Indentation() { Left = "360", Hanging = "360" };

            previousParagraphProperties19.Append(tabs19);
            previousParagraphProperties19.Append(indentation19);

            NumberingSymbolRunProperties numberingSymbolRunProperties19 = new NumberingSymbolRunProperties();
            RunFonts runFonts19 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize10 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties19.Append(runFonts19);
            numberingSymbolRunProperties19.Append(fontSize10);

            level19.Append(startNumberingValue19);
            level19.Append(numberingFormat19);
            level19.Append(levelText19);
            level19.Append(levelJustification19);
            level19.Append(previousParagraphProperties19);
            level19.Append(numberingSymbolRunProperties19);

            Level level20 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue20 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat20 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText20 = new LevelText() { Val = "n" };
            LevelJustification levelJustification20 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties20 = new PreviousParagraphProperties();

            Tabs tabs20 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

            tabs20.Append(tabStop20);
            Indentation indentation20 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties20.Append(tabs20);
            previousParagraphProperties20.Append(indentation20);

            NumberingSymbolRunProperties numberingSymbolRunProperties20 = new NumberingSymbolRunProperties();
            RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize11 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties20.Append(runFonts20);
            numberingSymbolRunProperties20.Append(fontSize11);

            level20.Append(startNumberingValue20);
            level20.Append(numberingFormat20);
            level20.Append(levelText20);
            level20.Append(levelJustification20);
            level20.Append(previousParagraphProperties20);
            level20.Append(numberingSymbolRunProperties20);

            Level level21 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue21 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat21 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText21 = new LevelText() { Val = "n" };
            LevelJustification levelJustification21 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties21 = new PreviousParagraphProperties();

            Tabs tabs21 = new Tabs();
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Number, Position = 1080 };

            tabs21.Append(tabStop21);
            Indentation indentation21 = new Indentation() { Left = "1080", Hanging = "360" };

            previousParagraphProperties21.Append(tabs21);
            previousParagraphProperties21.Append(indentation21);

            NumberingSymbolRunProperties numberingSymbolRunProperties21 = new NumberingSymbolRunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize12 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties21.Append(runFonts21);
            numberingSymbolRunProperties21.Append(fontSize12);

            level21.Append(startNumberingValue21);
            level21.Append(numberingFormat21);
            level21.Append(levelText21);
            level21.Append(levelJustification21);
            level21.Append(previousParagraphProperties21);
            level21.Append(numberingSymbolRunProperties21);

            Level level22 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue22 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat22 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText22 = new LevelText() { Val = "n" };
            LevelJustification levelJustification22 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties22 = new PreviousParagraphProperties();

            Tabs tabs22 = new Tabs();
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

            tabs22.Append(tabStop22);
            Indentation indentation22 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties22.Append(tabs22);
            previousParagraphProperties22.Append(indentation22);

            NumberingSymbolRunProperties numberingSymbolRunProperties22 = new NumberingSymbolRunProperties();
            RunFonts runFonts22 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize13 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties22.Append(runFonts22);
            numberingSymbolRunProperties22.Append(fontSize13);

            level22.Append(startNumberingValue22);
            level22.Append(numberingFormat22);
            level22.Append(levelText22);
            level22.Append(levelJustification22);
            level22.Append(previousParagraphProperties22);
            level22.Append(numberingSymbolRunProperties22);

            Level level23 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue23 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat23 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText23 = new LevelText() { Val = "n" };
            LevelJustification levelJustification23 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties23 = new PreviousParagraphProperties();

            Tabs tabs23 = new Tabs();
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Number, Position = 1800 };

            tabs23.Append(tabStop23);
            Indentation indentation23 = new Indentation() { Left = "1800", Hanging = "360" };

            previousParagraphProperties23.Append(tabs23);
            previousParagraphProperties23.Append(indentation23);

            NumberingSymbolRunProperties numberingSymbolRunProperties23 = new NumberingSymbolRunProperties();
            RunFonts runFonts23 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize14 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties23.Append(runFonts23);
            numberingSymbolRunProperties23.Append(fontSize14);

            level23.Append(startNumberingValue23);
            level23.Append(numberingFormat23);
            level23.Append(levelText23);
            level23.Append(levelJustification23);
            level23.Append(previousParagraphProperties23);
            level23.Append(numberingSymbolRunProperties23);

            Level level24 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue24 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat24 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText24 = new LevelText() { Val = "n" };
            LevelJustification levelJustification24 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties24 = new PreviousParagraphProperties();

            Tabs tabs24 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

            tabs24.Append(tabStop24);
            Indentation indentation24 = new Indentation() { Left = "2160", Hanging = "360" };

            previousParagraphProperties24.Append(tabs24);
            previousParagraphProperties24.Append(indentation24);

            NumberingSymbolRunProperties numberingSymbolRunProperties24 = new NumberingSymbolRunProperties();
            RunFonts runFonts24 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize15 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties24.Append(runFonts24);
            numberingSymbolRunProperties24.Append(fontSize15);

            level24.Append(startNumberingValue24);
            level24.Append(numberingFormat24);
            level24.Append(levelText24);
            level24.Append(levelJustification24);
            level24.Append(previousParagraphProperties24);
            level24.Append(numberingSymbolRunProperties24);

            Level level25 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue25 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat25 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText25 = new LevelText() { Val = "n" };
            LevelJustification levelJustification25 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties25 = new PreviousParagraphProperties();

            Tabs tabs25 = new Tabs();
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Number, Position = 2520 };

            tabs25.Append(tabStop25);
            Indentation indentation25 = new Indentation() { Left = "2520", Hanging = "360" };

            previousParagraphProperties25.Append(tabs25);
            previousParagraphProperties25.Append(indentation25);

            NumberingSymbolRunProperties numberingSymbolRunProperties25 = new NumberingSymbolRunProperties();
            RunFonts runFonts25 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize16 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties25.Append(runFonts25);
            numberingSymbolRunProperties25.Append(fontSize16);

            level25.Append(startNumberingValue25);
            level25.Append(numberingFormat25);
            level25.Append(levelText25);
            level25.Append(levelJustification25);
            level25.Append(previousParagraphProperties25);
            level25.Append(numberingSymbolRunProperties25);

            Level level26 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue26 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat26 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText26 = new LevelText() { Val = "n" };
            LevelJustification levelJustification26 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties26 = new PreviousParagraphProperties();

            Tabs tabs26 = new Tabs();
            TabStop tabStop26 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs26.Append(tabStop26);
            Indentation indentation26 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties26.Append(tabs26);
            previousParagraphProperties26.Append(indentation26);

            NumberingSymbolRunProperties numberingSymbolRunProperties26 = new NumberingSymbolRunProperties();
            RunFonts runFonts26 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize17 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties26.Append(runFonts26);
            numberingSymbolRunProperties26.Append(fontSize17);

            level26.Append(startNumberingValue26);
            level26.Append(numberingFormat26);
            level26.Append(levelText26);
            level26.Append(levelJustification26);
            level26.Append(previousParagraphProperties26);
            level26.Append(numberingSymbolRunProperties26);

            Level level27 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue27 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat27 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText27 = new LevelText() { Val = "n" };
            LevelJustification levelJustification27 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties27 = new PreviousParagraphProperties();

            Tabs tabs27 = new Tabs();
            TabStop tabStop27 = new TabStop() { Val = TabStopValues.Number, Position = 3240 };

            tabs27.Append(tabStop27);
            Indentation indentation27 = new Indentation() { Left = "3240", Hanging = "360" };

            previousParagraphProperties27.Append(tabs27);
            previousParagraphProperties27.Append(indentation27);

            NumberingSymbolRunProperties numberingSymbolRunProperties27 = new NumberingSymbolRunProperties();
            RunFonts runFonts27 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };
            FontSize fontSize18 = new FontSize() { Val = "16" };

            numberingSymbolRunProperties27.Append(runFonts27);
            numberingSymbolRunProperties27.Append(fontSize18);

            level27.Append(startNumberingValue27);
            level27.Append(numberingFormat27);
            level27.Append(levelText27);
            level27.Append(levelJustification27);
            level27.Append(previousParagraphProperties27);
            level27.Append(numberingSymbolRunProperties27);

            abstractNum3.Append(nsid3);
            abstractNum3.Append(multiLevelType3);
            abstractNum3.Append(templateCode3);
            abstractNum3.Append(abstractNumDefinitionName1);
            abstractNum3.Append(level19);
            abstractNum3.Append(level20);
            abstractNum3.Append(level21);
            abstractNum3.Append(level22);
            abstractNum3.Append(level23);
            abstractNum3.Append(level24);
            abstractNum3.Append(level25);
            abstractNum3.Append(level26);
            abstractNum3.Append(level27);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 2 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 0 };

            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 1 };

            numberingInstance3.Append(abstractNumId3);

            NumberingInstance numberingInstance4 = new NumberingInstance() { NumberID = 4 };
            AbstractNumId abstractNumId4 = new AbstractNumId() { Val = 0 };

            numberingInstance4.Append(abstractNumId4);

            NumberingInstance numberingInstance5 = new NumberingInstance() { NumberID = 5 };
            AbstractNumId abstractNumId5 = new AbstractNumId() { Val = 0 };

            LevelOverride levelOverride1 = new LevelOverride() { LevelIndex = 0 };
            StartOverrideNumberingValue startOverrideNumberingValue1 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride1.Append(startOverrideNumberingValue1);

            LevelOverride levelOverride2 = new LevelOverride() { LevelIndex = 1 };
            StartOverrideNumberingValue startOverrideNumberingValue2 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride2.Append(startOverrideNumberingValue2);

            LevelOverride levelOverride3 = new LevelOverride() { LevelIndex = 2 };
            StartOverrideNumberingValue startOverrideNumberingValue3 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride3.Append(startOverrideNumberingValue3);

            LevelOverride levelOverride4 = new LevelOverride() { LevelIndex = 3 };
            StartOverrideNumberingValue startOverrideNumberingValue4 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride4.Append(startOverrideNumberingValue4);

            LevelOverride levelOverride5 = new LevelOverride() { LevelIndex = 4 };
            StartOverrideNumberingValue startOverrideNumberingValue5 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride5.Append(startOverrideNumberingValue5);

            LevelOverride levelOverride6 = new LevelOverride() { LevelIndex = 5 };
            StartOverrideNumberingValue startOverrideNumberingValue6 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride6.Append(startOverrideNumberingValue6);

            LevelOverride levelOverride7 = new LevelOverride() { LevelIndex = 6 };
            StartOverrideNumberingValue startOverrideNumberingValue7 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride7.Append(startOverrideNumberingValue7);

            LevelOverride levelOverride8 = new LevelOverride() { LevelIndex = 7 };
            StartOverrideNumberingValue startOverrideNumberingValue8 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride8.Append(startOverrideNumberingValue8);

            LevelOverride levelOverride9 = new LevelOverride() { LevelIndex = 8 };
            StartOverrideNumberingValue startOverrideNumberingValue9 = new StartOverrideNumberingValue() { Val = 1 };

            levelOverride9.Append(startOverrideNumberingValue9);

            numberingInstance4.Append(abstractNumId5);
            numberingInstance4.Append(levelOverride1);
            numberingInstance4.Append(levelOverride2);
            numberingInstance4.Append(levelOverride3);
            numberingInstance4.Append(levelOverride4);
            numberingInstance4.Append(levelOverride5);
            numberingInstance4.Append(levelOverride6);
            numberingInstance4.Append(levelOverride7);
            numberingInstance4.Append(levelOverride8);
            numberingInstance4.Append(levelOverride9);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(abstractNum3);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);
            numbering1.Append(numberingInstance3);
            numbering1.Append(numberingInstance4);
            numbering1.Append(numberingInstance5);

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of endnotesPart1.
        private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes();

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "004805C1", RsidRunAdditionDefault = "004805C1" };

            Run run46 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run46.Append(separatorMark1);

            paragraph40.Append(run46);

            endnote1.Append(paragraph40);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph41 = new Paragraph() { RsidParagraphAddition = "004805C1", RsidRunAdditionDefault = "004805C1" };

            Run run47 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run47.Append(continuationSeparatorMark1);

            paragraph41.Append(run47);

            endnote2.Append(paragraph41);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of footerPart2.
        private void GenerateFooterPart2Content(FooterPart footerPart2)
        {
            Footer footer2 = new Footer();

            Paragraph paragraph42 = new Paragraph() { RsidParagraphAddition = "00F34666", RsidParagraphProperties = "00F34666", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId37 = new ParagraphStyleId() { Val = "FooterRankLegend" };

            Tabs tabs24 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Left, Position = 3315 };

            tabs24.Append(tabStop24);
            SpacingBetweenLines spacingBetweenLines49 = new SpacingBetweenLines() { After = "240" };

            paragraphProperties39.Append(paragraphStyleId37);
            paragraphProperties39.Append(tabs24);
            paragraphProperties39.Append(spacingBetweenLines49);

            Run run48 = new Run();

            RunProperties runProperties26 = new RunProperties();
            NoProof noProof13 = new NoProof();
            Languages languages32 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties26.Append(noProof13);
            runProperties26.Append(languages32);

            Drawing drawing13 = new Drawing();

            Wp.Inline inline11 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent13 = new Wp.Extent() { Cx = 1447800L, Cy = 314325L };
            Wp.EffectExtent effectExtent13 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties13 = new Wp.DocProperties() { Id = (UInt32Value)12U, Name = "Image 12", Description = "RADAR_RankLegend" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties13 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks13 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties13.Append(graphicFrameLocks13);

            A.Graphic graphic13 = new A.Graphic();

            A.GraphicData graphicData13 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture13 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties13 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties13 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 12", Description = "RADAR_RankLegend" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties13 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks13 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties13.Append(pictureLocks13);

            nonVisualPictureProperties13.Append(nonVisualDrawingProperties13);
            nonVisualPictureProperties13.Append(nonVisualPictureDrawingProperties13);

            Pic.BlipFill blipFill13 = new Pic.BlipFill();
            A.Blip blip13 = new A.Blip() { Embed = "rId1" };
            A.SourceRectangle sourceRectangle13 = new A.SourceRectangle();

            A.Stretch stretch13 = new A.Stretch();
            A.FillRectangle fillRectangle13 = new A.FillRectangle();

            stretch13.Append(fillRectangle13);

            blipFill13.Append(blip13);
            blipFill13.Append(sourceRectangle13);
            blipFill13.Append(stretch13);

            Pic.ShapeProperties shapeProperties13 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset13 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents13 = new A.Extents() { Cx = 1447800L, Cy = 314325L };

            transform2D13.Append(offset13);
            transform2D13.Append(extents13);

            A.PresetGeometry presetGeometry13 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList13 = new A.AdjustValueList();

            presetGeometry13.Append(adjustValueList13);
            A.NoFill noFill23 = new A.NoFill();

            A.Outline outline11 = new A.Outline() { Width = 9525 };
            A.NoFill noFill24 = new A.NoFill();
            A.Miter miter11 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd11 = new A.HeadEnd();
            A.TailEnd tailEnd11 = new A.TailEnd();

            outline11.Append(noFill24);
            outline11.Append(miter11);
            outline11.Append(headEnd11);
            outline11.Append(tailEnd11);

            shapeProperties13.Append(transform2D13);
            shapeProperties13.Append(presetGeometry13);
            shapeProperties13.Append(noFill23);
            shapeProperties13.Append(outline11);

            picture13.Append(nonVisualPictureProperties13);
            picture13.Append(blipFill13);
            picture13.Append(shapeProperties13);

            graphicData13.Append(picture13);

            graphic13.Append(graphicData13);

            inline11.Append(extent13);
            inline11.Append(effectExtent13);
            inline11.Append(docProperties13);
            inline11.Append(nonVisualGraphicFrameDrawingProperties13);
            inline11.Append(graphic13);

            drawing13.Append(inline11);

            run48.Append(runProperties26);
            run48.Append(drawing13);

            Run run49 = new Run() { RsidRunAddition = "00F34666" };
            TabChar tabChar2 = new TabChar();

            run49.Append(tabChar2);

            paragraph42.Append(paragraphProperties39);
            paragraph42.Append(run48);
            paragraph42.Append(run49);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphAddition = "00F34666", RsidParagraphProperties = "00782598", RsidRunAdditionDefault = "00F34666" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId38 = new ParagraphStyleId() { Val = "Disclaimer" };

            ParagraphBorders paragraphBorders12 = new ParagraphBorders();
            TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "66AADD", Size = (UInt32Value)48U, Space = (UInt32Value)1U };

            paragraphBorders12.Append(topBorder10);

            paragraphProperties40.Append(paragraphStyleId38);
            paragraphProperties40.Append(paragraphBorders12);

            paragraph43.Append(paragraphProperties40);

            Table table4 = new Table();

            TableProperties tableProperties4 = new TableProperties();
            TableStyle tableStyle4 = new TableStyle() { Val = "Grilledutableau" };
            TableWidth tableWidth4 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders5 = new TableBorders();
            TopBorder topBorder11 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder5 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder5 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders5.Append(topBorder11);
            tableBorders5.Append(leftBorder6);
            tableBorders5.Append(bottomBorder12);
            tableBorders5.Append(rightBorder6);
            tableBorders5.Append(insideHorizontalBorder5);
            tableBorders5.Append(insideVerticalBorder5);
            TableLayout tableLayout4 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook4 = new TableLook() { Val = "01E0" };

            tableProperties4.Append(tableStyle4);
            tableProperties4.Append(tableWidth4);
            tableProperties4.Append(tableBorders5);
            tableProperties4.Append(tableLayout4);
            tableProperties4.Append(tableLook4);

            TableGrid tableGrid4 = new TableGrid();
            GridColumn gridColumn10 = new GridColumn() { Width = "8388" };
            GridColumn gridColumn11 = new GridColumn() { Width = "540" };

            tableGrid4.Append(gridColumn10);
            tableGrid4.Append(gridColumn11);

            TableRow tableRow5 = new TableRow() { RsidTableRowAddition = "002A248B", RsidTableRowProperties = "002A248B" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)618U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "8388", Type = TableWidthUnitValues.Dxa };

            tableCellProperties14.Append(tableCellWidth14);

            CustomXmlBlock customXmlBlock20 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "reportdoc" };

            CustomXmlBlock customXmlBlock21 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "footer" };

            CustomXmlBlock customXmlBlock22 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "ShortDisclaimer" };

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "003F1967", RsidParagraphAddition = "002A248B", RsidParagraphProperties = "00C87F09", RsidRunAdditionDefault = "002A248B" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId39 = new ParagraphStyleId() { Val = "Disclaimer" };

            ParagraphBorders paragraphBorders13 = new ParagraphBorders();
            TopBorder topBorder12 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders13.Append(topBorder12);

            paragraphProperties41.Append(paragraphStyleId39);
            paragraphProperties41.Append(paragraphBorders13);

            Run run50 = new Run() { RsidRunProperties = "004B56C1" };
            Text text32 = new Text();
            text32.Text = "Confidential Proprietary Information of Russell Investments not to be distributed to third party without the express written consent of Russell Investments. Please see Important Legal Information for further information on this material.";

            run50.Append(text32);

            paragraph44.Append(paragraphProperties41);
            paragraph44.Append(run50);

            customXmlBlock22.Append(paragraph44);

            customXmlBlock21.Append(customXmlBlock22);

            customXmlBlock20.Append(customXmlBlock21);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphAddition = "002A248B", RsidParagraphProperties = "00C87F09", RsidRunAdditionDefault = "002A248B" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId40 = new ParagraphStyleId() { Val = "FooterLogo" };

            paragraphProperties42.Append(paragraphStyleId40);

            paragraph45.Append(paragraphProperties42);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(customXmlBlock20);
            tableCell14.Append(paragraph45);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "540", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(tableCellVerticalAlignment2);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "00FB4EAB", RsidParagraphAddition = "002A248B", RsidParagraphProperties = "00FB4EAB", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId41 = new ParagraphStyleId() { Val = "FooterLogo" };
            Justification justification9 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties43.Append(paragraphStyleId41);
            paragraphProperties43.Append(justification9);

            Run run51 = new Run();

            RunProperties runProperties27 = new RunProperties();
            NoProof noProof14 = new NoProof();
            Languages languages33 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties27.Append(noProof14);
            runProperties27.Append(languages33);

            Drawing drawing14 = new Drawing();

            Wp.Anchor anchor3 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251656192U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
            Wp.SimplePosition simplePosition3 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition3 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset5 = new Wp.PositionOffset();
            positionOffset5.Text = "388620";

            horizontalPosition3.Append(positionOffset5);

            Wp.VerticalPosition verticalPosition3 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset6 = new Wp.PositionOffset();
            positionOffset6.Text = "-2077720";

            verticalPosition3.Append(positionOffset6);
            Wp.Extent extent14 = new Wp.Extent() { Cx = 1085850L, Cy = 323850L };
            Wp.EffectExtent effectExtent14 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone3 = new Wp.WrapNone();
            Wp.DocProperties docProperties14 = new Wp.DocProperties() { Id = (UInt32Value)62U, Name = "Image 62", Description = "RADAR_RLogo" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties14 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks14 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties14.Append(graphicFrameLocks14);

            A.Graphic graphic14 = new A.Graphic();

            A.GraphicData graphicData14 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture14 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties14 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties14 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 62", Description = "RADAR_RLogo" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties14 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks14 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties14.Append(pictureLocks14);

            nonVisualPictureProperties14.Append(nonVisualDrawingProperties14);
            nonVisualPictureProperties14.Append(nonVisualPictureDrawingProperties14);

            Pic.BlipFill blipFill14 = new Pic.BlipFill();
            A.Blip blip14 = new A.Blip() { Embed = "rId2" };
            A.SourceRectangle sourceRectangle14 = new A.SourceRectangle();

            A.Stretch stretch14 = new A.Stretch();
            A.FillRectangle fillRectangle14 = new A.FillRectangle();

            stretch14.Append(fillRectangle14);

            blipFill14.Append(blip14);
            blipFill14.Append(sourceRectangle14);
            blipFill14.Append(stretch14);

            Pic.ShapeProperties shapeProperties14 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D14 = new A.Transform2D();
            A.Offset offset14 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents14 = new A.Extents() { Cx = 1085850L, Cy = 323850L };

            transform2D14.Append(offset14);
            transform2D14.Append(extents14);

            A.PresetGeometry presetGeometry14 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList14 = new A.AdjustValueList();

            presetGeometry14.Append(adjustValueList14);
            A.NoFill noFill25 = new A.NoFill();

            shapeProperties14.Append(transform2D14);
            shapeProperties14.Append(presetGeometry14);
            shapeProperties14.Append(noFill25);

            picture14.Append(nonVisualPictureProperties14);
            picture14.Append(blipFill14);
            picture14.Append(shapeProperties14);

            graphicData14.Append(picture14);

            graphic14.Append(graphicData14);

            anchor3.Append(simplePosition3);
            anchor3.Append(horizontalPosition3);
            anchor3.Append(verticalPosition3);
            anchor3.Append(extent14);
            anchor3.Append(effectExtent14);
            anchor3.Append(wrapNone3);
            anchor3.Append(docProperties14);
            anchor3.Append(nonVisualGraphicFrameDrawingProperties14);
            anchor3.Append(graphic14);

            drawing14.Append(anchor3);

            run51.Append(runProperties27);
            run51.Append(drawing14);

            paragraph46.Append(paragraphProperties43);
            paragraph46.Append(run51);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph46);

            tableRow5.Append(tableRowProperties3);
            tableRow5.Append(tableCell14);
            tableRow5.Append(tableCell15);

            table4.Append(tableProperties4);
            table4.Append(tableGrid4);
            table4.Append(tableRow5);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphAddition = "00F34666", RsidParagraphProperties = "00F34666", RsidRunAdditionDefault = "00F34666" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId42 = new ParagraphStyleId() { Val = "FooterLogo" };

            paragraphProperties44.Append(paragraphStyleId42);

            paragraph47.Append(paragraphProperties44);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphMarkRevision = "00F34666", RsidParagraphAddition = "003E4D99", RsidParagraphProperties = "00F34666", RsidRunAdditionDefault = "003E4D99" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId43 = new ParagraphStyleId() { Val = "FooterPageNumber" };
            SpacingBetweenLines spacingBetweenLines50 = new SpacingBetweenLines() { After = "320" };

            paragraphProperties45.Append(paragraphStyleId43);
            paragraphProperties45.Append(spacingBetweenLines50);

            paragraph48.Append(paragraphProperties45);

            footer2.Append(paragraph42);
            footer2.Append(paragraph43);
            footer2.Append(table4);
            footer2.Append(paragraph47);
            footer2.Append(paragraph48);

            footerPart2.Footer = footer2;
        }

        // Generates content of footnotesPart1.
        private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes();

            Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph49 = new Paragraph() { RsidParagraphAddition = "004805C1", RsidRunAdditionDefault = "004805C1" };

            Run run52 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run52.Append(separatorMark2);

            paragraph49.Append(run52);

            footnote1.Append(paragraph49);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph50 = new Paragraph() { RsidParagraphAddition = "004805C1", RsidRunAdditionDefault = "004805C1" };

            Run run53 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run53.Append(continuationSeparatorMark2);

            paragraph50.Append(run53);

            footnote2.Append(paragraph50);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Thème Office" };

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline12 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline12.Append(solidFill2);
            outline12.Append(presetDash1);

            A.Outline outline13 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline13.Append(solidFill3);
            outline13.Append(presetDash2);

            A.Outline outline14 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline14.Append(solidFill4);
            outline14.Append(presetDash3);

            lineStyleList1.Append(outline12);
            lineStyleList1.Append(outline13);
            lineStyleList1.Append(outline14);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of headerPart2.
        private void GenerateHeaderPart2Content(HeaderPart headerPart2)
        {
            Header header2 = new Header();

            Paragraph paragraph51 = new Paragraph() { RsidParagraphAddition = "00C913B8", RsidParagraphProperties = "00C913B8", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId44 = new ParagraphStyleId() { Val = "En-tte" };

            ParagraphBorders paragraphBorders14 = new ParagraphBorders();
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders14.Append(bottomBorder13);
            SpacingBetweenLines spacingBetweenLines51 = new SpacingBetweenLines() { After = "60" };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunStyle runStyle10 = new RunStyle() { Val = "DateCar" };
            RunFonts runFonts52 = new RunFonts() { EastAsia = "MS Mincho" };
            FontSize fontSize49 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties17.Append(runStyle10);
            paragraphMarkRunProperties17.Append(runFonts52);
            paragraphMarkRunProperties17.Append(fontSize49);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript50);

            paragraphProperties46.Append(paragraphStyleId44);
            paragraphProperties46.Append(paragraphBorders14);
            paragraphProperties46.Append(spacingBetweenLines51);
            paragraphProperties46.Append(paragraphMarkRunProperties17);

            Run run54 = new Run();

            RunProperties runProperties28 = new RunProperties();
            NoProof noProof15 = new NoProof();
            Languages languages34 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties28.Append(noProof15);
            runProperties28.Append(languages34);

            Drawing drawing15 = new Drawing();

            Wp.Anchor anchor4 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = true, Locked = false, LayoutInCell = true, AllowOverlap = true };
            Wp.SimplePosition simplePosition4 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition4 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset7 = new Wp.PositionOffset();
            positionOffset7.Text = "8890";

            horizontalPosition4.Append(positionOffset7);

            Wp.VerticalPosition verticalPosition4 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset8 = new Wp.PositionOffset();
            positionOffset8.Text = "8890";

            verticalPosition4.Append(positionOffset8);
            Wp.Extent extent15 = new Wp.Extent() { Cx = 6848475L, Cy = 438150L };
            Wp.EffectExtent effectExtent15 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 9525L, BottomEdge = 0L };
            Wp.WrapNone wrapNone4 = new Wp.WrapNone();
            Wp.DocProperties docProperties15 = new Wp.DocProperties() { Id = (UInt32Value)65U, Name = "Image 65", Description = "RADAR_Opinion_Page2_BNR" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties15 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks15 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties15.Append(graphicFrameLocks15);

            A.Graphic graphic15 = new A.Graphic();

            A.GraphicData graphicData15 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture15 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties15 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties15 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 65", Description = "RADAR_Opinion_Page2_BNR" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties15 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks15 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties15.Append(pictureLocks15);

            nonVisualPictureProperties15.Append(nonVisualDrawingProperties15);
            nonVisualPictureProperties15.Append(nonVisualPictureDrawingProperties15);

            Pic.BlipFill blipFill15 = new Pic.BlipFill();
            A.Blip blip15 = new A.Blip() { Embed = "rId1" };
            A.SourceRectangle sourceRectangle15 = new A.SourceRectangle();

            A.Stretch stretch15 = new A.Stretch();
            A.FillRectangle fillRectangle15 = new A.FillRectangle();

            stretch15.Append(fillRectangle15);

            blipFill15.Append(blip15);
            blipFill15.Append(sourceRectangle15);
            blipFill15.Append(stretch15);

            Pic.ShapeProperties shapeProperties15 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset15 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents15 = new A.Extents() { Cx = 6848475L, Cy = 438150L };

            transform2D15.Append(offset15);
            transform2D15.Append(extents15);

            A.PresetGeometry presetGeometry15 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList15 = new A.AdjustValueList();

            presetGeometry15.Append(adjustValueList15);
            A.NoFill noFill26 = new A.NoFill();

            shapeProperties15.Append(transform2D15);
            shapeProperties15.Append(presetGeometry15);
            shapeProperties15.Append(noFill26);

            picture15.Append(nonVisualPictureProperties15);
            picture15.Append(blipFill15);
            picture15.Append(shapeProperties15);

            graphicData15.Append(picture15);

            graphic15.Append(graphicData15);

            anchor4.Append(simplePosition4);
            anchor4.Append(horizontalPosition4);
            anchor4.Append(verticalPosition4);
            anchor4.Append(extent15);
            anchor4.Append(effectExtent15);
            anchor4.Append(wrapNone4);
            anchor4.Append(docProperties15);
            anchor4.Append(nonVisualGraphicFrameDrawingProperties15);
            anchor4.Append(graphic15);

            drawing15.Append(anchor4);

            run54.Append(runProperties28);
            run54.Append(drawing15);

            paragraph51.Append(paragraphProperties46);
            paragraph51.Append(run54);

            CustomXmlBlock customXmlBlock23 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "reportdoc" };

            CustomXmlBlock customXmlBlock24 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "header" };

            CustomXmlBlock customXmlBlock25 = new CustomXmlBlock() { Uri = "http://hubblereports.com/namespace", Element = "ReportDate" };

            Paragraph paragraph52 = new Paragraph() { RsidParagraphMarkRevision = "006B1D99", RsidParagraphAddition = "003907B3", RsidParagraphProperties = "003907B3", RsidRunAdditionDefault = "003907B3" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId45 = new ParagraphStyleId() { Val = "En-tte" };

            ParagraphBorders paragraphBorders15 = new ParagraphBorders();
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders15.Append(bottomBorder14);
            SpacingBetweenLines spacingBetweenLines52 = new SpacingBetweenLines() { After = "60" };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunStyle runStyle11 = new RunStyle() { Val = "DateCar" };
            RunFonts runFonts53 = new RunFonts() { EastAsia = "MS Mincho" };
            FontSize fontSize50 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties18.Append(runStyle11);
            paragraphMarkRunProperties18.Append(runFonts53);
            paragraphMarkRunProperties18.Append(fontSize50);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript51);

            paragraphProperties47.Append(paragraphStyleId45);
            paragraphProperties47.Append(paragraphBorders15);
            paragraphProperties47.Append(spacingBetweenLines52);
            paragraphProperties47.Append(paragraphMarkRunProperties18);

            Run run55 = new Run() { RsidRunProperties = "006B1D99" };

            RunProperties runProperties29 = new RunProperties();
            RunStyle runStyle12 = new RunStyle() { Val = "DateCar" };
            RunFonts runFonts54 = new RunFonts() { EastAsia = "MS Mincho" };
            FontSize fontSize51 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "20" };

            runProperties29.Append(runStyle12);
            runProperties29.Append(runFonts54);
            runProperties29.Append(fontSize51);
            runProperties29.Append(fontSizeComplexScript52);
            Text text33 = new Text();
            text33.Text = "NOVEMBER 30, 2005";

            run55.Append(runProperties29);
            run55.Append(text33);

            paragraph52.Append(paragraphProperties47);
            paragraph52.Append(run55);

            customXmlBlock25.Append(paragraph52);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "00F34666", RsidParagraphAddition = "002E7D22", RsidParagraphProperties = "00C913B8", RsidRunAdditionDefault = "00740A1C" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId46 = new ParagraphStyleId() { Val = "Titre1" };

            paragraphProperties48.Append(paragraphStyleId46);

            Run run56 = new Run();
            Text text34 = new Text();
            text34.Text = "MORGAN STANLEY INVESTMENT MANAGEMENT, INC";

            run56.Append(text34);
            CustomXmlRun customXmlRun10 = new CustomXmlRun() { Uri = "errors@http://hubblereports.com/namespace", Element = "ContentTypeDesc" };
            CustomXmlRun customXmlRun11 = new CustomXmlRun() { Uri = "http://hubblereports.com/namespace", Element = "ManagerName" };

            paragraph53.Append(paragraphProperties48);
            paragraph53.Append(run56);
            paragraph53.Append(customXmlRun10);
            paragraph53.Append(customXmlRun11);

            customXmlBlock24.Append(customXmlBlock25);
            customXmlBlock24.Append(paragraph53);

            customXmlBlock23.Append(customXmlBlock24);

            header2.Append(paragraph51);
            header2.Append(customXmlBlock23);

            headerPart2.Header = header2;
        }

        // Generates content of imagePart6.
        private void GenerateImagePart6Content(ImagePart imagePart6)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart6Data);
            imagePart6.FeedData(data);
            data.Close();
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings();

            Divs divs1 = new Divs();

            Div div1 = new Div() { Id = "110056252" };
            BodyDiv bodyDiv1 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv1 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv1 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder1 = new DivBorder();
            TopBorder topBorder13 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder1.Append(topBorder13);
            divBorder1.Append(leftBorder7);
            divBorder1.Append(bottomBorder15);
            divBorder1.Append(rightBorder7);

            div1.Append(bodyDiv1);
            div1.Append(leftMarginDiv1);
            div1.Append(rightMarginDiv1);
            div1.Append(topMarginDiv1);
            div1.Append(bottomMarginDiv1);
            div1.Append(divBorder1);

            Div div2 = new Div() { Id = "454059354" };
            BodyDiv bodyDiv2 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv2 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv2 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv2 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv2 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder2 = new DivBorder();
            TopBorder topBorder14 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder2.Append(topBorder14);
            divBorder2.Append(leftBorder8);
            divBorder2.Append(bottomBorder16);
            divBorder2.Append(rightBorder8);

            div2.Append(bodyDiv2);
            div2.Append(leftMarginDiv2);
            div2.Append(rightMarginDiv2);
            div2.Append(topMarginDiv2);
            div2.Append(bottomMarginDiv2);
            div2.Append(divBorder2);

            Div div3 = new Div() { Id = "802694763" };
            BodyDiv bodyDiv3 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv3 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv3 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv3 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv3 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder3 = new DivBorder();
            TopBorder topBorder15 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder17 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder3.Append(topBorder15);
            divBorder3.Append(leftBorder9);
            divBorder3.Append(bottomBorder17);
            divBorder3.Append(rightBorder9);

            div3.Append(bodyDiv3);
            div3.Append(leftMarginDiv3);
            div3.Append(rightMarginDiv3);
            div3.Append(topMarginDiv3);
            div3.Append(bottomMarginDiv3);
            div3.Append(divBorder3);

            Div div4 = new Div() { Id = "1232883998" };
            BodyDiv bodyDiv4 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv4 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv4 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv4 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv4 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder4 = new DivBorder();
            TopBorder topBorder16 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder18 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder4.Append(topBorder16);
            divBorder4.Append(leftBorder10);
            divBorder4.Append(bottomBorder18);
            divBorder4.Append(rightBorder10);

            div4.Append(bodyDiv4);
            div4.Append(leftMarginDiv4);
            div4.Append(rightMarginDiv4);
            div4.Append(topMarginDiv4);
            div4.Append(bottomMarginDiv4);
            div4.Append(divBorder4);

            divs1.Append(div1);
            divs1.Append(div2);
            divs1.Append(div3);
            divs1.Append(div4);
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(divs1);
            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of imagePart7.
        private void GenerateImagePartTopicRankContent(ImagePart imagePart, string imageStringData)
        {
            System.IO.Stream data = GetBinaryDataStream(imageStringData);
            imagePart.FeedData(data);
            data.Close();
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();

            Font font1 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "20002A87", UnicodeSignature1 = "80000000", UnicodeSignature2 = "00000008", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Wingdings" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "05000000000000000000" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Auto };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "MS Mincho" };
            AltName altName1 = new AltName() { Val = "ＭＳ 明朝" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "02020609040205080304" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Modern };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "A00002BF", UnicodeSignature1 = "68C7FCFB", UnicodeSignature2 = "00000010", UnicodeSignature3 = "00000000", CodePageSignature0 = "0002009F", CodePageSignature1 = "00000000" };

            font3.Append(altName1);
            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Arial" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "020B0604020202020204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "20002A87", UnicodeSignature1 = "80000000", UnicodeSignature2 = "00000008", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Tahoma" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020B0604030504040204" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            NotTrueType notTrueType1 = new NotTrueType();
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00000001", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(notTrueType1);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Arial Unicode MS" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020B0604020202020204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Roman };
            NotTrueType notTrueType2 = new NotTrueType();
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00000001", CodePageSignature1 = "00000000" };

            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(notTrueType2);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            Font font7 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number7 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily7 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch7 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature7 = new FontSignature() { UnicodeSignature0 = "A00002EF", UnicodeSignature1 = "4000004B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000009F", CodePageSignature1 = "00000000" };

            font7.Append(panose1Number7);
            font7.Append(fontCharSet7);
            font7.Append(fontFamily7);
            font7.Append(pitch7);
            font7.Append(fontSignature7);

            Font font8 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number8 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily8 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch8 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature8 = new FontSignature() { UnicodeSignature0 = "A00002EF", UnicodeSignature1 = "4000207B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000009F", CodePageSignature1 = "00000000" };

            font8.Append(panose1Number8);
            font8.Append(fontCharSet8);
            font8.Append(fontFamily8);
            font8.Append(pitch8);
            font8.Append(fontSignature8);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);

            fontTablePart1.Fonts = fonts1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "ppelletier";
            document.PackageProperties.Title = "Product to review";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "4";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2010-01-19T15:22:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2010-01-19T15:24:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Julien Blin";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2006-09-26T13:33:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data
        private string imagePart1Data = "R0lGODlhAQABAJEAAAAAAP///////wAAACH5BAUUAAIALAAAAAABAAEAAAICVAEAOw==";

        private string imagePart2Data = "R0lGODlhvgA4APcAAAAAAIAAAACAAICAAAAAgIAAgACAgICAgMDAwP8AAAD/AP//AAAA//8A/wD//////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwAAZgAAmQAAzAAA/wAzAAAzMwAzZgAzmQAzzAAz/wBmAABmMwBmZgBmmQBmzABm/wCZAACZMwCZZgCZmQCZzACZ/wDMAADMMwDMZgDMmQDMzADM/wD/AAD/MwD/ZgD/mQD/zAD//zMAADMAMzMAZjMAmTMAzDMA/zMzADMzMzMzZjMzmTMzzDMz/zNmADNmMzNmZjNmmTNmzDNm/zOZADOZMzOZZjOZmTOZzDOZ/zPMADPMMzPMZjPMmTPMzDPM/zP/ADP/MzP/ZjP/mTP/zDP//2YAAGYAM2YAZmYAmWYAzGYA/2YzAGYzM2YzZmYzmWYzzGYz/2ZmAGZmM2ZmZmZmmWZmzGZm/2aZAGaZM2aZZmaZmWaZzGaZ/2bMAGbMM2bMZmbMmWbMzGbM/2b/AGb/M2b/Zmb/mWb/zGb//5kAAJkAM5kAZpkAmZkAzJkA/5kzAJkzM5kzZpkzmZkzzJkz/5lmAJlmM5lmZplmmZlmzJlm/5mZAJmZM5mZZpmZmZmZzJmZ/5nMAJnMM5nMZpnMmZnMzJnM/5n/AJn/M5n/Zpn/mZn/zJn//8wAAMwAM8wAZswAmcwAzMwA/8wzAMwzM8wzZswzmcwzzMwz/8xmAMxmM8xmZsxmmcxmzMxm/8yZAMyZM8yZZsyZmcyZzMyZ/8zMAMzMM8zMZszMmczMzMzM/8z/AMz/M8z/Zsz/mcz/zMz///8AAP8AM/8AZv8Amf8AzP8A//8zAP8zM/8zZv8zmf8zzP8z//9mAP9mM/9mZv9mmf9mzP9m//+ZAP+ZM/+ZZv+Zmf+ZzP+Z///MAP/MM//MZv/Mmf/MzP/M////AP//M///Zv//mf//zP///yH5BAEAABAALAAAAAC+ADgAAAj/AP8JHEiwoMGDCBMqXMiwocOHECNKnEixosWLGDNq3Mixo8ePIEOKHEmypMmTKFOqXKmQ2kCXAmH+k0nzpc2YN2cerImTpUNqQG0G9RnRj58pSJNOOaoUKdOmT5VGTXrUj8yZTZ36SXWV6ECjSFG4pBZ2qVeIU1CoXcu2rVu2ad/KFUuQWqq4KJbi9XOWYCo/a1P9S5UqcF+HZdemXTxXLWPGjRd3Hez45eIphweulQkYBd/MCzvrrTq6NGnSTk2b3mxQNGfWmeMKFigb9MKws0dKNlgYad28KHInnKyz5ULiwF+btZ2wNknYBF2/XMuXsNHPM/9ufXlX69W/VLuC/0cq/J9z82rLMyeIu+Tugr0xT08L9KhjmcljpuX6j+nAtHzZ5Rl7S9kFoExhKYfdegV1pt5H0H1133+VzUTNWtPlJRg1gE3BmXydcScfeiMWFhyFJ9KWHoMHtUfSe37lNaJ9HmZYo0D5Xegbe1jlVZB8d9GlWVr/0Sfhciyyt+Jzak3WGVxT8JchCjYKZqKQOF2oVkEwxXWVgyqmiJ6YSdK2n3t5TRafjgNyOSGORpK1mFU/AkdYQXZyxZWLCR65YJlgMomlhGle2ZWWVMbk2GxPLkqoY1GOpVhWy513XpnmnTkSWU3yVmFcIyqapmadCnQlqF/hVRmiOo1V3D/Qdf/2Z5JxzQphqdFNGGSboiY6k6N+NSpcd4FxOqhBli6JaaZkhvRmjDvCCp2WN8KKFHHM2irrmNjCieW2y/anrEgRCiTdr2phhyhMJkaZ3Xfp9gdfuhzG25yyl5bpom6jzhutuKVS2yVrZM1LF53zbfgsVkN+a++y+Upkq0Hljulur8G5BCpYtWI1Yr01FojiQO1u2Juk8QJFXbe2RQzRUSzj2K+pwDGWKnABgooZjVTaxxSv7T4Fn6opJrZfzTMnGehEnD6IJ64WAiX1VVPLRNhsVcdkF2GHXj0Ul1tLqXXVZIe7r8QyMgRjuGxH5HJDbPqoUMVt163Q2a0lxPOxdVb/a/ffeo/rZqhTUjc31IAnTmCz0DqJ9L/IJt1q1Iqz/fapBgMn+NOD7mo4aAi3HPPdmrbm2KBGy43t2jN1KNboKBW2HswVgZt5hZS15S7tFCMOMOFndQY7Sl5WhLe5um+I15zoIic5ZZCfFZZt6xq/OaJsCdZofucSVHH3QEl5tVCvZsndVg/ahT5Mxh5qV0zony92XX/NbyF/HN7pve/W6Y/Q8WE53WdAhTNvIYlUfqOZ6vRSoLJ4aDxL+QyHnOIS8HToT915SpcAVCBiPbBmOcNZV5jClNncpYRyyou6FGOUVDFwYgArj2iy15/laeh3ctsf34IkH6CEBTBbuYwP/1dloU4JT2VECpNlQuiZOw3RM8F5knZ6wysSZeeGSAxOcHaFHepsRWViOkpzShcmzckHMiocjA3LszAFhopNWJuQz+piliBZLY1N88tsptcgDJUxjtHqUK6E5DoTLqZXCEqiZu5FpiuxpT5v6ZHuOOe4mV1JhyjD0l1UVC1WaWlB72tfnYD2LF2Na46/Yx/uYGMs+IyRTEhL2eNOxCbGdIV10Hsj7kg0sIv1J0BpWxy6fCnMPvoySAsCDvL8piMQBbNHNaqemTLWkKVtb1REw6ZbekgqvokmRjmUFiF1tbtJJaVYbQmdKJUkpnbtr0bmPOeWALbCeYrzNZByWpFSxP+px1TnMfFql2K0hx1cxuc3/4KRnShDM/Jc7WqfoSKoNugrQuUGmfhJYq3qt5WOLlNdu5Fm6x6nT+dck0r9bAu6JqklmHyPXiR72D1juqU2fTM0aJTkLUt30G66hG6ECtUqjQROtkysPbVcaFKNJNH89GaAz+upohI6s35eCJDA28kE48XHXB0TQN5LYof0Ka5oWXWlxBEQLmNoMcX81IbyiWWbavZT/u2Sh4u7CgHZt1ZT/ehAxWPnRXcpR5maLpzPnBbV6OZAgO7OsS7h2W4kSyM1PTOXSmLmQlGknrFgcEVdtSjJDtSwhnXrgpZRHXQSuNZswmWlr8WeYnInl0q1VsudPqUj1BCVmwn+Sq9bEiWClIVbmQnpktPZI2HtSTCy/MmwKLJhpxzJmifhhZNvSaB90mUlKQ4saYLk0g85Npb91OdayPMMzLY6KmOJTJzbyV2UjNImz+FvoyTailiAGD6nBO6c57SKnJrClQEzRnmUSsp3phJEAtNGg9OxFXupaS7rqhNUBG1KpqRiJqawS2eMSg3vGIgZ925ouzesnIpXzOIWu/jFMI6xjGdM4xrbuEwBAQA7";

        private string imagePart3Data = "R0lGODlhmAAhAPcAAFxcXMnJycbf8tzs922u35/K6sTe8qHL6oK543ez4Weq3aPM6om95ePw+a3R7e72+/7//8Ld8mdnZ9TU1HJycszj9PT5/bjY73Cw3/X5/fT09IeHh52dndbo9o/B5rOzs+v0+oW75KioqLjX797e3r6+vnq14n19feDu+Onp6ZKSkqPM68Ld8dDQ0K3S7f39/c/Pz5nG6Pr6+v7+/u/v79LS0tjY2Nzc3Pn5+d/f39PT09vb2/v7+/X19evr69bW1u3t7eTk5NnZ2dXV1dHR0fz8/Nra2urq6vf39+fn5/j4+Gaq3c7Ozv///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACwAAAAAmAAhAAAI/wCZCIxRYmBBJgQNKkR4MCHDhQ4jNpwIkeLDixIrasRoMSPHjR4dCiTRpCSNkk1OmkSpMiXLlytjupTZsibMmTht0rypMyfPnzuD+jQ5QOAAlEiTKl3KtKnTp1CjSp1KdSoNgVWzat3KtatXpyVKfM16VGrZqGehpn26dixUGjGYtHTbdICIJUtMiGiLsgDevDH4lvSL10QMp4QBI/5rggldqUaVMjk7gERZyyVJNqkM1cSSAjQGGPBcQCkNDJ9Dj15iwDRe0KI9t6XxOjQT2UoNLBFxwfZuwUlJqHTMeTPmJpqLsxQ4tyQAx02YAJjOoYkEEdElnARQffHRAoc3L/9Z6/d7+PJJPQ+G7lepXpTh1SddcngA+KPvoXLfDMD6hiYcTAdADCQAQBIHEkgmllLPRccdXACIwAEFAAIwAA0GOjXAeCXdpVJ+Jel21G7ygdhEebTFkJ8JDCCFXhMelmRCayjlV4AIoKXEmn7VYUiDBBvEYGF0FAww5IRKxdVcEwCIhSBKIkhQoHUUcMBEgk6piNKGWzYGpQkdjnjSaC19llkTWkZnQpkXlHTbEvCBiRKHTeQFHZpyOgWABHz2JwGC3VkmXZ8UNBdZUg0+2WGC3EkgXZVPGcYSnHOWVhIGNA5GaZ1n7Riipxt2qhIGtG25aamD6UZje08BQAETIgD/8GOA2AUIwAYDMrEBlsvdiVSD0olFwp8VVjfdkknFYKl4Z3lZ0hJtlhRqjZ2mdZtJm3LahGx0MrEpl0ileW2r3fnJAYInYViCrABSmFRYS+1pYazTuSukYxL099SNk6o07aXYPZsjs3Meptu2AbOqo2N/7TZYwDq2ZkBpLMoIcVP7+QhkExTIy8EAThwVYFJwySWZCOGREBdSJjORaVMkLAEduHiyTGnDjqXpppwx4FWWmXEitSnQHYK5IV7S0gmWZti57ObKaJ6EVVKHPqZUjEipmtRdki3RnAgYXJ2teHlCOfaGy0orQtlWk8xc20utHdhmPfs6p2EXbugsUnoH/0bDBXYqpfdkf+e15OBF+cU23CwvyHhSBqBWGLKf/jWw4LpZTvkADvyFY10MeJ724ywpSfrpqKe+VNWqt+76Y1fZ/frstAf3MlPw1q777kWZzFTJhlZWlPCTCUr88cYnP7zyxS/vfPPQI/+89NEzT/311mc/vEAjNcV67NyHL/745Jdv/vnop6/++uy3P5lpb7NMAvnzj1+/+PeHnz/3+49E///2AyD+BKg/AvLPgEwoQVsSSLKQHCR8DhxfBMU3QQh25IHcq2AGLyhBDlIQg6yDj+9QUr3paa+EKMSeCVeYwhOqsIUsvJ7+vPe+3dnwdQkcIfxkd8MePq6GYHGcD0iHqDvgEfGItQshEpdIutil6yVPXEkUXTLFJ9qkilSEok+wyMUtevGKX9QiGMcoRhJuECQenKAa08hGNLrxI3Bc40JyCEcmBAQAOw==";

        private string imagePart4Data = "iVBORw0KGgoAAAANSUhEUgAAAiUAAACfCAIAAACpwT6GAAAAAXNSR0IArs4c6QAAAAlwSFlzAAAuIAAALiAB1RweGwAAZi1JREFUeF7tvQmYXNl1HobuWntvoNEAGvs6M5jh7MPhcBlSJCWKn2hHEq3IkkKJURbysx3Fkb4v/mI5jigpUuzYX+RYW0RLTkQrdiJLImVJFMVIorlI5OwbZzAABhgsDXSj0Wt1ddfenf/c8+rW6XvfVtvrauAV3tRUV923nfPu+e9/zrnn9m1ubu6KX7EEYgnEEoglEEugyxLo7/Lx48PHEoglEEsglkAsAZJAjDfxcxBLIJZALIFYAlFIIMabKKQcnyOWQCyBWAKxBPp6Kn5TKFXzhdJasbJWKK+XKtiK5WqxVCtXq+XKRm1jAwrr6+9L9velkol0MjGQTQ5l0iOD6dGhzO7hAXwI1GhfX19gm7iBlMBGtVQtrldL+WpxDe+10nq1XKiVCxuV8sZGdddGbVd/or8/2Z9KJ1IDifRAIjOYzAwls9iGk5nB/mRGHi2Wf/eerlhT3ZNtm0eOVcMC3H68WV0vLeeLK2ullTV6r1RrtQ2A4CbeN7BtOu/yG3wG9NDv4r2/v3/v6MDe8cH940NTEyP7dw+FN3OxETS6E6ClvLZUzi+UVueLK7eBMRu1yuZGVSmjBvHvIs1s7nKSTTYxCsCzRK/+/r7+hNqSfYlkMj2YGZ3Elh7ekx7aDRCy+20s/HZsWaypdqTX1X1j1bh09u3iN4urhaVcYTFfWMwVgTCMK4QxsGf4UHM+NCCHsYeQiMCG3ukjfSB8gv3T4LSxCRN2aHLk6L7R4wfGT07t9nqqXC3d3Wz+gDGl3O3C0kxhcRq9pVYtb1TLm7WqAhuopEbiVnogpFFgQzro438gjoQ3CnfqqJNI9idStKUyID0Dew5lx6cyo3uBPbHw2zF2sabakV5X941V4yPeqPlNvlBeyK3PrxTw3qAsG5vVjQ1gDHCHthr9SfghEEjDjIKWXQpoyN7xZ/zbAj8MQpt0nFNTu08f2nPvkYmpiWGH07m51Gzzd/cATxUuzJVbhYXp/K23a+X1WqUE+r9RhQezArBhpGGMcd6V+F2eKhKZAh6GHJAeRXfgbesD5CSxZRKpbCKdHdp3YnDiSGZsH0CIWLalkbtH+E3ZvlhTTYkrysaxasJIOzq8AaGZXwbSrCM8w0ig4GSjWtuyEdjgm/qv0qvGbEa9bTF3hDY00lb/aMTN2OPQIOZMowOZ+47tffDEvvuOThi2LNDY3cG2r5xfXF+czs9eLubm4DSrVYobABsFM9AMERrFPFnkLGYStPMSkIOPEjOY6zCSAHWY7iQU8CTT/ck0ICehXG3D+06C9GRGJvTDGqidMI/1ndcm1lTP6jRWTXjVRIE3oDJzS+tzy3kACfvKGGYq1Y1KrVZV70gHoG/wZx1+HJcaNXacOBpvGrap7spRlk1t+OcYRQeW2EHHbAkW7+FT+x49feChk/vYrvlbtzvY9iEws3b76urN8+W1ZQzNADMAm80aOA0TmnqEpoEudak7uOLIWWdf1MGHsX8r+6nzHkIdYA+cbIQ6GXaypQbHRqbuGZo8nh5pDAXuYMmH75zcMtZUsxKLrH2smmZF3V28QSLAraW1uaU8oIWQRsAM8gJKFcAMcs+cd3wD1GEEYtTR4RwaZDOv2Qo1HCwgx40Tq278WccdGo5z4IchB+P2VKL/iXumnrr/0H1H9/LxbOyR9u4Os30AmLXbb+emz5Xyi8RpkGxGSMNBGgkzxFmUdEi6TmxGucscgUjgqScOsI9T8VAmoeyCIwhyVNeI7gB1lIcNXCczhIjOyKGzQ5PH8EGOA3wU0eyzvuPax5rqWZXFqmlNNd3CG2Q231rKzyzkkdAMO8/4AdQBupTKQJoqNvyEz5TxXEG6M6EOEAioA8hR+NQI5NhgowBmlwoOUEoUPDUJOGz4A20qTwqNFJRg0yEfDg7hfXw4/Z4Hjrz/oaPIZNNGzQAeH6TZiU62Wrm4Nn81N/16YWm2WszXQzWUEaAiNAQPanMCMApakGkGVKd3ZicIyWCasCPWBqorUqPyNuhd57A5WQYE9c5PDvCoU6hMNuY6yJxG/nRmbP/oobODk8cQ5pG6uNtQJ9ZUa+Ysgr1i1bQj5K7gzdzS2szi6tJqkQP+Cj8IS4AuhXKlWKrifb1I7/yZUIchp1onN/V8AZ0dsMWvo/xmZAbpnfGGQgTJBKEO3p3NASHGHUdKOg2Byda9hyc++Oixpx88wmPqMMCzQ21fYXlm9cabqzMXK4VVONBoAo3KPXNyARhmFIsBBmBKjYq4JFWoP8nJzXjn7xXjoagMc0N6dwiNCvYQG4UblP1yVfUB03T4A+UdcJ6bw3jodKQ6lUqQAdFJDYwM7T89cvDegd1TBvXcoZJvtn/GmmpWYpG1j1XTpqg7jDegNTMLqzcWclXYFkVr4CIDlcH3CmNoQ77AeqmMd4ANZnQqrrMVZth4qZd/NVEdolZAoVBHURwGnlQigQ8pZOQq4GG3G7l7lHtNh5GSib7vfOzER9958sCeYT2mdh1ce9m7Huc6GJHl5y4tvf0SApsAG/KhVUubVZXiDHgQhIZxpR/QkgDnSKn3NKWWpQdTw3uSA2P9meH+1EA/yEcCv6Y2lavN8XXiaBT+qWxWisScwJ8KK9W1xVqlsFkr0+lqZfqVzssI5OS8OS47aIniOoAcIjrpoT1jxx4e2n+Sic5dwnViTbVpzrq3e6yajsi2k3iDDLQbt3MLuYJDaxSnYZhZK5bXChUkQ9NWLOMbIBCQBpwG7EeZfmeEHIgxXrfN2ON4guD1wThccB3GHkYdbslJBDo94ZFT+z721GnEdaR1uwNQB3M2V6bfWJk+V1lfUT40RWvI4gPRATbar6VgBq4tYEwqi0g+3jPjB1Mjk4mB8UR2pE+VCVBKckYDPCaQQbUtATbFFzeAPaXVjfWVSv52eWV2s1rcRKZ1paSQqbxLUR9nNg/TS4K6lCpSAKIzOnLwPgR1kL3mOg7gJ6HHwT58L401FV5WEbeMVdMpgXcMb24urE7fzgFI2E8FpEFghpEmt1ZeLZTy6/QOWlMAp6EQDqWiobEDNEFUJvwNS+BhxkNI009Ehz8w6jDR0SwHF3Ngz9D3vueej73rlAE5/qjTy06e9fnrS1dfRh4awMZJdwa92IQ7y0AaxO3T/WkiLqAv2d2H0kCa4b392WEnJ1DjTB1jBAFtJHFIvDFgAH/WCrna+mJlZaaSmwEOYQMTAurAz0beNp5DSqIHRaX5oUiYTg2MDew9Mnb0IUzWYaUEjgDCPyc91TLWVE+pQ15MrJoOqqYDeAOicH1u5drcCuFHjaI1FKcpEZvJrZdoW6N3/IkvKUeAkYbi02pWR+eQRsrFoDuAGQYeHdph24Ur0HOAkKQA39rfevreH/rg2Ww6JcfUrpbOC2l6YcQN7pa/9dbCW8+WcvMVGHoEbMAqVEEadqBxyRnOEAPSwLgnBkYHJk+kxw4mh/YgNlMfCQio8WA2Bq3ZqoVGtTpHLPDirS9Wc7fKi1drxdUNeNuAOhRJIj+bE9dRwTiapkO+tRHUIxg/8fjQ/lMYKmjUkcCzo4lOrKkOmrPOHipWTWflSd4IH2MR5mRwiAFpwGzYN4WAPxAFnGZ1vUz10PKoilYiWlNQtKbuPXOyZSNZ6c2Z8q4qrThOtnqAh5IMkLqmUqV1Bh3w8vvec+YT3/nAxOigTXQMx4705/QO0UFpAEysWXjrGZTW4IANDDrV1lS0hnPMlN8MHAJIM5QcHB/Yfyo9fig5uJthJhBsDE9amEdFy4c/bBRXqiszpYW3wXs2y7hCuNoU6hD9YqKjUtfSAwjnpAZ3j598fHjqngS+ofyP+lzSrZOoegHpw4hCt4k11ZS4omwcq6Yb0k585jOfafm4gJa3Z1em53MwUeRDK9cQ/0f9TWSmLayiYk0BER1MwcE3QCDOdSZas3UiTctnD7+jMwVRTQjhMboiV1ySQFYC4wJum29cmV9dL546uHsom9JnwQ4wZ4YfyWeUvV22D9GRlRuvg9lQdoByo6mATY2jNSrrDAGSLJVwHhgF0gwdemDo8EOZiWO7EhmoB0jD78bLpDn+iRy+umEZIiCUGNqbhNculdmsFijnTaE37cq8l7OruZoO/G75RZAeTA6lNDnxMlSwXWIP/zQ2wCbWVAtSi2SXuBN1Scyt4w2BzcwSstEo47m6oWhNBX4zYAyQBhtQB+QGSFMoVytUPoCMBxuT7XnBe6agjiLlziRQ5xsYOYYixXWo8vTF6wura4XTgJwB9zUOjNG6K9GJ3vZRP7n++uKl5zGXswo3GrIDVKlNx4fGCWDEGEaQbDa479TI0Uezk6d2JbOuGGOwHAm98n5b0KaG7b5kNjm8DykJRLWRQdAQmR4dqIk7qrJOZW0ZCdkp1PpUjjXjvLZGWriwyHaJNdXmI9Q9TcWq6Z5qWsQbuNEuzywjRwBgw8wGuLICZpMralqTK5SQL6BoDVUKiJ7WuD+RjDrOGFqnWjnTexhy1Gvj4vTCerGECToDmQbLsY/pjNa9HTuRoQ48AMvTry9dZrBZdVLROOMZ3kSqmEkTXJIDI6nB8ZGjjwxOnU2QA428Z/JlUBk1RNgyRtBuRpaGjKkEfnZRCkhXZjgxuAepChulvDMNiGN7ToUCUoeaLloDYyOWM7yHpgGJs+vP+viRib0Fwxdrqnecz4b6YtV0VTWt4A3A4/LM4vU5J2bDzAbRGqwsgFJpiznlQ6PsAEp33n5aY9kDZ2RcL/pJdKfhVeMMAva7bVy4Nl+pVh88uR9ZBl6GzBhWu46yI7B9AI3cjXOLSBBQYFMtr9NMF8RCCBCQloeMryyWn0HGV3bP0ZHjj2UnT2z2JX2QxoYZDS36A4dS+CU/G6gjMcn43JBqIgXfGhKvd21gpk6FmnGxA8fvyeVXiX9W1pdxL6lhlPik89jmvseJTqwpuytF0EHCDAti1XRbNa3gzeWbS1dnVzAsJjdapYpcAPjNllD+ObcOfqN8aBV8r1dO2zYHmv8jVic6DDbKwAnqo/AGP7xxZRaBhUfO0Fz3wJdPLKHbPWp15sLChW8iyAE3mhvYDCSw4Obg+OCBe4YOP5gcnjTiNDankTcrEYVxxXg3gIeSAX2hiFFHAhifrj8zkhgcJ41UCuqLeokdh+g4ldmqgBxkSyOPTr3sEVkvQ06sKam1ntJUrJpuq6ZpvEE22qWbi44bTYHN8loRYKMCNgUBNsr5H3lqQCAkbDWjjk3jWi5qJR0ygirGw/EcGlOfuzo7nE3dc8Qp7mnbONeTGthjmMWmrjOwMaYIzF/4a1SrpQSBcsFkNsgOoBSv8eHDDwwevL8vPexPa2ykkQCjsYQ/GK8ELTrQePGONkRpAmQDDyI6/QPjcJptlFadAUCd4jjJHmpUUC2uUr7DwKg/5HQb5gNVYzSINeU1ONh2TcWqiUA1zeHN3PLapRuLiP/DS4Y35UYrLSMbDalouQLm2eAbVA1ARKf3waZupxyDwAEK9s44oQMycxTcqFar128tTe0dOTw5ZpgPr07iRXS60alKqwuLl54rLN5oZKNJNxqlog2D2QwffXjgwH2b/SkvsFE3rjLHBPPQgOGKLvJLA2n0TzYZkt42TYO2nBcZdMggSABy8iqOo6aCMvmsh94o6FQtI4Eb6Q87BXJiTfWspmLVRKOaJvBmrVB+68Yiss5AbgA2KLjppD4jYINUNEp6JrCBk22b89CaHXPWcYaTCBovZmfkaarl8oVcfv3s8X3DA1TWRZtmgwpIEiMhxx44tHCNrrugrNPS1VdWZ84jqlErrmFdTlUSjRYRqMdsRghsjjyc3X8vyp3JPDTF31Q5Z5ERIMFGY4aNJfob/uAFNv7kxg7zNASIQw6MIXV7o7ACtkk1QEX6AKMOUu9wm+nhCVTBce0tPeWriTXVs5qKVROZaprAm7duLt6Yz6nsZ5rUiRmdyAsAreG8Z5WKtjPBxrFVmuAolxpDj5OqRl616bklZEQ9cZYqSUtcMTAGexhmzieo0z7qYF7n0tsvVjCvk+bqo4KATn1O0crNyHtWYDNw4F4YbB2z0UjD96fvSMdpGCcMUPF3nRm7aB+aATkGxsg/Tami9lB2FAnQG8WcojhqHqhKt+OyFNATCoMikJMeabg6DQbZO5ATa6pnNdUjqsGEgIVX/3Tt5uuF2fN6W599kzd8kxmZRPV0nx7k4sPvsU4UFm+mb6+cv7bAiz1T9jPlCCBsU8RsG+VGw3TOnQw2DoTU7b9eJ5TMmpOJCy7w1vTcvvHhEwcbix/b/MZAI1eW0ymvWmFphsM2TvYzKgg4kzqTKhsNbrSxoUP3kxttV5/NbAxaI11nBmuRwCOpjFczw5km4zeGg87uPIw6jtwAeVnyYdYhRzvWeEjgQE5ycA9uVuuiByEn1hQPDnpQU72jmrWb58qrc6royZZXfY3DPqyBq/HGiJLaroLe7ESh8AalzwA2qOsMvMF8GhQRwLxOBhskCyD1GQi0I91oFr+w02t5Lg5VvEFdyWplZXXtHaemRgYdr5ph41z9Zl2CHHICXHl5fV7V4lTZz2rsDw9TvfLYwOjg/jMDBx/Y1Z80pnNKWqPtu0YCjSI+qOMazrGxxI7icBvdPeRnm+vQNxigZYapzg1iOZS+odfO4ZxCZ82e1PAk6lt7qcMe97XPLMMfIdZUz2qqd1RTKyyvXH7O6QISccT0tsHJExJvdmInCsYb2KZLN5auzcGTphKgy9V8obScpzoCap5NCfNsUHOMHOzbWDsgfO8PaikgR9dWwVCaairDWzU7vzKQTj5yz2EvjuIFObZjp02Wk599a/nKS6iQhlUGyJPGyzaTF4zq+Sezo5k9RwYPPdiXHtRVaqQbTZpmg9kAZgzI8YrQSPDQfjPZDTSc+KSo2ZBjOtbUAjyUO1ApQhE899MpRqSaohQpio0id8AQvr4SfcA2ZR707Lj/HmuqvqZrz2mqd1SzfOnZainvPLE0jKpv6isi8n27hvefBI83vAW8y07pRAF4Awt1e2X93NXbtPImrTJAtTgpJy3vFEYD10Gtmh4qH9CaSdi6V2MZN+d7KnFDFEexnJtzy6cOTxyYMDNxbUQx7Wb9uK4tm7rwUn4JOWnFlVtUtAaF/Wv1GpdY5YcWj1EJacceRYEyWRJNT7KRYCOTAhhpfPAGCweUlmdKKzP5G28U5q+ik6xOv46tvDxbWJzGxYCIYGVoWkqn/nKFHNce4spvHPqVxApv6Wp+XhW2UespOCvFqVvp60dtaVqqJ5WVdE3CTPsyb0pBuvG2aAqzlwrzV4rzV9dvveXo6AapCVslv1hcvkmurcwgFlSVMr/bNLUtqpHhTI0c63OX87MXbE/alh7Rt2to34n04KgxyJNtjM+6L/ROJwqoDw0L9eKFmbdnl8FfsFInajzDgbawUri9sobZNljYBnSHsp9p4ndr/XHXJ7/7kZ/4+JMt7PzEpz/bwl4hd+GlCoiwqSoqSL2FWQeNoJUxa5WPPnX/T33iu/SoX4c0JFGwTaqEHzZ/rQ23oRR40jC7kybcFHIb5aKqqYwyAqo8GsI2Q7uHURht/30+YGOwEB//GLcs5+bQKwoL1zlnvLFAt8sE/11De48NHzidGdsvpa0TE3RGnDHJVJemNorr8J9oXJh9s3jj1Y3C8gb8hyi2BsihxXKwbMEQ1lAYOfbY8OF3aLCUrjxJdFoWe8gnx7jliDW1fvttqKmytqhDkD6aygxPQE3YtGr44u8GTW1LJ7JHVJA2bMvMC39Eyw8GvfY98CHdp3ZoJ/LjN7gllON8/eptrtVfr1tDE26AOkgZKKEQp8p+bhlsIOFHTh9419lDQaJ2+f2zf/xCC3u1sgslq1HkgOMHGGJfm5k/fnDP4f00v10OnPl54lPYiCK/aWe4DZihtQaQJqBz0hxPGpEb1OLM7j0xcODsZl/CSEhjU2IMaXVemSQ3MnIDz/LypecwNK4Wc+r2+E3lhwmwaXgAdvUhpLQ29zbG0Wm4ubK0Sre0+FoO/CUuwBl/yXb1z1Jf5FXDygXlvFqfTQdyVBSnvx/p0URx0gPG8W11aOG38jA0s0+Umiot3Vi68A2QTngdHeEFaQpTgzFtC5rKDO9GYMDQgnyG7zxNRakaezDHjz2/566/VlyZdZ5S6UkTn/lX+NM4frNzVdMoC2b0I8ZPhG1UDTRarxPONCQ985rQajEbYjbGyKiZzrgj2jrmlVYJ7cMCZRAXtr5iqfLV589jHiithV2v3q8/1JOoXRaS0YNHYwgZXhbYcR3rdeaX4BKpLzTAkzRppE8LdGaGVclnZ3EBfTFeYGPAjOFSW5+9cPvbf17K3XKgVGCMBJjGIFrcSXl18dZrf7F89WUwQgNKZLY0fpKkxOt6sAugKzN5At5C1CDo60+RG43uCquRV0DyUOSmuDhtr6egR4L60qJ5aKPUVH76teW3voVHooHroTWF+B/UtHTlJTe4b5hFVtmdoakoVcNeEJnVqcEGHzCTITd9TnclLzvgNNiqoZ3YiTzxBnc+u5i/dmuFl7/EnBvkBSBFDWBD5dHKqMW5QdymtyvWhLfjni2p36pVqWlBTPrAo/pvvHj++dffZrwJuWCMkRXWmtVDhbTczAWEFhEk36jxMsyUJkBpaSms1Dk4sP9McmS/4ZLSYKPHsLobGNEa2TFW3n5h5dorDUKjZGT2DU10PD7kpt+Y+/ZfcglOYwTNHU+GDQLJVmbsYGbvKSzC1pdMY2oOLkblD1bhXtsoF4q3LyOBwqfWdWsyb+0pikZToHqLb34tf/O8cnHWqWfzmoKaUOzVGVXc6ZqKRjWGR9dOqIG0AfPcqcJuO1w17njD3fL67Ryv/QxoUUtEE7+hlTrLtMoARW3o1Vpn3FF7Eb70E+qA4hD2UM8ulMt/9fJ5jTdeK5XJEIVNblgK4YVIMQzUrVlb5rUGVOTcWQqzDytDY4LnwFhq/LBcLk1fAJ9Lj4mkx8wYf/FPABtEAhzXmbBfjuYsf5qPRkv5BYygNcuRwMMI5Bo9MjIXnDaJRGbiKG4TkSoqK4C1SvHCowiKUykiYQHpDBpvCInqLzsm0dVHMDJNrV5/bQv7bENTq7MXFy4+I8cEd6SmIlON8VRLNy8LFiu+F5dnpVM68JmUfmmtqR3UiTz5zVK+eH2OikAzuQGhgQ8NeEOetDJmolAEN1A6d0YDGjYCa2g9arUx4Oza9cyrFy9dv+UFORJppNGX7p3wSMOSrJbW8rcuV0vrIDeYDlQvXUNuDrU+9EB270kUu9QBdnle/aTKNAfbecWdJHft1bW5S7onMKfZgjRNqpYhhzBSTOc0Okwg8DAQYiUCus30IPyHDsWh+Bq8auUNPJ4L16qo6yNgxlBEkxfeYvNoNLU2cz54TNDMHQBy1uav6j1suqNV5j9E6GVNRaMag9zYjz1GSEtvK3LT0muHqsYFb9gITs/lkH6GCZ7ICACbAd4g9RlbY6GBJsblLUm0V3ZiS0uAo/gNURyeZji3mHvu22+xZfNZiZlzq7ysnh1d8LpvtCytzCEHulYBuVGrdjrlNUG5Un1cvWZsyvYmudoI6UYzbEdp5dbqzXN8Ge0jjb4dQI7sYD6uG1fgkY6+9Pih/uwIURysLc0UB0krlEZYqqwtlFdv+xTA1l7NZsE+/PMYjaYQr1q5+koHxwR8g6A4PCzwQR0ZR7CBp5c1FY1qtItYCsrpUOorfMY6vIacwz9gO1c1nvzm5vxqndxsIOcZMAN/GjYiNypN4K7wpDUUq6M4KpZTpzgvvH5Zx2/sKA77c2ykCY8xxiO4vjBNnjSqk0bkRlV0UTPwlTMtvedYf3bMOKMGG370jeilNA26kyxdek6DjXMBW1PRWugYvIsxfNbXpruifbWajckrT49MZCeOg8/Bi0ixK8YbiLoKilMs52btjGoNMy1ffFM7RqCplWuvSrDplKZgBKEm+2aN8cHO1VQEqpFeL0YX41FHjsbK9OtNPVE+jXeQaky8YVMIsJlGaU5VLQ3ONGBMEWBTpjQBBG54dmenhLUTjiOy1JQ/TW30GL385tsvnduSNWCkD+gQgkQdMo+iKnMY+EEbRMLzc5dV5MYhNypCTJkCnJmWHD2gp7AYwwGbMbjOU4FNX71xjuY5aweaN9L0J9MjB87sPvHo2OEHMJMjpB5dfQi6wxgDZ75svEtc5M/p8YP96Sy8iLSwdCNRrYps4DLNPF3xzxoII/OQd2Q0i0ZTmA5VUkm0/pqCgibPPr3/wQ9jmzjzLvwZ5qZWZ1zwxjCaO1FT0ajGIDcG0rAYOU7W2deO6ETu82/OX1+4ubDqLKqmlh5AWhpqpqm6nJwo0DFZZdKJyzeXnz13w9gCJ+VEN/9GdWv8pwqpIU6AKTioqYCcKGy1yd3DD545yk+VMeizv9HN7N5r/GTLd33hOgyBKihQoFL87EyjhaIzyIHOTBxPT57Cd5Lf6OvxitnYIc3b576KQzhnd5vLiZ+ANOPHHtr/wAcHJw5nR/cN7J4amTqTGZusrucAh/5PBobPicyAKz7JkaCUpDxgo00yA1CBTwlhG+Va3CBkJIBKgvRg+ieXtzFGmoZGAmXe2lMegaYwLECSlc9cTuho6uGPQC/pwfFUdgQbZA59DU4cKlMyvZ+aoCNoE7t43f7O1VQEqpGPnOsDhghZa+SG5t94K0WbFPnBuBizTeSdaAu/0YO+mYU8wIaCN7UaktPYn8ZFOfFlZ8kNYOa3/+xle2utq3dxLydtAFnRilWoF4jK6xev6uCNa2K060BbX6dNdOxb4DYoZIvygs6cG/qKM9Mw7SbVl8ykRjGZv88LbOw8Y4k0+qFEOSmq++k8lZ6yPPDgh8FpjJ8HxqcwiMZ7oAryM281a8hcggTJVHp8ClhLt69UwguykqcRufpuIRwWDk6tuSZ/Drzg8A0i0xRcXv5gAx1hURP7yvEl1AQ08r8pgJl/Ay/I6VlNRaYaiTFaSlqY6L/tpAmEeRR7WTUu8Zv5lXUiN44zjcCGJnuWa/hwN0Zu6tbXmd6gvDsqZcCZiHPu8vSlazMyimNU/jfiN7bJC/MMIecKmdB4WDEbqp4pwM60BONN/8AeDTYGIWCvlNdL+6ywl57n7JOjCQeaqyHD7rBi8N7IpQFcbw2JA/Bfh4ccdlDY158emaQQjspSY/emiuLQXJxK7lattO6VNRBG4C23iUBT6/PXfMAGVz525AEvHbGaoET/GywuOTPefZrZdq3HNRWBavzBBr9ioOD18If0doZ5MntWNQ280XZwbmkN1Tl5mqeK32D+DZgNER1VBbqT48EwsuuVNjybro+iBfyJvWyVcuXC2zf889OMGL4eaPOthaE4VAe6hOU7y8Q62ZOmagrAeABvQG76ULRfvPi555eXFbCJ//rta3RBHm40/AIssZmNVBBsGeIEgSqTSbd2Y6O3yHuRwJMeGk+PHcAZVZYa4w38iXB4VpEYXSuu2EhviCjwOltoEIGmaqU158I8NBVoudDAf1gQMnVqZ2kqAtXYApGPEJDGh9xglNDC8+a1S2+qxoXf3Fpag02rlxVQeIOUATXxUzv2OyiXHXUohTRqtK1mgJIlBAC/dZXwRkKO4VgzDB9jjPTq+AiB0QXrq1NaGpaBUbEKdqZRalYisQsLEAxNuB5QQo7hQOM/dQN8oKlnvmCDH8N0CbjUAikOSt346132Fq+7QKGQDIA2mYIQlEuN5coUB8WXlumPrRNxJAVkiWm8b/85jExTheUZH01B/oHuMuzu7/kEBw0pkB2hqchUw0LTMjFk6JMmgGFcYK8JqRHdrAdV44Y3y3kKiyvIQfwGG2Z3qrS0uy8NeouG2aPmJKexESTI2dx1+bqTgOsDOfYcTBk5CIwioDoncgTAOVUpflhJZltOclpiYNxAL28b7eCORBp+Lsll781sWBJhwjNohjC1f98IM3x27S2asfFtpEf2qFk4VEvN4WUOxXHwRovFdQ5ssx04TPsINAXDhBwN6MJ1G5yk7JXAV3KgsSJqYOMWBgc9qKkIVOMDNuD0zkDBkibGB2FGci2oqdc6kYM3eqC3ul6+tbROwZu6P43BhgsK3L3OND1sYSvv5EOzed68dvPW3MKy4VLTg2s2eTJTWUvbAAmv4TbSBEq525wG7dQUUKMoZKfBn4ZM6L50w5kmRzfSk6YdazbS8C6BIWJaoFos2+zz9A+MH/DvG14dz9hLjhONy2a8SQ2OJajQQFJlRavFNdQCEqgqVltfxPTPQJdaINI31cmj0dTeM09xirPrFuhMa+qOQjbufU1FoxovZuOfJgCwCUNJQ+qilzuRyW+waidWGWCwcVxqoDhV/vNujdwIBTpI4wRvHDpQKJamZ2lOu3Sj+bvUmnKmVcvrqGFDbiKmmE7whiIzCF0kBnbDpaa9QzzCkvyGkUZjj/6Jb4t7CN4DOUdgLqaWk0+wutk+oy9PX628O4ANytugcjRBLzNO8E2Vto6Jn8j6lWBvuxw12LSPOnyEaDTVrAxd21cLnikbaB9+QpUc3+jHyXgIoZdt1FQvqMYnTYDnsXVEp14H6Z1OZOYLLObWVbVHCksTxqjZJoQ3d0Mp6DA6d7IGlGlz4IYM3MzcvA4VeFWM9qc4PrEEVRCsQot4InhD+RoqeMOePRQXyI7Y6GX407wgRz6IITlHGCF1to2+SFcQRdaA4jdOUTsSD8KMtKHQ7LrNLCW57Ox1Et5EoqmOXLa/ulsbbveyprZRNf5pAsgVbE3aTT0GPaIawhs5uEOZTpp5Q6bRCeEw0akPrJu6xzuvMRDGWQKHy7/zHUKAc7eX5Po3Mkxtl7SR8OA6sja+rJXyapJpjfyZjfkiVMmGtvSQJDdyaKkzAozUADbchnqOP/0J/w3em23XqOFVw59YX67uT+PBEwkILjVFcYjfsII08Bgets7eUTSaav+aATY+Kek4fnZ3gEc08Bp6TVPbqBqfNAE4qLtNbmxNbaNqzPmey/kScxqVL0DAU4ef2JnmKE4xG2GsFd+4vbjMGONDbrxSBjRaeGFPpZjngmmNmf8KVRTFSfQrvNFPlR7IeIENA5LxHmg+mmoQSJVC5h3YN6UvW7sHaf1QgG69wpBCHJVSAXhWs+i9Qjh6rNC+M00faqdoKnf9DX+FBkbgvHa3H61e0BRUvF2q8UkTgAy7lCbgqp1eUM2WfAFcZW69SPmkiN8Q0jgUh+tzNmVx7tjGataNYjYNbxo8OEvLOWN2YUiKY4ONHVTA1EXOFOCxupItExQFOcksW1VpnV1dT3Jc01UFBc4WbMeBoMkZf0hkBrmEmhoEqJQBhTmUolEp+oCNQQrbecL1vjtCUxhu+w8IELzpSASuFzS1varxTxOAnKMnN9zxt0s1/UY3Q34a0IW8aWoTn4Ujp6u2qrcPLpCG53s4W2511QtvjLCNtIAGTshbl3qpqrg3D9sdVFH8RFEczMJJBoKNwaD5gdPvHRe5/3ROnA55zM2e1Bid6TtK0HxP8BtBOVlWEFa1qCVsAw9fQDswY99CL2uKCz9PP/cF1/LP8l6Gp043qx3Zvjc1tS2q8UkTgMTGTzzSjpxb2HfbVWPmC+QLSH4WYKMyolTZqZjf8MCADb2yVA7akIRW805oOgzLkREdNnn+LrWNcpEz0xS4benaDsWpf6efJ9tFq79p4TFtaheAjX9sAEdr2V2jr6RB4IA3omK3MwrgxOj6Oi6SMhrA09StBTbuNU2V1TJ32GZe+tNr3/xdMJsQqpnq7KC7RzQVvWr80wR47lTgE9XVBtGrxozfYPlOlUxK82+I4VDIQDnTYrhpmLr6p3ooGohRKBTs8Iz8RrMckqmoYcOmUL7LJ4yNI+aR1JmUfRHSiaQA0ePFP+kG+s+OP9A+5Tj5XHCmteaucb1+rP2j7pjHASj3oLWDpIGKFqxBK7XM+UP7LKc3NVXDqH55BlvIegEhyxEFPjM9pantUo3/ogOjR+4PFGM3Gmyvasz8NEy+4coC7ExjqFFwEwOO0L7CCLJUVFmGEKFULklEMZDGSFGz+Y3xYBnmj9Y+IGcat6orgp1pfe6KcQWd7mGMvn62bv79pCPDZ91tEI7myI19UgC1xHJJIiUOdbBX7yBN2XcNsEFV6ZBTesMLrUc0FbFq/NME0AW2ndzIAWhkncjkN6Uqih0ytaHuqeI3ceTGpXMx8UCmgHI3Yu4/IKFRREB61eT3cpStPzswYo2yNerUawrIZGg9it9ybdJpZjjQdISQd5DxjvC2I7BlmIWkAqvd+JxFXrZxR3Kv+vCoDtFbSaQh+cCbCt9gB2nKuCnErr2WMAh/+7Jlr2kqStUELjoQZVqarb5tVI2Zn4ZRopMmoMBGWbiY2WxVWZ3xcXFIFcbfSPR7TvIIzBcw/Dl66K1xiOITpIQtQ3hmnnYPl8yGcUXaZR8b3ZpZkXthFanA8EBgZeIwl2FiZ2MfnU7BnkOX8ueSO9qiDnN2o40YFmzuFE3ZtwmH2+rMW4E1JpqVz/ZqartU458m0JEu0KwiXFFnC/BE0onMDomlehXeNFw+ypUWQ05dGyQJojXOJA8qnUK8JpVMSO4ik6FdOY3hz5GuHpcnA4sONNKvOCLeyPpVuVjOSzcz/GkMPO0/oz5HQGg6zEJSI+3lPhkXgJvC6gNWGoVqRZl7CSWqBsuRYjfApiPCodVFe15TXncKK4nMAuixI6LoNU1FphoMuVauv+4lwzCLD3VD/j7HjLITmfnQyWQ/+9O0Jw1g01jcKWJJ9N7pGHzZmUaT2J1tI5tJ6ZiNDM94zfFkO+hj8uQwnGarONEag+LQYTBLSsrJYDP8k/ahyz87K92Fi88GHhAju9YyBYwb3HIXqqxcvWx2A3jpphMp/lsySCnYwAtutsGO0JTPTYHlzL72F4EkNaRY5OiHdtlWTUWmGoy6fGhiV0tzhtSLNgiND1GpZgu/QVfMppJcHFqN2+k/dU2N2i3hb+nObElgQ9ZNrbPCG5VOGRrIePEYEqOzlEPjgwQbe9CtRcc/9aezjXr7zm/s6VSxNRrgmy+NOl11oOmzwpMWmAHVpZEd9W0nU5yGAmTjVPoAvZIZKRdXoqMBqc3HdadoKvA2Ic/b574e2KyFBtulqShVg2QZn8lnEZTmbEEvaiQQUScy60MPZpLsTMOQkQpCK1eER/pPa7e20/dS7IbzxOtgg8/Dg4MGbBjMhiFHj7U1OPnYO20fE1g1Wb1oAemG/PhKkE2IbOktL3NcqX7snj8NfSyMJ60bdQlxU6oItFqmoTEXljP3UDZ6C95o4bvKvB3e07OaQhKUrImH1b7DJAdi6IABRGc76nZpKmLV+OfL9Ai5sc1FZJ3IyYfWpnBkMO0YUk71rUduYpeaUhK70nhCrCqgCV8WAUt1bGTIFUuMTGiDA/Eu2vxJQNLPBL7sTw+oevtqcKDDA3QsroJcX13YKlOhD+KKQB0xKCGHw0N7j4WxdOEvSd9RDUWgSRFUUMAJ5BA0U/Fs1PJ0PaCEFulqC392r8P2sqZwzdACVvtG3dXAekKIQHQqd6AXNBVNJ/LPlwlci73Nx6+F3aNXjclvxoYzKAitiqc1eiKPreOXggdOFae1gdSCNPQBkLN7fNTGEj210wASA2N8TB63TGaGuCRlI+GKwUZFjzYKOVaNq5Lkl93QInwvgYYJPQ1mrkvPT7WQo5QBLoKhxMCxLlU5e0DCdpcuQB62lzWlrxOkB6nP/tKATgOLEjUrz+3VVLdV458mAFltbw60v7IiU42JNxMjAwQ29frQTvxGe8WbfcTusPYqeKNcWGoD3mAmplopYGL3mIE3rt4zDS3SnxboyUkovNklJjayY04tLFbdKK16idlIE+i4NuBGC5zdiaH0vrNPBw6oW7g2vrvy2jIDv8gOd1Zq6E818IaPL5G+hTMG7tKzmjKuHFkbcG/6305gnYhAaegGvaCpbqvGP00Aoli/fY1rC/ls/iJdfvtlY9/2kwkjVo2ZDz05NujUF1CxbR2/if1pnOXEJp43zFimSctYCW2ztm8vFaB0JS4Sh7iNTWj87WASJZCdJZPhJlJDeKe6gfKnFZY3ahX5pHaDx9g9AemzYbz8sGvt56R59UP4ncurtwn1CW9UCKfuTKOV6LbijY3rgUgf3qRyy97UlOtdwL3pPwhAFCeQuYaXz7Zrqquq8U8TYClx6Q3/zV+e0IixO+oVhVdBL3SiLfM90f327x5MJvp5ZU8nR00tJonlE2PIoWoCitzAvmOj8lwb+FDNZNL79+1tltAYwCOpjzEYR8l9LMerJhCoNT3rP5M/DciHVcVECKf95y/MEULOthk7/EBnwzbGtVXWV7AEgEL9+uAIYKyW2e5PZvtS7vEbQ7xh7jdkmx7UlNeVh8mVCiSvIcWCZtuuqa6qJkxZjfCyirhllKqx/GmjA4f2DlFQQoVwuFqLcoe7VqiKWDLbfDo15YbAhpAGkFPFOtvkTzs0tX/32CguzvCSGXRHN2j2NpBnlR7ZyxRHcBcVwiH8q9TWFuUxOz5sNy4Yw17M0ggc/AJpAp02zYrCaA/Yw2VAAoDeOrmBhOB7TCSGdvcl0j7H7wYL7DVN+Ys3cBHPTk3EwWVsu6a6p5owZTXafM67unuUqjHxBjd2dHKMpuDUl5HmW+Vs3Ls4a4ChBEIBxgBp1FZTlm7XxpGDtPiuq5X3itM0Fb/BwdPDe/oTKYri7ALT5OUQlHNPOfQ21glvvPxFHYcfOJEDwQb1uLqXI1AnKBvF5VkUz3YWP3We1H6ADWZ6JgbGA3tpNx7ontKUvwQCS0aWV7eMYwLl6dUAA6Ne0FSXVBO4umDLcotgx4hVY8ZvcIfHD4yyP01XtalDTgS338OnIDpBZAIwA2tLw2qCHAqcHD00pZ1pdVO4pVKkvis7eGOAgRc2UFdJZmBGFeRwsiDnq9d2gd/kb21W1o1T2392BHjgOgic2gmwQdJtt3WJTIHCwjTpgvLTlDQoLY2caQpvxowLUGmWTp5lN5CGT9c7mmpf/oGjipCn6BFN3UmqCSn5wGYRq8asD43rOzU1jt5IWdE6MXrXJqXiKooTeAN3ZgPiEkQmyLpVHLwhyNmoppLJ40cOGohCaKBe/EG/y2ZSUK5tZIPU4DgtnMx40wjhcDypslkt1fLz8iz67IY62oQcuA4CV4dEYADzCruRkGZIrLQyV6sUlBbU/BvOFFDkBpkCfZmRbXkUu6opyN9/25Zb9j8pHrke0VRXVdODkg+8pOhVQ/XTDBt0/MDYmUPjWOazkTXABSFVCOeudKlpZxrcaACb0kZFbSo55OTxI4ctf1pI1sIPhKsfTH/Jo3Jkc2bGp2DEKYSDIbyzwijxG/ImVcvV1Vs6IVjvKz+0iTS4zjB1BLq0hordc3DX6wvXa+Ui4019pqciN8lUanS/nQytD2LMSeoU14lAU6szF5F367M1FXFpP5s20KLR490DmopANWFE0WttoleNGb9hq3T/0QkO4WBz1vd04gZ3Lb8BuYE4yjWFNHivAXVqhDenTxz1eoykibc9aYHMw7CDmTHYUHKpEb9xYH8TU01RP22zVq6t3toorkg6pTlT+0iDw8KQhSmrBWbTvexnKbFS7vb6/LVapUjONGfyjZp2k0j2J9KJoQkDyAM9ae2gTmSaSmUDSFugq3OLDPMB4ZmOrL22vZqKTDW9hiVhrid61fRVkWGlXnKJsBcvzv6Df/VVzC/EcDGBlKhdWKaX/EmcRLB11ZUw99VKm+d/41P+uz3x6c+2ctym9+EF1SrV0nplbaW0eru8Oo84anltSS3zvOsnP/1j9505iQXy8EqIVzKZxF94T6XgcqOX/sB/cgO8eEc+Ar+4k+i8AFZNcXVh7uUvllZuIX9xo1LEJSnSCRWlMJE+kR1PTz2QnXqAj6mPz2fRp9AxjKbMK4AWOQKBtgwJAl3NfpZ0EGGkhYvfKq0uAAhZEQQ2qWwiO5oc3jt0+v196WGu78CSdNWOFLuUeVPPiMZ13YO6pykwG/85T5B/+DQN6NQ/4xnphchob0oasjFLZhs1FaVqWpaS3PHK13/H5ziIiQameIS8jO1SjUt+Gq744ZOTD5/cW0G0AmN6TGqkeu/cb5VH7W4jOVTynyM3xY1yETPXVNiAbNx9p0/cc+p4SB232Qze5+zEsf5UViWqQXH1hXCQFQ2KUy1Wl66jlpomUvIDTu0V0QlzVTBzgWAD2xQB2PDVllbnV2cuVMtO8Ma5hT7AagpZFanxQ/2ZES9WJ7MGwtx7C226p6n0CE0r9nn5r/Qld4QzLXB6Tfv8ptc01T3VtPCcbO8u26Iad7yBIB47vZ/MbA3JT3AkcTk1JZ+7C2wUuUGMBGlp8KEBaWgD5Dglmc/eczKCh0YPvbO7DwJv+upRHHVqdXkqhIPCNpWl63w9Gl0MP14LqBMmRwBI085AuFkZrs1dIZKHytDIlXDKChAVJ8mksqmR/fqAMgxGD6/IT2uZ0Phcbbc1hYKbgbIKU9EOg6cw6xW1P5ruHU11WzWBeum1BtuiGk+8eee9+yfHMlSPsgZfkgrkqHUk76YUNVXOh+a4QARlwAxcathQkHijWoQoUKPz/ntO+T9GNhkMOb424g38Z3p0Mj0ySVGcfkRx6onRGAhQJTfkZxcri1drxVWD4hgIpP8M0wFQtDFwrQGuOhzmaO23wa3B6Zy7+SbcaCqERpU6neqciNwkM8mhPYmhPQasSmiR8jcgpzXmHrGmAkkkmKg/5CiwCU5qh1rbSTLsBU1FrJr2H+9ojrCNqjHxRmvoxIHxp+6dUpmmVJyKJuSoRXG4pNpd4lJTcS3cOqUJ1MprtdIavZcL/Fg8ePaegwf2BY52uQGbOR+LFmjs0AA1brMTR6m2DVMcJzGarlIVtiltFHPlhSuuLjXtyw7/TMPlEqZQB7wugXm6doOmMqkEX9nI37pUzi/Rih1VqlznCBfV0hS5SY8f1ssQ4JYNQqMBJiTqh5eVbNlVTYVZkBuOspsvfdE1c51/ClP7ObD6gL9w8Ez2oKa6qprWnpbo99pG1VBo2nUMDim8+4GDmWSfroKsUgWcQE70Mor8jGy0qaCAApv1anGtWsqD33AadDqdeujsPYaV0X8aIvXHkkAOJHUEl1piYJS8akhUk1EcpjiVQmXhcnVtwYYcXJuGHK/YhrwddrmEmfEHLPFP0nX9tVL0LGvtpWtcNiZ4YmkWIjcAVypjw6MfmnMDmElkR5Ij+2xyY6AOw78eBHTw0YpAU0j/C+NVg4gwVrj2zd9FUgCDPf6cfu4L+DMM0mMMEUikfOTWg5qKQDUdfJC6d6jtVc0WfmO4Fx47M/W+dxziJV7qG9fwFKXfuyeY7TyyClWxJw1xEXjSADbFfLW4Xis55OaRd9x335kTPlChn29Dqpru2FjlDzx8wNTQ7sF9p0BxqNwAV/BU10oUh0q6IaNhrXTrAs1KEetYS/gRXMEpguAq6TAul4hVVC2u5m6cqxRyyqWJG6wXhFaetP70YGbieH+WFiLS4CpHACxAQy+dQh1pzrqtqfArpUJKPGsKW/hUAsikzcVaekpTUaom4h7Rwum2VzUNvHEdkr/vwaOZZL9a7kWv9cKrKOqthVveAbuoHGiy4EhFq5YANqtqy3MWMsjNIw/cJ5HDABU5oDZwRT79eqAtzaKrdOR4PLP7UHJwrE5xRBQHziVKoitUVm6Wl6Y1xthrjAZCDobDYVwuUSoS+lidvQQXjSI3SAenZe5U5AbVa9KQBmhfauygDTY2zNiQ08EbiUBTIB9dLYSKNIH2yE2PaioC1XTwQerGoba9E7nnC2jFPHX/4Q88fLTuUlNrvVA4Ryer0TiyG3LZ1mNqTxrApgRCQ0hTwJZH8IYv7PGH7n/w7BkDJGx2osfONsVxHWh73bXsJ/gMj8rA/tMYzpsUh4JtlU0kbZfyxZuvVdYWNcXhqSHSn+bjUguTIxC9gtZuI3PhxUphlRYgQOSGagqo6UcgNykiN9m9J/sHxgxnmh4TSNRpSvhN3WlkmgIetAMJPjcFMMOk3abu2mjcm5qKTDXtiK7b+267ajzz0/Sdf+jRExMjWV5bTC004qxtRQarXpSy22KK8Pjqpuql0lTYJl9Zz1UKK9rrPT468sQj7zAuyXiaNdJIe2d8GXhTEsDkkByfs3uOppCophIHdvUn6xXVqNwApQhXCrXS6vr0KxRUt7xqNEbwLuwWMkcg8OI726CYu7185ZXK2jKnpTkFBaieAFUT6E8PYI5ncuygZHIS7G2wsXXhOlwIeRfboqluzK5tfzHWXtPUtqgm5GMTcbNeUE3iZ37mZ/i29YBXjxD5w+T4YC5feO3SDCWeOulpzvIE9QmgjlHtrPg+9Tcf9z/gZ//4hc6eke9O2Wea3UlgU8ihjgCyoaiaQIVyoPH6wLufeOrxh8mRU68FwBUB+F1PWZe1BuTkdj3VX7dka2gczfD5yGE7PhPM7Oor57C0JTFOZ2lLdXmN+VH4HvZ4eK8qEeFcnrSztluJF7bRCXidFu+W4w3vPxlYnYV3gD9z6fIL+VtvldewutqaM+eGbgk5aRlUlktmxwYPPpAYntQcznki61LVhRuMggKGWNq8XzZt0WiKL3Vw4jCtmxlUlibkfaGqN1b+bqccUc9qKnrVhJS50Wz82EM+W8j+4nrqHlGNS/zGGOjhz488efodx9GZnaVfeKkxXsH3DmI5TpBZzZ2EJ61YK64Rs1lbUevfOZ40VEt78rGH2Gpro6YVLL/0ceBoQ++6o+vjIl1ADGyZ3Ycze48l0kMUyJGJA07F6GINXrXZN0soOoDwW71ekUwckAMLPinmbYRJXmqtL7W2FyBw+eprmHBDYIM0AcJXWlqN6pVTHegsJJCZOJYcoyUh9Euey4fctMNpvG4nGk3ps4PlYGtnogwfCjlvKJfSDtj0vqYiVk1rD3w39uod1QTEb1hDUxOjH33qnnR/HxXk59Vf8A7UgWnG5HZGHXrt3FiOAhvKEaAlOylsA7ApAGmWaCtQHUy80qnUu594eHJiN9spyQ8C0UU20Pvq44R5yOQRiCSls4P7TicGx6n2fjIDKqO9aqQUSqtbx9zPtasvlFZmdfxGAo88Kb5HClNggZMw19nBNlDH8rVvL195Gat0APJR4gFPICVHAnHrOWmJwbH0xPFd/Smb3BhKYZw2FNHBq5WjB32WbmjKuGYEcg4++j0th3OQHQCkaXMJiZ2iqW53om48Tm0es6dU44I3dofENx987NT3vPseit8QuakvOEbLKqulR9TAcsdCjgYbWisafjPMs0HSbTm/jA38RjkS6fXedz36xMNm+UKJPYboYGtsWOJv+ICBto+P4LVLenTf4NRZTDpBAAM5WrSOslNUTa0zjen3mJ1aWFm99Nfl1dv+FCc/+xbmtbT5ZHd2d/AYMJvFy8/Dmck5aRjf0JgGUlOla2hF+uzIwNT9qAbtSm4M4+L6YBvibfkWpE5t5XZQUzi4neuBID9YzuF3fh9QJyTXQTM0BtK0XwWyxzW1vapp+YnqyI69ppo+VEfjJ5iHh7JKNH7SL5SRnp5b+tXf+/oLb14nV0Y/ZjygNmIKw2yMNHlRFlqyACuzKPPYEWFFchAJNlSRE1NtKuvLpdwClYLOzSMfmi8DpdJ+4G98ZO8eWoxOB2yMwIAu+WwXh9b1obl4sxHdkZWhZX1onEvH9qV2WC9QSqW0vnrtlcLsm7X1ZRAamvLpLAZDl7kLHieUjs6MJId2j9/zNEayOLcuSm2M9yWqRSJ5v5MA+B2wyS+qCTe6mkAjbANuN3DgvuzU/QhnqSJ/9ABrWyxV4yVwyVDbv+Xe0RR4KoI6qGKOoSHK2/CkXQAMIjR4R9FPrHTZfm00ltiO0FTvqKb9xyz8EXpQNWa+gFQM45B+DWVTQwPp81dmc2tYupgC6yq8rs0UD9q5hPROgRwbbNYRs0GOFvpqJb+IgAFrd3Lvno9+x3uPHTko/STGZzZwuvi/RBSNMTpxQAKMBjBXn5vr49UwrED41EANNfkrSEUD16wBoNQuSjGOCkF3Niqrt/rTQ8mBUVquTax3YNhcPRgM/1h3vCXQRU1RfAFGE2UIGmADblNnNsmBkfTuwwMHzuL2jQlGfEdaHcZaD4bkpSg6fiO6B6meEZ2mEFjOju4bmjyGpAzM3OQQND7gT3yJn9qJPEsp7WhNxZ0o+k7UwBv9GOkRosQe/jw1MdK3a/PVC9MYWjuzuxXw8L51y6Uhp8dRh7PROGZDM/NRPgCpAWpguADIgVeN7wthm49+6L2PK0+atmXaokmY8cIbOb7WbSRJ0kfTkKPPJQ2irRo0Q8x8VyJdxWwbqu9CMagG5JBqVJI0FVirVfK3ccMoyU4rGlghKD6jfu+G8Q1zTMh/8dLzK9deRfYzzbbB1E5dJw0JaYmUcqMNp4b3Dh5+KDG4xwAbvn4NKhJ1jOQ0ySPDXFhgm1hTPaupWDU9oprEZz7zGdmRNOZr02b051OH95YrldcuYvo6pQkR6ji1BtSqbOxJc5hOL/vWGGyoGinAhrLRSnCjOWCjVvFqVPf6rg+8+8NPP6UNse1M42+MxdP0umcG0TFccK7kJoyfR2JPf2YY8RtADi0vTbSTlULDaxoNqG9IX4i/rS9V15aTQ+OoiKP1btAdPXoItLAdb7B2+8rCW8/kZy8ibIY5tgQ2CBk6RTkJbGiqTXYYnrShI4+kxg5p96+Whk1uJMwYOejR4GusqZ7VVKyaiFXT4Dc6VEBWqk5ZpD9Nfz59eHJtvXjh6gwnp7E508lpzhCZCU8v+tbqxbXIVnHqMyoIENjAgQakIWYjwOb97378uz/4XgQ9JOGQc2UMfqM9ZgbM6D+NobfNcgRT9FvbTrNPNvqAHHjoq+vLShdKKRpyWKP4XlX7RpinuHQTfnxKNEDsbatvTfKbKH1rmB+AIjq3z/91cWkGuqBVOzlBgJM1sJYagU02kRlOEtg8jFJpVMuvHrPR46SQ5MbgkR0HTmMMJ/tUrCl+dHtBU3EnirgTOXjDlkXDDPcWw5+muzRyAk4cmlhdK1y+fkvhDblreByNo6iQjgrh1JMGdDZWV3t1uINrHxqVfqlP6kQFgeUKloje6kbDATGvE2AzNDigrbAruZEQYsdv4ExDA3631422wwleAKCtv+FVc5RFkDMKu4yENCosxnjTiOVsUuIBQQ451uB5K6/MVYs54gqZwW2GnM3NtYVrS5dfXL76KlLRkB1QpXKcSH1GJotacskAm0MPpidObGz2SXKjTZgRubHJjSSUElzDPT8BrQzCFGuqdzQVq2bLiG2bOpF7/EbCvqQ4zkB5czOTSh47sGc1v/72DYIchTc8/bPuXqO0VTUwVchTJzrqi+151WmNWiqGAjbKhwanTYUqCCAbbR54oxMEcI1PPvogwAbVa2Sowwi02HECf35jkxsGMMOrxnYwkF7I8QF9RtgiO0KQA35GSEO6qMdySHV1x5qSwEalVlgtLFwH6BLqIAgkErV9PndWe1g8bfnaawsXvrW+OA3UhzqI1qD2cyPLjlKfyY2GLLvBcdQRyE6e2uxLyBlF0plmBGxsWqmZTTTONDl0a3yONbU1ocM1XTMaTcWdSFt4fj7tEVsHOxHhjWHXDI5psxxGnYFM6tjUxHqxdPn6LFXwhIEg5wahTj1moGwcAQxMIZm++omihxxBazBlFUunlKnqs0pFWyJao5gNxtTamDKz2TM+aoCNjxPAZjaMPZwsIMmNa+Ba9i7bFHoZR+NZIb9TdhSOtQ2GHMVymLo6t8bOT15Gj6ZPVXDjxcXpzWoVRAflYQw4kc9GIP41BUUA+NyNNxYuPINCNYB8SnpGIU5nVRs8Rfzgq3k2FLMZoZjN4Qeze09tAFBV9rMsQqoR2hXvjRwNjeWdvSPdXV1pU6wpO8Zm+5y7qqm4E217J9rCb6APL7CRwzTdcwaz6VOHJ5E+cPHKTc6AYuChTUcRtKWrW6No3WsO0sDIYkSP6SlMa5D4hNE0p6KpmA1F2rW55JiNZDb4ySAiNlOR+QKuwRttCjWbkXEg/lIymxatIUHOGGbe1Io5BhuIoI7wOlVarbZQRx3IpLRyC1wH6a0JqrWcNU6t/2zxkgQQ4aTF5Vu56TfmLz6Tn3kLs0PIgYaFUxu0RgVs6hUEeFInmM3wUcRsTgBsbDeaBBttsOSgzE4TiJLceMJwrKn6OMzItdF9rasjA7/hUaya7qhmS76AVIANPK5Otkw6ed+Jg327Ni5evamSpCkJiiGH6Q4HeBqBBOXWUlzHGQ42NShupvEWpFEONCwusI75HFXQGuVDcwI2IjsAqc/IRgPYDA44NlfyGzt4Y/gBvHIEAvmN7GB6dOxq9ENZfPCCgTGamILqlkiPJqnVp+PUaQ6zH9IX1ShSNSOAOrm5tbm3gb4QF61gRlVB3ZOkQ13GVm2BSqIcOorTLF56Lq/OQpymSOtzb1TgQENqAGU8Kk+iyg5IcS3O0fTw3qGjj1CCgIrZ2FM7+SJdPWlGzEyTSFcK0szT5WuswqeVx5qq95loNNXEOCNWTRdUo6LI6kU2qf4KrDXAk9vl60vfeOlP/uOLM/PLmNZGbhDYCxSKhzME4+UUfaBKBFSPAMUIVD0CzH6npjw5lLp/h3q7GsKzLVXzH6nmm5peU8MynbRyWp5iNrzEQCHnLKOizj05sec73vvEe9/5qDRG0s0l8UZ6zzgdQLvO8Nl46Z9cs9Q03ZGjOVeDbkC+VzEIXRWivDJTmrtQy93aKOexLg7o3S4qeadGAM6LTk4zQKEUVS2CrDwVwcQ2iPLD2fED2bEDqaGxhMIe20z7Aw8wvrS2WM4tFJZn1+evIRUQWiAqA+xHRgD0Ul/ewvG8Ajj44cFlpInZpMcPDh64Lzl6gGFGgo0dtjGcaa4JGl4I2qHHr47nW1d8iDWlHx7tGLDHatJnIIGhhcGNjzbjTrSNncjBGw02+KBzTH06CZDGhpwXXr/0p1978ZXzV3iIigkpNEqlsjc0Uq6/0zd9/Sn8qlAH70Ca9oGnATOctkAhcUzdQLozlRErIuWJsgOwmI1aOQ1IowvV8KN59szJ9z/1OIrWeIGNdIJp/4zOOvPBGxts7Cy18M40O5zGapLFh/RnyvhGbZ75y5XFK5slhhzkfVWIgzZS2HmOJ6EOin4COh3g4bECFJfKJDNDmZG9KIJCG5YWTWcJmWjoUE9DJHRXMq+UKTBWyFFsbG2xlJvHnxSVoZp79F5DLgDDDE2sYe5bxzG1eBomEgHw+rk22r7TGayilh3VSKNHQtqgMEgHkhs5brBRs7Ngw0eLNUXPBFIlxdCmRzQVq2a7VLMFbzTq8NXwuzZn0qgZeKOx58athS9/46U//+bL5UqV6MtW+6VQB5uCHCJAMFgaeIgWUYVjEJ66ry2I9DgxcHp6KC7uwAzyaGlpOGXg1Ggai49hW0O9AEVu8rpwAHcE+NDe8+Qj73vyMdRG8wEb3VW0aZOjaY03zGwYY+Rn13CCJky2n8drTKcH9ZKMuuqIUYe0hiVSlqfLcxc3i7mNyvpmFahTgRtN5BMq4qJEz1pzGA+tZqbU1HjHWIFGEjRQoMZbWSl7UIlT4ryEKMpTR1kJDsDQen0c3lPpc3WkUSdFckUKha6RHUButAGkot2f2n0YtdHkuEf3E3zQQnMdMruCupZqZ4fMrnAVa4rBpgc1Fatmu1SzpRykxBsDcrQ3Q1bwlBSHP/P71557/Svfevnc5WllkhTqODarXuKTXTfJuiHbQnfUQNthPGr4LJKp6Qqdom2Oy4xnzrMVU6vycCiiTGsKVAp1ZsN4s4b4jTNZvW4hTh0/giUGdNVnwx5xAF+iggxHa7wxyI0P2BiBHMOTpn0I/tZQOgT0sIC8h3WPk+Q6mhlU8/OVxavVlelNlPVEsTVU9qT6Nxxma0zXVcnrqtgX/k+52gw/9Y34aEIpB5Ihg69NLYfLeEkkB1TqkbzGvB9qU4cZhe2MNKr8K4gUnHgDmA6ZnTienTyJqs82rdH37gU2WilaU5oA6cFEBGBjUBz8GWvKGBZsr6biTuQzDuiearYkpBk0k8eVmuUY5kyjiyvq3Lw1/9VnX/vac99eWslvRR0aMqvFS8irRpADBw6No+txHR5co7S+MnaK7MhxdCMLQDlwyIfjBGkwiHbi3orZIEgAclNGwIA2+NNE0IKsAdLPnnz0He989EGsZ8PWQYINGykNNja5kfEbV2eaDOd45Urp/DR9Og05rkNmacXswYF0rBkeNociVErV3Ax8axvriwQ5DtGpcBY7rSvhMI66PBxIUOjCioBSSB38p4NNzqUyzXTigLxEReO9Pk5Aply92BEOopGGwAZIM5ganhzYfyY1NsXr2RjMhimd1pQeBxiZ6F6R5yjJTaypHteUTXEMd449eos7kTZT0lr6WCr7Jz+88aI4bMsYb7woDv/0+oUr33j+28+8fK5UrnAKUX2kTG407U+jmAGHc3gD2NCwVw1+aRxNd8fTeHQmgKQ1yoEGp019LThFbihBABSHcp+cdaD1zafTqccePPv4Q/efOXlMfykNvT/YaHMmq3B6QY49/0bm5nrlfQYOwF0dAqwvo59o4NGjB8wGrSnUoYgOIIdm8iOPQKUU8vQpJ/ZQn7LDWKzRxeE06gvxQOkJPo25Pg4X5dQ4B2YUKeLHQI0zENujDOxBJNQN7DuVHjvUlxk2Bjracxhowgz2adBHY1TRVD9puXGsKddhQS9oKlZN9Kpx8KZuYpwUNT129qc40ofm+plR58VvX3jmlXMvvHpBoY4Tmq6jC6+dwxvFn8ldo/7HY2rH0HFlfWf4jJE4eY+c0ixwo1FUgJa4RqTCSRCoEMWhFNutLyDNw/ff+/AD977jvtMG0mjE1o4a6bFhy6Whwsh7dvWnueakGWCj8YbPHobc6MuWDgG2yJINSNSRkNMYQxRXqrnZ2vL1DUYdQI7CbKr4qVYK1xVxnEyM5i2uTj1kPToZiRT+URyXU0hSA4mB0ezeE8hDw7QhfRdyrCPvVGrH1bcpc5wM8bY8KGv+1rfsEWuqZzUVqyZi1bjgjQYbbcVsWyYjBLZjTfMe3Qwm77U3L73w2oWX37iwuJxjFxkZIOU6U7yn4UPTSKPsbz0DSqUE0POhTKszuWdraFqtPYr17U2YwR3tHht94L7TgJl7Tx139atoW2/gjR4CuCpGgoqRHWCs8WWHr1smNzbkcLeRzk+D6MhAiB5DYBcUW9tYn6+t3NwoLDPRUdlrVUF32IXllCqgU+thofqjfjEiq10TIEZQJxTkDCkE0mSTQxOZiSPJkf2J7KiNNAat0ZCsWYsMhhmBMa21FoC8TXRx3V3btVhTvaapWDW2ZdODNu5rHexEnnjjRXFsj40Xs9EON4k6127eevWNi6+fv3Tu4hWnZzqzcOrjX22hHIcNGzI2ecrWcR6aKmdAJpbyoKj4ppeZuOfUsXtPnTh75sTBA/tkGwN1JL3gz65pAjqSpimOvaanP7ORAGZwmqb4jbT80jTLwYHNcjQH2rILYl3F5Y38HKI7uwh1ADmK69BkHRJyI5NNlAHdgjs6+iVCbmI8QVNqONeZvGfJTGbPkdToAeDNrmRGDms0reEv5T1Kuikjz0ZvMUJiWsLbRW74kdOyijXlE11zHQh2A/7lMWPV2OO27nWiBt7oXiEBX3cPI3hrhKNlLEeGdjTRMaII+PONi2+ff+vtC5euXrpyvULJ0zoNTfjQdHyAXWkOucFFOWEGr2cRGHDi6KGTxw6fPn7k9Imj+jk2jI58vg1aoyHH8M8wzOBLg75Il5or3rj6efRJ9YXJSw3T01x7i2uWhxwo6AaGWd8s50F0kE1Qy9+Gk42IDvkqeea/AngHb3TBCC6WsyU2s8V1Rj5SitOQ9yw1kB47kByeTA7uRl1RaXn5egxCo29Nw7AX1+w1E+ZDcfQwTqIsdy49OIg1FebJ72CbuBOxdZI+GMMesrSbtU62jvzwRlIcOWo2uoemL/6oY0MOH+f6jdnLV6evTd+4fnP2xsxcqUxrrbfwyqTTU/snD03tO3Jw/5FDU4c82IwhOJvWGORG4oQGGxm/CUQaw5MmaZMkNM2SGy2ikL3FsGIyMieZhPNIVYsbpVVsm4WVWmGZSA8vRSPrsTLpVM9hI41Q5QI4k3gQp0nBabYHAIN0ADjNUGLHdTTjhTT8iLsyGy9aY3spO9JPWngg7V1iTRmOGqncTlm01jQVqyYa1WzBm0CK4+WrscM5dgjHHr7JYThbnMWlldnb87fnFxeWlheXV3K5fH5tfa1QKBYxf7SCNDRQi1Qylc2g1srA8NDg6MgQAjO7x0f37tm9b2L3+Nioftpcodj2oRkWzTBt7FXzitwYPMYO2Hi5qu2xQ8tgw/crAUNSUltfdiBHcgt9qC3XA6ShNLYipVDT6s5wuJV2oUwAgj1q1QCVb6aqEiQzfUk1hyY9AGjpTw3Sh0TKvjxjmG/cgjQ9NvzLrA0f1qhhpv1BWWsmzHWvWFO2r6ZHNBWrJgLVBOCNtAtGOFqzHBmekTEbO34jIUc66Ngs8ru/JdLDEGlnfcyBYWtc+YR0ahkxG4Pc2MEbG3KMNrZBlANwo6e1Yxlde4vNSl3xxnBn2fKULFAjtG0mXK9fjhy1cm2MNE5q0BoN/Abrl+I1BsttQngHMcY4VKypntVUrJpuq8bEG1dDwDAgnf5stqAeBhvtK/MiOkYzvbtGHQ05NupII6UvT37wNw0GpzHMpbakGgZkkIA/y+mENuT40BrsqHdni8nvuAZ+t/GvZTPnatbtIIEX3rhGUGwMMJiHvgX5wdW8yvGBcak+SCOFpqElfKafz1W1LOf2d4w1JUd4dvdsX8ItHyFWTbdVE4A32tZ7eWlkVICBB7RGw49BcYwQDrc3AgkabyTXCRwOS/jx8aQZ9p3/1ABgu27kaNoneKMzCDQa2cNwbTqNEYTNErrRW1yJjqSYWgta0fzB9WKkhF0ZpN7LPoLXMfU4wIfZGGTRpv8Gp2mHL7ashTA7+ti1WFNhBNi9NrFqutqJXPDG1bgbeCN7hRGYkRSH4cfGIe2Lk2DjZfIMkmswG3/7pZ9Lm0lIu6Y/G0EzV7yRNaFtumMEe6RN7JInTfY9KQ0NGAZ3tGHG0IIr5ITBntasgFaNoRFbF67RGpaqhitD461dUgR7xZrqWU3FqumeavzwxgAe6VVzHYXZjjXbzyYpjs1vpMtOWj1Ns8KDDYvMHolrB5pt3ZiCGP40f2eaHa2xs9Gkp07Tqa4OIuQYTaO1F+oYqtTNtIdNCp+l6oPx4Z9Um2sa/FKrw1aKbGmDTc/SGhvGYk1FAO2tnSJWTWty89/LHW8MpNEWR2KAxgY5XjYgxwtvZBRBo45t6eyBts11AoUSntlolxcbOCMPSgKPzoeWeGP70DR02VxKwmE37GOgW0DDjMQb40svyA8JOYZqvLA/JPBL7Ri7yIFFN4QZ+Iy10yDWVDvS6+q+sWo6Lt4AvDGAxxVvbN+ahBmD0OjEAa/AtTEMxwXIQI79BATA6daVfbWd0jxDkg9jQG2jjhHCMcDGwBsv+6iNo7aMXTKRrr2FFaq5ixw0uIKNbsmqdx2IBD6Uxp16OdAMiUlao3/SH1iMXaWJgffVqQaxpjolyY4fJ1ZNZ0XqiTdyDGsYGo06rtZK8xW7mIrrFBwvrxqfxdWrI6/Ha6ztY+O0qZIBFS9PmhGPcYUcA5kMv5wcjEcGNvyU6N4iccJgjRJ7DH+mATY23bHP4vp0tkYxbYzxoTWS4nS2h0RztFhT0ci5hbPEqmlBaF67+OGN12DWleUYMWdjqo3Mh9bMRiKNdMrZA3C+Ev2yL8zz9gS/se1+ILmxwUZzGokx2vMmGZI0l14j8S4xG0MarqTQB3Wk/I3PkuUYhzWwx9V7ZpASFouB+saX8ldDjN0miB3sZiEPFWsqpKCibxarpiMyD4U3BtfRRkeOf+1MM4PoGDAjwcY1MdprcK2xJ3BwLe1RUz4cwzPmWs/OdZa7ZDbabmoja9vHaPDGh+gYQG7IPAy/Mbim/VAa5CZQEXpYYCCNgVX6RJHJsCP9LfAgXqPpWFOBout2g1g17Us4AG8MpNEmPiTkMAh5AQ//hEP5uNRsk9dBvJH8xk5LgyGzQcW/mIp2yvUU2NjAbDtIDXNmuEwNWiP/NMDM9YmUGCPRV39vkxiNOjbM2Myp/W7QU0dwtWvysTe4Kf8pe4qhTa0jeeRYUy0oPVZNC0KTuwTjjQ052lrxB4Pl8KPv8/JiNgZDcvXkSJeafy+Sw2rDxknHmoQZ6Q0zQjJ27tnf/Mwf2aL/2i99Qg7PfWzldo3KfTqMK+T4fGnL3zBnBp+ziY5UhKuO+EvjncW+XQJss7+F3z3WVHhZRdwyVk3LAu975o3pv/sv/qTl/eWO/+8//v6pPUMy5uyFOjat0S01XEm84WPaozaJhfYtGP406f13jdzYeMMwI4kO4xBe3/OPv2Cf8ev/4kdtZmOby223lWE6jJS2/dkAG5+Bsw0YhktNIo0XQt/xnMarA8aa6ohp6sZBYtW0INVO4s2//5mPH9g9JNm9TGDzJz2u+QIydQr3hj8Nw2cPseXg18Yb/Aqo4HcNDAwzrmCj0cX+8N0//fu2uP/qX35Sn0LCmxyPbxfY/PafvfzLf/Cscc2/+ve/58mzhyRs+7jaXDHG7nX6FDZIuGKP7W3zEtd2ia6FftXZXVyFHGuqs0Ju7WixapqSGxnfTr1o6WBhx3UwQxtruWyM/szLMLu+UqkUf48PaK//5G/0wjP6T91YLu1s7CWPI3+yL8D1avW9uAqNQctrkN4pOXfwOK4U0BWA7aQJV/n4f6ndkvZKDRrypZPNuLwO3vjOOpQeu+hHS7ptjaFSrKkolRurpilpdxJveAKeQQVcp+VzGX+NCvJPxgAbCVyxwcAS/GljkgEkXkeW37PRNK5QWlIGFfvlOlSXzXpwhC47jIGUEnhcjRrjh8yhMKRkN5CPh3F8Q3osN+Pymnq477DGsaZ6VqGxakKqppP+tN//2f/04MSwnTxjRHR0YMYrQUB/byfe8De4N/2BP3uZftuAGgND7VWT42uDmdlmV39jDMZd/UX62rYXbFz9ab/2332M/Wn6ZQvTx2Pgs5fXXUvK4iOZ7ZVVyM6zvc1iTW2v/H3OHqvGSzh++Wm/+Dtf/4OvnzP2/ImPP/ljH3nYNjR2PN8AHq/JOhJ+7BwBCTkaZmywMRQs7b6GHD0GMdDCK5ZjuJVcgcoHb/i8PQI2uIyQeOMPIf4dyacH2vjhiigxzLRgQ13HW7GmWpBkx3eJVWOItBW8+eR3P+IfJZPY8+a1+efOzzyP7cKMPPf7Hzx85uD4+x44uH/3oMwL8OI0BnoZtMYVb/h0MGFf+fbssxfnX7+2zN+c2D/y7nv3fe9Tx20I+fw3337l8vzLl29zy0dP7/vAg4c/9q5TGpNcmY2GNIkx/mADIH/z6ryEc0j1XWcPGWxDS+zGfO7PX7iMP2/cXrUHAfgeO2L3kcHMx58+69ptmsUbeZA3rt5+9tyNZ9+8gXf5/Xc+fvK+o3s//NiJQ3tHmgIb2TjGmA6aOVcD5z+GCK+LWFPtaCpWDZlHHyl48RtYRi9zL5HmufM3f+0Lz795fcFfSQCeH/+u+4E6Emk+9ctfubW0LnfcPz7wLz/1Hj7+n79y4/Ls6l++epMb/MMfeOi3/r8LcytF2X7fWPaX/st3AWN+k34q2NeAA37mP3snEuoYdf763Oyv/vGr+ULFbnnP4T3//NMfHB3MSDbzgz/3+ZsLq7IxbO7nf/5v8zdo+X3/4/9zY95oMPqHv/BDMNm/8DtfB364igWyBYM0fnJVhI9UcQTWEV44XVP57nJf3v2XP//Muavz/koE8Pw33/+kP+poybTTaeN9m5WAv5nzOVqMLs2Kutn2d6FqWs8X0O4pNq/SmuDPX/8PL/y3v/xngWCDvb722vRP/Pp/vDybk6Hmhh9K6BAB/DemV/7+v/rmb375vAYb/M5TZAxl45vf/asrv/h7r7iCDRrfWi787L99vlDZKJQ3/tnvv/RP//0LrmCDlhemF//73/iKkQdhnbBxfp+OioxkWH8vsMEhXFlIs88xn6XZvez2fJxAsMGO4F4/9ouff/Pagu1gNL5p/6riIzQlgUCNeDVo6ixx4xYkcBeqpnW8MYarEnL+yb/9q899+dXwCsgXyr/w754plGuw6ZzlDAiz8eMPn7nyc//uBeCE8VOiH+3NF5r94TNX/a9hdmn9f/29F3/ys1//xusOVfJqf3568Y+++Zb0qnm1FGBj3gJgBnASKBa08QGkwN25AXhJmHP5HA2kqqkj5NZL/8Nn/xzvIa8wbhZLIJbA3SaBdvEG8pJIg8+f+/Irn//Gm83KcWZx7be+9G2dPG2zB2DD5/7ivOth+xP9uIhmz8jtX750e3ZpLcy+L1yclcEb110MnhfmsK5t/uBrTQvQPs5vf+mVlq0/kMY1SuR/R4DJX7GmlLYshHjHWAKxBO4wCXQAbyTkIGLxuT8zmQ1Cyn/ve5/41q/+F9/8lR/H9mf/9Ef+zn/yuC3Hr756fa1UZa9aU1Imf5rbDsMDqf/8u+7/k5/7Xmxf/Pnv+/GPPOB12MfO7P+ZH33vX/6zH/rKP//hP/qff+A7Hj5qt0Q4SvJfnyv08achCIQAyfO/8SnekI589the+1DPvjltf8mhHeyid+cP+BLhE7s9wMYI74eUKvGwL71iNMZ1yiv/y1/6pB1nwi5wrLUMciEvL24WSyCWwA6VQOv5Aq43/H996eVf+fyWoimIIf/Bz/0gN9bZBPj8b7786q/9hxeMg/zsJ5/+wENH0OxH/pc/mlnIG79+z5Mnp3YP/uAH7pXH4TY//r99eXZxC005sGfo//ypj2gsZAz437/w0hefpSwv+frhD5791Mce0VjCu/ztn/+CkQ6A7zHB6PDkKO/78f/pd610gJE//IUf1kf+3n+EfIEtSQGH9lK+gHF2tEFLW5gAkqYeKdecgp/+xNM6XS18fprd0vXKcXmux/wnn/pOV/xr6nbixrEEYgnceRLoDL/RcvmLF01r/v1P36czCyQ/+L733WdLE141p0Sm9dvUnuGf+ltP/PCH7peFcPRnm9/gG64RINMQHj+z3z7p6FBWz5Bnjxle77x3ym6JOBOjkQeDacWnB1OOrf0Hy9XEr7YUTeHEa/n6+PtdlIUG3++We91+8Kl9acRHiCUQS6AHJdBJvIEjxc5l+pXPP/fk3/lNbO/6u7/F21N/719j+8g/+L9tcdxcyHNM3v4J3xl1U2StFDfJOqvXyIIrY8NZtyMTfhj1uw7VeYxsv1oo+/jKWtbuoUmXySs+VhtCBrHg7Ud/8Q+e+PRnsXUkIQ234KpEJKrxWYztQz/52/ZdY4ZQy6KId4wlEEvgDpZAJ/EGExjblBRAhU2/nZ+GI+tqj0xBNJBQvMeGKIVPGqK4PaiMK95osNGVbHyoije/afHuD4fjN6AdSACDxQfGAAB4C5Os3NRlta/Epk4XN44lEEvg7pFAJ/Fm2mMOY3hpcoVpwhs3e69hRk6FYVBx9acZzfDnmCveKCQzKgi4JrwpNNSnasV7Fl4UsiW4DugLwMb2dLV2QJ+92ldixy8pPmAsgVgCd4YEOok3HZCIjNqbhzMrT0uQcD21xJv6ZzeQqJ9UQw6O5j+dsxteNS/pEdj80hdbyzTrgEbiQ8QSiCUQS6BDEugk3rj6hWQSLXKunvs//mu5Pfvr/5Xc/uGPvNfLWwUA0ACjgUSH9938b05URu4FX52rP83OaHAVb8c9aWGU6FX8htOjtXiRJx3maIFtwijRSMg2/kRSXOBZ4gaxBGIJ3IUS6CTeoFikLUEjwGAQGP8/jaNxY8PxVYcTiw0pfNLHl9zFOqyTcqZZi/oQnbvM57FznUODqTA8/QWQo+ukderZDaPETp0rPk4sgVgCd5UEOok3sIOYz2iIDyEH11QrL6TxcVVJ8JC+Mq9djDb+7EQfPEpfWeCj5upGQxayLefAQ4Vs0JQSQx4zbhZLIJZALAFIoJN4g8O5zgJB+MEnu5cNPargfPinPsfm1cviS0gwPvu4v0K0pCb6CPX2PfF4uM7Vb21Wjev9nLvmrLwgf21Bibw78rORIR2Hmnri0YkvIpZA70mgw3jzyY82lmLTN8tT6JG8a6RXwZjCQmFiPHJ88Sv+NEiPIS6vX8MwkpBHDnOoKJXoymNQbEYWN4PHMkxBaFdHGQ6llcIrF0AjbSoxSvnE54olEEtgB0mgw3iDqfJeEQUYMp4+ojeMhWEoPepCumY4O4K1iU4gv/FRSaulPqPQsuvyawBmBmneMB0nTCFn10QALuqsZ4wyNemcEqMQUXyOWAKxBHaKBDqMN7htxLG9FqlsUyi+/jQXfHJr3+YlRL07+E2nhInjhI/6dE+JUUswPl8sgVgCPSOBzuMNbg25uR3Pm+oZiUV9If/oE0+Hxwn/i3N1lHntEisxak3H54slcKdLoCt4wywHBqupsTkgyrWS2J2ugoD7g3frVz2WLdB7Qs6oyhwoKNfFqn32ipUYKNK4QSyBWALhJdAtvMEVwAgCcnimCDbXQbpc0wVtOlIpOfzN75SWyFH+Nz/9cVcfFySG+ZWQc0gOBIFjTQQcyl53h1e4OXt0UoolVuJOeUji64wl0PsS8Fv/pvevPr7CWAKxBGIJxBLYKRLoIr/ZKSKIrzOWQCyBWAKxBCKQQIw3EQg5PkUsgVgCsQRiCXS6vkAs0VgCsQRiCcQSiCXgKoGY38QPRiyBWAKxBGIJRCGBGG+ikHJ8jlgCsQRiCcQSiPEmfgZiCcQSiCUQSyAKCcR4E4WU43PEEoglEEsglkCMN/EzEEsglkAsgVgCUUggxpsopByfI5ZALIFYArEE/n/WRV8gcv9WHgAAAABJRU5ErkJggg==";

        private string imagePart5Data = "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";

        private string imagePart6Data = "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";

        private string imagePartTopicRank1Data = "iVBORw0KGgoAAAANSUhEUgAAAMYAAAAlCAIAAACWKQEmAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAh1QAAIdUBBJy0nQAAG4ZJREFUeF7tnPlT23d6x7d/TTttd3NtdrvbzrTd7XR3u509Op3pTLbZxIAAG9+x48SxkzhxnDi7uWwjCQHmNrexjcHYYA5jDl0IBOISQtw3CBCXAIm+Pl9hoVsydvrDTr7zHQZbX32fz/F+nuf9HB/+amdn53vfXd+twHNcASD13fXdCjzHFfhelO/a3naubWwOTS40tlvKH5tK6ow3G4xVLSa9yTI9N7/h2HQ6nVG+KvrHXM7tbcf6+uL0yvTg0lj30pjJPj3IP7cda3y0s+OK/lXRPykJXVtfnLFPWySh3UKobWp7c93l3Pq2hLoQu7m1trQ6P7Y80b80aloa71mdHdlctTm3HHwY/fijf1LMdHN9Y3l2ZcaKOITapyzrtsmtjRXX9taOa5/LGxlSjs3twfG54jrjWdX9353N/ack1Ssy+QsxyS/GJv8oXv7zo6rXPsj+PKuytqV9embOtc1OP4dre2OVHZ3qqrU8TDPd/LQj/6wh+5Qh51THjbOmm5fMNarJjgf2iX4m/xyEPXkFb2NxJ401CO0u+xRZQmi2JLT04sDD1Mn2KvvUwNbG6r6XO3C07CubujBoGNPc6q342lj0YXvuGUPWW4ac08aiD3rvfjHclD9nVq/ZJsQ2P6cLJK3ODc/0NA7WZfTcvtxRcI61RWjHjXe7Si+a78vH2+4ujnVvrS3vY6YRIDWzsJJbZXjjUsk/Hk59KU7xYpzypTjly7Ld+yWZ+OeLsYpXZIqfHVW99WVRQ0vb2tozbTNLbJ+2jrQUdxZ/qLt+VKOM16Yk6q8facs6Ke6MY9q0g5qUBN31I52F54cf57PHru3NZ1xqhAKm4aaCrtKPdelHNCnxGlWC/vpRt1D9rtB4XdphRjXcXLg80etyPQfl2Vq3A5f+Krkh521t6kE1k007pM885pYrRqJK4Gaze27/acZU71hZ2Mceey8OBm9tYXxMe8dUdglB7pnq0g9LEt/SZx7XpR0Sy5t2CHgNNWQvjnRtbzmeanlDQgrjpOsdj/2s7NWEFAlGKa/Eh7t5AKj965GUjxTFo6Ojzn2olMu1ubY00VZpLDivVSUCmq7iC0MNOdOmusWRzuVJs33SvDTWM2duHdWU9ZZ/gWKxDRiwUXXps6w1351oqzAWvq9NFYvbVfKxtTFnqguhJiQiVxKqHmkt7b37ZVv2W2C6Pe+dcf1drMu+NxiTszJt7X8g12Ue16QeNOS+3XP3C3Z6fkC7NNbrlmsbMk531g7WZnQWoV1H0Kveiq9sQwbn1j5VCIc+1VnLBFk3XXqSseB9S33mlPGhbajDs7zzA7pxXXl/5RWMJY8xsKHG3DXbZPTONzikHFvbd5p7fn8uHwv0clgk+eFMIC9OcexybmdXl1Mwj6e41hensBPYBo0qsav4I7wMlgMu5RJOnR/bTwwDNmVrY2lmbkDTf/8qa61NPzRYd311fnQfG7yxPGdtyNJL9sB08xOAsjI7JGhTAJOA64Chuf7WvntX9BlHMSH4RzT+KWb45FFetWDR9dz6TJ0i02ccH6hJXbDomVGga2Mj8bP2qUE0rbPoA1bGWHBuxvQIcDytXGjZuO4uFledEm8sPD+qLoO0bTlWA7HC+qJm4MxSrcJ6ASxzVTIoj3J5g0CKDaxtG/z3U9n4NT/ECDcn3N/ujSvEMgWi6mWZIuGj9GHrQPTTxj7h19kqLL/5vmJ5wgwt5etMD6jNm9Xj+js4eJvVsLm66J6b+GhpBgqCJmGoB6qVG/bZ6CXypMM+P/QoD8vE180P5NBwYOP7BpeQ5QUv93jG2yrca215mCoM5FNeNms7hkebEt+ed2a8rZJhuHyDG+wQ8N1YnOanUwQiOyBrcbSrr/Ibbeqh9rx3p7vqpAAl2gtuimkXY1YlYA4XR4wClGEJOErlWJ6HWQqnkZoIq1ubH4tGnj+knE6Xunv012dyX4gNxIryF6dz/ut8gef+zdkbPz6YGugQsVXfP5B84nLGzORINNDeWl8ZepzvduojLSUCNNIFXxYLkX2yVX6g+ZvXmq+8plbEGAtR03oU3f0Mqw+3xfGzWJba65swyugutH+wIVtwl4xjI60lYNr7ewAavZzqejjcXDSmK5+36IGOx3Txy2x/C7rOmK2PcqJHFSZhaaKvs/B9jVLGVoGSIJbJuY1cveBSh3kGYu4ZGHtsbcxjg9syT2DnokQVcxnTl7stsfVRtvDXXsEyTBRbS3S5Oud1z4447IK3MeDl8V6iE02KrL/qKuoUMeb1h9T47PKxq/deDMATuHkhVn6jxjg4YfPcrabRX7+Th+kKRNVLMtWP469dL7iztuKzVYHbDThmepvghnCUQTCxYnM/g8HAqQMU8KRWxLYmv8mtVsaq5QdwdjAqz4LyC74AhouRQ30x2xFBxVdm+5p0GYKfWh/lOlZ3he5CeX1l0nDPkHsaR6OWx7D9+vTDfVVX7VNmj5tgn4iY2rJPtWUcnzI+iHJ3wQRuDqHG/Pfm+luCEhQSFhC7lmtvtCS/iTr52QYAMVCdolHFd9/+fHUusq9HhG3E2H7jXYT23v1qY3neb3FIW8CuCGzbc9723MwL7kiSQTzscs4P6jsKzutSD42qb0b0uT6Qcrpc6RX6HyWmBrozQPODWHmV2seXjc4s/eZsflBICQjKlP95Uq7RqMPjmnXpvvUZE8aqr83tmVaUoz33bQ0YUsqAy0hzobUhpy3zOP8EYSw6pMezOhh28AflJGRbnuyPCCk0ktwEQnvu/JlEl6992oAmg28BX2Uc4Y+Gn4o4DCRkC8azJ3RzY1R7i0jCmH9uZWY4olBgN9VRjVEksJoy1jj9naxkdJ3bo60lrckCx8yUFVib96NrLigmASCoGm4pYuLh5QLB/vvJzILBQ578HsYKLQy2iWBTEatmjsrdGzUebMjyWFChP6Z64gPGszCoDy/RB1IT8/bfncv/QYw8aHD3Qqwi+3573+gcj7lfGh5SuL8XYhTnv8rcXA/pjOAQY7o7mCLiZPigR2thpVB1MU9lHEaeOIiPhGnpaRIGXCljI0lNedsGiHZv+VesnbUx1+MWg05esn8IFYHbwoDOz1SAVEP2aUTDWvCkc33NpCraMo6K5VbEThgqvd8Jwnru/InxDLcUQ+oj7O7STFfJRR42V6d4jLHPV1wuUo5YJgQB9xCQEiRyzqzRXT+Mu7eP9YYXitkGxKjilBdb8HwF0IwbKrHEWlU8TAvu775585j2lvfyQkLM1UoG31+VHB7He5DCRFVpzC/JFKGsDhD51du5vz2b/2VRczSQEoYqLuVXR6+MDHSHMlQwmO5blzEA/VXXvAeKWvTcutxy9Y+t8jd7y//seOINoVnS8yhTrKU23XsXQQYht7AZhecllx/y2licAge8ZKBGFUiDsB/oaMu1P3bfvIQPgk9srS6zmm7nO1CT4p0GY0smO6qxjtgA4YbCXC4XQYYuLQldx2sHpZikqZhUq1AkWRhIIYQFEVBWJY6py8KE96yPueoayoCXJJQJHB1ebLA+U9Kfg+P6cqZA4CzuuVFpZXwS6GRSgCYUhSx0mInuQWpxZf1ceu33DwQ3UW67Bar+9o1rZ1Kqo4SUQFVM8u3yihAO2GUbNkKM4KEzPY99vA8Ey1QPqRw3VGC9PFZne2MF8AlIKWLRdb/XUkwgFYlHwLSgy6GmPW/RoLhEeaxR4H4sjfcSWiJ6cbjTs/G4QkHm5Af6Kr72Eyp8aNknAGW2pzFMgntrfdlal4k96L59ORSdn+1tAnMCAWWfwt5CWSnmxVZPtj/AjprKPt2ER4e4KGSRSAOdE4Z7QZG3YZ8j58mo8BJMXCo6hQwkIex9lVfF29oq3fF40GsPUuax+f/5sIj8ePiU5t+9mXxGVRM9pL4fq7xwJds2F8RsMCzyQCxcZ+EHK7ORuQhCYQZdpGcUUJx4Mux+E9veQOcyWhUxOM1tR3CSgasdURcjFPPuHUyFUTswRAqKdQdVlroMv0Xf2rDjasHxUGMeiA/1Hom9XQIuw80FQfcDI01qt/XaG+QYye4asnB/ccG41K6E5Ym+jhvv4K1I5QcVyjh5DwslnhkP/szKrBVxCELH+u5dHWq8MdP9aN0rxvTR802JZZKmuq/w5pR+0nchhddr6hr5t5NZEbPkTwupl2Qpr51VDZh7oJ5+svF6kBW1PJbJ+MVcITbGNW2qFzUKQjDhPjR+pggGRoIU3SW0IeEb9CUIJXYDHySISfSFJyJkg6C0pBgI6wRVVyXO9rf6fQWhpKQ1qYk95V+shfB9WIgFqwF8SPb4UZDsIl67q1adDGoPEIhgIN0Sw0BqY2kW36dLTZrufhTUtGw51kA5CmYq/ZiHg8503qKDKohARBGjTj6ArsJrBZEnxAkw8wx7fkDTlnmys/jC4qgplMN9Aimn625L30+T0sKbKD59WkjxlV8cVxoM2sAhOpZn8CNQB0I5DEz43cUHrc4MQRul2CS2r+JKoPvAAkFTiKdIJEpxX5Ba+ursMCvCblkbsiMWYtcWJlBiKB1WjWwQKahgHtxFAtadQ8JHhzAYWyCGwoAh+y3bcEcAkXKtTFuYGtvPS3DfpCJFmjsspMjAkU1Af8iDA/1AuURFFHxa5TF4q6213YjK77GR1mLMGIYffegsuQC8UFeMcUfeu8vjfYGEDwUT1fScU7N9wTMgvH8XUttOV2lD96uJqlcSItTy9gGpnx1RaNRNJCX95sPaQRuFL2jKjxguba4ukaoWyoQmpR5cHDYGIkYYA4tOpFjyztpGu4JRYBf+ov3GWSAy9PhGxGQSATwgIEUkxZgJfZVXglZgqK2SEiPTQbkmKKTIx1JG1KYl8QwP+w2MuZMbY2u1KQkUAyCO1DQjQor8MFyepRAON1h9hlCGmiBOAD/Fw4EDI4Mw3FIIa0RziAyIaUZbS90BNazR8jA9sJgIOQP0jA2XGmr1fCGV8O1BCr7sz/t2IaWIY3ep5YWxUmQvR1vLsBOkiPgJiwo6H2GZLRpCEiCFZQ4FKWHqgFRjXkRIUYQnhw7ho+YjEkXyGLhaIBPCTwlI5Z6B7weH1PYm2QcCQ1pWAiGFegg3p4gl5ehObPpAamEve+79cik8TNOmJUqrF8TGAymcgHDxISDF27ZWl7wTLmQKKGMIV6iIRTMDycPKlIUknASp+giQog5T2Wr+aVL6t+H4fnlCaWjXB0KKTBLpzVZF3GBdehhiC1mZ7X3MPFuVseyrpS5k1UWk0Xsa8S+4NppeggbqhMckhzCNqHiUJX1B0Yw1oBknAtVFU31w43KJLFF6Et5WwCXYhYed7q6nQkwQvmBlKfY8Mu6pr+IrUSFQxhJJERiCFfJwtPEIx5fztn1mUKrHBTDR1cX++9ewUlDmoHxfeMYH5D5iaKDw1LgisIudHQgDzEHkk5UyEfP6ThVFFe1cuafhlH51Sc+Du1YKG/it0XPVa++pBgaIOPyZzS49Z87lXzjsc8E3Q8r+USiV0kJvQEjDlJlQuHHtbTwU9FyklIJdOFCiAcBB3nxzPTjDCPyew27TiX4mGb5pYUDrEwdtO7BABA295V9S3Ai+Z4KetxtyRLsIea+9UpLLSeJA6pSC+yeYSi72lP+Zu6vkI2GSU2TkFOiJG2oqCMQEcRk8WpuehBbtBIv8AeJQYz7rRgtDlLEtg8d1YJ/wwjDIWV+jK6WaGzFRaCwtoKGqt3tJBCp3r10s+TaSCJ/Jc5YWguTZQMBE+z1Rb8k9Q5NQ0M3AEVCzlJRGZGvAE99C6cUdsI7oZb8ElxEqFZvB+T7KM6q5JRV5TtMX60/vFiZG1KVs4XBTkXecTGzI7pKaJywiYeatHkQJZoxBSjyteWHqX9Ay6k5sMPkIz9jIiQALUbuUYg5err72prgpLNIcJ25Z69XXSbpKtV6fa8HaRp6d7afVKURMIBAgOs/SjyxY2gKfoWlTdKj6XjRzYiBFolURQ8eEj6uVQkg01lyT6lgObgJ4fg9Sy6sbFzJq//5A8nPMS9FrRT/x3YrKoPUsNkYiDUfZKno8AoNSXDtBDXMT1FWVaH6goAiDSRB3WwXExc9zARHMMm8jTx2mA4K2Amw7VmHSWO1XuiGopP5DcghKTozmQe3CUIc7nU3plPjOs9AAlAxhB/Ti+tHZvuYw5EykrxqyIcsk4TwcH9Tildpz34HePbnPSiHVaXYOPDFr4gP6dvzCWwRBKBlSd9lnYZwa05GKSwRAhX4zZeT4Yl7ibaphtENNu1yKfSE69obU6sJEV9kllnei/X4YzuBTkKls7aMXKrBNyhtkTxXxvUhB5siVIbMpZEFmZdF06zOqpN23PhfdFF4XgTE5X53qIPxpV2UJcSniSuqLHvfc+cI7kMFuYX7Yho58CjI+lWB/U2SbotojZSK+9ivdQEpoAGfjSW3Q+k1KgmLc0ngfzggzyTjZb+Ijzwt5niiJje8svUgyMxxTwcf1tWLqdGkHpzqr3e2EbPOajcaSYb+beAqMIhEjZLO2UUHyy3fgnuhSZzXoaA3FacT7N9cx20CKJOrKzF6VnY9QV8rJIm9Zo4IhgUuWYkx9U5goEfHFkFbwydu5nPR/QhlF4hSeGvryKRsvLK/94WJJqLKxG1jRQwoT9XKc/LIid3szdEbR5cLwEF1r0pIAkKcDjhVnWbHY7K7UEeC+vUrlyW+SgPAcZxB9SGM9HTfeo/+JBozwjeHgAKFUPAiY6YXy47Y2wrfME5BxxGGTiPVIqwquI7oSZGPa23sNn5ioMUFX2YNxfUX4WjVbADrdNWboEd0EYeyoMN5AShELPQ8Mu8izj6hvMpiu4g/9YwW/nSZ06G/GztFYQWAI9/d8DqOgeAXaQA9Ah3tI5Wppmoo4Zm0bavdqbXUR1ojnRVhzPXxY498vlfOg4yeHOLkQsiwDpN5NfegeGZ0Iv30vdHNLnPL3pxT6Nl043d3ZwQuIVo0UWWfxR7aRTvc0SNXg3az1mWRBg94UOyc7ML+7KT4HXUSiDymBmhfNmeEl8qkk9HNW0Jh/Xipo7IUOzk3o9j1QJXJgcnGz6ML50uJXk+rdb0Q3JiwKiAj+G03Hoyhs14lzCmlJxOphugVFYJVDqp328HPeRpGRSylsbUf+eyLWU98kRR5+spAed39VR947uGbPw2jvDEGomGasiDd3pynMc1vWCXyid/qUQIoaF8UDUnoiVRv28ofU1Lz9lLwqDKSwYf97sbSssaeoviv1rv7nJ7OCOkrqMD9JkOeW3l1fjRBVYdLJEMIe2DOiMIImQapE07Uj/O3xBSKb0ngDk9OWdZxyRzSpAXeend0VjWmVX9PT6M3ksDdLo93m+8kdBe+xvpyYICad6W7gnJ0HfJAb+DhJc3wTh8Miprjcu8C3qBJKh15O0tUTgmLisBzQLNw3rNzb5SFleayXeFCjkNHv64e24BstTGk3HX/MFBoH3fRqXdyyT1polSYAYhbk8+g2Q1cpCEonFncvYg560cic4TRoCopY5/CHFEaiwzz13+8XhHF/NE799etX/+b1q3QlBG3WA080IJz/JndhDtoR+YQhg4YGUVuQGvGuUAqIWCrZna7LRbKA3mIUGkIzTNIv6pN9UFHIKUDki5xQgGX7wcJNdNgDUO7deAP4SHkMPcrFN+EZR5qLo+9ORlVweV03P8G2wVroYcfURbNEkuXeIDIg7MU3UfvDOUZ5aoVuHOHoiUiUcT23PsejeUMZ7WIMHNehKIRJ81kE2s+l40Ok00RG8GG6NNoIV5DjDFvbzmrtwM9PZP0gVvnDpzke4yZbfOXFWLnsQprFHKT0GGo4Qn0pL3BGTxXfVfoJ9XAyfuHH7traRP/oCBPn3VIT+u/Lw7dJBb4NS8DxGCq+os5V+jEpIqkDPZwOIJS9xJVIXSgJAw9SSNhG01/vkc6ezVva8GhSd+Ex9kmyHBFOE4Hayfb79LLCEAjipjofMpJIm7v3OV9H8cTpI2UcFRWK6+IMRYTjDE775MBgXaY+QzBL2tSi6UtGZPBDVxxXf2wcfv2TUpxaUDsUKtHA86/GK89fLRgaGoxShzzzZjvpAoNLuqt4+Bp0iziFYAQzhi/ghmORbhZnZgY0ltoMPBfJd3EeobmItu7ol3jPqq/bacx1M1M6qCBYpCLptJGErqPNbqGiBGabJPtAVgkc8DBfgQ/5RalRDgBUYRTFcRep9IHJGXqUQ3USiCPILZSfzBpGT78bjfAMTLRgAP2Sj2jtjRgKBI4EhyV6ZrJPAWXk0mkzpi3n5XTzScvrmekKbQs2a8dwYz58DqqOc+fw0ir9ytEdaQ95NHTb6ey2zhy/WvkPB1XkPyOe5ntZxmOKX55I+Sb77vT0U5wk9J48/mXa1NB18xKQQjOYDIla0lEgZlxbjken6k5BhvWlZUw6KZtI9ARP945lotxXz2Nk/KhYweulNg+ZOMdccoFS62hLybj2jiS0lDUl2QHcOf0sDtMVfcj2bK7tnuR5Wok8j76R/CRTBYkUPZwi/3Sqp/xPg/VZY5oyMVntHRixuUpuLPwAtw4O9FknCEGWx3qi5G2Bo4JlcrCHLlnJxNJbkWQsep+WRgRR1XEvr7U+G0YrwCQddMakUcmO3juHtFLu0WAYpxbsBbXGP1ws5nAVrVSS0UoR2QH3Lf0uuLxM+c+HU05/VdikM9pXcFiR+VOobUD/sEy0KNFxIQ6wSxODZkkZnSR+EQ0YUkYY38FaoGehSG70O41QMkM0KnEkVZw1FSLcQjni5xHKse7DnJciNU8teR92ImA8Lvg+HUvEASK3yclxhKoSn0z2sNAr8T8HMcZkHyi8SN7qmf7kBh6WSg6FcPRWf/34k5keJI8jVpjlRag4wJ7EoRoOsJOaiXgkxm9eEf4mggDWjmtucaWypfdSTt3rHxf9y7H0HyeqfhgvaNZPDqr+4630xMuFyYU1mnbT+lqEjrbo9xhFJJyhKYcKgGjl4Y9P5L3L6QPIBFlHlJv1xbU/+19D8B4SZ4wg+5SoObJNUpvUABJJbSMdxaV1c8bUIALS5/SnRHb1lqyjY4Wdg9/AzKi9kCAgS0Qy3VT6CWjjWAHt1FKL4v4V1W/lMRZ4N9BM6amv4hsUiQYpQ96ZzqL3yfZx9I0ufo447295I0PKY7FW1jb6R6arNX1Ftca8B4a86vZbDcaW9v7RyWnHxkaUjjZ6VLm9A3UMsEUMSFckfw6Apcdf4Ob2bfwjDoA0KYQGodAdIixufgFJ/Oe3J5RRYfZE5DVjpZ2B+t3icIcUgs0GHICOOIOoH+CkxsYqHJGz3TQPIpTlRVHhkc8y02gh5RkmLcXQLPg7P7nCRw1RTy7ig08OkkfHECO+LqoH3EfXfQ+wR/XFZ3xoT+hzM0sRRvRcZ/rUkHrG5fru63/xK/AdpP7it/j/e4L/B2oQbl33rkngAAAAAElFTkSuQmCC";

        private string imagePartTopicRank2Data = "iVBORw0KGgoAAAANSUhEUgAAAMoAAAAkCAIAAABHSTINAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAh1QAAIdUBBJy0nQAAHIJJREFUeF7tnPdzVFeWx2f/mP1la8Zje4IneWtnPDtTu1u7tVs1W7NVMw4ok6OxwQY8NsaBbCvnhAICISEJhAAhhEBSdyt0K+dWaKWWWq1WB3VQs5/7HogOr4MA7y/rrleUkN57995zzj3ne77n3P67x48f/+D7z/cS+I4kgHl9//leAt+RBH6whfd6vd4Nj9tpc6zMWecnLLPDa3Oj9uUZl23V63Hxty28Krpbvd7Hbs+G2breO7HYrNXfaBupax2+rRnvHJqdXbI4XO6NDW90b9ryXd6NjQ2XY91itBr1a3MjlrkRq3FqfdW44Xbxpy2/LvoHvBsI02Uz25anEa9ldsi6OO4wz3mcNoT/GIl8Fx9JswzBQNZFWbNj9qVpprHxYpqNyrwY22ldXpnUGTQ1w/XJvVc+7b50tKvw/e7iD3ounxioPa9vKTEOPbKvzG241l/K8rEbo9n2QDt58UpbzFfVvz9Y+Isd2a/Fpb8al/6TxMzf7sv/04nyDzPvVLcMjhqW112elzKoeInX63asrS2Mz3U3jN/N7r/2ha7s467CwyxWW/px37Uvxpvy5nV3rPPj7nXby1U2orOZDMvDrQhzoPactvw44hVCLjnSe/UzxG7oqEEFzrUlYWcv6cOrsCHzpG5GfX2kIUVotvTIM83WnNM/uLQ4+OC5NRvZvFz21aWRtpHbGd2XPtRk7WjPSFBnbdfk7unM39+Rt0+Tu1udmaTKSOrMPzBYc2a2q85hnmczvMjybQ7Xvc6JDzJu/+FQ4c+SMl+NS+PCtl6Lly5+wM5i016Lz3hzT947p64VNmjnltdebEzJtDxu83TvZEtJz5VP1Tm7VBkJqszEjtzdnfn7uMRKs7bzS03O7t6Kv021XbHMDqKdF1mp/CziWl9dNHTUDVSf7iw4gDBVGYma7J2IF6l25O1VZ+/kN+rsHRjcyK3U5VGV227huRcc2uVYWx5pH2vMQbO8/KlmWe+Bjvx9LBzNcnXk7x28fna284Z92bBVzYYzL0zbuqjHsDoLDwhB5+/vq/yCvbXYd9803mme6lnR65ZG1chl9G5mT/kJTc4ursGac0sj7e5163MsfsPrJeqdKm5+a3/+6/EZGNPrCRk/eXrxc+B/uScu/Zc7srefr23q1j+/G5MVrKnpvnRExebJ3tlz+ZPRxpzZ7vrlUQ2be2VKZxrrWOhtmrhfhBA68/eqc3bg2HjEubb8Im4MsGEcaRuoOavO2Yn5dl06PHTzm+n2a0tDrSt6rRDyRLdx8OF0eyWC7S4+rM7e3ll0aKwxe21+7LndmNCscXL0TjbWjAHhKfqunZrAUfU2yZplycuj6tnOm2N3c3oqPtHk7ESz/dVfEaMwyug1G9K8Ntzr7JLeylPtGXFsoOH6FONwm8O84HUH7VcCt9NuM04YOmr7Kk+qs5Lwrmxuokz08+BOz8ZG18jczos3XotLez0+3deq8FWvxKS+EsO/aT+KSftxLDf4mR1G9rv9BRX3+qyOrbsTr9e6MM4C8c0YTf+1U7MddbalKY/THrxZN9zOdfMC+h6sO6fJ24PcR+9kgleez5c4baszqsquwoPt6fG60o/1LaXmqV6X3RIM7/iNkyg21Tv5oAQjU2Um9VZ+jgWgpi0JmZvxuMtjHQR63BWaHaq9YBxuJeYAKwNfhWZdDtvStEFT2195Slh/wcGp1gqndSXKQUOYl9dLQGQHt2fEay8dMWgIeQsRIS3oARQ80pCG0NU5u7G2LfmwgUnjti+rsJvXE3xsKz4dk3rrYMGO87WfFjSdLLy/L/nm7/bn80tfE8QWgWW/P1hw5X7/utMd5eLlyER2MnTjGzYx0554ULw2PxoRQbL7eWpGXd1VdBihD99MdqwubGFQ6VaXw8IbOvP2qTMSQVe4DeEY/GM8M1lfW8ZBivxJ+hNh0Tj0sK/ylCorSVd+gkCxpYDFzSv67p6Kv6nS0ewHeF+7yRCdZkfHGnNFgMrdPa26FqXvUDAvTHZ1pl9bdlSFbZV+xGzYsrLsWC1LfXZZljxOh59YvV7nmmmqrbKj4AB4ZbbrJj4gotwRm2HJknSuVhiNbzSMT/9ZQuaxnEYyxzW7k9jndHns6+6e8YXPCu8TE4mem07uiYUdKrzRPhK1xL2OldmR2+kiQBQcnFFVefwBOwtH+jiJue5biwPNa7PDvoiev7LvEZEqMwHRA7qj92Eel2NGU4OIhHXeSrMu6INTb+6ZVlV1X/oAbDRYe2ET52ENgJahGxfVmYnakiMWw2CUabvQrGGgt+JTECTTNo1pAjwWy2cVfiqW1C2G9npxqzPqKpASoNDQWSdkFemjYF4IdLDuIrYFtgD6+aqKHYZL6yk/Ll/aS0eNAw+ChxDzUFWpsxIRDY9EmsPjVdv62cuPBNLyNxcs5rOCpiWzwjJgK5KvtRM0/cKocHWp/3WsTL9gjjioCMdOm/5hOcC5I2/PjOa62+GHF/kveFZbehS405ayDZWAUTAjVPv4KQsDUYMvAbGxrQ2a67wwmnFRlUnf1VV0CMBOUMYRKuwH4Wa0YKPW5He4sIbNTS4PISys9jxqws7I7KIZF0ZpuP5b2bZAUQFOa2PDQ3LKnzb1yw+6smNo3DI/Ir/fbV/D42qydmqLDi8Nt0UEnYHmxRoYg+Sos2D/nLbB4080LA604FTb02LbU2O4Wr95GwUoLox0d7DuQnt63FDdBX4Os3jgPKj8D4eKSA/9XFFs2n9+XNanX5TDhdPthvpq7JxYtT5xpUur9pivq16JTfV9Cmv70bbU8xVt0YAws17XVXiIFY3eyXJZTb6TxHMQ3DXZO9rSYtpS32tLfhcLa0+NZe2sSMD5px/cM1sZxKkr/cis74lGzTwO0YBwdOXHIbcUlQSuGLmVhpCxBu5ExwHmhSsCg2MBWL+hszZyQPdiPbXkgx25e8DsgWEH03Hah+outn37tqRcsVKuttRtjA7m21yX02oie1Wlxw1UfeUrB8WFB5oX4AmnxZLG7xfihAKeAZB1Sb6RRJJ7kPhsV30ogeK0e8qO4/zne++GSeDNVsfuizfwOr5hEYv5h3eTU6vUREPeT2S8eLWNX4LMdl2ow3WJSL3hrWzuf2N7ZhAIS3vrQEGzdjI86YrRIybkCMtgN04GrMI83Ue4bE+PVWcCrb6d66pH2R3ZO5E1F/bke7/bYSHdA55zp2c9AhhAFPO6u5ocmJ0DC72Nytmfd2Ox/z4xtz0tLpR5CQl4XCSzEgg7tjb7xMGEUofNOCnieEY8VABkU/Bt65YlonB76jZiLmQErlrinvbiQVGl7/3AUzYGr5rrrg8w+oDX+pkXS8XAyUp4o6/Bbj7jsq0QvI2DLcKxp8eHNy9i8+TDco1QTzJpf6hldwwZ/nFvnuBLfVCXiHpx6S09U7gurr6JxX8/UvKqRHr98L2Uu53jcjTpGJ5960Ahv/R99qeJGa/Gpp8ue2h1PPFzwUPLW7+r6H2sn5w/0DFseMiP2LitKe+yR4kpTMJpWSIMiQ2dsg0zCtgwxBpJE/vh98MjMCoBI/UpqvQEXAX5mqJYEFfvlc+YAEIOY148y9tIC1SZ2w3q6jAIjMBHLBLxveDA8phGcVCMBkamLeU9TMc0rkHRQHD5CkjRcNjIB8AKnxJqCfIQfublshLRzrelxvBvmKQPcAaijGheuBfQQ1fx4e7iD2HIFBdPzSf9uuaVWL+wiK3ATfwsMTO/vrtBPXazfRTX9Zvdubi3nyZk/P1fvsm92eWRykEDU8Y/Hiaq+pkXj/84NvWvJ69OhUZgAlmrqoA+RDTEF2AQgJL53sbROxmgfuNQq9cjseReL48gfbTef+1LgfF9Pkh54PppHNtc181wXCuIalKrhcbMShJbX4mVBW5DJRJ9yGSFnEMER3lwpDrdehWPMFD9tcS1Kn+IaEO3UoRma8+HimiASFbX+u3b4/fyYJjD4SpShKlefCEO3jTRGSbx9DUv76phUMq0k2a19WGesZtmojIvANOaiZyfxc+oqhW96JLZnnDmOpHR1/1s/gzFha/i2gRYwP8fxaTe7ybPEqT1o96pf9qXH5A/ytb55u6cBs1YKHE7VhcH6s7jikbvZsEnhbrN9/fMf7ypEFyCDoaE9/KjP4A+M5pqwgqgiigT6oXEByhToAXcFd5C6TYvkai76DBGzCaHPpTBbjD22nx2dbpXW3xYOM7ZIeVxsYbpPm3pEfwNYV3RppnY5MOytpR329Ni4G8hLMiUYbxCxT6or9HbGbxw6lGFJzTB+cy88J+LAw/YMVBnkBFhPHz05sXkGJ79B8pRjPeDU8b/OFoaAOqfEapwYFI6uYmugO1UG00WwYZQdL50R/fTxMwA0CY/TsQ8XfoQqlZRhXgsCT0k4CfCowcxkMdFcjffcxetC+yVBva6GfBar9ezNKoi9MBxr84MhNqchBUZ8MGAu2x+yYT8Qo97fexuFjcwt8X+ZjycCoidHs68qEP3V31JlMcgFEOEgGg9d+GrRGQkkVdqPgDpU5JiLLymgPOpMeqsHT3lx6hbKBIQ0na6jvsn0SGahdpOPuYlcsY6Nha5qHUh5L7nRdGbFw52vucONt5beTI4SIOf2vtnfivqP4HRTdGZ4aV+tTOn8sGAnEuu2hzxZ6qJg8qeb1sqGYNymWhjY6HvXmfhQfaSyM+9ESrE9pVZ4D+eCfCuyoQI+AbEEyxQGg10ZR+BShf67isw4NIDBCZEQZBi62+4/SlDKdSxsREXDnKw+jQbcrbrhsjgwpqXy06OkkaIwC8qbhVMdlpdJWn2Y5GoKhUrMZFukKggXyh07hUROY2NFMtrYTEVdovXS+lCk7Vdd/kEC49sXsAR8BquGFBJM0aYeLEF83r8GI+ohiYp+chmnAra8d6mrolf7sxStI/gX5I2ni1/ZLGJtBHQdvler1QdUjZNYui2r6oUzQthIW52M4V5cG5ETpLyHLQWKkfoOInxxjw6c4LlQ2mIEAbAZ1uHognIEiDSMC/wTbApEFWBUDgPylPLIyqGgJeOaF7krZSlQfcT9wsViTc0O/mojMlD1tNfpKhZioyDtedI2hwr8ziCxb4memHIDSU3tn3VoBB2l8c1pAKkRyRJUZiXk0mUs0ExL4LuSzWvHVSWUFKwecFj/WJ7ZPPCsH64LSXpXM30osiocXu6sfk/HS8nwQxlmpjXu19UKRaIJAB0VZNLfSMq88KLzHc3DN9KEZ0L6fG4MfZh8IaG5xTV39w9gv13BXkmafEUCcD1iuZFCKM2IDoXoNZqz8tYIkrzGm/MATFPNOUrBjICH20gRL2eik+DtSArxWlZdphmN3eaaB6Z6mEHYpSkz5OPLgcXNyH9EQhEFaYZ2bzYcJQg2tMTYOqsCxMvx7xEcGzE27OtaecINi+ohzf35EYMjhAN/33i8uDUkhwWRw2m+DPXFSGXT1qQmnS2JlRwnNPeogEEnwT1HNF7SfmZhyob4JeQgXEAYqh5ByzHMisow46Cg3O626HwHG6vp+KECI53swIURlldV3Yc2wKk41N5AwY321n3FHudgOJWJMlEcGwgOCZiBIpmLYpLbWSXkmbngD1RdfKwZvoyJAodpHgmKN/0givwsuSP5pmBKMzL4wa9iqp40fu0rb0kaO+aar3CDsDx0pAYZF6PdWML/3L4UjCz4OuTgFyQXjD7Mk26bLEfSKkXlW+fpolgH0b58njuPWKowsqpyYx3UMmhlox7iAjtN99gN83BsoLu8WHUkv3e7PXQYAKeE7W8iU7yJEWJY6NQA0Sc/utf+xYzsJuxe3kC66THk0AQ5qbbrnAN1JwhnnDxZgIctHbwbIllJAqCvtbd9m4olPMBgtDC7CVaM0wTXRErOZsz5ymJe4uDKAjA70xjBoo0I5GsAsI2snmRUFAIo48M/ywmGrq5OXrsRfpKok4Ip+lRcWPNGC1/OVmJKYSEX1S1kzJL7/Ssu4TgTGuOr0pbIPSJlWSRREBFz0faSMwtbewJ1WNoW5rpu/Yl6qSGGJzS4mNQg2mi2zzV51s8QaYoScSL5HdgKX13IK4IRpoI1X/9NGlmqM2JQ9I/IE4lUI21+YQIcilwjDAvehhlWC2V3WRaVb7akt/BSwXzkavT/aLTM3c3E1YcF9ulKk+XFCY4230rgFKRwYaiuumMwNEyB7I9u2nW14Yg2KmkYV607yqmwPLNfrQqwBMKmFVBK9PoE8okozUvr9fMykuOYLJkGYqO3eF00/uAoSiaF/7px3Fpp8seudzCGeCKbmvGEs7UxJ++zr+JZ2ve+7LyzT05wVESpuPfPrzUr1cA4PKi8CITzYVyLc8SxD8tj3cgOHRM7RVmeFNndPBJrbkJiIjyq698yCVFjTU1VkSo0E0iaNE40EKxBU0v9DRuPHU2S8OtpA78Ur7IyKRrhzycfKkzRG000Ly8XkrpPALPQiNaKJU5TAaoYKbNbncHVYSYPFE4ILVnIKCkZOjbBm9chHL3eTkU6RAjUr4UabI/Beg7Bz/zIrqPNxfxRnLUMNlmlOaFlKm3sHIa3wSiVPIk7Jqb7SNw8cHUqGiwiU3fdeEG/avyjCl+U21cMFlpw5cv7dj8/3xaEUybkVEeSr1FGA0lbqCrRPLtxmfANAqS2udDTYZSNxu3LS0Wh4Hz53gFvxxrFPELGwK/bzYRyM/hHkQfc9b25TF1WDDntS5O0u9KvBu5lbJJcEDzUmRkE/pfj+AvVBmUHeOgYaHBLIaBAFIUOp4DAe1p8SIVDd2ZLdD9w3JW1FX8fjASXxFRay+5i2VumD4RwfOtzE21XpbNujX5XZEL+7xckF7qaqIcm5NzCWGibWBJG1AitXsn0DkZqi4UpXlRfWN4qBQmGmZDzxhX//y3ioDGB1HYiUmDQe0cmcOqQlnJ5IL5r59fxZgCmibe2JF9panfJRdzQnwAE31XT0FS9109SdLkexdVoIkHJYLlyogn3pFK0xPcX/211PNOpfW9gZpzvrFe6m5IVaXFcSeFijCD8ieSOxJPDT29RYcgAp46dZE9kI0GXEDDZ7SqaLryg5LMc3GwBZsWoGpMHXZcilE6OC00Cx7w16x38mEp1JrUYHhkqD6ZxeIR8J2CmEiLAfbZRLh/9llbmIDjYKdNtJSE7/oKNC+iBsOLhuaiQ4v9Td6njYS+bwdbcKaAgYEgtHYorgoyhk4mdhUWFpxk+T7i9nhL7vb8aicHgZ71NwP2f749q6yxp2d8UTe+oHjRVNjYOU5GGeC9AGT7k+vnlyO0YqPP+d57MKtInA76gPZLfBXzJx2TaHp6cmIk/glQH9d35TOi5OYS2MrkoaIFPn8flcrwtiX/1WbUY9MoT3BAQXSg7xtoSJHgVzzqDG7IofYg+k6Zf1N+cHtLwEw2nA48HJplnpD4vnmraIqUishSB47AfPJi2WCguuWRNj992S3EaFrkgT1W40T49Qa1Ez7tImLSfZWfr870BT9PYzTpD8022pKji/0K7YRYNBSOJmeHaLLrqFVMZ3xfixPaeaHOt7AN6gLRc1Loj+8X0QoW6vr9gYI3kvxoMzzZr3fntPRORXME0klfcd15eKzOQqj2e35MgZdG1nlgeHfx+4KvT4vB1PDrdNZbDEObOJIf8Pc0JQOSRm9nKtKtwQIklcMt4QtBDhP3C9ZFTq3soeHSoDm4BIHiX+ACa4vOC6ry5cdoH4qGXiEJkMjSBMzaPEne82RQ5kODzGhDBpzIk8WK3or9dBmR4vhWIKBt0WxH7l40C9kR8aCUQrcqoF68ghQpM4m0EzI6AJVLxx5NNKgAHYKinjgnyONkRurMhNGGdMX6SYDEaX+gOP3PBwsDwtyTk2fy+bMQly89IQG4zK/LWuRUIPJH6unQXf4EiUNZLfTfD/DWrA5PA+gB+XI+SoYmm1ADOfB4H/0wKKzi5OpMtE3JTIye46GbKaIJO28vnOe6eVEZwUhJnbADH4SAJZGEjd3JwmVSGZzXNUR5oAPNzmsbRGKRkchRDtDLM83KBzznRoCAtIgu9DXR30BO7WO1XmFb2gbR7pGZCCQVfUqRPspHOeiJm2wpJS8g8EP1kh0I1irSOWysDhQ83lTQkbeb/ADbF2ExuvOHACyKibQ/hCHiw9eOYCjo9DqRd2/etIUTSuw/GCxERjggw51WVYs2gUinFxE60XOu5zbFLgAAkIXjYls6FobFsG+pB0uFpiSiAeqU9moEzpPIQH85PCLelNRE/7A0mtMMm2ZAHxHkOc3fQrOXP8Fni/0fSUfisK1x6qlJxNFSgWajOdAQ8iAao049KqeCJnjkwoNUtai24rQUX0rm5TDNUb2GLcTnsXKAsG1R4XhCGHOn+6/4jo5o+BwWJvm2tI+y7k7Om6NZtu802PpsWcAHK+X07OD1Mwu99xzmRcV8Wxy6s64QEMnwOeQDFCM/JwmNaJHBC8ccwXCDNWdlhAcLMNN+Df8BSFI0MiyDgvR0W2V3yVFQUUf+HkiQMM0/oUQNxUBuAQKTKgQHRhrS2RuSZhUoaIANfAfttXSz4SnZDP1VX1GfiCIWi/HDHKMV3pKIIHqjpdZncnVkMaO+Bj1NTxLdt5AX5uke+Bt9cxFYlYOgHFwjhSbrtC/NRNwTCqjOKRrq//xJxU/EGdpnSD+M35JP23IE7Wz5w8UVW6R9qCxzULNZr4W4EtxmejxegWOl4vx7fzOdUhwQQqDgaNgpihAIoSv/gHTbLron4Paew7aezINjcCYDCE9u+QeKQfpzJtnQfROOVHzXA0I2DOGu6MyBONBxriSLk9PbgVA0UkvpSFQVnoBlE+JZmu7ycXwBZUBasOg7pQ7LyR26iZ5odqqXbaZvLia/1mQLMpnmwYnmIslvRfstGxG+BIAdxnjUKKSFUQ+hIyURl0bhgoCiLf6QY0lSui4IErYC3h7iR8p7n2fZSAHfQ23xy5KWf/2gWJiOdPY/mDglFMp/+s2uXNKCpi49Xx0QCQmE+zvjQj5Nq6/1Xj0pSrkSXc6GQfEsk8VSKxNogZ0mGRakP/1L7OytOsugSXgRF7QZZWxBCeHJJDaETvzuog8Yl3/FebVsIXx+j5fVN1+iczBiwhReGsJ3zg1hLtqyY+LLDZ5odg/LfKrZg+KbENITuDAsWreXBlujPN64OXTk75jAVIF4tPByPhgGT26TktjkRCAe7AiLJ0BwQABKELj3gsuWLOwx9GnbwMzJovu0Rfx6Vw7UAx2tMA7yv2QAP0/K5HtN9iffvNo8QGUpmjwxGuMDxxCw6LLCRcmVFlwFK5UWyzdr7KZSCU6CiaWzJVRbRDQDBdwjzuWuzNIPSC8hhWfxlRaZRCIRNMT53uydbGNod76BwaTvFse4o/YfYbeU0CzeV99Sxm6BimIguU4gNJu9Az9COWikIZVKg9BsaHY+1CiRzWvTjRMCgB3kwHxFDEnpRHMxHQRkGeSu0JJkJS9nzU9nipFxTGjRZH2g1Z+//GhfSj21oLivq7efq6WOVNk8MDS1RE3pZRmWr4Ck/ggLfDSVHw6OEiW5YK7hMGEK6P3dEoqP3toQICwawiQaIljEO/GgiJgFTQVD5LStUDx4DsgRYQJSOwhsMGEKjAW2Y1CGZgvxXQHEbqb03JqN2ryezVFMBy6EUYEsEMcvHB0iLZ/z8m7Pms25YrGbLHYcG11c3/Wg8pxEpdcjrZRLfNPVc0b86C1sc1yCgOjJkccV/Rf/F0P7axYhb3Xigfc/h3m96JDfP///RwL/Cy1o/10BnlvVAAAAAElFTkSuQmCC";

        private string imagePartTopicRank3Data = "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";

        private string imagePartTopicRank4Data = "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";

        private string imagePartTopicRank5Data = "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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }
        #endregion
    }
}
