using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Ent = Russell.RADAR.POC.Entities;
using System.IO;

namespace Russell.RADAR.POC.PublishingServices
{
    public partial class OpinionDocumentGenerator
    {
        private void AddAltChunks(MainDocumentPart mainDocumentPart)
        {
            CreateAltChunk(mainDocumentPart, opDoc.Discussion, "Discussion");
            CreateAltChunk(mainDocumentPart, opDoc.InvestmentStaff.Content, "InvestmentStaff");
            CreateAltChunk(mainDocumentPart, opDoc.OrganizationalStability.Content, "OrganizationalStability");
            CreateAltChunk(mainDocumentPart, opDoc.AssetAllocation.Content, "AssetAllocation");
            CreateAltChunk(mainDocumentPart, opDoc.Research.Content, "Research");
            CreateAltChunk(mainDocumentPart, opDoc.CountrySelection.Content, "CountrySelection");
            CreateAltChunk(mainDocumentPart, opDoc.PortfolioConstruction.Content, "PortfolioConstruction");
            CreateAltChunk(mainDocumentPart, opDoc.CurrencyManagement.Content, "CurrencyManagement");
            CreateAltChunk(mainDocumentPart, opDoc.Implementation.Content, "Implementation");
            CreateAltChunk(mainDocumentPart, opDoc.SecuritySelection.Content, "SecuritySelection");
            CreateAltChunk(mainDocumentPart, opDoc.SellDiscipline.Content, "SellDiscipline");
        }

        private void CreateAltChunk(MainDocumentPart mainDocumentPart, string value, string sectionKey)
        {
            var altChunkPart = mainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml, "altChunk" + sectionKey);
            using (var altChunkStream = altChunkPart.GetStream())
            using (var stringStream = new StreamWriter(altChunkStream))
            {
                stringStream.Write(value.ToCompleteXHTML());
            }
        }

        private static Paragraph CreateTopicTitleParagraph(string topicTitle)
        {
            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00233025", RsidParagraphAddition = "00F8047A", RsidParagraphProperties = "00DD5BAE", RsidRunAdditionDefault = "00F8047A", ParagraphId = "19CE1D53", TextId = "661746E5" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId20 = new ParagraphStyleId() { Val = "StyleBefore9ptAfter0pt" };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "StyleCategoryRankGraphic10pt" };

            paragraphMarkRunProperties12.Append(runStyle1);

            paragraphProperties20.Append(paragraphStyleId20);
            paragraphProperties20.Append(paragraphMarkRunProperties12);

            Run run33 = new Run() { RsidRunProperties = "00233025" };

            RunProperties runProperties20 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "Style10ptBold" };

            runProperties20.Append(runStyle2);
            Text text23 = new Text();
            text23.Text = topicTitle;

            run33.Append(runProperties20);
            run33.Append(text23);

            Run run34 = new Run() { RsidRunAddition = "003C0519" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { ComplexScript = "Arial" };
            Bold bold12 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Caps caps1 = new Caps();
            NoProof noProof11 = new NoProof();
            Kern kern1 = new Kern() { Val = (UInt32Value)20U };
            Position position1 = new Position() { Val = "-4" };
            FontSize fontSize9 = new FontSize() { Val = "20" };
            Languages languages11 = new Languages() { Val = "fr-CA", EastAsia = "fr-CA" };

            runProperties21.Append(runFonts4);
            runProperties21.Append(bold12);
            runProperties21.Append(boldComplexScript1);
            runProperties21.Append(caps1);
            runProperties21.Append(noProof11);
            runProperties21.Append(kern1);
            runProperties21.Append(position1);
            runProperties21.Append(fontSize9);
            runProperties21.Append(languages11);

            Drawing drawing11 = new Drawing();

            Wp.Inline inline11 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "356B07CE" };
            Wp.Extent extent11 = new Wp.Extent() { Cx = 838200L, Cy = 152400L };
            Wp.EffectExtent effectExtent11 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties11 = new Wp.DocProperties() { Id = (UInt32Value)11U, Name = "Image 11", Description = "rank_category_5" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties11 = new Wp.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks11 = new A.GraphicFrameLocks() { NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties11.Append(graphicFrameLocks11);

            A.Graphic graphic11 = new A.Graphic();

            A.GraphicData graphicData11 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture11 = new Pic.Picture();

            Pic.NonVisualPictureProperties nonVisualPictureProperties11 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties11 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Image 11", Description = "rank_category_5" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties11 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks11 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties11.Append(pictureLocks11);

            nonVisualPictureProperties11.Append(nonVisualDrawingProperties11);
            nonVisualPictureProperties11.Append(nonVisualPictureDrawingProperties11);

            Pic.BlipFill blipFill11 = new Pic.BlipFill();

            A.Blip blip11 = new A.Blip() { Embed = "rId10", CompressionState = A.BlipCompressionValues.Print };

            A.BlipExtensionList blipExtensionList11 = new A.BlipExtensionList();

            A.BlipExtension blipExtension11 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            A14.UseLocalDpi useLocalDpi11 = new A14.UseLocalDpi() { Val = false };

            blipExtension11.Append(useLocalDpi11);

            blipExtensionList11.Append(blipExtension11);

            blip11.Append(blipExtensionList11);
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
            A.Extents extents11 = new A.Extents() { Cx = 838200L, Cy = 152400L };

            transform2D11.Append(offset11);
            transform2D11.Append(extents11);

            A.PresetGeometry presetGeometry11 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList11 = new A.AdjustValueList();

            presetGeometry11.Append(adjustValueList11);
            A.NoFill noFill21 = new A.NoFill();

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill22 = new A.NoFill();

            outline11.Append(noFill22);

            A.ShapePropertiesExtensionList shapePropertiesExtensionList11 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension21 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

            A14.HiddenFillProperties hiddenFillProperties11 = new A14.HiddenFillProperties();

            A.SolidFill solidFill21 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex21 = new A.RgbColorModelHex() { Val = "FFFFFF", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "" } };

            solidFill21.Append(rgbColorModelHex21);

            hiddenFillProperties11.Append(solidFill21);

            shapePropertiesExtension21.Append(hiddenFillProperties11);

            A.ShapePropertiesExtension shapePropertiesExtension22 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

            A14.HiddenLineProperties hiddenLineProperties11 = new A14.HiddenLineProperties() { Width = 9525 };

            A.SolidFill solidFill22 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex22 = new A.RgbColorModelHex() { Val = "000000", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "" } };

            solidFill22.Append(rgbColorModelHex22);
            A.Miter miter11 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd11 = new A.HeadEnd();
            A.TailEnd tailEnd11 = new A.TailEnd();

            hiddenLineProperties11.Append(solidFill22);
            hiddenLineProperties11.Append(miter11);
            hiddenLineProperties11.Append(headEnd11);
            hiddenLineProperties11.Append(tailEnd11);

            shapePropertiesExtension22.Append(hiddenLineProperties11);

            shapePropertiesExtensionList11.Append(shapePropertiesExtension21);
            shapePropertiesExtensionList11.Append(shapePropertiesExtension22);

            shapeProperties11.Append(transform2D11);
            shapeProperties11.Append(presetGeometry11);
            shapeProperties11.Append(noFill21);
            shapeProperties11.Append(outline11);
            shapeProperties11.Append(shapePropertiesExtensionList11);

            picture11.Append(nonVisualPictureProperties11);
            picture11.Append(blipFill11);
            picture11.Append(shapeProperties11);

            graphicData11.Append(picture11);

            graphic11.Append(graphicData11);

            inline11.Append(extent11);
            inline11.Append(effectExtent11);
            inline11.Append(docProperties11);
            inline11.Append(nonVisualGraphicFrameDrawingProperties11);
            inline11.Append(graphic11);

            drawing11.Append(inline11);

            run34.Append(runProperties21);
            run34.Append(drawing11);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run33);
            paragraph20.Append(run34);
            return paragraph20;
        }

        private static Paragraph CreateTopicContentParagraph(string sectionKey)
        {
            Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "00385C36", RsidParagraphProperties = "002462E2", RsidRunAdditionDefault = "00F8047A", ParagraphId = "1347ABCF", TextId = "77777777" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "StyleAfter0pt" };
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { After = "360" };

            paragraphProperties21.Append(paragraphStyleId21);
            paragraphProperties21.Append(spacingBetweenLines16);

            // Seems to need an empty run before...
            Run run35 = new Run();
            Text text24 = new Text();
            text24.Text = "";
            run35.Append(text24);

            var altChunk = new AltChunk();
            altChunk.Id = "altChunk" + sectionKey;

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run35);
            paragraph21.Append(altChunk);
            return paragraph21;
        }
    }
}
