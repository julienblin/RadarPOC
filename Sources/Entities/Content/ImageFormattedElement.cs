using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ImageFormattedElement : BaseFormattedElement, IWidthSpecifier, IHeightSpecifier
    {
        public string Source { get; set; }

        public System.Web.UI.WebControls.Unit Width { get; set; }

        public System.Web.UI.WebControls.Unit Height { get; set; }

        public ImageFormattedElement()
        {
            Width = System.Web.UI.WebControls.Unit.Empty;
            Height = System.Web.UI.WebControls.Unit.Empty;
        }

        public override void ToXHTML(StringBuilder builder)
        {
            builder.AppendFormat("<img src=\"{0}\" / style=\"", Source);

            if (!Width.IsEmpty)
                builder.AppendFormat(" width: {0};", Width);

            if (!Height.IsEmpty)
                builder.AppendFormat(" height: {0};", Height);

            builder.AppendFormat("\">", Source);
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart)
        {
            var imageRefId = "img" + IdHelper.GenerateRandomId();
            var imagePart = AddImagePart(mainDocumentPart, imageRefId);

            var result = new DocumentFormat.OpenXml.Wordprocessing.Run();

            var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
            result.Append(drawing);

            var inLine = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
            drawing.Append(inLine);

            var inLineExtent = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = UnitHelper.ConvertPixelsToEMUS(Width), Cy = UnitHelper.ConvertPixelsToEMUS(Height) };
            inLine.Append(inLineExtent);

            var docProperties = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)1U, Name = imageRefId, Description = imageRefId };
            inLine.Append(docProperties);

            var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
            inLine.Append(graphic);

            var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
            graphic.Append(graphicData);

            var picture = new DocumentFormat.OpenXml.Drawing.Pictures.Picture();
            graphicData.Append(picture);

            var nonVisualPictureProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties();
            picture.Append(nonVisualPictureProperties);

            var nonVisualDrawingProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = imageRefId };
            nonVisualPictureProperties.Append(nonVisualDrawingProperties);
            
            var nonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties();
            nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);

            var blipFill = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill();
            picture.Append(blipFill);

            var blip = new DocumentFormat.OpenXml.Drawing.Blip() { Embed = imageRefId };
            blipFill.Append(blip);

            var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
            blipFill.Append(stretch);

            var fillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();
            stretch.Append(fillRectangle);

            var shapeProperties = new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties();
            picture.Append(shapeProperties);

            var transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
            shapeProperties.Append(transform2D);

            var offset = new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L };
            transform2D.Append(offset);

            var extents = new DocumentFormat.OpenXml.Drawing.Extents() { Cx = UnitHelper.ConvertPixelsToEMUS(Width), Cy = UnitHelper.ConvertPixelsToEMUS(Height) };
            transform2D.Append(extents);

            var presetGeometry = new DocumentFormat.OpenXml.Drawing.PresetGeometry() { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
            shapeProperties.Append(presetGeometry);

            var adjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
            presetGeometry.Append(adjustValueList);

            return new List<OpenXmlElement> { result };
        }

        private DocumentFormat.OpenXml.Packaging.ImagePart AddImagePart(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart, string imageRefId)
        {
            var absoluteImagePath = System.Configuration.ConfigurationManager.AppSettings["BaseImageDirectory"] + Source.Replace("/", "\\");
            var extension = System.IO.Path.GetExtension(absoluteImagePath);

            DocumentFormat.OpenXml.Packaging.ImagePart imagePart;
            switch (extension)
            {
                case ".jpg":
                case ".jpeg":
                    imagePart = mainDocumentPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg, imageRefId);
                    break;
                case ".gif":
                    imagePart = mainDocumentPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Gif, imageRefId);
                    break;
                case ".png":
                    imagePart = mainDocumentPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png, imageRefId);
                    break;
                case ".bmp":
                    imagePart = mainDocumentPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Bmp, imageRefId);
                    break;
                default:
                    throw new NotSupportedException("Unsupported image type : " + extension);
            }

            using (var fileStream = new System.IO.FileStream(absoluteImagePath, System.IO.FileMode.Open))
            {
                imagePart.FeedData(fileStream);
            }

            return imagePart;
        }

        public override object Clone()
        {
            var clone = new ImageFormattedElement();
            clone.Source = Source;
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
