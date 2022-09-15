using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace BigExcelCreator.CommentsManager
{
    internal class CommentManager
    {
        private List<CommentReference> CommentsToBeAdded { get; set; }

        internal CommentManager()
        {
            CommentsToBeAdded = new();
        }

        internal void SaveComments(WorksheetPart worksheetPart,Worksheet worksheet)
        {
            VmlDrawingPart vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
            WorksheetCommentsPart worksheetCommentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();

            LegacyDrawing legacyDrawing = new LegacyDrawing() { Id = worksheetPart.GetIdOfPart(vmlDrawingPart) };

            worksheetPart.Worksheet.SheetDimension = new SheetDimension() { Reference = "A1:D5" };
            worksheetPart.Worksheet.Append(new SheetData());
            worksheetPart.Worksheet.Append(legacyDrawing);

            var comments = new Comments();
            Authors authors = new Authors();
            Author author = new Author();
            author.Text = "DR-IT";
            authors.Append(author);
            comments.Append(authors);

            CommentList commentList = new CommentList();
            Comment comment = new Comment() { Reference = "A2", AuthorId = (UInt32Value)0U };
            CommentText commentTextElement = new CommentText();

            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            FontSize fontSize = new FontSize() { Val = 9D };
            Color color = new Color() { Rgb = new HexBinaryValue("FF000000") };
            var family = new FontFamily() { Val = 2 };
            RunFont runFont = new RunFont() { Val = "Tahoma" };

            runProperties.Append(fontSize);
            runProperties.Append(color);
            runProperties.Append(runFont);
            runProperties.Append(family);
            Text text = new Text();
            text.Text = "Comment";

            run.Append(runProperties);
            run.Append(text);

            commentTextElement.Append(run);
            comment.Append(commentTextElement);
            commentList.Append(comment);
            comments.Append(commentList);
            worksheetCommentsPart.Comments = comments;
            worksheetCommentsPart.Comments.Save();

            BuildVmlDrawingPart(vmlDrawingPart);


            worksheetPart.Worksheet.Save();
        }

        private static void BuildVmlDrawingPart(VmlDrawingPart vmlDrawingPart, )
        {

            var writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8);
            writer.WriteStartElement("xml");


            var shapeType = new DocumentFormat.OpenXml.Vml.Shapetype();
            shapeType.Id = "_x0000_t202";
            shapeType.CoordinateSize = "21600,21600";
            shapeType.OptionalNumber = 202;
            shapeType.EdgePath = "m,l,21600r21600,l21600,xe";

            //var stroke = new DocumentFormat.OpenXml.Vml.Stroke();
            //stroke.JoinStyle = new EnumValue<StrokeJoinStyleValues>(StrokeJoinStyleValues.Miter);
            //var path = new DocumentFormat.OpenXml.Vml.Path();
            //path.AllowGradientShape = TrueFalseValue.FromBoolean(true);
            //path.ConnectionPointType = new EnumValue<ConnectValues>(ConnectValues.Rectangle);

            //shapeType.Append(stroke);
            //shapeType.Append(path);

            var shape = new DocumentFormat.OpenXml.Vml.Shape();
            shape.Id = "_x0000_s1025";
            shape.Style =
                        "position:absolute;margin-left:55.5pt;margin-top:1pt;width:104pt;height:61.5pt;z-index:2;visibility:hidden";
            shape.InsetMode = new EnumValue<InsetMarginValues>(InsetMarginValues.Auto);
            shape.FillColor = "infoBackground [80]";
            shape.StrokeColor = "none [81]";
            //shape.StrokeWeight = "0.75pt";
            shape.Type = "#_x0000_t202";

            shape.Append(new DocumentFormat.OpenXml.Vml.Fill() { Color2 = "infoBackground [80]" });
            //shape.Append(new DocumentFormat.OpenXml.Vml.Stroke() { LineStyle = new EnumValue<StrokeLineStyleValues>(StrokeLineStyleValues.Single), DashStyle = "solid" });
            shape.Append(new DocumentFormat.OpenXml.Vml.Shadow
            { Obscured = TrueFalseValue.FromBoolean(true), Color = "none [81]" });
            shape.Append(new DocumentFormat.OpenXml.Vml.Path() { ConnectionPointType = new EnumValue<ConnectValues>(ConnectValues.None) });
            var textBox = new DocumentFormat.OpenXml.Vml.TextBox();
            textBox.Style = "mso-direction-alt:auto";
            shape.Append(textBox);
            var clientData = new DocumentFormat.OpenXml.Vml.Spreadsheet.ClientData();
            //Set the Note Type
            clientData.ObjectType = new EnumValue<ObjectValues>(ObjectValues.Note);
            clientData.Append(new MoveWithCells());
            clientData.Append(new ResizeWithCells());
            clientData.Append(new Anchor("1, 15, 0, 8, 3, 33, 4, 7"));
            clientData.Append(new AutoFill("False"));
            clientData.Append(new CommentRowTarget("1"));
            clientData.Append(new CommentColumnTarget("0"));

            shape.Append(clientData);

            shapeType.WriteTo(writer);
            shape.WriteTo(writer);

            writer.WriteEndElement();
            writer.Flush();
            writer.Close();
        }
    }

    
}
