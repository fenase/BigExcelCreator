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
using System.Linq;
using BigExcelCreator.Ranges;
using System.Globalization;

namespace BigExcelCreator.CommentsManager
{
    internal class CommentManager
    {
        private List<CommentReference> CommentsToBeAdded { get; set; }

        private List<string> AuthorsList { get; set; }


        internal CommentManager()
        {
            CommentsToBeAdded = new();
            AuthorsList = new();
        }

        internal void Add(CommentReference commentReference)
        {
            CommentsToBeAdded.Add(commentReference);
        }

        internal void SaveComments(WorksheetPart worksheetPart)
        {
            VmlDrawingPart vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
            WorksheetCommentsPart worksheetCommentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();

            LegacyDrawing legacyDrawing = new LegacyDrawing() { Id = worksheetPart.GetIdOfPart(vmlDrawingPart) };

            worksheetPart.Worksheet.Append(legacyDrawing);


            var wrtiter = BuildVmlDrawingPartBegin(vmlDrawingPart);
            Authors authors = new Authors();
            var comments = new Comments();
            CommentList commentList = new CommentList();


            foreach (var CommentToBeAdded in CommentsToBeAdded.OrderBy(x => x.CellRange))
            {
                if (!AuthorsList.Contains(CommentToBeAdded.Author))
                {
                    Author author = new Author();
                    author.Text = CommentToBeAdded.Author;
                    authors.Append(author);
                    AuthorsList.Add(CommentToBeAdded.Author);
                }


                Comment comment;
                if (!string.IsNullOrEmpty(CommentToBeAdded.Author))
                {
                    comment = new Comment() { Reference = CommentToBeAdded.Cell, AuthorId = (UInt32Value)(uint)AuthorsList.IndexOf(CommentToBeAdded.Author) };
                }
                else
                {
                    comment = new Comment() { Reference = CommentToBeAdded.Cell };
                }
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
                text.Text = CommentToBeAdded.Text;

                run.Append(runProperties);
                run.Append(text);

                commentTextElement.Append(run);
                comment.Append(commentTextElement);
                commentList.Append(comment);

                CellRange cell = CommentToBeAdded.CellRange;
                BuildVmlDrawingPartAdd(wrtiter, cell.StartingColumn.Value, cell.StartingRow.Value);
            }

            comments.Append(authors);
            comments.Append(commentList);
            worksheetCommentsPart.Comments = comments;
            worksheetCommentsPart.Comments.Save();

            BuildVmlDrawingPartEnd(wrtiter);


            worksheetPart.Worksheet.Save();
        }



        private static XmlTextWriter BuildVmlDrawingPartBegin(VmlDrawingPart vmlDrawingPart)
        {
            var writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8);
            writer.WriteStartElement("xml");
            return writer;
        }

        private static void BuildVmlDrawingPartAdd(XmlTextWriter writer, int rowId, int colId)
        {
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
            clientData.Append(new CommentRowTarget((rowId - 1).ToString(CultureInfo.InvariantCulture)));
            clientData.Append(new CommentColumnTarget((colId - 1).ToString(CultureInfo.InvariantCulture)));

            shape.Append(clientData);

            shapeType.WriteTo(writer);
            shape.WriteTo(writer);
        }


        private static void BuildVmlDrawingPartEnd(XmlTextWriter writer)
        {
            writer.WriteEndElement();
            writer.Flush();
            writer.Close();
        }
    }

    
}
