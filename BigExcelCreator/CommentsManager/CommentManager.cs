// Ignore Spelling: xml Tahoma xe mso

using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace BigExcelCreator.CommentsManager
{
    internal class CommentManager
    {
        // Inspired by (and mostly copied from): https://www.dritsoftware.com/docs/netspreadsheet/openxmlsdk/CommentsFormatting.html

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

            LegacyDrawing legacyDrawing = new() { Id = worksheetPart.GetIdOfPart(vmlDrawingPart) };

            worksheetPart.Worksheet.Append(legacyDrawing);


            using (XmlWriter writer = BuildVmlDrawingPartBegin(vmlDrawingPart))
            {
                Authors authors = new();
                Comments comments = new();
                CommentList commentList = new();


                foreach (CommentReference CommentToBeAdded in CommentsToBeAdded.OrderBy(x => x.CellRange))
                {
                    if (!AuthorsList.Contains(CommentToBeAdded.Author))
                    {
                        Author author = new() { Text = CommentToBeAdded.Author };
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
                    CommentText commentTextElement = new();

                    Run run = new();
                    RunProperties runProperties = new();
                    FontSize fontSize = new() { Val = 9D };
                    Color color = new() { Rgb = new HexBinaryValue("FF000000") };
                    FontFamily family = new() { Val = 2 };
                    RunFont runFont = new() { Val = "Tahoma" };

                    runProperties.Append(fontSize);
                    runProperties.Append(color);
                    runProperties.Append(runFont);
                    runProperties.Append(family);
                    Text text = new() { Text = CommentToBeAdded.Text };

                    run.Append(runProperties);
                    run.Append(text);

                    commentTextElement.Append(run);
                    comment.Append(commentTextElement);
                    commentList.Append(comment);

                    CellRange cell = CommentToBeAdded.CellRange;
                    BuildVmlDrawingPartAdd(writer, cell.StartingRow.Value, cell.StartingColumn.Value);
                }

                comments.Append(authors);
                comments.Append(commentList);
                worksheetCommentsPart.Comments = comments;
                worksheetCommentsPart.Comments.Save();

                BuildVmlDrawingPartEnd(writer);
            }


            worksheetPart.Worksheet.Save();
        }



        private static XmlWriter BuildVmlDrawingPartBegin(VmlDrawingPart vmlDrawingPart)
        {
            XmlWriterSettings xmlWriterSettings = new() { Encoding = Encoding.UTF8 };
            XmlWriter writer = XmlWriter.Create(vmlDrawingPart.GetStream(FileMode.Create), xmlWriterSettings);
            writer.WriteStartElement("xml");

            Shapetype shapeType = new()
            {
                Id = "_x0000_t202",
                CoordinateSize = "21600,21600",
                OptionalNumber = 202,
                EdgePath = "m,l,21600r21600,l21600,xe"
            };

            shapeType.WriteTo(writer);

            return writer;
        }

        private static void BuildVmlDrawingPartAdd(XmlWriter writer, int rowId, int colId)
        {
            Shape shape = new()
            {
                Id = "_x0000_s1025",
                Style =
                        "position:absolute;margin-left:55.5pt;margin-top:1pt;width:104pt;height:61.5pt;z-index:2;visibility:hidden",
                InsetMode = new EnumValue<InsetMarginValues>(InsetMarginValues.Auto),
                FillColor = "infoBackground [80]",
                StrokeColor = "none [81]",
                Type = "#_x0000_t202"
            };

            shape.Append(new DocumentFormat.OpenXml.Vml.Fill() { Color2 = "infoBackground [80]" });
            shape.Append(new DocumentFormat.OpenXml.Vml.Shadow { Obscured = TrueFalseValue.FromBoolean(true), Color = "none [81]" });
            shape.Append(new DocumentFormat.OpenXml.Vml.Path() { ConnectionPointType = new EnumValue<ConnectValues>(ConnectValues.None) });
            TextBox textBox = new() { Style = "mso-direction-alt:auto" };
            shape.Append(textBox);
            ClientData clientData = new()
            {
                //Set the Note Type
                ObjectType = new EnumValue<ObjectValues>(ObjectValues.Note)
            };
            clientData.Append(new MoveWithCells());
            clientData.Append(new ResizeWithCells());
            clientData.Append(new Anchor($"{colId}, 10, {rowId - 1}, 10, {colId + 2}, 0, {rowId + 3}, 0"));
            clientData.Append(new AutoFill("False"));
            clientData.Append(new CommentRowTarget((rowId - 1).ToString(CultureInfo.InvariantCulture)));
            clientData.Append(new CommentColumnTarget((colId - 1).ToString(CultureInfo.InvariantCulture)));

            shape.Append(clientData);

            shape.WriteTo(writer);
        }


        private static void BuildVmlDrawingPartEnd(XmlWriter writer)
        {
            writer.WriteEndElement();
        }
    }
}
