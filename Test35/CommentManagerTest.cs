using BigExcelCreator;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static Test35.TestHelperMethods;

namespace Test
{
    internal class CommentManagerTest
    {
        [SetUp]
        public void Setup()
        {
            // Method intentionally left empty.
        }

        [TearDown]
        public void Teardown()
        {
            // Method intentionally left empty.
        }

        const string hasComment = "this should be commented";
        const string hasNoComment = "this should not be commented";



        [Test]
        public void Comments()
        {
            MemoryStream memoryStream;
            using (BigExcelwriter writer = GetwriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("name");
                writer.WriteTextRow(new[] { hasComment, hasNoComment, hasComment });
                writer.Comment("first comment", "A1");
                writer.Comment("second comment", "C1");
            }


            using (SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false))
            {
                WorksheetPart worksheetPart = reader.WorkbookPart.WorksheetParts.First();
                Row row = GetRows(worksheetPart.Worksheet).First();
                IEnumerable<Cell> cells = GetCells(row);
                Assert.Multiple(() =>
                {
                    Assert.That(GetCellRealValue(cells.ElementAt(0), reader.WorkbookPart), Is.EqualTo(hasComment));
                    Assert.That(GetCellRealValue(cells.ElementAt(1), reader.WorkbookPart), Is.EqualTo(hasNoComment));
                    Assert.That(GetCellRealValue(cells.ElementAt(2), reader.WorkbookPart), Is.EqualTo(hasComment));

                    var worksheetCommentsPart = worksheetPart.WorksheetCommentsPart;
                    Assert.That(worksheetCommentsPart, Is.Not.Null);
                    Assert.That(worksheetCommentsPart.Comments, Is.Not.Null);
                    Assert.That(worksheetCommentsPart.Comments.CommentList, Is.Not.Null);
                    Assert.That(worksheetCommentsPart.Comments.CommentList.Count(), Is.EqualTo(2));
                    Assert.That(worksheetCommentsPart.Comments.CommentList.ElementAt(0), Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(0)).Reference, Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(0)).Reference.Value, Is.EqualTo("A1"));
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(0)).CommentText, Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(0)).CommentText.InnerText, Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(0)).CommentText.InnerText, Is.EqualTo("first comment"));
                    Assert.That(worksheetCommentsPart.Comments.CommentList.ElementAt(1), Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(1)).Reference, Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(1)).Reference.Value, Is.EqualTo("C1"));
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(1)).CommentText, Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(1)).CommentText.InnerText, Is.Not.Null);
                    Assert.That(((Comment)worksheetCommentsPart.Comments.CommentList.ElementAt(1)).CommentText.InnerText, Is.EqualTo("second comment"));
                });
            }

        }

    }
}
