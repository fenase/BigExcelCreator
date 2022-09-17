using BigExcelCreator.Ranges;

namespace BigExcelCreator.CommentsManager
{
    internal class CommentReference
    {
        public string Cell { get; set; }
        public string Text { get; set; }
        public string Author { get; set; }

        internal CellRange CellRange { get => new(Cell); }
    }
}
