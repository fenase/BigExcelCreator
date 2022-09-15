using System;
using System.Collections.Generic;
using System.Text;

namespace BigExcelCreator.CommentsManager
{
    internal class CommentReference
    {
        public int col { get; set; }
        public int Row { get; set; }
        public string Text { get; set; }
        public string Author{ get; set; }
    }
}
