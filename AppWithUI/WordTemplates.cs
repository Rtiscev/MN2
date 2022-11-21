using Word = Microsoft.Office.Interop.Word;
namespace AppWithUI
{
    public class WordTemplates
    {
        public void MarginsOfPage(ref Word._Document document, int left, int right, int top, int bottom)
        {
            document.PageSetup.LeftMargin = left;
            document.PageSetup.RightMargin = right;
            document.PageSetup.TopMargin = top;
            document.PageSetup.BottomMargin = bottom;
        }
        public Word.Paragraph ParagraphText(Word._Document doc, string text, int font_size, int alignment, int spaceafter, ref object endofdoc)
        {
            object oRng = doc.Bookmarks.get_Item(ref endofdoc).Range;
            Word.Paragraph para;
            para = doc.Content.Paragraphs.Add(ref oRng);
            para.Range.Text = text;
            para.Range.Font.Size = font_size;
            para.Alignment = (Word.WdParagraphAlignment)alignment;
            para.Format.SpaceAfter = spaceafter;
            para.Range.InsertParagraphAfter();
            return para;
        }

    }
}

