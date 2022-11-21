/*
using Word = Microsoft.Office.Interop.Word;
using Apa;
using WTemplates;

#region Basic setup
Console.OutputEncoding = System.Text.Encoding.UTF8;
ApaData data = new();
WordTemplates wordTemplates = new();
object oMissing = System.Reflection.Missing.Value;
object oEndOfDoc = "\\endofdoc";  //endofdoc is a predefined bookmark 
object oStartofDoc = "\\StartOfDoc"; // start of it
#endregion  

#region Start Word and create a new document
Word._Application oWord;
Word._Document oDoc;
oWord = new Word.Application();
oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
#endregion

#region Setting up the window itself
oWord.Height = oWord.System.VerticalResolution;
oWord.Width = oWord.System.HorizontalResolution;
oWord.Visible = true;
wordTemplates.MarginsOfPage(ref oDoc, 30, 30, 30, 30);
oDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
#endregion

Word.Section section = oDoc.Sections.Add();
Word.Sections allsecods = oDoc.Sections;

var abvvv = allsecods.First;
abvvv.Borders.Enable = 1;
abvvv.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth075pt;

#region Insert Text
Word.Paragraph opara56;
opara56 = wordTemplates.ParagraphText(oDoc, "ТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ МОЛДОВЫ\nФакультет Информатики\nВычислительной техники и Микроэлектроники", 48, (int)Word.WdParagraphAlignment.wdAlignParagraphCenter, 10, ref oStartofDoc);
// Insert a paragraph at the beginning of the document.
Word.Paragraph oPara1;
oPara1 = wordTemplates.ParagraphText(oDoc, "Report", 24, (int)Word.WdParagraphAlignment.wdAlignParagraphCenter, 10, ref oStartofDoc);

// Insert a paragraph at the end of the document.
Word.Paragraph oPara2;
oPara2 = wordTemplates.ParagraphText(oDoc, "Use the Paragraphs property to return the Paragraphs collection. Use the Add(Object), InsertParagraph InsertParagraphAfter(), or InsertParagraphBefore() method to add a new paragraph to a document. Use Paragraphs(index), where index is the index number, to return a single Paragraph object. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.", 12, (int)Word.WdParagraphAlignment.wdAlignParagraphJustify, 0, ref oEndOfDoc);

// Insert another paragraph.
Word.Paragraph oPara3;
oPara3 = wordTemplates.ParagraphText(oDoc, "This is a sentence of normal text. now here is a table:", 12, (int)Word.WdParagraphAlignment.wdAlignParagraphCenter, 0, ref oEndOfDoc);
#endregion

Word.Section section1 = oDoc.Sections.Add();
section1.Borders.Enable = 0;

#region WorkingWithTable
Word.Table oTable;
Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
oTable = oDoc.Tables.Add(wrdRng, 15, 7, 1, 2);
oTable.Range.ParagraphFormat.SpaceAfter = 6;
int r, c;

Dictionary<string, Func<double, double>> myDic = new()
{
    { "LogN", data.Log2_N },
    { "LnN", data.Ln_N },
    { "\u221AN", data.Sqrt_N },
    { "N", data.N },
    { "1/N", data.One_Div_N },
    { "NLogN", data.N_Times_Log2N },
    { "N^2", data.N_Square },
    { "N^3", data.N_Cube },
    { "2^N", data.Two_In_Power_N },
    { "N!", data.FactorialSecond },
    { "N/LogN", data.N_Div_LogN },
    { "LogLogN", data.Log2_Log2N },
    { "LogN!", data.Log2N_Factorial },
    { "N^N", data.N_Times_N },
};
Dictionary<string, int> myDic2 = new()
{
    { "wdColorAqua", 13421619 },
    { "wdColorAutomatic", -16777216 },
    { "wdColorBlack", 0 },
    { "wdColorBlue", 16711680 },
    { "wdColorBlueGray", 10053222 },
    { "wdColorBrightGreen", 65280 },
    { "wdColorBrown", 13209 },
    { "wdColorDarkBlue", 8388608 },
    { "wdColorDarkGreen", 13056 },
    { "wdColorDarkRed", 128 },
    { "wdColorDarkTeal", 6697728 },
    { "wdColorDarkYellow", 32896 },
    { "wdColorGold", 52479 },
    { "wdColorGray05", 15987699 },
    { "wdColorGray10", 15132390 },
    { "wdColorGray125", 14737632 },
    { "wdColorGray15", 14277081 },
    { "wdColorGray20", 13421772 },
    { "wdColorGray25", 12632256 },
    { "wdColorGray30", 11776947 },
    { "wdColorGray35", 10921638 },
    { "wdColorGray375", 10526880 },
    { "wdColorGray40", 10066329 },
    { "wdColorGray45", 9211020 },
    { "wdColorGray50", 8421504 },
    { "wdColorGray55", 7566195 },
    { "wdColorGray60", 6710886 },
    { "wdColorGray625", 6316128 },
    { "wdColorGray65", 5855577 },
    { "wdColorGray70", 5000268 },
    { "wdColorGray75", 4210752 },
    { "wdColorGray80", 3355443 },
    { "wdColorGray85", 2500134 },
    { "wdColorGray875", 2105376 },
    { "wdColorGray90", 1644825 },
    { "wdColorGray95", 789516 },
    { "wdColorGreen", 32768 },
    { "wdColorIndigo", 10040115 },
    { "wdColorLavender", 16751052 },
    { "wdColorLightBlue", 16737843 },
    { "wdColorLightGreen", 13434828 },
    { "wdColorLightOrange", 39423 },
    { "wdColorLightTurquoise", 16777164 },
    { "wdColorLightYellow", 10092543 },
    { "wdColorLime", 52377 },
    { "wdColorOliveGreen", 13107 },
    { "wdColorOrange", 26367 },
    { "wdColorPaleBlue", 16764057 },
    { "wdColorPink", 16711935 },
    { "wdColorPlum", 6697881 },
    { "wdColorRed", 255 },
    { "wdColorRose", 13408767 },
    { "wdColorSeaGreen", 6723891 },
    { "wdColorSkyBlue", 16763904 },
    { "wdColorTan", 10079487 },
    { "wdColorTeal", 8421376 },
    { "wdColorTurquoise", 16776960 },
    { "wdColorViolet", 8388736 },
    { "wdColorWhite", 16777215 },
    { "wdColorYellow", 65535 },
};

string[] s2 = new string[] { "LogN", "LnN", "\u221AN", "N", "1/N", "NLogN", "N^2", "N^3", "2^N", "N!", "N/LogN", "LogLogN", "LogN!", "N^N" };

for (c = 2; c <= 7; c++)
{
    int power = (int)Math.Pow(2, c - 2);
    oTable.Cell(1, c).Range.Text = power.ToString();
    oTable.Cell(1, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
}

for (r = 2; r <= 15; r++)
{
    Word.Range tempR = oTable.Cell(r, 1).Range;
    tempR.Text = s2[r - 2];
    oTable.Cell(r, 1).Range.OMaths.Add(tempR);
    oTable.Cell(r, 1).Range.OMaths.BuildUp();
}

int rows;

Random randomnum = new();
for (c = 2; c <= 7; c++)
{
    rows = 2;
    foreach (var item in myDic)
    {
        Word.Range tempr = oTable.Cell(rows++, c).Range;
        string temp_text = "";
        if (data.Number_of_Digits(item.Value.Invoke(Math.Pow(2, c - 2))) > 4)
        {
            temp_text = Math.Round(item.Value.Invoke(Math.Pow(2, c - 2)), 3).ToString("0.###E+0");
        }
        else
        {
            temp_text = Math.Round(item.Value.Invoke(Math.Pow(2, c - 2)), 3).ToString();
        }
        tempr.Text = temp_text;
        oTable.Cell(rows, c).Range.OMaths.Add(tempr);
        oTable.Cell(rows, c).Range.OMaths.BuildUp();
        oTable.Cell(rows, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
    }
}

for (r = 2; r <= 15; r++)
{
    for (c = 2; c <= 7; c++)
    {
        var value = randomnum.Next(myDic2.Count);
        KeyValuePair<string, int> pair = myDic2.ElementAt(value);
        oTable.Cell(r, c).Shading.BackgroundPatternColor = (Word.WdColor)pair.Value;
    }
}


oTable.Rows[1].Range.Font.Bold = 1;
oTable.Rows[1].Range.Font.Italic = 1;
#endregion

#region Text
// Insert P
Word.Paragraph oPara4;
object oRngdd = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
oPara4.Range.Text = "This is a sentence of normal text. Now here is a table:";
oPara4.Range.Font.Bold = 0;
oPara4.Format.SpaceAfter = 24;
oPara4.Range.InsertParagraphAfter();

// Insert P
wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
wrdRng.InsertParagraphAfter();
wrdRng.InsertAfter("THE END.");
wrdRng.InsertParagraphAfter();
#endregion


//frame.Borders.InsideLineStyle = (Word.WdLineStyle)Word.WdLineWidth.wdLineWidth050pt;
// Insert Picture
object rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
Word.InlineShape aaa;
aaa = oDoc.InlineShapes.AddPicture(@"C:\Users\djuls\Pictures\jinx2.png", ref oMissing, ref oMissing, ref rng);
aaa.ScaleHeight = 50;
aaa.ScaleWidth = 50;

DirectoryInfo mydir = new DirectoryInfo(@"C:\Users\djuls\Documents\graphs");
FileInfo[] f = mydir.GetFiles();


foreach (FileInfo file in f)
{
    object thif = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
    Word.InlineShape img;
    img = oDoc.InlineShapes.AddPicture(file.FullName, ref oMissing, ref oMissing, ref thif);
    img.ScaleHeight = 40;
    img.ScaleWidth = 40;
    Console.WriteLine("File Name: {0} Size: {1}", file.Name, file.Length);
}
// Save in .docx and .pdf
oDoc.SaveAs2("report", 17);
oDoc.SaveAs2("report", 16);
oWord.Visible = true;

var ghtue = allsecods.Last;
ghtue.Borders.Enable = 0;
//abvvv.Borders.Enable = 1;
//abvvv.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth075pt;

// Close instances
oDoc.Close();
oWord.Quit();
*/