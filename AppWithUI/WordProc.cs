using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace AppWithUI
{
    public class WordProc
    {
        public void word(ref List<List<TextBox>> A, ref List<TextBox> B)
        {
            #region Basic Setup

            int n, m;
            n = A.Count;
            m = A[0].Count - 1;
            List<bool> used = new List<bool>();
            for (int i = 0; i < m; i++)
            {
                used.Add(false);
            }

            #endregion

            #region Basic Logic

            List<string> Marks = new();
            Fuac sample = new();
            bool should_i_go = true;
            sample.InsertTable(ref A, ref B, ref Marks, 0, -1);

            for (int main = 0; main < A.Count && should_i_go; main++)
            {
                bool chek = check_if_alike(ref A);
                if (chek)
                {
                    sample.InsertText("Система не совместна, т.к. есть одинаковые строки");
                    should_i_go = false;
                    break;
                }

                int RsLocation = get_min(ref A, main, ref used);
                Marks.Add("X" + (RsLocation + 1));

                double Rs = Convert.ToDouble(A[main][RsLocation].Text);
                sample.InsertText($"Разрешающий элемент : {Rs}");

                sample.InsertTable(ref A, ref B, ref Marks, main, RsLocation);
                for (int i = 0; i < A.Count; i++)
                {
                    if (i == main)
                        continue;

                    double temq = Convert.ToDouble(B[i].Text);
                    temq -= Convert.ToDouble(B[main].Text) * Convert.ToDouble(A[i][RsLocation].Text) / Rs;
                    temq = Math.Round(temq, 4);
                    B[i].Text = temq.ToString();

                    for (int j = 0; j < A[i].Count - 1; j++)
                    {
                        if ((Convert.ToDouble(A[i][j].Text) == 0) || (j == RsLocation))
                            continue;

                        double qik = Convert.ToDouble(A[i][j].Text);
                        qik -= Convert.ToDouble(A[main][j].Text) * Convert.ToDouble(A[i][RsLocation].Text) / Rs;
                        qik = Math.Round(qik, 4);
                        A[i][j].Text = qik.ToString();
                    }
                }
                divide_the_row(ref A, ref B, RsLocation, main);
                sample.InsertText($"После обнуления x{RsLocation + 1} столбца");
                nullify_the_column(ref A, RsLocation, main);
                sample.InsertTable(ref A, ref B, ref Marks, main, RsLocation);
                IfHasLineIsZero(ref A, ref B);
            }

            if (should_i_go)
            {
                sample.InsertText("Решённая таблица:");
                sample.InsertTable(ref A, ref B, ref Marks, 0, -1);
                sample.InsertText("Ответ:");
                result(A, B);
            }

            sample.Close();

            #endregion

            #region Functions

            int get_min(ref List<List<TextBox>> vec, int i, ref List<bool> mark)
            {
                double min = double.MaxValue;
                int location = 0;
                for (int j = 0; j < vec[i].Count - 1; j++)
                {
                    if (Math.Abs(min) > Math.Abs(Convert.ToDouble(vec[i][j].Text)) && Convert.ToDouble(vec[i][j].Text) != 0)
                    {
                        min = Convert.ToDouble(vec[i][j].Text);
                        location = j;
                    }
                }
                mark[location] = true;
                return location;
            }

            void divide_the_row(ref List<List<TextBox>> vec1, ref List<TextBox> vec2, int element, int position)
            {
                double save, saveAgain;
                double aaa = Convert.ToDouble(vec1[position][element].Text);
                for (int j = 0; j < vec1.Count; j++)
                {
                    save = Convert.ToDouble(vec1[position][j].Text);
                    save /= aaa;
                    save = Math.Round(save, 4);
                    vec1[position][j].Text = save.ToString();
                }
                saveAgain = Convert.ToDouble(vec2[position].Text);
                saveAgain /= aaa;
                saveAgain = Math.Round(saveAgain, 4);
                vec2[position].Text = saveAgain.ToString();
            }

            bool check_if_alike(ref List<List<TextBox>> A)
            {
                bool check = false;
                for (int main = 0; main < A.Count; main++)
                {
                    int count = 0;
                    if (check) break;
                    for (int i = main + 1; i < A.Count; i++)
                    {
                        count = 0;
                        for (int j = 0; j < A[i].Count - 1; j++)
                        {
                            if (A[main][j].Text == A[i][j].Text)
                            {
                                count++;
                            }
                        }
                        if (count == A.Count)
                        {
                            check = true;
                            break;
                        }
                    }
                }
                return check;
            }

            void nullify_the_column(ref List<List<TextBox>> vec1, int column, int row)
            {
                for (int i = 0; i < vec1.Count; i++)
                {
                    if (i == row)
                        continue;
                    vec1[i][column].Text = "0";
                }
            }

            void IfHasLineIsZero(ref List<List<TextBox>> vec1, ref List<TextBox> vec2)
            {
                for (int i = 0; i < vec1.Count; i++)
                {
                    int zeroSum = 0;
                    for (int j = 0; j < vec1[i].Count - 1; j++)
                    {
                        if (Convert.ToDouble(vec1[i][j].Text) == 0)
                            zeroSum++;
                    }

                    if (Convert.ToDouble(vec2[i].Text) == 0)
                        zeroSum++;
                    if (zeroSum == vec1.Count + 1)
                    {
                        vec2.RemoveAt(i);
                        vec1.RemoveAt(i);
                    }
                }
            }

            void result(List<List<TextBox>> a, List<TextBox> b)
            {
                for (int i = 0; i < a.Count; i++)
                {

                    string holder = "";
                    bool inserted = false;
                    string tempik = "";
                    for (int j = 0; j < a[i].Count - 1; j++)
                    {
                        if (a[i][j].Text != "0")
                        {
                            if (inserted)
                            {
                                holder += "+";
                            }
                            if (string.Equals(a[i][j].Text, "1"))
                            {
                                tempik = "";
                            }
                            else if (string.Equals(a[i][j].Text, "-1"))
                            {
                                tempik = "-";
                            }
                            else if (string.Equals(a[i][j].Text, "-0"))
                            {
                                continue;
                            }
                            else
                            {
                                tempik = a[i][j].Text;
                            }

                            holder += tempik + "X" + Utf_Version((j + 1).ToString());
                            inserted = true;
                        }
                    }

                    holder += " = " + b[i].Text + "\n";
                    sample.InsertMathExp(holder);
                }
            }

            string Utf_Version(string text)
            {
                if (string.Equals(text, "0"))
                    return "\u2080";
                else if (string.Equals(text, "1"))
                    return "\u2081";
                else if (string.Equals(text, "2"))
                    return "\u2082";
                else if (string.Equals(text, "3"))
                    return "\u2083";
                else if (string.Equals(text, "4"))
                    return "\u2084";
                else if (string.Equals(text, "5"))
                    return "\u2085";
                else if (string.Equals(text, "6"))
                    return "\u2086";
                else if (string.Equals(text, "7"))
                    return "\u2087";
                else if (string.Equals(text, "8"))
                    return "\u2088";
                else if (string.Equals(text, "9"))
                    return "\u2089";
                else return "";
            }

            #endregion
        }


        public class Fuac
        {
            public Fuac()
            {
                wordGen.oword.Visible = true;
                wordTemplates.MarginsOfPage(ref wordGen.odoc, 30, 30, 30, 30);
                wordGen.odoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;

                Word.Paragraph opara1;
                opara1 = wordGen.odoc.Content.Paragraphs.Add(ref wordGen.omissing);
                opara1.Range.Text = "Решение системы:";
            }

            public void InsertText(string text)
            {
                Word.Paragraph para;
                para = wordTemplates.ParagraphText(wordGen.odoc, text, 16, (int)Word.WdParagraphAlignment.wdAlignParagraphLeft, 10, ref wordGen.oendofdoc);
            }
            public void InsertMathExp(string text)
            {
                Word.Range wrdrng = wordGen.odoc.Bookmarks.get_Item(ref wordGen.oendofdoc).Range;
                wrdrng.Text = text;
                wrdrng.OMaths.Add(wrdrng);
            }
            public void InsertTable(ref List<List<TextBox>> v1, ref List<TextBox> v2, ref List<string> marks, int step, int location)
            {
                int n = v1.Count;
                int m = v1[0].Count - 1;
                Word.Table otable;
                Word.Range wrdrng = wordGen.odoc.Bookmarks.get_Item(ref wordGen.oendofdoc).Range;
                otable = wordGen.odoc.Tables.Add(wrdrng, n + 1, m + 3, 1, 2);
                otable.Range.ParagraphFormat.SpaceAfter = 6;

                otable.Cell(1, 1).Range.Text = "№ Шага";
                otable.Cell(1, 2).Range.Text = "БП";
                for (int j = 3; j < m + 3; j++)
                {
                    otable.Cell(1, j).Range.Text = "x" + (j - 2);
                }
                otable.Cell(1, m + 3).Range.Text = "b";

                for (int j = 0; j < marks.Count; j++)
                {
                    otable.Cell(j + 2, 2).Range.Text = marks[j];
                }

                otable.Cell(2, 1).Range.Text = (step + 1).ToString();

                //for RS logic here
                //string conc = n.ToString() + " " + m.ToString();
                //otable.Cell(2, 2).Range.Text = conc;

                for (int i = 2; i < n + 2; i++)
                {
                    for (int j = 3; j < m + 3; j++)
                    {
                        otable.Cell(i, j).Range.Text = v1[i - 2][j - 3].Text;
                    }
                    otable.Cell(i, m + 3).Range.Text = v2[i - 2].Text;
                }
                if (location != -1)
                    otable.Cell(step + 2, location + 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorAqua;
            }
            public void Close()
            {
                wordGen.oword.Quit();
                wordGen.odoc.Close();
            }

            public AppWithUI.WordGen wordGen = new AppWithUI.WordGen();
            public WordTemplates wordTemplates = new();
        }
    }
}