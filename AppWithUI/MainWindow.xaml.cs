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

namespace AppWithUI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int r = int.Parse(Rows.Text);
            int c = int.Parse(Columns.Text) + 1;

            RowDefinition[] rowdef = new RowDefinition[r];
            ColumnDefinition[] coldef = new ColumnDefinition[c];

            for (int i = 0; i < r; i++)
            {
                rowdef[i] = new RowDefinition();
                GridManipulation.RowDefinitions.Add(rowdef[i]);
            }
            for (int j = 0; j < c; j++)
            {
                coldef[j] = new ColumnDefinition();
                GridManipulation.ColumnDefinitions.Add(coldef[j]);
            }

            for (int i = 0; i < r; i++)
            {
                List<TextBox> tempik = new List<TextBox>();
                for (int j = 0; j < c; j++)
                {
                    //TextBlock textBlock = new TextBlock();
                    //TextBox plus = new TextBox();
                    tempik.Add(new TextBox());
                    tempik[j].MinWidth = 50;
                    //plus.MinWidth = 20;

                    WrapPanel wrapPanel = new WrapPanel();

                    Grid.SetRow(wrapPanel, i);
                    Grid.SetColumn(wrapPanel, j);

                    if (j == c - 1)
                    {
                        //textBlock.Text = "=";
                        //tempik[j].Name = "x" + (j + 1).ToString();
                        TextBox aaae = new TextBox();
                        aaae.MinWidth = 50;
                        boxEnd.Add(aaae);
                        boxEnd[i].Name = "b" + (i + 1);
                        //wrapPanel.Children.Add(textBlock);
                        //wrapPanel.Children.Add(tempik[j]);
                        wrapPanel.Children.Add(boxEnd[i]);
                    }
                    else
                    {
                        //textBlock.Text = "x" + ((j + 1).ToString());
                        tempik[j].Name = "x" + ((j + 1).ToString());
                        wrapPanel.Children.Add(tempik[j]);

                        //wrapPanel.Children.Add(textBlock);
                        if (j != c - 2)
                        {
                            //wrapPanel.Children.Add(plus);
                        }
                    }
                    GridManipulation.Children.Add(wrapPanel);
                }
                box.Add(tempik);
            }
        }
        private void Solve_equation(object sender, RoutedEventArgs e)
        {
            WordProc wordproc = new WordProc();
            wordproc.word(ref box, ref boxEnd);
            //this.win
        }
        private void Random_fill(object sender, RoutedEventArgs e)
        {
            Random random = new Random();
            for (int i = 0; i < box.Count; i++)
            {
                for (int j = 0; j < box[i].Count; j++)
                {
                    box[i][j].Text = random.Next(10).ToString();
                }
                boxEnd[i].Text = random.Next(10).ToString();
            }
        }
        private void Fill_with_mine(object sender, RoutedEventArgs e)
        {
            //List<List<double>> uh = new List<double>() { new List<double>(){ 3, 2, 1, -1 },
            //new List<double>() { 1, 2, 3, -1 },new List<double>() {2, 3, 1, 1 }, new List<double>(){5, 5, 2, 0 }, new List<double>(){1, 1, 1, 2 } };
            //for (int i = 0; i < box.Count; i++)
            //{
            //for (int j = 0; j < box[i].Count; j++)
            //{
            //box[i][j].Text = uh[i]
            //}
            //boxEnd[i].Text = random.Next(101).ToString();
            //}
        }

        public List<List<TextBox>> box = new List<List<TextBox>>();
        public List<TextBox> boxEnd = new List<TextBox>();
    }
}/*
3 2 1 -1
1 2 3 -1
2 3 1 1
5 5 2 0
1 1 1 2
*/
