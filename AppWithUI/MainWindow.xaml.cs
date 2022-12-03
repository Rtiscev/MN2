using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

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
                List<TextBox> tempik = new();
                for (int j = 0; j < c; j++)
                {
                    tempik.Add(new TextBox());
                    tempik[j].MinWidth = 50;

                    WrapPanel wrapPanel = new();

                    Grid.SetRow(wrapPanel, i);
                    Grid.SetColumn(wrapPanel, j);

                    if (j == c - 1)
                    {
                        TextBox aaae = new();
                        aaae.MinWidth = 50;
                        boxEnd.Add(aaae);
                        boxEnd[i].Name = "b" + (i + 1);
                        wrapPanel.Children.Add(boxEnd[i]);
                    }
                    else
                    {
                        tempik[j].Name = "x" + ((j + 1).ToString());
                        wrapPanel.Children.Add(tempik[j]);
                    }
                    GridManipulation.Children.Add(wrapPanel);
                }
                box.Add(tempik);
            }
        }
        private void Solve_equation(object sender, RoutedEventArgs e)
        {
            WordProc wordproc = new();
            string chosenMethod = (method_Jacobi.IsChecked == true ? "Jacobi" : "Jordan-Gaus");

            wordproc.word(ref box, ref boxEnd, chosenMethod);
        }
        private void Random_fill(object sender, RoutedEventArgs e)
        {
            Random random = new();
            for (int i = 0; i < box.Count; i++)
            {
                for (int j = 0; j < box[i].Count; j++)
                {
                    box[i][j].Text = random.Next(10).ToString();
                }
                boxEnd[i].Text = random.Next(10).ToString();
            }
        }

        public List<List<TextBox>> box = new();
        public List<TextBox> boxEnd = new();
    }
}/*
3 2 1 -1
1 2 3 -1
2 3 1 1
5 5 2 0
1 1 1 2
*/
