using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PasteWithFormatting
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private object data;
         
        public Window1(object data)
        {
            this.data = data;
            InitializeComponent();
        }

        private void ApplyFormattingClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void textBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            var newText = "";

            foreach (var line in data.ToString().Replace("\r\n", "\n").Trim('\n').Split('\n'))
            {
                try { newText += String.Format(Formatting.Text, line.Split(',')) + "\r\n"; }
                catch (Exception) {  }
            }

            this.Preview.Text = newText;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            this.Formatting.Text = "";
            this.Close();
        }


    }
}
