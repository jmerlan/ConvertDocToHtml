using System;
using System.Collections.Generic;
using System.IO;
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
using System.Xml.Linq;

using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace ConvertDocxToHtml
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string docXFilePath = @"C:\Users\Jay\Amazon Drive\Business\BIM Extension\Projects\14074 - MyLewis 2\Documentation\Operations\Work Plans\Tower Crane Work Plan\Chapter 1 - Work Plan\Sec 1 - Overview\Overview.docx";
        string htmlDirectory = @"C:\Users\Jay\Amazon Drive\Business\BIM Extension\Projects\14074 - MyLewis 2\Documentation\HTML\Blank.html";
        string dirPath = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ConvertDirToHtml()
        {
            StringBuilder sb = new StringBuilder();
            txtDirPath.Text = @"C:\Users\Jay\Amazon Drive\Business\BIM Extension\Projects\14074 - MyLewis 2\Documentation\Operations\Work Plans\Tower Crane Work Plan";
            dirPath = txtDirPath.Text;
            string[] files = Directory.GetFiles(dirPath, "*", SearchOption.AllDirectories);

            foreach (string f in files)
            {
                sb.Append(System.IO.Path.GetFileName(f) + "\n");
            }

            MessageBox.Show(sb.ToString());
        }

        private void buttConvert_Click(object sender, RoutedEventArgs e)
        {
            ConvertDirToHtml();

            //byte[] byteArray = File.ReadAllBytes(docXFilePath);
            //using (MemoryStream memoryStream = new MemoryStream())
            //{
            //    memoryStream.Write(byteArray, 0, byteArray.Length);
            //    using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            //    {
            //        HtmlConverterSettings settings = new HtmlConverterSettings()
            //        {
            //            PageTitle = "My Page Title"
            //        };
            //        XElement html = HtmlConverter.ConvertToHtml(doc, settings);

            //        File.WriteAllText(htmlDirectory, html.ToStringNewLineOnAttributes());
            //    }
            //}
        }
    }
}
