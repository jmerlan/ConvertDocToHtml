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
using System.Windows.Forms;
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

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Gets a list of all files from the selected folder
        /// </summary>
        private string[] GetAllFiles(string dirPath)
        {
            // String for testing
            StringBuilder sb = new StringBuilder();

            // Get path of selected directory
            dirPath = txtDirPath.Text;
            string[] files = Directory.GetFiles(dirPath, "*", SearchOption.AllDirectories);

            // Print test string of all filenames
            foreach (string f in files)
            {
                sb.Append(System.IO.Path.GetFileName(f) + "\n");
            }
            System.Windows.MessageBox.Show(sb.ToString());

            return files;
        }

        /// <summary>
        /// Converts a file to HTML
        /// </summary>
        private void ConvertDirToHtml(string docXFilePath, string htmlFilePath)
        {
            //byte[] byteArray = File.ReadAllBytes(docXFilePath);
            byte[] byteArray = File.ReadAllBytes(@"C: \Users\jay.merlan\Desktop\Temp\Doc Conversion\Tower Crane Work Plan\Chapter 1 - Work Plan\Sec 1 - Overview\Overview.docx");

            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "My Page Title"
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                    File.WriteAllText(htmlFilePath, html.ToStringNewLineOnAttributes());
                }
            }
        }


        private string GenerateHtmlFilePath(string sourceFilePath)
        {
            string htmlPath = txtDirPathDest.Text;

            
            return htmlPath;
        }

        private void buttConvert_Click(object sender, RoutedEventArgs e)
        {
            // Get list of all files from the path in the text box
            string[] files = GetAllFiles(txtDirPath.Text);
            
            foreach (string f in files)
            {
                // Generate HTML filename
                string docFileName = System.IO.Path.GetFileName(f);

                System.Windows.MessageBox.Show(htmlFilePath);

                // Convert
                ConvertDirToHtml(f, htmlFilePath);
            }
        }

        private void btnFolderSource_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                txtDirPath.Text = dialog.SelectedPath;
            }

        }

        private void btnFolderDest_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                txtDirPathDest.Text = dialog.SelectedPath;
            }

        }
    }
}
