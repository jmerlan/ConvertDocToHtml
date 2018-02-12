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

            return files;
        }

        /// <summary>
        /// Converts a file to an HTML file
        /// </summary>
        /// <param name="docXFilePath">The file to be converted</param>
        /// <param name="htmlFilePath">The path of the newly created HTML file</param>
        private void ConvertDirToHtml(string docXFilePath, string htmlFilePath)
        {
            byte[] byteArray = File.ReadAllBytes(docXFilePath);
            //byte[] byteArray = File.ReadAllBytes(@"C:\Users\Jay\Amazon Drive\Business\BIM Extension\Projects\14074 - MyLewis 2\Documentation\Operations\Work Plans\Tower Crane Work Plan\Chapter 1 - Work Plan\Sec 1 - Overview\Overview.docx");

            // Get extension of file for format checking
            string fileExtension = System.IO.Path.GetExtension(docXFilePath);

            if (fileExtension == ".docx")
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    System.Windows.Forms.MessageBox.Show(docXFilePath);
                    memoryStream.Write(byteArray, 0, byteArray.Length);

                    try
                    {
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
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.ToString());
                    }
                }
            }
            else
            {
                // Do nothing?
            }
            
        }


        private string GenerateHtmlFilePath(string sourceFilePath)
        {
            //string htmlPath = @"C:\temp\HTML";

            string htmlPath = txtDirPathDest.Text;

            // Get filename of source
            // Generate HTML filename
            string docFileName = System.IO.Path.GetFileName(sourceFilePath);


            htmlPath = htmlPath + docFileName + ".html";

            return htmlPath;
        }

        private void buttConvert_Click(object sender, RoutedEventArgs e)
        {
            // Get list of all files from the path in the text box
            string[] files = GetAllFiles(txtDirPath.Text);
            
            foreach (string f in files)
            {
                // Convert
                ConvertDirToHtml(f, GenerateHtmlFilePath(f));
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
