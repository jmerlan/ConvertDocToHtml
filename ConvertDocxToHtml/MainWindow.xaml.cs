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
using Spire.Doc;

namespace ConvertDocxToHtml
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Fields to log statistics
        public static List<string> convertedDocs = new List<string>();
        public static List<string> docsNotConverted = new List<string>();
        StringBuilder stringConvertedDocs = new StringBuilder();
        StringBuilder stringNotConvertedDocs = new StringBuilder();


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
        /// <param name="docFilePath">The file to be converted</param>
        /// <param name="htmlFilePath">The path of the newly created HTML file</param>
        private void ConvertDirToHtml(string docFilePath, string htmlFilePath)
        {
            byte[] byteArray = File.ReadAllBytes(docFilePath);

            // Get file location (directory only)
            string fileFolder = System.IO.Path.GetDirectoryName(docFilePath);

            // Get extension of file for format checking
            string fileExtension = System.IO.Path.GetExtension(docFilePath);

            // Get filename of file for HTML title
            string fileName = System.IO.Path.GetFileName(docFilePath);

            // Temporary file location
            string tempFile = @"C:\temp\temp.docx";

            // Process if filetype is DocX
            if (fileExtension == ".docx" || fileExtension == ".doc")
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    //System.Windows.Forms.MessageBox.Show(docXFilePath);
                    memoryStream.Write(byteArray, 0, byteArray.Length);

                    try
                    {
                        using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                        {
                            HtmlConverterSettings settings = new HtmlConverterSettings()
                            {
                                PageTitle = fileName
                            };

                            doc.Clone(tempFile);

                            WordprocessingDocument tempDoc = WordprocessingDocument.Open(tempFile, true);

                            XElement html = HtmlConverter.ConvertToHtml(tempDoc, settings);


                            File.WriteAllText(htmlFilePath, html.ToStringNewLineOnAttributes());
                        }

                        convertedDocs.Add(fileName);
                    }
                    catch (Exception ex)
                    {
                        //System.Windows.Forms.MessageBox.Show(ex.ToString());
                        docsNotConverted.Add(fileName);
                    }
                }
            }
            else
            {
                // Do nothing?
            }

            foreach (string s in convertedDocs)
            {
                stringConvertedDocs.Append(s + "\n");
            }
            foreach (string s in docsNotConverted)
            {
                stringNotConvertedDocs.Append(s + "\n");
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

            //System.Windows.Forms.MessageBox.Show(stringConvertedDocs.ToString());
            //System.Windows.Forms.MessageBox.Show(stringNotConvertedDocs.ToString());

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
