using System;
using System.Collections.Generic;
using System.Configuration;
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
using LumenWorks.Framework.IO.Csv;
using Novacode;
using MessageBox = System.Windows.MessageBox;

namespace DotNetDocxMerge
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

        private void btnTemplate_Click(object sender, RoutedEventArgs e)
        {
            txtTemplate.Text = OpenSelectFileDialog("docx", "Word Document");
        }

        private void btnCsv_Click(object sender, RoutedEventArgs e)
        {
            txtCsv.Text = OpenSelectFileDialog("csv", "Text Files" );
        }

        private void btnDist_Click(object sender, RoutedEventArgs e)
        {
            txtDist.Text = OpenSaveFileDialog();
        }

        

        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            string dist = txtDist.Text;
            string csv = txtCsv.Text;
            string template = txtTemplate.Text;
            Task t = WriteFileAsync(template,csv,dist);
            await t;
        }

        #region WriteToFile

        private async Task WriteFileAsync(string pathTemplate, string pathCsv, string pathDist)
        {
            await Task.Run(() => WriteFile(pathTemplate, pathCsv, pathDist));
        }

        private void WriteFile(string pathTemplate, string pathCsv, string pathDist)
        {
            try
            {
                Task task = DisableAllButtonsAsync();
                using (DocX docuement = DocX.Create(pathDist))
                {
                    using (DocX template = DocX.Load(pathTemplate))
                    {
                        int lineCount = 0;
                        using (var csv = new CsvReader(new StreamReader(pathCsv), true))
                        {
                            lineCount = csv.Count();
                        }
                        using (var csv =
                            new CsvReader(new StreamReader(pathCsv), true))
                        {
                            var fieldCount = csv.FieldCount;

                            var headers = csv.GetFieldHeaders();
                            while (csv.ReadNextRecord())
                            {
                                if (csv.CurrentRecordIndex != 0) docuement.InsertSectionPageBreak();
                                docuement.InsertDocument(template);
                                for (var i = 0; i < fieldCount; i++)
                                {
                                    docuement.ReplaceText(String.Format("<<{0}>>", headers[i]), csv[i]);
                                }
                                Task t = UpdatePgbAsync((double) (csv.CurrentRecordIndex + 1)/lineCount*100);

                            }

                        }

                    }
                    docuement.Save();
                    if (docuement.Text.Contains("<<") || docuement.Text.Contains(">>"))
                    {
                        MessageBox.Show("There is some << / >> not merged.");
                    }
                    Task t2 = FinishMergeAsync();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("{0}", ex.Message), "Warning", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            finally
            {
                Task task = EnableAllButtonsAsync();
            }
            
        }

        #endregion

        #region Open FileDialog

        private string OpenSelectFileDialog(string format, String formatDesc)
        {
            string fileName = "";
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = String.Format(".{0}", format); // Default file extension
            dlg.Filter = String.Format("{0}|*.{1}", formatDesc, format); // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                fileName = dlg.FileName;
            }

            return fileName;
        }  

        private string OpenSaveFileDialog()
        {
            string fileName = "";
            // Configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".docx"; // Default file extension
            dlg.Filter = "Word document|*.docx"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                fileName = dlg.FileName;
            }

            return fileName;
        }

        #endregion

        #region Set enable button

        private async Task DisableAllButtonsAsync() { await Task.Run(() => Dispatcher.Invoke(DisableAllButtons)); }
        private async Task EnableAllButtonsAsync() { await Task.Run(() => Dispatcher.Invoke(EnableAllButtons)); }
        private void DisableAllButtons() { SetEnableToAllButton(false); }
        private void EnableAllButtons() { SetEnableToAllButton(true); }

        private void SetEnableToAllButton(bool boolVal)
        {
            btnCsv.IsEnabled = boolVal;
            btnDist.IsEnabled = boolVal;
            btnTemplate.IsEnabled = boolVal;
            btnStart.IsEnabled = boolVal;
        }

        #endregion

        #region Handle finish merge

        private async Task FinishMergeAsync()
        {
            await Task.Run(() => Dispatcher.Invoke(FinishMerge));
        }

        private void FinishMerge()
        {
            UpdatePgb(0D);
            MessageBox.Show("Merge Done");
        }

        #endregion

        #region Handle progressBar update

        private async Task UpdatePgbAsync(double percentage)
        {
            Task t = Task.Run(() =>
                Dispatcher.Invoke(() =>
                {
                    UpdatePgb(percentage);
                })
            );
            await t;
        }

        private void UpdatePgb(double percentage)
        {
            txtpgb.Text = String.Format("{0:0.00}%", percentage);
            pgb.Value = percentage;
        }

        #endregion

        
        /// <summary>
        /// This is the function to init the txtbox with config value
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Initialized(object sender, EventArgs e)
        {
            txtTemplate.Text = DotNetDocxMerge.Properties.Settings.Default.template;
            txtCsv.Text = DotNetDocxMerge.Properties.Settings.Default.csv;
            txtDist.Text = DotNetDocxMerge.Properties.Settings.Default.dist;
        }

        
    }
}
