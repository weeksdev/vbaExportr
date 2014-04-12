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
using MahApps.Metro.Controls;
using Microsoft.Win32;

namespace vbaExportr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            this.exportFlyout.IsOpen = false;
            //var testExcel = Export.Open(@"C:\Users\WeeksDev\Documents\Book1.xlsm");
        }

        public class ExcelFile
        {
            public List<Export.ResponseObject> Projects { get; set; }
            public string fileName { get; set; }
            public string filePath { get; set; }
        }

        ExcelFile excelFile = new ExcelFile();

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.AddExtension = true;
            ofd.Filter = "Excel Files (*.xlsm,*.xls)|*.xlsm;*.xls";
            ofd.ShowDialog();
            this.uploadFileFld.Text = ofd.FileName;
            if (ofd.FileName != "")
            {
                this.exportFlyout.IsOpen = true;
                this.excelFile.Projects = Export.Open(ofd.FileName);
                this.excelFile.fileName = ofd.FileName.Split('\\').Last();
                this.fileNameFld.Content = "File Name: " + this.excelFile.fileName;
                this.excelFile.filePath = ofd.FileName;
                this.numberOfProjectsFld.Content = "Project Count: " + this.excelFile.Projects.Count;
                int totalBasFiles = 0;
                bool protectedContent = false;
                foreach (var project in this.excelFile.Projects)
                {
                    if (project.isProtected)
                        protectedContent = true;
                    foreach (var file in project.basFiles)
                    {
                        totalBasFiles += 1;
                    }
                }
                this.numberOfModules.Content = "Module Count: " + totalBasFiles;
                if (protectedContent)
                    System.Windows.MessageBox.Show("Protected content detected. Please remove protection to extract content.");
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog();
            dlg.RootFolder = Environment.SpecialFolder.Personal;
            dlg.Description = "Please select extract folder...";
            dlg.ShowDialog();
            string selectedPath = dlg.SelectedPath;
            if (selectedPath == "")
            {
                System.Windows.MessageBox.Show("No path selected.");
            }
            else 
            {
                try
                {
                    if (this.deleteBasFiles.IsChecked == true)
                        foreach (var matchingFile in System.IO.Directory.GetFiles(selectedPath, "*.bas"))
                            System.IO.File.Delete(matchingFile);
                }
                catch
                {
                    System.Windows.MessageBox.Show("Error deleting previous bas files.");
                    return;
                }

                foreach (var project in this.excelFile.Projects)
                {
                    foreach (var file in project.basFiles)
                    {
                        System.IO.File.WriteAllText(System.IO.Path.Combine(selectedPath + "\\" + project.projectName + "." + file.componentName + ".bas"), file.code);
                    }
                }
                if (this.includeExcelFile.IsChecked==true)
                    System.IO.File.Copy(this.excelFile.filePath, System.IO.Path.Combine(selectedPath + "\\" + this.excelFile.fileName), true);

                System.Windows.MessageBox.Show("Extract Complete!");
                System.Diagnostics.Process.Start(selectedPath);
            }
        }
    }
}
