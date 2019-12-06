using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Xceed.Words.NET;

namespace DocMergeUtil
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<MacrosItem> MacrosList = new ObservableCollection<MacrosItem>();
        private ObservableCollection<string> FileList = new ObservableCollection<string>();
        private CommonOpenFileDialog OpenCatalogDialog = new CommonOpenFileDialog()
        {
            IsFolderPicker = true
        };
        private CommonOpenFileDialog HandleSaveCatalogDialog = new CommonOpenFileDialog()
        {
            IsFolderPicker = true
        };

        public MainWindow()
        {
            InitializeComponent();
            Init();
        }

        private void Init()
        {
            MacrosGrid.ItemsSource = MacrosList;
            CatalogList.ItemsSource = FileList;

            MacrosList.Add(new MacrosItem() { MacrosName = "Code", NewValue = "12344321" });
            MacrosList.Add(new MacrosItem() { MacrosName = "Service", NewValue = "Пожарные услуги" });
            MacrosList.Add(new MacrosItem() { MacrosName = "FIO", NewValue = "Иванов Иван Иванович" });
        }

        private string GetSavePath()
        {
            return OpenCatalogDialog.FileName + "\\Result_" + DateTime.Now.ToString("DD-MM-YYYY_hh-mm-ss");
        }

        private void AddClick(object sender, RoutedEventArgs e)
        {
            MacrosList.Add(new MacrosItem());
        }

        private void DeleteClick(object sender, RoutedEventArgs e)
        {
            MacrosGrid.SelectedItems.Cast<MacrosItem>()?.ToList().ForEach(p =>
            {
                MacrosList.Remove(p);
            });
        }

        private void OpenCatalogClick(object sender, RoutedEventArgs e)
        {
            if (OpenCatalogDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                Directory.GetFiles(OpenCatalogDialog.FileName).ToList().ForEach(path =>
                {
                    FileList.Add(path);
                });
            }
        }

        private void HandleAsClick(object sender, RoutedEventArgs e)
        {
            if (HandleSaveCatalogDialog.ShowDialog() != CommonFileDialogResult.Ok)
            {
                return;
            }
            var saveDirectoryPath = HandleSaveCatalogDialog.FileName;
            FileList.ToList().ForEach(path =>
            {
                using (DocX doc = DocX.Load(path))
                {

                    MacrosList.ToList().ForEach(macros =>
                    {
                        doc.ReplaceText("<" + macros.MacrosName + ">", macros.NewValue);
                    });

                    doc.SaveAs(saveDirectoryPath + "\\" + Path.GetFileName(path));
                }
            });
        }

        private void NewProjectClick(object sender, RoutedEventArgs e)
        {

        }

        private void OpenProjectClick(object sender, RoutedEventArgs e)
        {

        }

        private void SaveClick(object sender, RoutedEventArgs e)
        {

        }

        private void SaveAsClick(object sender, RoutedEventArgs e)
        {

        }

        private void ExitClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
