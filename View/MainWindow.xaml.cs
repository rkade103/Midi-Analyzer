﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.WindowsAPICodePack.Dialogs;
using Midi_Analyzer.Logic;

namespace Midi_Analyzer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string sourceFileType;
        private Analyzer analyzer;

        public MainWindow()
        {
            InitializeComponent();
            sourceFileType = "MIDI";
            this.errorDetection.IsEnabled = false;
            this.results.IsEnabled = false;
        }

        /// <summary>
        /// This method is meant to clear the contents of the source path and array should the user pick a different file type.
        /// It also checks which radio button is now checked, and assigns that to the sourceFileType variable.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonCheckChange(object sender, RoutedEventArgs e)
        {
            ListBox path = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            path.Items.Clear();
            RadioButton midiButton = (RadioButton)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("midiButton");
            if(midiButton.IsChecked == true)
            {
                sourceFileType = "MIDI";
            }
            else
            {
                sourceFileType = "CSV";
            }
        }

        /// <summary>
        /// Populates the listbox of source files with the selected files from the user. Also opens the dialog to select the data.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PopulateSourceListbox(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            if (sourceFileType == "MIDI")
            {
                dlg.Filter = "MIDI files|*.MID;*.MIDI";
            }
            else
            {
                dlg.DefaultExt = ".csv";
                dlg.Filter = "CSV Files (*.csv)|*.csv";
            }
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true && dlg.FileNames.Length != 0)
            {
                ListBox sourcePathBox = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
                //Clear the existing items from the box.
                sourcePathBox.Items.Clear();
                foreach(string file in dlg.FileNames)
                {
                    sourcePathBox.Items.Add(file);
                }
            }
        }

        /// <summary>
        /// Opens a dialog to browse for a file. Enforces the selection of either midi or csv files.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForFile(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            if (sourceFileType == "MIDI")
            {
                dlg.Filter = "MIDI files|*.MID;*.MIDI";
            }
            else
            {
                dlg.DefaultExt = ".csv";
                dlg.Filter = "CSV Files (*.csv)|*.csv";
            }
            Nullable<bool> result = dlg.ShowDialog();
            if(result == true && dlg.FileNames.Length != 0)
            {
                TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
                path.Text = String.Join(";\n", dlg.FileNames);
            }
        }

        /// <summary>
        /// Opens a dialog that can only select .mid files.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForMidi(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "MIDI files|*.MID;*.MIDI";
            Nullable<bool> result = dlg.ShowDialog();
            if(result == true && dlg.FileNames.Length != 0)
            {
                TextBox modelBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("modelBox");
                modelBox.Text = dlg.FileName;
            }
        }

        /// <summary>
        /// Opens a dialog that can only select .csv files. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForCSV(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.DefaultExt = ".csv";
            dlg.Filter = "XLSX Files (*.xlsx)|*.xlsx";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true && dlg.FileNames.Length != 0)
            {
                TextBox excerptBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("excerptBox");
                excerptBox.Text = dlg.FileName;
            }
        }

        /// <summary>
        ///  Opens a dialog that can only select image files. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForImage(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Image Files |*.jpg;*.jpeg;*.png;*.bmp";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true && dlg.FileNames.Length != 0)
            {
                TextBox imageBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("imageBox");
                imageBox.Text = dlg.FileName;
            }
        }

        /// <summary>
        /// Opens a dialog that only allows the selection of folders.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForFolder(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
                path.Text = dialog.FileName;
            }
        }

        /// <summary>
        /// Opens a dialog to select a folder using windows forms.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseForFolderForms(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
                path.Text = dlg.SelectedPath;
            }
        }

        /// <summary>
        /// Converts the source files into either their csv or midi counterparts. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConvertFile(object sender, RoutedEventArgs e)
        {
            ListBox sPath = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
            string[] sourceFilesArray = new string[sPath.Items.Count];
            sPath.Items.CopyTo(sourceFilesArray, 0);
            List<string> sourceFiles = sourceFilesArray.ToList();
            string destinationFolder = destPath.Text;
            Converter converter = new Converter();
            
            if(sourceFileType == "CSV")         //If the source file is a csv, convert it into midi.
            {
                Console.WriteLine("Running conversion to MIDI...");
                converter.RunMIDIBatchFile(sourceFiles, destinationFolder);
            }
            else if(sourceFileType == "MIDI")   //If the source file is a mid, convert it into csv.
            {
                Console.WriteLine("Running conversion to CSV...");
                converter.RunCSVBatchFile(sourceFiles, destinationFolder);
            }
            else
            {
                Console.WriteLine("There was an error with the source file type selection.");
            }
        }

        /// <summary>
        /// Runs the first part of the analysis on the source files.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AnalyzeFile(object sender, RoutedEventArgs e)
        {
            //Get the source paths and the destination path.
            ListBox sPath = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");

            //Get the excerpt file, model and image paths.
            TextBox excerptBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("excerptBox");
            TextBox modelBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("modelBox");
            TextBox imageBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("imageBox");
            string excerptCSV = excerptBox.Text;
            string modelMidi = modelBox.Text;
            string image = imageBox.Text;

            //Make an array of source files.
            string[] sourceFilesArray = new string[sPath.Items.Count + 1]; //Add one for the model.
            sourceFilesArray[sourceFilesArray.Length - 1] = modelMidi;
            sPath.Items.CopyTo(sourceFilesArray, 0);
            List<string> sourceFiles = sourceFilesArray.ToList();
            if (!CheckSourceFiles(sourceFiles, sPath))
            {
                return; //An error was detected when checking source files.
            }
            string destinationFolder = destPath.Text;
            if (!CheckDestinationFolder(destinationFolder))
            {
                return; //An error was detected when checking the destination folder.
            }

            //Get the converter and run it on the source files.
            Converter converter = new Converter();
            converter.RunCSVBatchFile(sourceFiles, destinationFolder, false);

            //Run the first part of the analyzer and get the bad files.
            analyzer = new Analyzer(sourceFiles, destinationFolder, excerptCSV, modelMidi, image);
            List<string> badSheets = analyzer.AnalyzeCSVFilesStep1();

            //Populate next tab with the names of the bad sheets.
            ListBox xlsList = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("xlsFileList");
            xlsList.Items.Clear();
            foreach(string name in badSheets)
            {
                xlsList.Items.Add(name);
            }

            //Switch the focus to the next tab.
            TabControl tabControl = (TabControl)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("tabController");
            this.errorDetection.IsEnabled = true;
            this.results.IsEnabled = false;     //You do this in case the person has rerun the tool without closing it.
            tabControl.Items.OfType<TabItem>().SingleOrDefault(n => n.Name == "errorDetection").Focus();
        }

        /// <summary>
        /// Opens the analyzed.xlsx file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFile(object sender, MouseButtonEventArgs e)
        {
            var list = sender as ListBoxItem;   //This is programmed to be called from the listbox of bad sheets.
            TextBox destPath = this.destinationPath;
            string file = destPath.Text + "//analyzedFile.xlsx";

            Process.Start(@"" + file);
        }

        /// <summary>
        /// Runs the second part of the analyzer, where the IOI and articulation rows are created, as well as all graphs. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateGraphs(object sender, RoutedEventArgs e)
        {
            analyzer.AnalyzeCSVFilesStep2();
            this.results.IsEnabled = true;
            TabControl tabControl = (TabControl)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("tabController");
            tabControl.Items.OfType<TabItem>().SingleOrDefault(n => n.Name == "results").Focus();
        }

        /// <summary>
        /// Allows the user to delete an item from the sourcepath list box.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteItem(object sender, System.Windows.Input.KeyEventArgs e)
        {
            ListBox sourcePath = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            if(e.Key.Equals(Key.Delete) || e.Key.Equals(Key.Back))
            {
                if(sourcePath.SelectedItems.Count != 0)
                {
                    var selectedItems = sourcePath.SelectedItems;
                    for (int i = selectedItems.Count - 1; i > -1; i--)
                    {
                        sourcePath.Items.Remove(selectedItems[i]);
                    }
                }
            }
        }

        /// <summary>
        /// Opens the analyzed file worksheet.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenAnalyzedFile(object sender, RoutedEventArgs e)
        {
            TextBox destPath = this.destinationPath;
            string file = destPath.Text + "//analyzedFile.xlsx";

            Process.Start(@"" + file);
        }

        /// <summary>
        /// Checks for common exceptions related to source file and model midi file input.
        /// </summary>
        /// <param name="paths">A string array containing paths to files</param>
        /// <returns>bool representing success or failure.</returns>
        private bool CheckSourceFiles(List<string> paths, ListBox sPath)
        {
            string message = "";
            //Check if list is actually empty.
            if(paths.Count == 1)
            {
                if(paths[0].Trim() == "" || paths[0] == null)
                {
                    message = "No model or source files was provided for analysis.\n" +
                        "Please provide a model midi file, and optionally source files.";
                    MessageBoxResult result = MessageBox.Show(message, "No files given", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
            FileChecker fileChecker = new FileChecker();
            //Check for duplicate model file path in the source files, and delete it from the source path list.
            string modelPath = paths[paths.Count - 1];
            int index = paths.IndexOf(modelPath);
            if (index != (paths.Count - 1))
            {
                int delete2 = sPath.Items.Count;
                paths = RemoveElementAtIndex(paths, index);
                sPath.Items.Clear();
                for(int i = 0; i < paths.Count-1; i++)  //The limit is reduced by 1 to avoid the last element, the 
                {
                    sPath.Items.Add(paths[i]);
                }
                int delete = paths.Count;
                message = "Duplicate of model found in source files.\n" +
                        "One duplicate will be removed before processing.";
                MessageBoxResult result = MessageBox.Show(message, "Duplicate Found", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            List<string> nonExistentFiles = new List<string>();
            foreach(string path in paths){
                if (!fileChecker.FileExists(path))
                {
                    nonExistentFiles.Add(path);
                }
            }
            if(nonExistentFiles.Count > 0)
            {
                message = "The following files could not be found:\n";
                foreach(string path in nonExistentFiles)
                {
                    message = message + path + "\n";
                }
                MessageBoxResult result = MessageBox.Show(message, "Files not found", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        private bool CheckDestinationFolder(string path)
        {
            FileChecker fileChecker = new FileChecker();
            string message;
            if(path.Trim() == "" || path == null)
            {
                message = "No destination folder was provided.\n" +
                    "Please provide a destination folder.";
                MessageBoxResult result = MessageBox.Show(message, "No folder given", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            else
            {
                if (!fileChecker.FolderExists(path))
                {
                    message = "The destination folder could not be found.";
                    MessageBoxResult result = MessageBox.Show(message, "Folder not found", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
            return true;
        }

        private List<string> RemoveElementAtIndex(List<string> list, int index)
        {
            list.RemoveAt(index);
            List<string> newList = new List<string>();
            foreach(string item in list)
            {
                if(item != null && item.Trim() != "")
                {
                    newList.Add(item);
                }
            }
            return newList;
        }
    }
}
