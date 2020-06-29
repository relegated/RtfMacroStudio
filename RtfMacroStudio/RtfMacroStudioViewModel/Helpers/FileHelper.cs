using Microsoft.Win32;
using RtfMacroStudioViewModel.Interfaces;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace RtfMacroStudioViewModel.Helpers
{
    public class FileHelper : IFileHelper
    {
        public void LoadFile(RichTextBox richTextBox)
        {
            if (richTextBox == null)
            {
                throw new Exception($"Rich text box must be provided");
            }
            
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Rich Text files (*.rtf)|*.rtf|Text files (*.txt)|*.txt";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                var fileName = openFileDialog.FileName;
                if (File.Exists(fileName))
                {
                    var textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                    using (var fileStream = new FileStream(fileName, FileMode.OpenOrCreate))
                    {
                        textRange.Load(fileStream, DataFormats.Rtf);
                    }
                }
                else
                {
                    throw new FileNotFoundException($"{fileName} could not be found");
                }
            }
        }

        public void SaveFile(RichTextBox richTextBox)
        {
            if (richTextBox == null)
            {
                throw new Exception($"Rich text box must be provided");
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text files (*.rtf)|*.rtf|Text files (*.txt)|*.txt";
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (saveFileDialog.ShowDialog() == true)
            {
                var fileName = saveFileDialog.FileName;
                var format = GetFormat(saveFileDialog.Filter);
                if (File.Exists(fileName))
                {
                    if (MessageBox.Show($"{fileName} already exists. Overwrite?", "Rtf Macro Studio", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        ProcessSaveFile(richTextBox, fileName, format);
                    }
                }
                else
                {
                    ProcessSaveFile(richTextBox, fileName, format);
                }
            }
        }

        private string GetFormat(string filter)
        {
            if (filter.Contains("rtf"))
            {
                return DataFormats.Rtf;
            }
            else
            {
                return DataFormats.Text;
            }
        }

        private void ProcessSaveFile(RichTextBox richTextBox, string fileName, string format)
        {
            var textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
            using (var fileStream = new FileStream(fileName, FileMode.Create))
            {
                textRange.Save(fileStream, format);
            }
        }
    }
}
