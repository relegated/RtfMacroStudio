using RtfMacroStudioViewModel.Interfaces;
using RtfMacroStudioViewModel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.ViewModel
{
    public class StudioViewModel
    {

        #region Properties

        public FlowDocument CurrentRichText { get; set; } = new FlowDocument();
        public List<MacroTask> CurrentTaskList { get; set; } = new List<MacroTask>();

        public List<ESpecialKey> SupportedSpecialKeys { get; set; } = new List<ESpecialKey>();
        public TextPointer CaretPosition { get; set; }

        public delegate void NotifyPropertyChanged(string PropertyName);

        public event NotifyPropertyChanged PropertyChanged;

        public RichTextBox RichTextBoxControl { get; set; } = new RichTextBox();
        public IEditingCommandHelper EditingCommandHelper { get; }
        public List<EFormatType> SupportedFormattingOptions { get; set; } = new List<EFormatType>();
        public List<string> AvailableFonts { get; set; } = new List<string>();
        public string SelectedFont { get; set; } = string.Empty;
        public bool CurrentBoldFlag { get; set; } = false;
        public bool CurrentItalicFlag { get; set; } = false;
        public bool CurrentUnderlineFlag { get; set; } = false;
        public Color CurrentColor { get; set; } = Colors.Black;

        private double currentTextSize;
        public double CurrentTextSize 
        {
            get
            { return currentTextSize; }
            set
            {
                var newValue = value;
                if (newValue < 4)
                    newValue = 4;
                if (newValue > 128)
                    newValue = 128;
                currentTextSize = newValue;
            }
        }
        public List<double> AvailableTextSizes { get; set; } = new List<double>();



        #endregion

        #region Constructor

        public StudioViewModel(IEditingCommandHelper editingCommandHelper)
        {
            GetAvailableFonts();
            SelectedFont = "Segoe UI";
            GetAvailableTextSizes();
            CurrentTextSize = 12;
            SupportedSpecialKeys = Enum.GetValues(typeof(ESpecialKey)).Cast<ESpecialKey>().ToList();
            SetupSupportedFormatTypes();
            CaretPosition = CurrentRichText.ContentStart;
            EditingCommandHelper = editingCommandHelper;
        }

        private void GetAvailableFonts()
        {
            foreach (var fontFamily in Fonts.SystemFontFamilies)
            {
                AvailableFonts.Add(fontFamily.Source);
            }
        }

        private void GetAvailableTextSizes()
        {
            AvailableTextSizes.Add(10);
            AvailableTextSizes.Add(12);
            AvailableTextSizes.Add(14);
            AvailableTextSizes.Add(16);
            AvailableTextSizes.Add(18);
            AvailableTextSizes.Add(20);
            AvailableTextSizes.Add(24);
            AvailableTextSizes.Add(28);
            AvailableTextSizes.Add(32);
            AvailableTextSizes.Add(64);
        }

        private void SetupSupportedFormatTypes()
        {
            SupportedFormattingOptions = Enum.GetValues(typeof(EFormatType)).Cast<EFormatType>().ToList();
            SupportedFormattingOptions.Remove(EFormatType.Color);
            SupportedFormattingOptions.Remove(EFormatType.Font);
            SupportedFormattingOptions.Remove(EFormatType.TextSize);
        }

        #endregion

        public void RunMacro()
        {
            foreach (var task in CurrentTaskList)
            {
                switch (task.MacroTaskType)
                {
                    case EMacroTaskType.Text:
                        ProcessText(task.Line);
                        break;
                    case EMacroTaskType.SpecialKey:
                        ProcessSpecialKey(task.SpecialKey);
                        break;
                    case EMacroTaskType.Format:
                        ProcessFormat(task);
                        break;
                    default:
                        break;
                }
            }
        }

        private void ProcessFormat(MacroTask task)
        {
            switch (task.FormatType)
            {
                case EFormatType.Bold:
                    EditingCommandHelper.ToggleBold(RichTextBoxControl);
                    break;
                case EFormatType.Italic:
                    EditingCommandHelper.ToggleItalic(RichTextBoxControl);
                    break;
                case EFormatType.Underline:
                    EditingCommandHelper.ToggleUnderline(RichTextBoxControl);
                    break;
                case EFormatType.Font:
                    SetFont(RichTextBoxControl, task);
                    break;
                case EFormatType.Color:
                    SetColor(RichTextBoxControl, task);
                    break;
                case EFormatType.TextSize:
                    SetSize(RichTextBoxControl, task);
                    break;
                case EFormatType.AlignCenter:
                    EditingCommandHelper.AlignCenter(RichTextBoxControl);
                    break;
                case EFormatType.AlignJustify:
                    EditingCommandHelper.AlignJustify(RichTextBoxControl);
                    break;
                case EFormatType.AlignLeft:
                    EditingCommandHelper.AlignLeft(RichTextBoxControl);
                    break;
                case EFormatType.AlignRight:
                    EditingCommandHelper.AlignRight(RichTextBoxControl);
                    break;
                default:
                    break;
            }
        }

        private void SetSize(RichTextBox richTextBoxControl, MacroTask task)
        {
            var selection = richTextBoxControl.Selection;
            selection.ApplyPropertyValue(TextElement.FontSizeProperty, task.TextSize);
        }

        private void SetFont(RichTextBox richTextBoxControl, MacroTask task)
        {
            var selection = richTextBoxControl.Selection;
            selection.ApplyPropertyValue(TextElement.FontFamilyProperty, task.TextFont);
        }

        private void SetColor(RichTextBox richTextBoxControl, MacroTask task)
        {
            var selection = richTextBoxControl.Selection;
            selection.ApplyPropertyValue(TextElement.ForegroundProperty, GetBrushFromColor(task.TextColor));
        }

        private SolidColorBrush GetBrushFromColor(Color textColor)
        {
            return new SolidColorBrush(textColor);
        }

        private Color GetColorFromBrush(SolidColorBrush brush)
        {
            return brush.Color;
        }

        private void ProcessSpecialKey(ESpecialKey specialKey)
        {
            RichTextBoxControl.Document = CurrentRichText;
            RichTextBoxControl.CaretPosition = CaretPosition;

            switch (specialKey)
            {
                case ESpecialKey.LeftArrow:
                    EditingCommandHelper.MoveLeftByCharacter(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.RightArrow:
                    EditingCommandHelper.MoveRightByCharacter(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.UpArrow:
                    EditingCommandHelper.MoveUpByLine(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.DownArrow:
                    EditingCommandHelper.MoveDownByLine(RichTextBoxControl);
                    EditingCommands.MoveDownByLine.Execute(null, RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.Home:
                    EditingCommandHelper.MoveToLineStart(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.End:
                    EditingCommandHelper.MoveToLineEnd(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ShiftHome:
                    EditingCommandHelper.SelectToLineStart(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ShiftEnd:
                    EditingCommandHelper.SelectToLineEnd(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlHome:
                    EditingCommandHelper.MoveToDocumentStart(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlEnd:
                    EditingCommandHelper.MoveToDocumentEnd(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlShiftHome:
                    EditingCommandHelper.SelectToDocumentStart(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlShiftEnd:
                    EditingCommandHelper.SelectToDocumentEnd(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.Delete:
                    EditingCommandHelper.Delete(RichTextBoxControl);
                    break;
                case ESpecialKey.Backspace:
                    EditingCommandHelper.Backspace(RichTextBoxControl);
                    break;
                case ESpecialKey.Enter:
                    EditingCommandHelper.EnterParagraphBreak(RichTextBoxControl);
                    break;
                case ESpecialKey.ShiftLeftArrow:
                    EditingCommandHelper.SelectLeftByCharacter(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ShiftRightArrow:
                    EditingCommandHelper.SelectRightByCharacter(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ShiftUpArrow:
                    EditingCommandHelper.SelectUpByLine(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ShiftDownArrow:
                    EditingCommandHelper.SelectDownByLine(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlLeftArrow:
                    EditingCommandHelper.MoveLeftByWord(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlRightArrow:
                    EditingCommandHelper.MoveRightByWord(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlShiftLeftArrow:
                    EditingCommandHelper.SelectLeftByWord(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                case ESpecialKey.ControlShiftRightArrow:
                    EditingCommandHelper.SelectRightByWord(RichTextBoxControl);
                    CaretPosition = RichTextBoxControl.CaretPosition;
                    break;
                default:
                    break;
            }
        }


        private void ProcessText(Paragraph line)
        {
            CurrentRichText.Blocks.Add(line);
            CaretPosition = CurrentRichText.ContentEnd;
        }

        public void AddTextInputMacroTask(string textInput)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run(textInput));
            CurrentTaskList.Add(new MacroTask()
            {
                Line = paragraph,
                MacroTaskType = EMacroTaskType.Text
            });
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public void AddSpecialKeyMacroTask(ESpecialKey specialKey)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.SpecialKey,
                SpecialKey = specialKey,
            });
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }
        public void AddFormatMacroTask(MacroTask task)
        {
            CurrentTaskList.Add(task);
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public void RemoveTaskAt(int taskIndex)
        {
            try
            {
                CurrentTaskList.RemoveAt(taskIndex);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine($"Cannot remove task with index of {taskIndex} - {ex.Message}");
            }
        }

        public void ClearAllTasks()
        {
            CurrentTaskList.Clear();
        }

        public void MoveTask(int sourceIndex, int destinationIndex)
        {
            if (sourceIndex < 0 || sourceIndex >= CurrentTaskList.Count 
                || destinationIndex < 0 || destinationIndex >= CurrentTaskList.Count)
            {
                return;
            }

            var taskToMove = CurrentTaskList[sourceIndex];
            CurrentTaskList.RemoveAt(sourceIndex);
            CurrentTaskList.Insert(destinationIndex, taskToMove);
        }

        private void RaisePropertyChangedEvent(string propertyName)
        {
            PropertyChanged?.Invoke(propertyName);
        }

        public static string GetTextFromMacroTask(MacroTask macroTask)
        {
            string returnValue = string.Empty;

            foreach (var item in macroTask.Line.Inlines)
            {
                if (item is Run)
                {
                    returnValue += ((Run)item).Text;
                }
            }

            return returnValue;
        }

        public void RefreshCurrentFormatting()
        {
            //Select one character
            EditingCommandHelper.SelectRightByCharacter(RichTextBoxControl);
            var selectionRange = new TextRange(RichTextBoxControl.Selection.Start, RichTextBoxControl.Selection.End);

            //Read formatting
            CurrentBoldFlag = (FontWeight)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontWeightProperty) == FontWeights.Bold;
            CurrentItalicFlag = (FontStyle)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontStyleProperty) == FontStyles.Italic;
            var selectedProperty = (TextDecorationCollection)selectionRange.GetPropertyValue(TextBlock.TextDecorationsProperty);
            CurrentUnderlineFlag = selectedProperty.Where(x => ((TextDecoration)x).Location == TextDecorationLocation.Underline).ToList().Count > 0;
            SelectedFont = ((FontFamily)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontFamilyProperty)).Source;
            var currentBrush = (SolidColorBrush)RichTextBoxControl.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            CurrentColor = GetColorFromBrush(currentBrush);
            CurrentTextSize = (double)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontSizeProperty);

            //Undo select
            EditingCommandHelper.MoveLeftByCharacter(RichTextBoxControl);
        }
    }
}
