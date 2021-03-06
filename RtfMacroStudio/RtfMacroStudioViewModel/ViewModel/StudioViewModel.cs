﻿using RtfMacroStudioViewModel.Interfaces;
using RtfMacroStudioViewModel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.ViewModel
{
    public class StudioViewModel
    {

        #region Properties

        public FlowDocument CurrentRichText { get; set; } = new FlowDocument();
        public List<MacroTask> CurrentTaskList { get; set; } = new List<MacroTask>();

        public List<ESpecialKey> SupportedSpecialKeys { get; set; } = new List<ESpecialKey>();

        public delegate void NotifyPropertyChanged(string PropertyName);

        public event NotifyPropertyChanged PropertyChanged;

        public Dictionary<string, Variable> RegisteredVariables { get; set; } = new Dictionary<string, Variable>();

        public RichTextBox RichTextBoxControl { get; set; }
        public IEditingCommandHelper EditingCommandHelper { get; }
        public IFileHelper FileHelper { get; }
        public IMacroTaskEditPresenter MacroTaskEditPresenter { get; }
        public IMacroRunPresenter MacroRunPresenter { get; }
        public IProgressPresenter ProgressPresenter { get; }
        public List<EFormatType> SupportedFormattingOptions { get; set; } = new List<EFormatType>();
        public List<string> AvailableFonts { get; set; } = new List<string>();

        private string selectedFont = "Segoe UI";
        public string SelectedFont
        {
            get => selectedFont;
            set
            {
                selectedFont = value;
                ApplyCurrentFormatting();

                if (IsCurrentlyRecording)
                {
                    AddFormatMacroTask(new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                        FormatType = EFormatType.Font,
                        TextFont = new FontFamily(selectedFont),
                    });
                }
            }
        }

        private bool currentBoldFlag = false;
        public bool CurrentBoldFlag
        {
            get => currentBoldFlag;
            set
            {
                currentBoldFlag = value;
                ApplyCurrentFormatting();

                if (IsCurrentlyRecording)
                {
                    AddFormatMacroTask(new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                        FormatType = EFormatType.Bold,
                    });
                }
            }
        }

        private bool currentItalicFlag = false;
        public bool CurrentItalicFlag
        {
            get => currentItalicFlag;
            set
            {
                currentItalicFlag = value;
                ApplyCurrentFormatting();

                if (IsCurrentlyRecording)
                {
                    AddFormatMacroTask(new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                        FormatType = EFormatType.Italic,
                    });
                }
            }
        }

        private TextAlignment currentTextAlignment = TextAlignment.Left;
        public TextAlignment CurrentTextAlignment
        {
            get => currentTextAlignment;
            set
            {
                currentTextAlignment = value;
                ApplyCurrentFormatting();
                
                if (IsCurrentlyRecording)
                {
                    var newAlignmentTask = new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                    };

                    switch (currentTextAlignment)
                    {
                        case TextAlignment.Left:
                            newAlignmentTask.FormatType = EFormatType.AlignLeft;
                            break;
                        case TextAlignment.Right:
                            newAlignmentTask.FormatType = EFormatType.AlignRight;
                            break;
                        case TextAlignment.Center:
                            newAlignmentTask.FormatType = EFormatType.AlignCenter;
                            break;
                        case TextAlignment.Justify:
                            newAlignmentTask.FormatType = EFormatType.AlignJustify;
                            break;
                        default:
                            break;
                    }

                    AddFormatMacroTask(newAlignmentTask);
                }
            }
        }

        private bool currentUnderlineFlag = false;
        public bool CurrentUnderlineFlag
        {
            get => currentUnderlineFlag;
            set
            {
                currentUnderlineFlag = value;
                ApplyCurrentFormatting();

                if (IsCurrentlyRecording)
                {
                    AddFormatMacroTask(new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                        FormatType = EFormatType.Underline,
                    });
                }
            }
        }

        private Color currentColor = Colors.Black;
        public Color CurrentColor
        {
            get => currentColor;
            set
            {
                currentColor = value;
                ApplyCurrentFormatting();

                if (IsCurrentlyRecording)
                {
                    AddFormatMacroTask(new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                        FormatType = EFormatType.Color,
                        TextColor = currentColor,
                    });
                }
            }
        }

        public DrawingImage ColorImageDrawing
        {
            get
            {
                RectangleGeometry rectangleGeometry = new RectangleGeometry(new Rect(0, 0, 25, 25));
                GeometryDrawing geometryDrawing = new GeometryDrawing(new SolidColorBrush(currentColor), new Pen(), rectangleGeometry);
                return new DrawingImage(geometryDrawing);
            }
        }

        public MacroTask MacroTaskInEdit { get; set; }

        private string currentCaptureString;

        private double currentTextSize = 12;
        public double CurrentTextSize
        {
            get => currentTextSize;
            set
            {
                var newValue = value;
                if (newValue < 4)
                    newValue = 4;
                if (newValue > 128)
                    newValue = 128;
                currentTextSize = newValue;
                ApplyCurrentFormatting();

                if (IsCurrentlyRecording)
                {
                    AddFormatMacroTask(new MacroTask()
                    {
                        MacroTaskType = EMacroTaskType.Format,
                        FormatType = EFormatType.TextSize,
                        TextSize = currentTextSize,
                    });
                }
            }
        }



        public List<double> AvailableTextSizes { get; set; } = new List<double>();
        public bool IsCurrentlyRecording { get; set; } = false;

        private bool isCapturingString = false;
        public bool IsCapturingString
        {
            get => isCapturingString;
            set
            {
                if (value == false && isCapturingString == true)
                {
                    if (!string.IsNullOrEmpty(currentCaptureString))
                    {
                        AddTextInputMacroTask(currentCaptureString);
                    }
                    currentCaptureString = string.Empty;
                }
                isCapturingString = value;
            }
        }

        public int NumberOfLinesInDocument
        {
            get
            {
                int lineCount = 0;

                foreach (Paragraph block in RichTextBoxControl.Document.Blocks)
                {
                    foreach (var inline in block.Inlines)
                    {
                        if (inline is LineBreak)
                        {
                            lineCount++;
                        }
                    }
                    lineCount++;
                }

                return lineCount;
            }
        }

        public int NumberOfLinesAboveCursor
        {
            get
            {
                int lineCount = 0;

                foreach (Paragraph block in RichTextBoxControl.Document.Blocks)
                {
                    if ((block.ContentStart.CompareTo(RichTextBoxControl.CaretPosition) == -1 
                        && block.ContentEnd.CompareTo(RichTextBoxControl.CaretPosition) == 1)
                        || block.ContentStart.CompareTo(RichTextBoxControl.CaretPosition) == 0)
                    {
                        return lineCount;
                    }
                    
                    foreach (var inline in block.Inlines)
                    {
                        if (inline is LineBreak)
                        {
                            lineCount++;
                        }
                    }
                    lineCount++;
                }

                return lineCount;
            }
        }

        private bool isRunningTask = false;
        private bool IsRunningTask 
        {
            get => isRunningTask;
            set
            {
                isRunningTask = value;
                if (isRunningTask)
                {
                    ProgressPresenter.ShowDialog();
                }
                else
                {
                    ProgressPresenter.Hide();
                }
            }
        }

        private int runNTimes = 1;
        public int RunNTimes 
        {
            get => runNTimes;
            private set
            {
                runNTimes = value;
                ProgressPresenter.RunNTimes = runNTimes;
            }
        }
        private int currentlyRunningXofN = 1;
        public int CurrentlyRunningXofN 
        {
            get => currentlyRunningXofN;
            private set
            {
                currentlyRunningXofN = value;
                ProgressPresenter.CurrentlyRunningXofN = currentlyRunningXofN;
            }
        }

        #endregion

        #region Constructor

        public StudioViewModel(IEditingCommandHelper editingCommandHelper, IFileHelper fileHelper, IMacroTaskEditPresenter macroTaskEditPresenter, IMacroRunPresenter macroRunPresenter, IProgressPresenter progressPresenter)
        {
            RichTextBoxControl = new RichTextBox();
            GetAvailableFonts();
            SelectedFont = "Segoe UI";
            GetAvailableTextSizes();
            CurrentTextSize = 12;
            SupportedSpecialKeys = Enum.GetValues(typeof(ESpecialKey)).Cast<ESpecialKey>().ToList();
            SetupSupportedFormatTypes();
            CurrentRichText.Blocks.Clear();
            RichTextBoxControl.Document.Blocks.Clear();
            EditingCommandHelper = editingCommandHelper;
            FileHelper = fileHelper;
            MacroTaskEditPresenter = macroTaskEditPresenter;
            MacroRunPresenter = macroRunPresenter;
            ProgressPresenter = progressPresenter;
        }

        public void EditMacroTaskBegin(MacroTask macroTask)
        {
            MacroTaskInEdit = macroTask;
            MacroTaskEditPresenter.ShowEditControl(this);
        }

        public void EditMacroTaskCancel()
        {
            MacroTaskInEdit = null;
        }

        public void EditMacroTaskComplete(MacroTask displayedTask, string text, object selectedItem)
        {           
            switch (displayedTask.MacroTaskType)
            {
                case EMacroTaskType.Text:
                    MacroTaskInEdit.Line = text;
                    break;
                case EMacroTaskType.SpecialKey:
                    ESpecialKey eSpecialKey = displayedTask.SpecialKey;
                    Enum.TryParse(selectedItem.ToString(), out eSpecialKey);
                    displayedTask.SpecialKey = eSpecialKey;
                    break;
                case EMacroTaskType.Format:
                    ProcessFormatEditMacroTask(displayedTask, selectedItem);
                    break;
                case EMacroTaskType.Variable:
                    ProcessVariableEditMacroTask(displayedTask, text, selectedItem);
                    break;
                default:
                    break;
            }

            UpdateTask(displayedTask);
            
        }

        /// <summary>
        /// Processes a Variable task
        /// </summary>
        /// <param name="displayedTask">The task object</param>
        /// <param name="text">Original Variable name</param>
        /// <param name="selectedItem">Tuple of the new name, value, increment value, whether to use place value, and the number of places</param>
        private void ProcessVariableEditMacroTask(MacroTask displayedTask, string text, object selectedItem)
        {
            var variableValues = selectedItem as Tuple<string, int, int, bool, int>;

            displayedTask.VarName = variableValues.Item1;
            displayedTask.VarValue = variableValues.Item2;
            displayedTask.VarIncrementValue = variableValues.Item3;
            displayedTask.VarUsePlaceValue = variableValues.Item4;
            displayedTask.VarPlaceValue = variableValues.Item5;

            RegisteredVariables.Remove(text);
            RegisteredVariables.Add(variableValues.Item1, new Variable()
            {
                Name = variableValues.Item1,
                Value = variableValues.Item2,
                IncrementByValue = variableValues.Item3,
                UsePlaceValues = variableValues.Item4,
                PlaceValuesToFill = variableValues.Item5,
            });

        }

        private static void ProcessFormatEditMacroTask(MacroTask displayedTask, object selectedItem)
        {
            EFormatType eFormatType = displayedTask.FormatType;
            if (eFormatType == EFormatType.Color)
            {
                if (selectedItem != null)
                {
                    displayedTask.TextColor = (Color)selectedItem;
                }
            }
            else if (eFormatType == EFormatType.Font)
            {
                displayedTask.TextFont = new FontFamily(selectedItem.ToString());
            }
            else if (eFormatType == EFormatType.TextSize)
            {
                double textSize = displayedTask.TextSize;
                double.TryParse(selectedItem.ToString(), out textSize);
                displayedTask.TextSize = textSize;
            }
            else
            {
                Enum.TryParse(selectedItem.ToString(), out eFormatType);
                displayedTask.FormatType = eFormatType;
            }    
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
        }

        #endregion

        public virtual void RunMacro(int numberOfTimes)
        {
            IsRunningTask = true;
            RunNTimes = numberOfTimes;
            for (int i = 0; i < numberOfTimes; i++)
            {
                CurrentlyRunningXofN = i;
                RunMacro();
            }
            IsRunningTask = false;
        }
        
        public virtual void RunMacroToEndOfText()
        {
            int numberOfLinesUntilEndOfDocument = NumberOfLinesInDocument - NumberOfLinesAboveCursor;

            RunMacro(numberOfLinesUntilEndOfDocument);
        }

        public virtual void RunMacro()
        {
            bool singleTask = IsRunningTask == false;
            
            if (singleTask)
            {
                IsRunningTask = true;
            }

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
                    case EMacroTaskType.Variable:
                        ProcessVariable(task.VarName);
                        break;
                    default:
                        break;
                }
            }

            if (singleTask)
            {
                IsRunningTask = false;
            }

            CurrentRichText = RichTextBoxControl.Document;
            PropertyChanged?.Invoke(nameof(CurrentRichText));
        }

        private void ProcessVariable(string varName)
        {
            if (RichTextBoxControl.Selection.IsEmpty == false)
            {
                ProcessSpecialKey(ESpecialKey.Delete);
            }

            var currentVariable = RegisteredVariables[varName];

            if (currentVariable.UsePlaceValues)
            {
                var formatString = $"D{currentVariable.PlaceValuesToFill}";
                RichTextBoxControl.CaretPosition.InsertTextInRun(currentVariable.Value.ToString(formatString));
            }
            else
            {
                RichTextBoxControl.CaretPosition.InsertTextInRun(currentVariable.Value.ToString());
            }
            
            UpdateCaretLocationBy(RegisteredVariables[varName].Value.ToString().Length);
            CurrentRichText = RichTextBoxControl.Document;

            IncrementVariable(varName);
        }

        private void IncrementVariable(string varName)
        {
            RegisteredVariables[varName].Value += RegisteredVariables[varName].IncrementByValue;
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

            switch (specialKey)
            {
                case ESpecialKey.LeftArrow:
                    EditingCommandHelper.MoveLeftByCharacter(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.RightArrow:
                    EditingCommandHelper.MoveRightByCharacter(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.UpArrow:
                    EditingCommandHelper.MoveUpByLine(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.DownArrow:
                    EditingCommandHelper.MoveDownByLine(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.Home:
                    EditingCommandHelper.MoveToLineStart(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.End:
                    EditingCommandHelper.MoveToLineEnd(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ShiftHome:
                    EditingCommandHelper.SelectToLineStart(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ShiftEnd:
                    EditingCommandHelper.SelectToLineEnd(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlHome:
                    EditingCommandHelper.MoveToDocumentStart(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlEnd:
                    EditingCommandHelper.MoveToDocumentEnd(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlShiftHome:
                    EditingCommandHelper.SelectToDocumentStart(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlShiftEnd:
                    EditingCommandHelper.SelectToDocumentEnd(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.Delete:
                    EditingCommandHelper.Delete(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.Backspace:
                    EditingCommandHelper.Backspace(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.Enter:
                    int numOfParagraphs = RichTextBoxControl.Document.Blocks.Count;
                    EditingCommandHelper.EnterParagraphBreak(RichTextBoxControl);

                    //sometimes, the first EnterParagraphBreak doesn't add a new paragraph
                    if (numOfParagraphs == RichTextBoxControl.Document.Blocks.Count)
                    {
                        Paragraph paragraph = new Paragraph();
                        Run run = new Run(string.Empty);
                        paragraph.Inlines.Add(run);
                        RichTextBoxControl.Document.Blocks.Add(paragraph);
                        RichTextBoxControl.CaretPosition = RichTextBoxControl.Document.ContentEnd;
                    }

                    
                    break;
                case ESpecialKey.ShiftLeftArrow:
                    EditingCommandHelper.SelectLeftByCharacter(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ShiftRightArrow:
                    EditingCommandHelper.SelectRightByCharacter(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ShiftUpArrow:
                    EditingCommandHelper.SelectUpByLine(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ShiftDownArrow:
                    EditingCommandHelper.SelectDownByLine(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlLeftArrow:
                    EditingCommandHelper.MoveLeftByWord(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlRightArrow:
                    EditingCommandHelper.MoveRightByWord(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlShiftLeftArrow:
                    EditingCommandHelper.SelectLeftByWord(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.ControlShiftRightArrow:
                    EditingCommandHelper.SelectRightByWord(RichTextBoxControl);
                    
                    break;
                case ESpecialKey.Cut:
                    CutToClipboard();
                    break;
                case ESpecialKey.Copy:
                    CopyToClipboard();
                    break;
                case ESpecialKey.Paste:
                    PasteFromClipboard();
                    break;
                default:
                    break;
            }

            CurrentRichText = RichTextBoxControl.Document;
        }


        private void ProcessText(string line)
        {
            if (RichTextBoxControl.Selection.IsEmpty == false)
            {
                ProcessSpecialKey(ESpecialKey.Delete);
            }
            
            RichTextBoxControl.CaretPosition.InsertTextInRun(line);
            UpdateCaretLocationBy(line.Length);
            CurrentRichText = RichTextBoxControl.Document;
        }

        private void UpdateCaretLocationBy(int length)
        {
            TextPointer startLocation = RichTextBoxControl.Document.ContentStart;
            int startOffset = startLocation.GetOffsetToPosition(RichTextBoxControl.CaretPosition);
            TextPointer newLocation = startLocation.GetPositionAtOffset(startOffset + length, LogicalDirection.Forward);
            if (newLocation == null)
            {
                RichTextBoxControl.CaretPosition = RichTextBoxControl.Document.ContentEnd;
            }
            else
            {
                RichTextBoxControl.CaretPosition = newLocation;
            }
            
           
        }

        public void AddTextInputMacroTask(string textInput)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                Index = CurrentTaskList.Count,
                Line = textInput,
                MacroTaskType = EMacroTaskType.Text
            });
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public void AddSpecialKeyMacroTask(ESpecialKey specialKey)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                Index = CurrentTaskList.Count,
                MacroTaskType = EMacroTaskType.SpecialKey,
                SpecialKey = specialKey,
            });
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }
        public void AddFormatMacroTask(string formatTask)
        {
            EFormatType formatType;
            if (Enum.TryParse<EFormatType>(formatTask, out formatType))
            {
                CurrentTaskList.Add(new MacroTask()
                {
                    Index = CurrentTaskList.Count,
                    MacroTaskType = EMacroTaskType.Format,
                    FormatType = formatType,
                    TextColor = Colors.Black,
                    TextFont = new FontFamily("Segoe UI"),
                    TextSize = 12,
                });
                RaisePropertyChangedEvent(nameof(CurrentTaskList));
            }
           
        }
        public void AddFormatMacroTask(MacroTask macroTask)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                Index = CurrentTaskList.Count,
                MacroTaskType = macroTask.MacroTaskType,
                FormatType = macroTask.FormatType,
                TextColor = macroTask.TextColor,
                TextSize = macroTask.TextSize,
                TextFont = macroTask.TextFont,
            });
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public void AddVariableMacroTask(string name, int value, int incrementByValue, bool usePlaceValue, int placeValue)
        {
            var newVariable = new Variable()
            {
                Name = name,
                Value = value,
                IncrementByValue = incrementByValue,
                UsePlaceValues = usePlaceValue,
                PlaceValuesToFill = placeValue,
            };

            RegisteredVariables.Add(name, newVariable);
            CurrentTaskList.Add(new MacroTask()
            {
                Index = CurrentTaskList.Count,
                MacroTaskType = EMacroTaskType.Variable,
                VarName = name,
                VarValue = value,
                VarIncrementValue = incrementByValue,
                VarUsePlaceValue = usePlaceValue,
                VarPlaceValue = placeValue,
            });
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public bool IsVariableNameInUse(string variableName)
        {
            return RegisteredVariables.ContainsKey(variableName);
        }

        public void RemoveTaskAt(int taskIndex)
        {
            try
            {
                if (CurrentTaskList[taskIndex].MacroTaskType == EMacroTaskType.Variable)
                {
                    RegisteredVariables.Remove(CurrentTaskList[taskIndex].VarName);
                }
                
                CurrentTaskList.RemoveAt(taskIndex);
                ReIndexTasks();

                RaisePropertyChangedEvent(nameof(CurrentTaskList));
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine($"Cannot remove task with index of {taskIndex} - {ex.Message}");
            }
        }

        private void ReIndexTasks()
        {
            for (int i = 0; i < CurrentTaskList.Count; i++)
            {
                CurrentTaskList[i].Index = i;
            }
        }

        private void UpdateTask(MacroTask displayedTask)
        {
            
            CurrentTaskList[displayedTask.Index] = displayedTask;
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public void ClearAllTasks()
        {
            CurrentTaskList.Clear();
            RegisteredVariables.Clear();
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
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
            ReIndexTasks();
            RaisePropertyChangedEvent(nameof(CurrentTaskList));
        }

        public void ProcessDroppedControl(object dropTask, object dragTask)
        {
            MoveTask(((MacroTask)dragTask).Index, ((MacroTask)dropTask).Index);
        }

        private void RaisePropertyChangedEvent(string propertyName)
        {
            PropertyChanged?.Invoke(propertyName);
        }

        public static string GetTextFromMacroTask(MacroTask macroTask)
        {
            return macroTask.Line;
        }

        public void RefreshCurrentFormatting()
        {
            //Select one character if nothing is selected
            if (RichTextBoxControl.Selection.IsEmpty)
            {
                EditingCommandHelper.SelectRightByCharacter(RichTextBoxControl);
            }
            
            var selectionRange = new TextRange(RichTextBoxControl.Selection.Start, RichTextBoxControl.Selection.End);

            //Read formatting
            currentBoldFlag = (FontWeight)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontWeightProperty) == FontWeights.Bold;
            currentItalicFlag = (FontStyle)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontStyleProperty) == FontStyles.Italic;
            var selectedProperty = (TextDecorationCollection)selectionRange.GetPropertyValue(TextBlock.TextDecorationsProperty);
            currentUnderlineFlag = selectedProperty.Where(x => ((TextDecoration)x).Location == TextDecorationLocation.Underline).ToList().Count > 0;
            selectedFont = ((FontFamily)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontFamilyProperty)).Source;
            var currentBrush = (SolidColorBrush)RichTextBoxControl.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            currentColor = GetColorFromBrush(currentBrush);
            currentTextSize = (double)RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontSizeProperty);

            var currentBlock = RichTextBoxControl.Document.Blocks
                .Where(x => x.ContentStart.CompareTo(RichTextBoxControl.CaretPosition) == -1 &&
                x.ContentEnd.CompareTo(RichTextBoxControl.CaretPosition) == 1)
                .FirstOrDefault();
            Paragraph currentParagraph = currentBlock as Paragraph;
            if (currentParagraph == null)
            {
                currentTextAlignment = TextAlignment.Left;
            }
            else
            {
                currentTextAlignment = currentParagraph.TextAlignment;
            }

            //Undo select
            EditingCommandHelper.MoveLeftByCharacter(RichTextBoxControl);
        }

        public void RecordMacroStart()
        {
            ClearAllTasks();
            IsCurrentlyRecording = true;
        }

        public void ProcessKey(Key keyInput, ModifierKeys[] modifierKeys)
        {
            if (IsCurrentlyRecording)
            {
                if (IsMovementKey(keyInput))
                {
                    IsCapturingString = false;
                    CreateMovementKeyTask(keyInput, modifierKeys);
                }
                else if (IsSpecialKey(keyInput))
                {
                    IsCapturingString = false;
                    CreateSpecialKeyTask(keyInput, modifierKeys);
                }
                else if (IsControlPressed(modifierKeys))
                {
                    IsCapturingString = false;
                    CreateFormatKeyOrClipboardTask(keyInput);
                }
                else
                {
                    IsCapturingString = true;
                    currentCaptureString += TranslateKeyToStringForCapture(keyInput, modifierKeys);
                }
                
                
            }
        }

        private void CreateSpecialKeyTask(Key keyInput, ModifierKeys[] modifierKeys)
        {
            if (keyInput == Key.Enter)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Enter);
            }
            else if (keyInput == Key.Back)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Backspace);
            }
            else if (keyInput == Key.Delete)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Delete);
            }
        }

        private bool IsSpecialKey(Key keyInput)
        {
            return keyInput == Key.Enter ||
                keyInput == Key.Delete ||
                keyInput == Key.Back;
        }

        private string TranslateKeyToStringForCapture(Key keyInput, ModifierKeys[] modifierKeys)
        {
            //letters can be a straight conversion to initial string
            KeyConverter kc = new KeyConverter();
            var keyString = kc.ConvertToString(keyInput);

            if (keyString.Length == 1)
            {
                var charVersion = keyString[0];

                if (char.IsLetter(charVersion))
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return keyString.ToUpper();
                    }

                    return keyString.ToLower();
                }
                else if (char.IsDigit(charVersion))
                {

                    //is it punctuation above the number keys?
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        if (keyInput == Key.D1)
                        {
                            return "!";
                        }
                        else if (keyInput == Key.D2)
                        {
                            return "@";
                        }
                        else if (keyInput == Key.D3)
                        {
                            return "#";
                        }
                        else if (keyInput == Key.D4)
                        {
                            return "$";
                        }
                        else if (keyInput == Key.D5)
                        {
                            return "%";
                        }
                        else if (keyInput == Key.D6)
                        {
                            return "^";
                        }
                        else if (keyInput == Key.D7)
                        {
                            return "&";
                        }
                        else if (keyInput == Key.D8)
                        {
                            return "*";
                        }
                        else if (keyInput == Key.D9)
                        {
                            return "(";
                        }
                        else if (keyInput == Key.D0)
                        {
                            return ")";
                        }
                    }

                    //it's a regular digit
                    return keyString;

                }

                if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                {
                    return keyString.ToUpper();
                }

                return keyString.ToLower();
            }
            else
            {
                //other special text input that doesn't break the string
                if (keyInput == Key.Space)
                {
                    return " ";
                }
                else if (keyInput == Key.Tab)
                {
                    return "\t";
                }
                else if (keyInput == Key.Oem3)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "~";
                    }
                    return "`";
                }
                else if (keyInput == Key.OemMinus)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "_";
                    }
                    return "-";
                }
                else if (keyInput == Key.OemPlus)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "+";
                    }
                    return "=";
                }
                else if (keyInput == Key.Oem4)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "{";
                    }
                    return "[";
                }
                else if (keyInput == Key.Oem6)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "}";
                    }
                    return "]";
                }
                else if (keyInput == Key.Oem5)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "|";
                    }
                    return "\\";
                }
                else if (keyInput == Key.Oem1)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return ":";
                    }
                    return ";";
                }
                else if (keyInput == Key.Oem7)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "\"";
                    }
                    return "'";
                }
                else if (keyInput == Key.OemComma)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "<";
                    }
                    return ",";
                }
                else if (keyInput == Key.OemPeriod)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return ">";
                    }
                    return ".";
                }
                else if (keyInput == Key.Oem2)
                {
                    if (modifierKeys != null && modifierKeys.Contains(ModifierKeys.Shift))
                    {
                        return "?";
                    }
                    return "/";
                }
            }

            return string.Empty;
        }

        public void RunMacroPresenter()
        {
            if (MacroRunPresenter.ShowRunOptionWindow())
            {
                switch (MacroRunPresenter.Option)
                {
                    case ERunPresenterOption.Once:
                        RunMacro();
                        break;
                    case ERunPresenterOption.NTimes:
                        RunMacro(MacroRunPresenter.N);
                        break;
                    case ERunPresenterOption.End:
                        RunMacroToEndOfText();
                        break;
                    default:
                        break;
                }
            }
        }

        private void CreateFormatKeyOrClipboardTask(Key keyInput)
        {
            if (keyInput == Key.B)
            {
                AddFormatMacroTask("Bold");
            }
            if (keyInput == Key.I)
            {
                AddFormatMacroTask("Italic");
            }
            if (keyInput == Key.U)
            {
                AddFormatMacroTask("Underline");
            }
            if (keyInput == Key.E)
            {
                AddFormatMacroTask("AlignCenter");
            }
            if (keyInput == Key.J)
            {
                AddFormatMacroTask("AlignJustify");
            }
            if (keyInput == Key.L)
            {
                AddFormatMacroTask("AlignLeft");
            }
            if (keyInput == Key.R)
            {
                AddFormatMacroTask("AlignRight");
            }

            //clipboard are not format but special keys tasks
            if (keyInput == Key.C)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Copy);
            }
            if (keyInput == Key.X)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Cut);
            }
            if (keyInput == Key.V)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Paste);
            }
        }

        private bool IsControlPressed(ModifierKeys[] modifierKeys)
        {
            if (modifierKeys == null)
            {
                return false;
            }

            return modifierKeys.Contains(ModifierKeys.Control);
        }

        private void CreateMovementKeyTask(Key keyInput, ModifierKeys[] modifierKeys)
        {
            if (modifierKeys == null) // NO MODIFIER KEYS
            {
                //process the basic key
                if (keyInput == Key.Left)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.LeftArrow);
                }
                else if (keyInput == Key.Right)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.RightArrow);
                }
                else if (keyInput == Key.Up)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.UpArrow);
                }
                else if (keyInput == Key.Down)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.DownArrow);
                }
                else if (keyInput == Key.Home)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.Home);
                }
                else if (keyInput == Key.End)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.End);
                }
            }
            else if (modifierKeys.Length == 1) // EITHER CONTROL OR ALT
            {
                switch (modifierKeys[0])
                {
                    case ModifierKeys.Control:
                        if (keyInput == Key.Left)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ControlLeftArrow);
                        }
                        else if (keyInput == Key.Right)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ControlRightArrow);
                        }
                        else if (keyInput == Key.Home)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ControlHome);
                        }
                        else if (keyInput == Key.End)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ControlEnd);
                        }
                        else if (keyInput == Key.Up)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.UpArrow);
                        }
                        else if (keyInput == Key.Down)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.DownArrow);
                        }
                        break;
                    case ModifierKeys.Shift:
                        if (keyInput == Key.Left)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ShiftLeftArrow);
                        }
                        else if (keyInput == Key.Right)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ShiftRightArrow);
                        }
                        else if (keyInput == Key.Home)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ShiftHome);
                        }
                        else if (keyInput == Key.End)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.ShiftEnd);
                        }
                        else if (keyInput == Key.Up)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.UpArrow);
                        }
                        else if (keyInput == Key.Down)
                        {
                            AddSpecialKeyMacroTask(ESpecialKey.DownArrow);
                        }
                        break;
                    default:
                        break;
                }
            }
            else // BOTH CONTROL AND SHIFT
            {
                if (keyInput == Key.Left)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.ControlShiftLeftArrow);
                }
                else if (keyInput == Key.Right)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.ControlShiftRightArrow);
                }
                else if (keyInput == Key.Home)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.ControlShiftHome);
                }
                else if (keyInput == Key.End)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.ControlShiftEnd);
                }
                else if (keyInput == Key.Up)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.UpArrow);
                }
                else if (keyInput == Key.Down)
                {
                    AddSpecialKeyMacroTask(ESpecialKey.DownArrow);
                }
            }
        }

        public void StopRecording()
        {
            IsCapturingString = false;
            IsCurrentlyRecording = false;
        }

        private bool IsMovementKey(Key keyInput)
        {
            return (keyInput == Key.Left ||
                keyInput == Key.Right ||
                keyInput == Key.Up ||
                keyInput == Key.Down ||
                keyInput == Key.Home ||
                keyInput == Key.End);
        }

        public void CopyToClipboard()
        {
            RichTextBoxControl.Copy();

            if (IsCurrentlyRecording)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Copy);
            }
        }

        public void CutToClipboard()
        {
            RichTextBoxControl.Cut();

            if (IsCurrentlyRecording)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Cut);
            }
        }

        public void PasteFromClipboard()
        {
            RichTextBoxControl.Paste();

            if (IsCurrentlyRecording)
            {
                AddSpecialKeyMacroTask(ESpecialKey.Paste);
            }
        }

        public void ApplyCurrentFormatting()
        {
            if (RichTextBoxControl.Selection.IsEmpty)
            {
                if (RichTextBoxControl.Selection.Start.Paragraph == null)
                {
                    ApplySettingsToNewParagraph();
                }
                else
                {
                    Block currentBlock = RichTextBoxControl.Document.Blocks
                        .Where(x => x.ContentStart.CompareTo(RichTextBoxControl.CaretPosition) == -1 && x.ContentEnd.CompareTo(RichTextBoxControl.CaretPosition) == 1)
                        .FirstOrDefault();
                    if (currentBlock != null)
                    {
                        Paragraph currentParagraph = currentBlock as Paragraph;
                        ApplySettingsToExistingParagraph(currentParagraph);
                    }
                    else if (RichTextBoxControl.Document.Blocks.FirstBlock.ContentStart.CompareTo(RichTextBoxControl.Document.Blocks.LastBlock.ContentStart) == 0)
                    {
                        Paragraph currentParagraph = RichTextBoxControl.Document.Blocks.FirstBlock as Paragraph;
                        ApplySettingsToExistingParagraph(currentParagraph);
                    }
                }
            }

            RichTextBoxControl.Focus();
        }

        private void ApplySettingsToNewParagraph()
        {
            Paragraph newParagraph = new Paragraph();
            newParagraph.FontFamily = new FontFamily(SelectedFont);
            newParagraph.FontSize = currentTextSize;
            newParagraph.FontStyle = (CurrentItalicFlag ? FontStyles.Italic : FontStyles.Normal);
            newParagraph.FontWeight = (CurrentBoldFlag ? FontWeights.Bold : FontWeights.Normal);
            newParagraph.TextDecorations = (CurrentUnderlineFlag ? TextDecorations.Underline : null);
            newParagraph.Foreground = new SolidColorBrush(CurrentColor);
            newParagraph.TextAlignment = CurrentTextAlignment;
            RichTextBoxControl.Document.Blocks.Add(newParagraph);
        }

        private void ApplySettingsToExistingParagraph(Paragraph currentParagraph)
        {
            Run newRun = new Run();
            newRun.FontFamily = new FontFamily(SelectedFont);
            newRun.FontSize = currentTextSize;
            newRun.FontStyle = (CurrentItalicFlag ? FontStyles.Italic : FontStyles.Normal);
            newRun.FontWeight = (CurrentBoldFlag ? FontWeights.Bold : FontWeights.Normal);
            newRun.TextDecorations = (CurrentUnderlineFlag ? TextDecorations.Underline : null);
            newRun.Foreground = new SolidColorBrush(CurrentColor);
            currentParagraph.TextAlignment = CurrentTextAlignment;
            currentParagraph.Inlines.Add(newRun);
            RichTextBoxControl.CaretPosition = newRun.ElementStart;
            
        }


        public void OpenFile()
        {
            FileHelper.LoadFile(RichTextBoxControl);
        }

        public void SaveFile()
        {
            FileHelper.SaveFile(RichTextBoxControl);
        }
     
        
        public void CloseProgram()
        {
            ProgressPresenter.Dispose();
        }
    }
}
