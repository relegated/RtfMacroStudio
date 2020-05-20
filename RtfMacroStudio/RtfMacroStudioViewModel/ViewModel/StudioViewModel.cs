using RtfMacroStudioViewModel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Threading;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.ViewModel
{
    public class StudioViewModel
    {
        #region Fields
        
        //RichTextBox is a control, but it provides logical methods to move and select in a document
        private RichTextBox richTextBox;
        Dispatcher dispatcher;
        
        #endregion

        #region Properties

        public FlowDocument CurrentRichText { get; set; } = new FlowDocument();
        public List<MacroTask> CurrentTaskList { get; set; } = new List<MacroTask>();

        public List<ESpecialKey> SupportedSpecialKeys { get; set; } = new List<ESpecialKey>();
        public TextPointer CaretPosition { get; set; }

        public delegate void NotifyPropertyChanged(string PropertyName);

        public event NotifyPropertyChanged PropertyChanged;

        

        

        #endregion

        #region Constructor

        public StudioViewModel()
        {
            SupportedSpecialKeys = Enum.GetValues(typeof(ESpecialKey)).Cast<ESpecialKey>().ToList();
            CaretPosition = CurrentRichText.ContentStart;
            dispatcher = Dispatcher.CurrentDispatcher;
            dispatcher.Invoke(() => richTextBox = new RichTextBox());
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
                        break;
                    default:
                        break;
                }

                
            }
        }

        private void ProcessSpecialKey(ESpecialKey specialKey)
        {
            richTextBox.Document = CurrentRichText;
            richTextBox.CaretPosition = CaretPosition;
            
            switch (specialKey)
            {
                case ESpecialKey.LeftArrow:
                    CaretPosition = CaretPosition.GetPositionAtOffset(-1);
                    break;
                case ESpecialKey.RightArrow:
                    CaretPosition = CaretPosition.GetPositionAtOffset(1);
                    break;
                case ESpecialKey.UpArrow:
                    NavigateByLineIndexOffset(-1);
                    break;
                case ESpecialKey.DownArrow:
                    NavigateByLineIndexOffset(1);
                    break;
                case ESpecialKey.Home:
                    CaretPosition = CurrentRichText.Blocks.ElementAt(GetCurrentBlockIndex()).ContentStart;
                    break;
                case ESpecialKey.End:
                    CaretPosition = CurrentRichText.Blocks.ElementAt(GetCurrentBlockIndex()).ContentEnd;
                    break;
                case ESpecialKey.ShiftHome:
                    CaretPosition = CurrentRichText.Blocks.ElementAt(GetCurrentBlockIndex()).ContentStart;
                    break;
                case ESpecialKey.ShiftEnd:
                    CaretPosition = CurrentRichText.Blocks.ElementAt(GetCurrentBlockIndex()).ContentEnd;
                    break;
                case ESpecialKey.Delete:
                    break;
                case ESpecialKey.Backspace:
                    break;
                case ESpecialKey.Enter:
                    break;
                case ESpecialKey.ShiftLeftArrow:
                    CaretPosition = CaretPosition.GetPositionAtOffset(-1);
                    break;
                case ESpecialKey.ShiftRightArrow:
                    CaretPosition = CaretPosition.GetPositionAtOffset(1);
                    break;
                case ESpecialKey.ShiftUpArrow:
                    NavigateByLineIndexOffset(-1);
                    break;
                case ESpecialKey.ShiftDownArrow:
                    NavigateByLineIndexOffset(1);
                    break;
                case ESpecialKey.ControlLeftArrow:
                    EditingCommands.MoveLeftByWord.Execute(null, richTextBox);
                    CaretPosition = richTextBox.CaretPosition;
                    break;
                case ESpecialKey.ControlRightArrow:
                    EditingCommands.MoveRightByWord.Execute(null, richTextBox);
                    CaretPosition = richTextBox.CaretPosition;
                    break;
                case ESpecialKey.ControlShiftLeftArrow:
                    break;
                case ESpecialKey.ControlShiftRightArrow:
                    break;
                case ESpecialKey.ControlShiftUpArrow:
                    break;
                case ESpecialKey.ControlShiftDownArrow:
                    break;
                default:
                    break;
            }
        }

        private void NavigateByLineIndexOffset(int lineIndexOffset)
        {
            var currentBlockIndex = GetCurrentBlockIndex();
            var offset = GetCurrentCaretOffset(currentBlockIndex);
            CaretPosition = CurrentRichText.Blocks.ElementAt(currentBlockIndex + lineIndexOffset).ContentStart.GetPositionAtOffset(offset);
        }

        private int GetCurrentCaretOffset(int currentBlockIndex)
        {
            var currentBlock = CurrentRichText.Blocks.ElementAt(currentBlockIndex);
            return currentBlock.ContentStart.GetOffsetToPosition(CaretPosition);
        }

        private int GetCurrentBlockIndex()
        {
            for (int i = 0; i < CurrentRichText.Blocks.Count; i++)
            {
                var blockStart = CurrentRichText.Blocks.ElementAt(i).ContentStart;
                var blockEnd = CurrentRichText.Blocks.ElementAt(i).ContentEnd;
                if (
                    (blockStart.CompareTo(CaretPosition) == -1 && blockEnd.CompareTo(CaretPosition) == 1) ||
                    (blockStart.CompareTo(CaretPosition) == 0 || blockEnd.CompareTo(CaretPosition) == 0))
                {
                    return i;
                }
            }

            return -1;
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
        public void AddFormatMacroTask(EFormatType formatType)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = formatType,
            });
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
    }
}
