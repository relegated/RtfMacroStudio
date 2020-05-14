using RtfMacroStudioViewModel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
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



        

        #endregion

        #region Constructor

        public StudioViewModel()
        {
            SupportedSpecialKeys = Enum.GetValues(typeof(ESpecialKey)).Cast<ESpecialKey>().ToList();
        }

        #endregion

        public void RunMacro()
        {
            foreach (var task in CurrentTaskList)
            {
                CurrentRichText.Blocks.Add(task.Line);
            }
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
                KeyStroke = specialKey,
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
