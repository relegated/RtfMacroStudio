using RtfMacroStudioViewModel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        public void AddTextInputMacroTask(string textInput)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run(textInput));
            CurrentTaskList.Add(new MacroTask()
            {
                Line = paragraph,
                MacroTaskType = EMacroTaskType.Text
            });
        }

        public void AddSpecialKeyMacroTask(Key specialKey)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.SpecialKey,
                KeyStroke = specialKey,
            });
        }

        #endregion

        #region Constructor

        public StudioViewModel()
        {

        }

        #endregion

        public void RunMacro()
        {
            foreach (var task in CurrentTaskList)
            {
                CurrentRichText.Blocks.Add(task.Line);
            }
        }

        public void AddFormatMacroTask(EFormatType formatType)
        {
            CurrentTaskList.Add(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = formatType,
            });
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
    }
}
