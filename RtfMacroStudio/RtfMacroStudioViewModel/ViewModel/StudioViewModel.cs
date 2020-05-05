using RtfMacroStudioViewModel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

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

        
    }
}
