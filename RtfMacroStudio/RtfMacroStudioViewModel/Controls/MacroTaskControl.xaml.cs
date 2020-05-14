using RtfMacroStudioViewModel.Enums;
using RtfMacroStudioViewModel.Models;
using RtfMacroStudioViewModel.ViewModel;
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
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Controls
{
    /// <summary>
    /// Interaction logic for MacroTaskControl.xaml
    /// </summary>
    public partial class MacroTaskControl : UserControl
    {
        public MacroTaskControl(MacroTask macroTask)
        {
            InitializeComponent();

            TaskTextTitle.Text = RtfEnumStringRetriever.GetFriendlyString(macroTask.MacroTaskType);

            switch (macroTask.MacroTaskType)
            {
                case EMacroTaskType.SpecialKey:
                    TaskText.Text = macroTask.KeyStroke.ToString();
                    break;
                case EMacroTaskType.Format:
                    TaskText.Text = macroTask.FormatType.ToString();
                    break;
                case EMacroTaskType.Undefined:
                case EMacroTaskType.Text:
                default:
                    TaskText.Text = StudioViewModel.GetTextFromMacroTask(macroTask);
                    break;
            }
        }

        
    }
}
