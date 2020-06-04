using RtfMacroStudioViewModel.Controls;
using RtfMacroStudioViewModel.Interfaces;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfMacroStudioViewModel.Presenters
{
    public class MacroTaskEditPresenter : IMacroTaskEditPresenter
    {
        public void ShowEditControl(StudioViewModel viewModel)
        {
            MacroTaskEditControl macroTaskEditControl = new MacroTaskEditControl(viewModel);
            macroTaskEditControl.ShowDialog();
        }
    }
}
