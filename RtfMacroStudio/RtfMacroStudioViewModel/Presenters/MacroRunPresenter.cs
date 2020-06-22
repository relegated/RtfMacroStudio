using RtfMacroStudioViewModel.Controls;
using RtfMacroStudioViewModel.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Presenters
{
    public class MacroRunPresenter : IMacroRunPresenter
    {
 
        private int n = 1;
        public int N
        {
            get => n;
            set
            {
             if (value >=  1)
                {
                    n = value;
                }
            }
        }

        public ERunPresenterOption Option { get; set; } = ERunPresenterOption.Once;

        public bool ShowRunOptionWindow()
        {
            MacroRunPresenterWindow presenterWindow = new MacroRunPresenterWindow(this);
            return presenterWindow.ShowDialog() ?? false;
        }
    }
}
