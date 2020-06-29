using RtfMacroStudioViewModel.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfMacroStudioViewModel.ViewModel
{
    public class StudioViewModelMock : StudioViewModel
    {
        public int NumTimesRunMacroCalled = 0;
        
        public StudioViewModelMock(IEditingCommandHelper editingCommandHelper, IFileHelper fileHelper, IMacroTaskEditPresenter macroTaskEditPresenter, IMacroRunPresenter macroRunPresenter) 
            : base(editingCommandHelper, fileHelper, macroTaskEditPresenter, macroRunPresenter)
        {
        }

        public override void RunMacro()
        {
            NumTimesRunMacroCalled++;
            base.RunMacro();
        }

    }
}
