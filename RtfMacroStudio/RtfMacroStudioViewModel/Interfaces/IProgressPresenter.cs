using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfMacroStudioViewModel.Interfaces
{
    public interface IProgressPresenter
    {
        int RunNTimes { get; set; }
        int CurrentlyRunningXofN { get; set; }

        void ShowDialog();
        void Hide();
        void Dispose();
    }
}
