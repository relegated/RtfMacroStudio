using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Interfaces
{
    public interface IMacroRunPresenter
    {
        bool ShowRunOptionWindow();

        ERunPresenterOption Option { get; set; }
        int N { get; set; }
    }
}
