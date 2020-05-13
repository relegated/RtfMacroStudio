using RtfMacroStudioViewModel.ViewModel;
using System.Windows;

namespace RtfMacroStudioViewModel.Controllers
{
    public class MacroStudioController
    {

        public void StartMacroStudio()
        {
            StudioViewModel viewModel = new StudioViewModel();
            Window mainWindow = new MainWindow(viewModel);
            mainWindow.Show();
        }

    }
}
