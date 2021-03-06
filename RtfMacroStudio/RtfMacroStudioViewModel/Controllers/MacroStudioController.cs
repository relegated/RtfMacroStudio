﻿using RtfMacroStudioViewModel.Controls;
using RtfMacroStudioViewModel.Helpers;
using RtfMacroStudioViewModel.Presenters;
using RtfMacroStudioViewModel.ViewModel;
using System.Windows;

namespace RtfMacroStudioViewModel.Controllers
{
    public class MacroStudioController
    {

        public void StartMacroStudio()
        {
            StudioViewModel viewModel = new StudioViewModel(new EditingCommandHelper(), new FileHelper(), new MacroTaskEditPresenter(), new MacroRunPresenter(), new ProgressPresenterWindow());
            Window mainWindow = new MainWindow(viewModel);
            mainWindow.Show();
        }

    }
}
