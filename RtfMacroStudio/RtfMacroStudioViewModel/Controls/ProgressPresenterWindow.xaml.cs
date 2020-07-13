using RtfMacroStudioViewModel.Interfaces;
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
using System.Windows.Shapes;

namespace RtfMacroStudioViewModel.Controls
{
    /// <summary>
    /// Interaction logic for ProgressPresenterWindow.xaml
    /// </summary>
    public partial class ProgressPresenterWindow : Window, IProgressPresenter
    {
        private bool canClose = false;
        public int RunNTimes { get; set; } = 1;
        private int currentlyRunningXofN = 1;
        public int CurrentlyRunningXofN 
        {
            get => currentlyRunningXofN;
            set
            {
                currentlyRunningXofN = value;
                ProgressBarTaskRunProgress.Value = (double)currentlyRunningXofN / (double)RunNTimes;
                LabelProgress.Content = $"Currently running task {currentlyRunningXofN} of {RunNTimes}...";
            }
        }

        public ProgressPresenterWindow()
        {
            InitializeComponent();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //This window cannot be closed by user
            if (canClose == false)
            {
                e.Cancel = true;
            }           
        }

        void IProgressPresenter.ShowDialog()
        {
            Show();
        }

        void IProgressPresenter.Hide()
        {
            Hide();
        }

        public void Dispose()
        {
            canClose = true;
            Close();
        }
    }
}
