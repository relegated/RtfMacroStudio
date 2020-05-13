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
using System.Windows.Controls.Ribbon;
using RtfMacroStudioViewModel.Controls;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        StudioViewModel viewModel;

        public MainWindow(StudioViewModel viewModel)
        {
            InitializeComponent();

            Bind(viewModel);
        }

        private void Bind(StudioViewModel viewModel)
        {
            this.viewModel = viewModel;

            RichTextBoxMain.Document = viewModel.CurrentRichText;
            RichTextBoxMain.TextChanged += RichTextBoxMain_TextChanged;

            viewModel.PropertyChanged += ViewModel_PropertyChanged;
        }

        private void ViewModel_PropertyChanged(string PropertyName)
        {
            switch (PropertyName)
            {
                case nameof(viewModel.CurrentTaskList):
                    //for now, refresh the list of tasks
                    TaskList.Children.Clear();
                    foreach (var macroTask in viewModel.CurrentTaskList)
                    {
                        TaskList.Children.Add(new MacroTaskControl(macroTask));
                    }
                    break;
                default:
                    break;
            }
        }

        private void RichTextBoxMain_TextChanged(object sender, TextChangedEventArgs e)
        {
            viewModel.CurrentRichText = RichTextBoxMain.Document;
        }

        private void RibbonButtonAddMacroKeystroke_Click(object sender, RoutedEventArgs e)
        {
            viewModel.AddTextInputMacroTask("Demonstration Text");
        }
    }
}
