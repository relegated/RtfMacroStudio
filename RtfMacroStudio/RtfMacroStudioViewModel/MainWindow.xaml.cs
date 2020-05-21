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
using RtfMacroStudioViewModel.Enums;

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

            CreateSpecialKeyMenuItems();

            RichTextBoxMain.Document = viewModel.CurrentRichText;
            RichTextBoxMain.TextChanged += RichTextBoxMain_TextChanged;

            viewModel.PropertyChanged += ViewModel_PropertyChanged;
        }

        private void CreateSpecialKeyMenuItems()
        {
            foreach (var specialKey in viewModel.SupportedSpecialKeys)
            {
                RibbonMenuItem ribbonMenuItem = new RibbonMenuItem() 
                { 
                    Header = specialKey.GetFriendlyString(),
                    Name = specialKey.ToString(),
                };

                ribbonMenuItem.Click += RibbonMenuItemAddSpecialKeyMacroTask_Click;

                RibbonMenuButtonSpecialKeys.Items.Add(ribbonMenuItem);
            }
        }

        private void RibbonMenuItemAddSpecialKeyMacroTask_Click(object sender, RoutedEventArgs e)
        {
            ESpecialKey specialKey;
            if (Enum.TryParse<ESpecialKey>(((RibbonMenuItem)sender).Name, out specialKey))
            {
                viewModel.AddSpecialKeyMacroTask(specialKey);
            }
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

        private void RichTextBoxMain_LostFocus(object sender, RoutedEventArgs e)
        {
            var position = RichTextBoxMain.CaretPosition;
            EditingCommands.MoveToLineStart.Execute(null, RichTextBoxMain);
            MessageBox.Show($"Difference: {position.GetOffsetToPosition(RichTextBoxMain.CaretPosition)}");
        }
    }
}
