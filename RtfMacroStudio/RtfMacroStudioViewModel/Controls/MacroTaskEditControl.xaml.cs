using RtfMacroStudioViewModel.Enums;
using RtfMacroStudioViewModel.Interfaces;
using RtfMacroStudioViewModel.Models;
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
using System.Windows.Shapes;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Controls
{
    /// <summary>
    /// Interaction logic for MacroTaskEditControl.xaml
    /// </summary>
    public partial class MacroTaskEditControl : Window
    {
        public MacroTask DisplayedTask { get; set; }
        private StudioViewModel viewModel;

        public MacroTaskEditControl()
        {
            InitializeComponent();
        }

        public MacroTaskEditControl(StudioViewModel viewModel)
        {
            InitializeComponent();

            DisplayedTask = viewModel.MacroTaskInEdit;
            this.viewModel = viewModel;
            LabelTaskType.Content = DisplayedTask.MacroTaskType.GetFriendlyString();
            BindMacroTask(DisplayedTask, viewModel);
        }

        private void BindMacroTask(MacroTask macroTask, StudioViewModel viewModel)
        {
            switch (macroTask.MacroTaskType)
            {
                case EMacroTaskType.Text:
                    ShowOnlyText();
                    TextBoxTaskText.Text = macroTask.Line;
                    break;
                case EMacroTaskType.SpecialKey:
                    ShowOnlyComboBox();
                    ComboBoxItems.ItemsSource = viewModel.SupportedSpecialKeys;
                    SelectComboBoxItem(macroTask.SpecialKey.ToString());
                    break;
                case EMacroTaskType.Format:
                    ShowAppropriateFormattingOptions();
                    break;
                default:
                    break;
            }
        }

        private void ShowOnlyText()
        {
            ComboBoxItems.Visibility = Visibility.Hidden;
            TextBoxTaskText.Visibility = Visibility.Visible;
            RectangleColor.Visibility = Visibility.Hidden;
        }

        private void ShowOnlyComboBox()
        {
            ComboBoxItems.Visibility = Visibility.Visible;
            TextBoxTaskText.Visibility = Visibility.Hidden;
            RectangleColor.Visibility = Visibility.Hidden;
        }

        private void ShowAppropriateFormattingOptions()
        {
            if (DisplayedTask.FormatType == EFormatType.Color)
            {
                ComboBoxItems.Visibility = Visibility.Hidden;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Visible;
                RectangleColor.Fill = new SolidColorBrush(DisplayedTask.TextColor);
            }
            else if (DisplayedTask.FormatType == EFormatType.Font)
            {
                ComboBoxItems.Visibility = Visibility.Visible;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Hidden;
                ComboBoxItems.ItemsSource = viewModel.AvailableFonts;
                SelectComboBoxItem(DisplayedTask.TextFont.Source);
            }
            else if (DisplayedTask.FormatType == EFormatType.TextSize)
            {
                ComboBoxItems.Visibility = Visibility.Visible;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Hidden;
                ComboBoxItems.ItemsSource = viewModel.AvailableTextSizes;
                SelectComboBoxItem(DisplayedTask.TextSize.ToString());
            }
            else
            {
                ComboBoxItems.Visibility = Visibility.Visible;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Hidden;
                ComboBoxItems.ItemsSource = viewModel.SupportedFormattingOptions;
                SelectComboBoxItem(DisplayedTask.FormatType.ToString());
            }
        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            viewModel.EditMacroTaskCancel();
            this.Close();
        }

        private void buttonOk_Click(object sender, RoutedEventArgs e)
        {
            viewModel.EditMacroTaskComplete(DisplayedTask, TextBoxTaskText.Text, ComboBoxItems.SelectedItem);
            this.Close();
        }

        private void SelectComboBoxItem(string text)
        {
            for (int i = 0; i < ComboBoxItems.Items.Count; i++)
            {
                if (ComboBoxItems.Items[i].ToString() == text)
                {
                    ComboBoxItems.SelectedIndex = i;
                    return;
                }
            }
        }
    }
}
