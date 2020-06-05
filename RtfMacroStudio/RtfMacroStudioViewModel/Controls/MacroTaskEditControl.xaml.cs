using RtfMacroStudioViewModel.Enums;
using RtfMacroStudioViewModel.Models;
using RtfMacroStudioViewModel.ViewModel;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
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
            TextSizeSelector.Visibility = Visibility.Hidden;
        }

        private void ShowOnlyComboBox()
        {
            ComboBoxItems.Visibility = Visibility.Visible;
            TextBoxTaskText.Visibility = Visibility.Hidden;
            RectangleColor.Visibility = Visibility.Hidden;
            TextSizeSelector.Visibility = Visibility.Hidden;
        }

        private void ShowAppropriateFormattingOptions()
        {
            if (DisplayedTask.FormatType == EFormatType.Color)
            {
                ComboBoxItems.Visibility = Visibility.Hidden;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Visible;
                TextSizeSelector.Visibility = Visibility.Hidden;
                RectangleColor.Fill = new SolidColorBrush(DisplayedTask.TextColor);
            }
            else if (DisplayedTask.FormatType == EFormatType.Font)
            {
                ComboBoxItems.Visibility = Visibility.Visible;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Hidden;
                TextSizeSelector.Visibility = Visibility.Hidden;
                ComboBoxItems.ItemsSource = viewModel.AvailableFonts;
                SelectComboBoxItem(DisplayedTask.TextFont.Source);
            }
            else if (DisplayedTask.FormatType == EFormatType.TextSize)
            {
                ComboBoxItems.Visibility = Visibility.Hidden;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Hidden;
                TextSizeSelector.Visibility = Visibility.Visible;
                TextSizeSelector.ComboBoxSizes.ItemsSource = viewModel.AvailableTextSizes;
                TextSizeSelector.SelectSize(DisplayedTask.TextSize.ToString());
            }
            else
            {
                ComboBoxItems.Visibility = Visibility.Visible;
                TextBoxTaskText.Visibility = Visibility.Hidden;
                RectangleColor.Visibility = Visibility.Hidden;
                TextSizeSelector.Visibility = Visibility.Hidden;
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
            if (DisplayedTask.MacroTaskType == EMacroTaskType.Format && 
                (DisplayedTask.FormatType == EFormatType.Color || DisplayedTask.FormatType == EFormatType.TextSize))
            {
                if (DisplayedTask.FormatType == EFormatType.Color)
                {
                    viewModel.EditMacroTaskComplete(DisplayedTask, string.Empty, ((SolidColorBrush)RectangleColor.Fill).Color);
                }
                else if (DisplayedTask.FormatType == EFormatType.TextSize)
                {
                    viewModel.EditMacroTaskComplete(DisplayedTask, string.Empty, TextSizeSelector.ComboBoxSizes.SelectedItem);
                }
            }
            else
            {
                viewModel.EditMacroTaskComplete(DisplayedTask, TextBoxTaskText.Text, ComboBoxItems.SelectedItem);
            }
            
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

        private void RectangleColor_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (RectangleColor.Visibility == Visibility.Hidden)
            {
                return;
            }

            ColorDialog colorDialog = new ColorDialog();

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DisplayedTask.TextColor = Color.FromRgb(colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
                RectangleColor.Fill = new SolidColorBrush(DisplayedTask.TextColor);
            }
        }
    }
}
