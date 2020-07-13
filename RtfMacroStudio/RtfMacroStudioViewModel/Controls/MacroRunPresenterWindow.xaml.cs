using RtfMacroStudioViewModel.Interfaces;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Controls
{
    /// <summary>
    /// Interaction logic for MacroRunPresenterWindow.xaml
    /// </summary>
    public partial class MacroRunPresenterWindow : Window
    {
        public MacroRunPresenterWindow(IMacroRunPresenter macroRunPresenter)
        {
            InitializeComponent();
            MacroRunPresenter = macroRunPresenter;
        }

        public IMacroRunPresenter MacroRunPresenter { get; }

        private void Radio_Checked(object sender, RoutedEventArgs e)
        {
            if (RadioOnce == null || RadioNTimes == null || RadioEnd == null || TextNTimes == null)
            {
                return;
            }

            var checkedOption = ((RadioButton)sender).Name;

            if (checkedOption == nameof(RadioOnce))
            {
                MacroRunPresenter.Option = ERunPresenterOption.Once;
                TextNTimes.IsReadOnly = true;
            }
            else if (checkedOption == nameof(RadioNTimes))
            {
                MacroRunPresenter.Option = ERunPresenterOption.NTimes;
                TextNTimes.IsReadOnly = false;

                TextNTimes.GotFocus -= TextNTimes_GotFocus;

                TextNTimes.Focus();

                TextNTimes.GotFocus += TextNTimes_GotFocus;
            }
            else if (checkedOption == nameof(RadioEnd))
            {
                MacroRunPresenter.Option = ERunPresenterOption.End;
                TextNTimes.IsReadOnly = true;
            }
        }

        private void TextNTimes_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            int n;

            e.Handled = (int.TryParse(e.Text, out n) == false);
        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            this.Close();
        }

        private void ButtonOK_Click(object sender, RoutedEventArgs e)
        {
            MacroRunPresenter.N = int.Parse(TextNTimes.Text);
            DialogResult = true;
            this.Close();
        }

        private void TextNTimes_GotFocus(object sender, RoutedEventArgs e)
        {
            RadioNTimes.IsChecked = true;
        }
    }
}
