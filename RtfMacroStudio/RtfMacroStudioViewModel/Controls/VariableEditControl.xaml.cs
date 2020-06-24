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

namespace RtfMacroStudioViewModel.Controls
{
    /// <summary>
    /// Interaction logic for VariableEditControl.xaml
    /// </summary>
    public partial class VariableEditControl : UserControl
    {
        public VariableEditControl()
        {
            InitializeComponent();
        }

        public delegate void NotifyVariableChange();
        public event NotifyVariableChange VariableNameChanged;

        private string variableName = string.Empty;
        public string VariableName
        {
            get
            {
                return variableName;
            }
            private set
            {
                variableName = value;
                VariableNameChanged?.Invoke();
            }
        }

        public int VariableValue
        {
            get
            {
                int n = 0;
                int.TryParse(TextBoxValue.Text, out n);
                return n;
            }
        }

        public int VariableIncrementValue
        {
            get
            {
                int n = 1;
                int.TryParse(TextBoxValue.Text, out n);
                return n;
            }
        }

        private void TextBoxIntValue_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            int n;

            if (int.TryParse(e.Text, out n) == false)
            {
                e.Handled = true;
            }
        }

        private void TextBoxName_TextChanged(object sender, TextChangedEventArgs e)
        {
            VariableName = TextBoxName.Text;
        }
    }
}
