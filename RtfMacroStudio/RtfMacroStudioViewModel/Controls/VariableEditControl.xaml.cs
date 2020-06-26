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

        public void SetInitialValues(Variable variable)
        {
            TextBoxName.Text = variable.Name;
            OriginalName = variable.Name;
            TextBoxValue.Text = variable.Value.ToString();
            TextBoxIncrementValue.Text = variable.IncrementByValue.ToString();
            TextBoxPlaceValues.Text = variable.PlaceValuesToFill.ToString();
            CheckBoxUsePlaceValue.IsChecked = variable.UsePlaceValues;
        }

        public string OriginalName { get; set; }

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
                int.TryParse(TextBoxIncrementValue.Text, out n);
                return n;
            }
        }

        public int VariablePlaceValue
        {
            get
            {
                int n = 1;
                int.TryParse(TextBoxPlaceValues.Text, out n);
                return n;
            }
        }

        public bool VariableUsePlaceValue
        {
            get
            {
                return CheckBoxUsePlaceValue.IsChecked ?? false;
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

        private void CheckBoxUsePlaceValue_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxPlaceValues.IsReadOnly = false;
            TextBoxPlaceValues.Background = new SolidColorBrush(Colors.White);
        }

        private void CheckBoxUsePlaceValue_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxPlaceValues.IsReadOnly = true;
            TextBoxPlaceValues.Background = new SolidColorBrush(Colors.Gray);
        }
    }
}
