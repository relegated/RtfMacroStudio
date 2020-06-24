using RtfMacroStudioViewModel.Enums;
using RtfMacroStudioViewModel.Models;
using RtfMacroStudioViewModel.ViewModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Controls
{
    /// <summary>
    /// Interaction logic for MacroTaskControl.xaml
    /// </summary>
    public partial class MacroTaskControl : UserControl
    {
        private readonly MacroTask macroTask;
        private readonly StudioViewModel viewModel;

        public MacroTaskControl(MacroTask macroTask, StudioViewModel viewModel)
        {
            InitializeComponent();

            this.macroTask = macroTask;
            this.viewModel = viewModel;
            this.DataContext = macroTask;
            this.AllowDrop = true;

            TaskTextTitle.Text = RtfEnumStringRetriever.GetFriendlyString(macroTask.MacroTaskType);

            switch (macroTask.MacroTaskType)
            {
                case EMacroTaskType.SpecialKey:
                    TaskText.Text = macroTask.SpecialKey.ToString();
                    HideColorFontAndSize();
                    break;
                case EMacroTaskType.Format:
                    TaskText.Text = macroTask.FormatType.ToString();
                    SetupFormatType();
                    break;
                case EMacroTaskType.Undefined:
                case EMacroTaskType.Text:
                case EMacroTaskType.Variable:
                default:
                    HideColorFontAndSize();
                    TaskText.Text = StudioViewModel.GetTextFromMacroTask(macroTask);
                    break;
            }

        }

        private void SetupFormatType()
        {
            if (macroTask.FormatType == EFormatType.Color)
            {
                ShowFontColor();
                TextColor.Fill = new SolidColorBrush(macroTask.TextColor);
            }
            else if (macroTask.FormatType == EFormatType.Font)
            {
                ShowFontFamily();
                FontFamily.FontFamily = macroTask.TextFont;
                FontFamily.Text = macroTask.TextFont.Source;
            }
            else if (macroTask.FormatType == EFormatType.TextSize)
            {
                ShowFontSize();
                FontSize.Text = macroTask.TextSize.ToString();
            }
            else
            {
                HideColorFontAndSize();
            }
        }

        private void HideColorFontAndSize()
        {
            TextColor.Visibility = Visibility.Hidden;
            FontFamily.Visibility = Visibility.Hidden;
            FontSize.Visibility = Visibility.Hidden;
        }

        private void ShowFontColor()
        {
            TextColor.Visibility = Visibility.Visible;
            FontFamily.Visibility = Visibility.Hidden;
            FontSize.Visibility = Visibility.Hidden;
        }

        private void ShowFontFamily()
        {
            TextColor.Visibility = Visibility.Hidden;
            FontFamily.Visibility = Visibility.Visible;
            FontSize.Visibility = Visibility.Hidden;
        }

        private void ShowFontSize()
        {
            TextColor.Visibility = Visibility.Hidden;
            FontFamily.Visibility = Visibility.Hidden;
            FontSize.Visibility = Visibility.Visible;
        }

        private void UserControl_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            viewModel.EditMacroTaskBegin(macroTask);
        }

        private void ContextMenuDelete_Click(object sender, RoutedEventArgs e)
        {
            viewModel.RemoveTaskAt(macroTask.Index);
        }

        private void Canvas_Drop(object sender, DragEventArgs e)
        {
            viewModel.ProcessDroppedControl(macroTask, e.Data.GetData("macroTask"));
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DataObject data = new DataObject();
                data.SetData("macroTask", macroTask);
                DragDrop.DoDragDrop(this, data, DragDropEffects.Move);
            }
        }

        protected override void OnDragOver(DragEventArgs e)
        {
            base.OnDragOver(e);
            RectangleBorder.StrokeThickness = 10;
        }

        protected override void OnDragLeave(DragEventArgs e)
        {
            base.OnDragLeave(e);
            RectangleBorder.StrokeThickness = 1;
        }
    }
}
