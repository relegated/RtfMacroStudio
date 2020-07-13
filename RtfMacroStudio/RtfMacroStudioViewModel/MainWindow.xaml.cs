using RtfMacroStudioViewModel.Controls;
using RtfMacroStudioViewModel.Enums;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Ribbon;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
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
            DataContext = viewModel;

            CreateSpecialKeyMenuItems();
            CreateFormatMenuItems();

            viewModel.RichTextBoxControl = RichTextBoxMain;
            RichTextBoxMain.Document = viewModel.CurrentRichText;
            RichTextBoxMain.TextChanged += RichTextBoxMain_TextChanged;

            RibbonButtonColor.LargeImageSource = viewModel.ColorImageDrawing;
            RibbonButtonColor.SmallImageSource = viewModel.ColorImageDrawing;

            viewModel.PropertyChanged += ViewModel_PropertyChanged;
        }

        private void CreateFormatMenuItems()
        {
            foreach (var formatOption in viewModel.SupportedFormattingOptions)
            {
                RibbonMenuItem ribbonMenuItem = new RibbonMenuItem()
                {
                    Header = formatOption.GetFriendlyString(),
                    Name = formatOption.ToString(),
                };

                ribbonMenuItem.Click += RibbonMenuItemAddFormatOption_Click;

                RibbonMenuButtonFormat.Items.Add(ribbonMenuItem);
            }
        }

        private void RibbonMenuItemAddFormatOption_Click(object sender, RoutedEventArgs e)
        {
            viewModel.AddFormatMacroTask(((RibbonMenuItem)sender).Name);
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
                        TaskList.Children.Add(new MacroTaskControl(macroTask, viewModel));
                    }
                    ScrollViewerTaskList.ScrollToBottom();
                    break;
                case nameof(viewModel.CurrentRichText):
                    RichTextBoxMain.Focus();
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
            viewModel.AddTextInputMacroTask("Double-click to Edit Text");
        }

        private void RichTextBoxMain_LostFocus(object sender, RoutedEventArgs e)
        {
            //Update Text in viewmodel
        }

        private void RibbonButtonRunMacro_Click(object sender, RoutedEventArgs e)
        {
            viewModel.RunMacroPresenter();
        }

        private void RichTextBoxMain_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            List<ModifierKeys> modifierKeys = GetListOfModifierKeys();

            if (modifierKeys.Count == 0)
            {
                viewModel.ProcessKey(e.Key, null);
            }
            else
            {
                viewModel.ProcessKey(e.Key, modifierKeys.ToArray());
            }
        }

        private List<ModifierKeys> GetListOfModifierKeys()
        {
            List<ModifierKeys> returnList = new List<ModifierKeys>();

            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                returnList.Add(ModifierKeys.Control);
            }

            if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
            {
                returnList.Add(ModifierKeys.Shift);
            }

            return returnList;
        }

        private void ToggleButtonAlignment_Checked(object sender, RoutedEventArgs e)
        {
            if (ToggleButtonAlignLeft == null || 
                ToggleButtonAlignCenter == null || 
                ToggleButtonAlignRight == null || 
                ToggleButtonAlignJustify == null)
            {
                return;
            }

            var buttonName = ((RibbonToggleButton)sender).Name;


            if (buttonName == nameof(ToggleButtonAlignLeft))
            {
                ToggleButtonAlignCenter.IsChecked = false;
                ToggleButtonAlignRight.IsChecked = false;
                ToggleButtonAlignJustify.IsChecked = false;
                viewModel.CurrentTextAlignment = TextAlignment.Left;
            }
            else if (buttonName == nameof(ToggleButtonAlignCenter))
            {
                ToggleButtonAlignLeft.IsChecked = false;
                ToggleButtonAlignRight.IsChecked = false;
                ToggleButtonAlignJustify.IsChecked = false;
                viewModel.CurrentTextAlignment = TextAlignment.Center;
            }
            else if (buttonName == nameof(ToggleButtonAlignRight))
            {
                ToggleButtonAlignLeft.IsChecked = false;
                ToggleButtonAlignCenter.IsChecked = false;
                ToggleButtonAlignJustify.IsChecked = false;
                viewModel.CurrentTextAlignment = TextAlignment.Right;
            }
            else if (buttonName == nameof(ToggleButtonAlignJustify))
            {
                ToggleButtonAlignLeft.IsChecked = false;
                ToggleButtonAlignCenter.IsChecked = false;
                ToggleButtonAlignRight.IsChecked = false;
                viewModel.CurrentTextAlignment = TextAlignment.Justify;
            }
        }

        private void RibbonButtonColor_Click(object sender, RoutedEventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                viewModel.CurrentColor = Color.FromRgb(colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
                RibbonButtonColor.LargeImageSource = viewModel.ColorImageDrawing;
                RibbonButtonColor.SmallImageSource = viewModel.ColorImageDrawing;
            }
        }

        private void RibbonButtonAddVariable_Click(object sender, RoutedEventArgs e)
        {
            viewModel.AddVariableMacroTask("Edit Variable", 0, 1, false, 1);
        }

        private void ToggleButtonRecordControl_Checked(object sender, RoutedEventArgs e)
        {
            if (ToggleButtonRecordMacro == null ||
                ToggleButtonStopRecording == null ||
                viewModel == null) 
            {
                return;
            }

            var buttonName = ((RibbonToggleButton)sender).Name;

            if (buttonName == nameof(ToggleButtonRecordMacro))
            {
                ToggleButtonStopRecording.IsChecked = false;
                viewModel.RecordMacroStart();
            }
            else if (buttonName == nameof(ToggleButtonStopRecording))
            {
                ToggleButtonRecordMacro.IsChecked = false;
                viewModel.StopRecording();
            }
        }

        private void OptionMenuItemOpen_Click(object sender, RoutedEventArgs e)
        {
            viewModel.OpenFile();
        }

        private void OptionMenuItemSave_Click(object sender, RoutedEventArgs e)
        {
            viewModel.SaveFile();
        }

        private void OptionMenuItemExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void RibbonButtonPaste_Click(object sender, RoutedEventArgs e)
        {
            viewModel.PasteFromClipboard();
        }

        private void RibbonButtonCut_Click(object sender, RoutedEventArgs e)
        {
            viewModel.CutToClipboard();
        }

        private void RibbonButtonCopy_Click(object sender, RoutedEventArgs e)
        {
            viewModel.CopyToClipboard();
        }

        private void RichTextStudioForm_Closed(object sender, EventArgs e)
        {
            viewModel.CloseProgram();
        }

        private void RibbonButtonClearAllTasks_Click(object sender, RoutedEventArgs e)
        {
            if (System.Windows.MessageBox.Show("This will delete all macro tasks.  Are you sure?", 
                "Confirm Clear Macro Tasks", MessageBoxButton.YesNo, MessageBoxImage.Warning)
                == MessageBoxResult.Yes)
            {
                viewModel.ClearAllTasks();
            }
        }
    }
}
