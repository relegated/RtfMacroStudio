using Moq;
using NUnit.Framework;
using RtfMacroStudioViewModel.Helpers;
using RtfMacroStudioViewModel.Interfaces;
using RtfMacroStudioViewModel.Models;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioView
{
    [TestFixture, RequiresSTA]
    public class StudioViewModelTests
    {
        StudioViewModel viewModel;
        string propertyChangedText;
        Mock<IEditingCommandHelper> mockEditingCommandHelper;

        [SetUp]
        public void Setup()
        {
            mockEditingCommandHelper = new Mock<IEditingCommandHelper>();
            viewModel = new StudioViewModel(mockEditingCommandHelper.Object);
            viewModel.PropertyChanged += ViewModel_PropertyChanged;
            propertyChangedText = string.Empty;
        }

        [Test]
        public void CanCreateViewModel()
        {
            Assert.That(viewModel != null);
        }

        [Test]
        public void TextPositionBeginsAtDocumentHome()
        {
            Assert.That(viewModel.CaretPosition.CompareTo(viewModel.CurrentRichText.ContentStart) == 0);
        }

        [Test]
        public void CanAddTextInputMacroTask()
        {
            viewModel.AddTextInputMacroTask("text to add");

            Assert.That(propertyChangedText == nameof(viewModel.CurrentTaskList));
            Assert.That(viewModel.CurrentTaskList.Count == 1);
            Assert.That(((Run)viewModel.CurrentTaskList[0].Line.Inlines.FirstInline).Text == "text to add");
        }

        [Test]
        public void CanAddSpecialCharacterMacroTask()
        {
            viewModel.AddSpecialKeyMacroTask(ESpecialKey.Home);
            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlRightArrow);

            Assert.That(propertyChangedText == nameof(viewModel.CurrentTaskList));
            Assert.That(viewModel.CurrentTaskList.Count == 2);
            Assert.That(viewModel.CurrentTaskList[0].SpecialKey == ESpecialKey.Home);
            Assert.That(viewModel.CurrentTaskList[1].SpecialKey == ESpecialKey.ControlRightArrow);
        }

        [Test]
        public void CanAddFormatMacroTask()
        {
            viewModel.AddFormatMacroTask(new MacroTask()
            {
                FormatType = EFormatType.Bold,
            });

            Assert.That(propertyChangedText == nameof(viewModel.CurrentTaskList));
            Assert.That(viewModel.CurrentTaskList.Count == 1);
            Assert.That(viewModel.CurrentTaskList[0].FormatType == EFormatType.Bold);
        }

        [Test]
        public void CanRemoveMacroTask()
        {
            GivenSomeBasicTasks();

            viewModel.RemoveTaskAt(1);

            Assert.That(viewModel.CurrentTaskList.Count == 2);
            Assert.That(viewModel.CurrentTaskList[0].MacroTaskType == EMacroTaskType.Format);
            Assert.That(viewModel.CurrentTaskList[1].MacroTaskType == EMacroTaskType.Text);
        }

        [TestCase(-1)]
        [TestCase(7)]
        public void HandleAttemptToRemoveTaskWithInvalidIndex(int taskIndex)
        {
            GivenSomeBasicTasks();

            viewModel.RemoveTaskAt(taskIndex);

            Assert.That(viewModel.CurrentTaskList.Count == 3);
        }

        [Test]
        public void CanClearTaskList()
        {
            GivenSomeBasicTasks();

            viewModel.ClearAllTasks();

            Assert.That(viewModel.CurrentTaskList.Count == 0);
        }

        [Test]
        public void CanMoveTask()
        {
            GivenSomeBasicTasks();

            viewModel.MoveTask(2, 0);

            Assert.That(viewModel.CurrentTaskList.Count == 3);
            Assert.That(viewModel.CurrentTaskList[0].MacroTaskType == EMacroTaskType.Text);
            Assert.That(viewModel.CurrentTaskList[1].MacroTaskType == EMacroTaskType.Format);
            Assert.That(viewModel.CurrentTaskList[2].MacroTaskType == EMacroTaskType.SpecialKey);
        }

        [TestCase(1, -1)]
        [TestCase(-1, 0)]
        [TestCase(6, 8)]
        public void HandleInvalidMoveInputs(int sourceIndex, int destinationIndex)
        {
            GivenSomeBasicTasks();

            viewModel.MoveTask(sourceIndex, destinationIndex);

            Assert.That(viewModel.CurrentTaskList.Count == 3);
            Assert.That(viewModel.CurrentTaskList[0].MacroTaskType == EMacroTaskType.Format);
            Assert.That(viewModel.CurrentTaskList[1].MacroTaskType == EMacroTaskType.SpecialKey);
            Assert.That(viewModel.CurrentTaskList[2].MacroTaskType == EMacroTaskType.Text);
        }

        [TestCase(2)]
        [TestCase(10)]
        public void RunMacroUpdatesTheTextBasedOnTask(int numberOfTasksToInsert)
        {
            for (int i = 0; i < numberOfTasksToInsert; i++)
            {
                viewModel.AddTextInputMacroTask("text to add"); 
            }

            //ensure the text has not been changed yet
            Assert.That(viewModel.CurrentRichText.Blocks.Count == 0);

            viewModel.RunMacro();
            Assert.That(viewModel.CurrentRichText.Blocks.Count == numberOfTasksToInsert);
        }

        [Test]
        public void CanGetTextFromMacroTask()
        {
            viewModel.AddTextInputMacroTask("This is a task");

            Assert.That(StudioViewModel.GetTextFromMacroTask(viewModel.CurrentTaskList[0]) == "This is a task");
        }

        [Test]
        public void SupportedSpecialKeysArePopulated()
        {
            Assert.That(viewModel.SupportedSpecialKeys.Count == 23);
        }

        [Test]
        public void TextInputChangesCaratPosition()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();

            Assert.That(viewModel.CaretPosition.CompareTo(viewModel.CurrentRichText.ContentEnd) == 0);
        }

        [Test]
        public void CanMoveRightByCharacter()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.RightArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveRightByCharacter(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveLeftByCharacter()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.LeftArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveLeftByCharacter(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectRightByCharacter()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ShiftRightArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectRightByCharacter(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectLeftByCharacter()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ShiftLeftArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectLeftByCharacter(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveRightByWord()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlRightArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveRightByWord(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveLeftByWord()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlLeftArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveLeftByWord(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectRightByWord()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlShiftRightArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectRightByWord(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectLeftByWord()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlShiftLeftArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectLeftByWord(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveUpByLine()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.UpArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveUpByLine(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveDownByLine()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.DownArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveDownByLine(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectUpByLine()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ShiftUpArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectUpByLine(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectDownByLine()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ShiftDownArrow);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectDownByLine(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveToLineStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.Home);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveToLineStart(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveToLineEnd()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.End);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveToLineEnd(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectToLineStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ShiftHome);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectToLineStart(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectToLineEnd()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ShiftEnd);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectToLineEnd(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveToDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlHome);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveToDocumentStart(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanMoveToDocumentEnd()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlEnd);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.MoveToDocumentEnd(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectToDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlShiftHome);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectToDocumentStart(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanSelectToDocumentEnd()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlShiftEnd);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.SelectToDocumentEnd(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanEnter()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.Enter);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.EnterParagraphBreak(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanBackspace()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.Backspace);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.Backspace(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanDelete()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(ESpecialKey.Delete);
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.Delete(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignCenter()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask() 
            {
               MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.AlignCenter,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignCenter(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignJustify()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.AlignJustify,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignJustify(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignLeft()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.AlignLeft,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignLeft(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignRight()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.AlignRight,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignRight(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanToggleBold()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.Bold,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.ToggleBold(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanToggleItalic()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.Italic,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.ToggleItalic(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanToggleUnderline()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.Underline,
            });
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.ToggleUnderline(It.IsAny<RichTextBox>()), Times.Once);
        }

        [TestCase(255, 0, 0)]
        [TestCase(0, 255, 0)]
        [TestCase(0, 0, 255)]
        public void CanChangeColor(byte r, byte g, byte b)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.Color,
                TextColor = Color.FromRgb(r, g, b),
            });
            viewModel.RunMacro();

            var brush = (SolidColorBrush)viewModel.RichTextBoxControl.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            Assert.That(brush.Color == Color.FromRgb(r, g, b));
        }

        [TestCase("Times New Roman")]
        [TestCase("Arial")]
        [TestCase("Courier New")]
        [TestCase("Georgia")]
        [TestCase("Tahoma")]
        public void CanChangeFont(string fontFamilySource)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.Font,
                TextFont = new FontFamily(fontFamilySource),
            });
            viewModel.RunMacro();

            var theFont = (FontFamily)viewModel.RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            Assert.That(theFont.Source == new FontFamily(fontFamilySource).Source);
        }

        public void CanChangeTextSize()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.TextSize,
                TextSize = 16,
            });
            viewModel.RunMacro();

            var theSize = (double)viewModel.RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            Assert.That(theSize == 16);
        }

        private void GivenFirstWordIsSelected()
        {
            viewModel.RichTextBoxControl.Document = viewModel.CurrentRichText;
            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
            editingCommandHelper.SelectRightByWord(viewModel.RichTextBoxControl);
        }

        private TextPointer GivenCaratPositionIsInTheMiddleOfTheDocument()
        {
            viewModel.CaretPosition = viewModel.CurrentRichText.Blocks.ElementAt(4).ContentStart.GetPositionAtOffset(14);
            return viewModel.CurrentRichText.Blocks.ElementAt(4).ContentStart.GetPositionAtOffset(14);
        }

        private void GivenLinesOfTextAreAddedToCurrentRichText()
        {
            viewModel.AddTextInputMacroTask("In the greenest of our valleys");
            viewModel.AddTextInputMacroTask("By good angels tenanted,");
            viewModel.AddTextInputMacroTask("Once a fair and stately palace-");
            viewModel.AddTextInputMacroTask("Radiant palace-reared its head.");
            viewModel.AddTextInputMacroTask("In the monarch Thought's dominion,");
            viewModel.AddTextInputMacroTask("It stood there!");
            viewModel.AddTextInputMacroTask("Never seraph spread a pinion");
            viewModel.AddTextInputMacroTask("Over fabric half so fair!");

            viewModel.RunMacro();

            viewModel.ClearAllTasks();
        }

        private void GivenSomeBasicTasks()
        {
            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.Italic,
            });
            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlRightArrow);
            viewModel.AddTextInputMacroTask("Hello world!");
        }

        private void ViewModel_PropertyChanged(string propertyName)
        {
            propertyChangedText = propertyName;
        }
    }
}
