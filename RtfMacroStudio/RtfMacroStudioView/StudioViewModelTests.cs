using Castle.Components.DictionaryAdapter;
using Moq;
using NUnit.Framework;
using RtfMacroStudioViewModel.Helpers;
using RtfMacroStudioViewModel.Interfaces;
using RtfMacroStudioViewModel.Models;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Drawing;
using System.Linq;
using System.Windows;
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
        Mock<IMacroTaskEditPresenter> mockTaskEditPresenter;

        [SetUp]
        public void Setup()
        {
            mockTaskEditPresenter = new Mock<IMacroTaskEditPresenter>();
            mockEditingCommandHelper = new Mock<IEditingCommandHelper>();
            viewModel = new StudioViewModel(mockEditingCommandHelper.Object, mockTaskEditPresenter.Object);
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
            //Assert.That(((Run)viewModel.CurrentTaskList[0].Line.Inlines.FirstInline).Text == "text to add");
            Assert.That(viewModel.CurrentTaskList[0].Line == "text to add");
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
            viewModel.AddFormatMacroTask(EFormatType.Bold.ToString());

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
            string textToAdd = "text to add";
            for (int i = 0; i < numberOfTasksToInsert; i++)
            {
                viewModel.AddTextInputMacroTask(textToAdd);
            }

            //ensure the text has not been changed yet
            Assert.That(viewModel.CurrentRichText.Blocks.Count == 0);

            viewModel.RunMacro();
            ThenTheMacroUpdatedTheTextBasedOnTheNumberOfTimesTextWasRun(numberOfTasksToInsert, textToAdd);
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
            Assert.That(viewModel.CurrentRichText.ContentStart.GetOffsetToPosition(viewModel.CaretPosition) == 0);
            
            GivenLinesOfTextAreAddedToCurrentRichText();

            Assert.That(viewModel.CurrentRichText.ContentStart.GetOffsetToPosition(viewModel.CaretPosition) == 238);
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

            viewModel.AddFormatMacroTask(EFormatType.AlignCenter.ToString());
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignCenter(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignJustify()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(EFormatType.AlignJustify.ToString());
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignJustify(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignLeft()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(EFormatType.AlignLeft.ToString());
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignLeft(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanAlignRight()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(EFormatType.AlignRight.ToString());
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.AlignRight(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanToggleBold()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(EFormatType.Bold.ToString());
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.ToggleBold(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanToggleItalic()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(EFormatType.Italic.ToString());
            viewModel.RunMacro();

            mockEditingCommandHelper.Verify(m => m.ToggleItalic(It.IsAny<RichTextBox>()), Times.Once);
        }

        [Test]
        public void CanToggleUnderline()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddFormatMacroTask(EFormatType.Underline.ToString());
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

        [TestCase(16)]
        [TestCase(10)]
        [TestCase(160)]
        [TestCase(4)]
        public void CanChangeTextSize(double size)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();

            viewModel.AddFormatMacroTask(new MacroTask()
            {
                MacroTaskType = EMacroTaskType.Format,
                FormatType = EFormatType.TextSize,
                TextSize = size,
            });
            viewModel.RunMacro();

            var theSize = (double)viewModel.RichTextBoxControl.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            Assert.That(theSize == size);
        }

        [Test]
        public void CanDisplayAvailableFonts()
        {
            Assert.That(viewModel.AvailableFonts.Count > 10);
        }

        [Test]
        public void CanRefreshCurrentBoldFormatting()
        {
            Assert.That(viewModel.CurrentBoldFlag == false);

            GivenFirstWordOfDocumentTextIsBoldAndCaretIsAtDocumentStart();

            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.CurrentBoldFlag == true);
        }

        [Test]
        public void CanRefreshCurrentItalicFormatting()
        {
            Assert.That(viewModel.CurrentItalicFlag == false);

            GivenFirstWordOfDocumentTextIsItalicAndCaretIsAtDocumentStart();

            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.CurrentItalicFlag == true);
        }

        [Test]
        public void CanRefreshCurrentUnderlineFormatting()
        {
            Assert.That(viewModel.CurrentUnderlineFlag == false);

            GivenFirstWordOfDocumentTextIsUnderlinedAndCaretIsAtDocumentStart();

            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.CurrentUnderlineFlag == true);
        }

        [Test]
        public void CanRefreshCurrentFontFormatting()
        {
            Assert.That(viewModel.SelectedFont == "Segoe UI");

            GivenFirstWordOfDocumentTextIsTimesNewRomanAndCaretIsAtDocumentStart();

            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.SelectedFont == "Times New Roman");
        }

        [Test]
        public void CanRefreshCurrentTextColorFormatting()
        {
            Assert.That(viewModel.CurrentColor == Colors.Black);

            GivenFirstWordOfDocumentTextIsRedAndCaretIsAtDocumentStart();

            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.CurrentColor == Colors.Red);
        }

        [Test]
        public void CanRefreshCurrentTextSizeFormatting()
        {
            Assert.That(viewModel.CurrentTextSize == 12);

            GivenFirstWordOfDocumentTexttIsSizeSixteenAndCaretIsAtDocumentStart();

            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.CurrentTextSize == 16);
        }

        [Test]
        public void RecordMacroStartClearsCurrentTasksAndSetsFlag()
        {
            Assert.That(viewModel.IsCurrentlyRecording == false);
            GivenLinesOfTextAreAddedToCurrentRichText();

            viewModel.RecordMacroStart();

            Assert.That(viewModel.CurrentTaskList.Count == 0);
            Assert.That(viewModel.IsCurrentlyRecording);
        }

        [Test]
        public void RecordMacroProcessesAKeyOnlyIfFlagIsSet()
        {
            viewModel.ProcessKey(Key.Left, new ModifierKeys[] { ModifierKeys.Control, ModifierKeys.Shift });

            Assert.That(viewModel.CurrentTaskList.Count == 0);
        }

        [TestCase(Key.Left, true, true)]
        [TestCase(Key.Left, false, true)]
        [TestCase(Key.Left, true, false)]
        [TestCase(Key.Left, false, false)]
        [TestCase(Key.Right, true, true)]
        [TestCase(Key.Home, false, true)]
        [TestCase(Key.End, true, false)]
        [TestCase(Key.Right, false, false)]
        [TestCase(Key.Up, true, true)]
        [TestCase(Key.Up, false, false)]
        [TestCase(Key.Down, false, false)]
        [TestCase(Key.Down, true, false)]
        public void RecordMacroProcessesAMovementKey(Key keyInput, bool withControl, bool withShift)
        {
            var modifierKeyArray = GetModifierKeyArray(withControl, withShift);
            viewModel.RecordMacroStart();

            viewModel.ProcessKey(keyInput, modifierKeyArray);

            Assert.That(viewModel.CurrentTaskList[0].MacroTaskType == EMacroTaskType.SpecialKey);
            Assert.That(viewModel.CurrentTaskList[0].SpecialKey == GetExpectedSpecialKey(keyInput, withControl, withShift));
        }

        [TestCase(Key.B, true)]
        [TestCase(Key.E, true)]
        [TestCase(Key.J, true)]
        [TestCase(Key.L, true)]
        [TestCase(Key.R, true)]
        [TestCase(Key.I, true)]
        [TestCase(Key.U, true)]
        public void RecordMacroProcessesAFormattingKey(Key keyInput, bool withControl)
        {
            var modifierKeys = GetControlOnlyModifierKeyArray(withControl);
            viewModel.RecordMacroStart();

            viewModel.ProcessKey(keyInput, modifierKeys);

            if (withControl)
            {
                Assert.That(viewModel.CurrentTaskList[0].MacroTaskType == EMacroTaskType.Format);
                Assert.That(viewModel.CurrentTaskList[0].FormatType == GetFormatType(keyInput));
            }
        }

        [Test]
        public void RecordMacroTextStartsStringCapture()
        {
            viewModel.RecordMacroStart();
            Assert.That(viewModel.IsCapturingString == false);

            viewModel.ProcessKey(Key.B, null);

            Assert.That(viewModel.IsCapturingString);
        }

        [TestCase(Key.Up, false, false, false)]
        [TestCase(Key.Home, true, true, false)]
        [TestCase(Key.B, true, false, false)]
        [TestCase(Key.B, false, true, true)]
        [TestCase(Key.B, false, false, true)]
        public void RecordMacroTextStopsStringCaptureWithFormatOrMoveKey(Key keyInput, bool withControl, bool withShift, bool expectedStringCaptureState)
        {
            viewModel.RecordMacroStart();
            viewModel.ProcessKey(Key.B, null);
            var modifierArray = GetModifierKeyArray(withControl, withShift);

            viewModel.ProcessKey(keyInput, modifierArray);

            Assert.That(viewModel.IsCapturingString == expectedStringCaptureState);
        }

        [Test]
        public void StopRecordingTurnsFlagsOff()
        {
            viewModel.RecordMacroStart();
            viewModel.ProcessKey(Key.B, null);

            viewModel.StopRecording();

            Assert.That(viewModel.IsCurrentlyRecording == false);
            Assert.That(viewModel.IsCapturingString == false);
        }

        [Test]
        public void TurningIsCapturingStringToFalseCreatesTextTask()
        {
            viewModel.RecordMacroStart();
            GivenUserTypesHotGarbageBackspaceBang();

            viewModel.StopRecording();

            Assert.That(viewModel.CurrentTaskList[0].MacroTaskType == EMacroTaskType.Text);
            Assert.That(viewModel.CurrentTaskList[0].Line == "Hot garbage!");
        }

        [Test]
        public void MacroRecordingCapturesPunctuationCharacters()
        {
            viewModel.RecordMacroStart();
            GivenUserTypesPunctuation();

            viewModel.StopRecording();

            Assert.That(viewModel.CurrentTaskList[0].Line == @"`~!@#$%^&*()-_=+[{]}\|;:'"",<.>/?");

        }

        [Test]
        public void MacroRecordingDoesNotAddEmptyStringText()
        {
            viewModel.RecordMacroStart();

            viewModel.StopRecording();

            Assert.That(viewModel.CurrentTaskList.Count == 0);
        }

        [Test]
        public void ApplyCurrentFormattingAppliesCurrentFormatting()
        {
            viewModel.CurrentBoldFlag = true;
            viewModel.CurrentColor = Colors.Red;
            viewModel.CurrentTextSize = 28;
            viewModel.CurrentUnderlineFlag = true;
            viewModel.CurrentItalicFlag = true;
            viewModel.SelectedFont = "Times New Roman";
            viewModel.CurrentTextAlignment = TextAlignment.Justify;
            //viewModel.ApplyCurrentFormatting();

            viewModel.AddTextInputMacroTask("HotGarbage");
            viewModel.RefreshCurrentFormatting();

            Assert.That(viewModel.CurrentBoldFlag == true);
            Assert.That(viewModel.CurrentColor == Colors.Red);
            Assert.That(viewModel.CurrentTextSize == 28);
            Assert.That(viewModel.CurrentUnderlineFlag == true);
            Assert.That(viewModel.CurrentItalicFlag == true);
            Assert.That(viewModel.SelectedFont == "Times New Roman");
            Assert.That(viewModel.CurrentTextAlignment == TextAlignment.Justify);
        }

        private void GivenUserTypesPunctuation()
        {
            viewModel.ProcessKey(Key.Oem3, null); //`
            viewModel.ProcessKey(Key.Oem3, new ModifierKeys[] { ModifierKeys.Shift }); //~
            viewModel.ProcessKey(Key.D1, new ModifierKeys[] { ModifierKeys.Shift }); //!
            viewModel.ProcessKey(Key.D2, new ModifierKeys[] { ModifierKeys.Shift }); //@
            viewModel.ProcessKey(Key.D3, new ModifierKeys[] { ModifierKeys.Shift }); //#
            viewModel.ProcessKey(Key.D4, new ModifierKeys[] { ModifierKeys.Shift }); //$
            viewModel.ProcessKey(Key.D5, new ModifierKeys[] { ModifierKeys.Shift }); //%
            viewModel.ProcessKey(Key.D6, new ModifierKeys[] { ModifierKeys.Shift }); //^
            viewModel.ProcessKey(Key.D7, new ModifierKeys[] { ModifierKeys.Shift }); //&
            viewModel.ProcessKey(Key.D8, new ModifierKeys[] { ModifierKeys.Shift }); //*
            viewModel.ProcessKey(Key.D9, new ModifierKeys[] { ModifierKeys.Shift }); //(
            viewModel.ProcessKey(Key.D0, new ModifierKeys[] { ModifierKeys.Shift }); //)
            viewModel.ProcessKey(Key.OemMinus, null); //-
            viewModel.ProcessKey(Key.OemMinus, new ModifierKeys[] { ModifierKeys.Shift }); //_
            viewModel.ProcessKey(Key.OemPlus, null); //=
            viewModel.ProcessKey(Key.OemPlus, new ModifierKeys[] { ModifierKeys.Shift }); //+
            viewModel.ProcessKey(Key.Oem4, null); //[
            viewModel.ProcessKey(Key.Oem4, new ModifierKeys[] { ModifierKeys.Shift }); //{
            viewModel.ProcessKey(Key.Oem6, null); //]
            viewModel.ProcessKey(Key.Oem6, new ModifierKeys[] { ModifierKeys.Shift }); //}
            viewModel.ProcessKey(Key.Oem5, null); //\
            viewModel.ProcessKey(Key.Oem5, new ModifierKeys[] { ModifierKeys.Shift }); //|
            viewModel.ProcessKey(Key.Oem1, null); //;
            viewModel.ProcessKey(Key.Oem1, new ModifierKeys[] { ModifierKeys.Shift }); //:
            viewModel.ProcessKey(Key.Oem7, null); //'
            viewModel.ProcessKey(Key.Oem7, new ModifierKeys[] { ModifierKeys.Shift }); //"
            viewModel.ProcessKey(Key.OemComma, null); //,
            viewModel.ProcessKey(Key.OemComma, new ModifierKeys[] { ModifierKeys.Shift }); //<
            viewModel.ProcessKey(Key.OemPeriod, null); //.
            viewModel.ProcessKey(Key.OemPeriod, new ModifierKeys[] { ModifierKeys.Shift }); //>
            viewModel.ProcessKey(Key.Oem2, null); ///
            viewModel.ProcessKey(Key.Oem2, new ModifierKeys[] { ModifierKeys.Shift }); //?
        }

        private void GivenUserTypesHotGarbageBackspaceBang()
        {
            viewModel.ProcessKey(Key.H, new ModifierKeys[] { ModifierKeys.Shift });
            viewModel.ProcessKey(Key.O, null);
            viewModel.ProcessKey(Key.T, null);
            viewModel.ProcessKey(Key.Space, null);
            viewModel.ProcessKey(Key.G, null);
            viewModel.ProcessKey(Key.A, null);
            viewModel.ProcessKey(Key.R, null);
            viewModel.ProcessKey(Key.B, null);
            viewModel.ProcessKey(Key.A, null);
            viewModel.ProcessKey(Key.G, null);
            viewModel.ProcessKey(Key.E, null);
            viewModel.ProcessKey(Key.D1, new ModifierKeys[] { ModifierKeys.Shift });
        }

        private EFormatType GetFormatType(Key keyInput)
        {
            if (keyInput == Key.B)
            {
                return EFormatType.Bold;
            }
            if (keyInput == Key.I)
            {
                return EFormatType.Italic;
            }
            if (keyInput == Key.U)
            {
                return EFormatType.Underline;
            }
            if (keyInput == Key.E)
            {
                return EFormatType.AlignCenter;
            }
            if (keyInput == Key.J)
            {
                return EFormatType.AlignJustify;
            }
            if (keyInput == Key.L)
            {
                return EFormatType.AlignLeft;
            }
            if (keyInput == Key.R)
            {
                return EFormatType.AlignRight;
            }
            return EFormatType.Color;
        }

        private ModifierKeys[] GetControlOnlyModifierKeyArray(bool withControl)
        {
            var modifierKeys = new ModifierKeys[] { ModifierKeys.Control };
            if (withControl == false)
            {
                modifierKeys = null;
            }
            return modifierKeys;
        }

        private ModifierKeys[] GetModifierKeyArray(bool withControl, bool withShift)
        {
            if (withControl && withShift)
            {
                return new ModifierKeys[] { ModifierKeys.Control, ModifierKeys.Shift };
            }
            else if (withControl)
            {
                return new ModifierKeys[] { ModifierKeys.Control };
            }
            else if (withShift)
            {
                return new ModifierKeys[] { ModifierKeys.Shift };
            }

            return null;
        }

        private ESpecialKey GetExpectedSpecialKey(Key keyInput, bool withControl, bool withShift)
        {
            if (withControl && withShift)
            {
                if (keyInput == Key.Left)
                {
                    return ESpecialKey.ControlShiftLeftArrow;
                }
                if (keyInput == Key.Right)
                {
                    return ESpecialKey.ControlShiftRightArrow;
                }
                if (keyInput == Key.Home)
                {
                    return ESpecialKey.ControlShiftHome;
                }    
                if (keyInput == Key.End)
                {
                    return ESpecialKey.ControlShiftEnd;
                }
                if (keyInput == Key.Up)
                {
                    return ESpecialKey.UpArrow;
                }
                if (keyInput == Key.Down)
                {
                    return ESpecialKey.DownArrow;
                }
            }
            else if (withControl)
            {
                if (keyInput == Key.Left)
                {
                    return ESpecialKey.ControlLeftArrow;
                }
                if (keyInput == Key.Right)
                {
                    return ESpecialKey.ControlRightArrow;
                }
                if (keyInput == Key.Home)
                {
                    return ESpecialKey.ControlHome;
                }
                if (keyInput == Key.End)
                {
                    return ESpecialKey.ControlEnd;
                }
                if (keyInput == Key.Up)
                {
                    return ESpecialKey.UpArrow;
                }
                if (keyInput == Key.Down)
                {
                    return ESpecialKey.DownArrow;
                }
            }
            else if (withShift)
            {
                if (keyInput == Key.Left)
                {
                    return ESpecialKey.ShiftLeftArrow;
                }
                if (keyInput == Key.Right)
                {
                    return ESpecialKey.ShiftRightArrow;
                }
                if (keyInput == Key.Home)
                {
                    return ESpecialKey.ShiftHome;
                }
                if (keyInput == Key.End)
                {
                    return ESpecialKey.ShiftEnd;
                }
                if (keyInput == Key.Up)
                {
                    return ESpecialKey.UpArrow;
                }
                if (keyInput == Key.Down)
                {
                    return ESpecialKey.DownArrow;
                }
            }
            else
            {
                if (keyInput == Key.Left)
                {
                    return ESpecialKey.LeftArrow;
                }
                if (keyInput == Key.Right)
                {
                    return ESpecialKey.RightArrow;
                }
                if (keyInput == Key.Home)
                {
                    return ESpecialKey.Home;
                }
                if (keyInput == Key.End)
                {
                    return ESpecialKey.End;
                }
                if (keyInput == Key.Up)
                {
                    return ESpecialKey.UpArrow;
                }
                if (keyInput == Key.Down)
                {
                    return ESpecialKey.DownArrow;
                }
            }
            return ESpecialKey.Backspace;
        }

        private void GivenFirstWordOfDocumentTexttIsSizeSixteenAndCaretIsAtDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();
            viewModel.RichTextBoxControl.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, "16");
            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
        }

        private void GivenFirstWordOfDocumentTextIsRedAndCaretIsAtDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();
            viewModel.RichTextBoxControl.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, "Red");
            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
        }

        private void GivenFirstWordOfDocumentTextIsTimesNewRomanAndCaretIsAtDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();
            viewModel.RichTextBoxControl.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, "Times New Roman");
            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
        }

        private void GivenFirstWordOfDocumentTextIsUnderlinedAndCaretIsAtDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();
            var selectionRange = new TextRange(viewModel.RichTextBoxControl.Selection.Start, viewModel.RichTextBoxControl.Selection.End);
            selectionRange.ApplyPropertyValue(TextBlock.TextDecorationsProperty,
                new TextDecorationCollection()
                { new TextDecoration()
                    {
                        Location = TextDecorationLocation.Underline,
                    }
                });

            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
        }

        private void GivenFirstWordOfDocumentTextIsBoldAndCaretIsAtDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();
            viewModel.RichTextBoxControl.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, "Bold");
            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
        }

        private void GivenFirstWordOfDocumentTextIsItalicAndCaretIsAtDocumentStart()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            GivenFirstWordIsSelected();
            viewModel.RichTextBoxControl.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, "Italic");
            EditingCommandHelper editingCommandHelper = new EditingCommandHelper();
            editingCommandHelper.MoveToDocumentStart(viewModel.RichTextBoxControl);
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
            viewModel.CaretPosition = viewModel.CurrentRichText.Blocks.ElementAt(0).ContentStart.GetPositionAtOffset(104);
            return viewModel.CaretPosition;
        }

        private void GivenLinesOfTextAreAddedToCurrentRichText()
        {
            viewModel.AddTextInputMacroTask("In the greenest of our valleys" + Environment.NewLine);
            viewModel.AddTextInputMacroTask("By good angels tenanted," + Environment.NewLine);
            viewModel.AddTextInputMacroTask("Once a fair and stately palace-" + Environment.NewLine);
            viewModel.AddTextInputMacroTask("Radiant palace-reared its head." + Environment.NewLine);
            viewModel.AddTextInputMacroTask("In the monarch Thought's dominion," + Environment.NewLine);
            viewModel.AddTextInputMacroTask("It stood there!" + Environment.NewLine);
            viewModel.AddTextInputMacroTask("Never seraph spread a pinion" + Environment.NewLine);
            viewModel.AddTextInputMacroTask("Over fabric half so fair!" + Environment.NewLine);

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

        private void ThenTheMacroUpdatedTheTextBasedOnTheNumberOfTimesTextWasRun(int numberOfTasksToInsert, string textToAdd)
        {
            var textLength = 0;
            foreach (Paragraph paragraph in viewModel.CurrentRichText.Blocks)
            {
                foreach (Run inline in paragraph.Inlines)
                {
                    textLength += inline.Text.Length;
                }
            }

            Assert.That(textLength == numberOfTasksToInsert * textToAdd.Length);
        }

        private void ViewModel_PropertyChanged(string propertyName)
        {
            propertyChangedText = propertyName;
        }
    }
}
