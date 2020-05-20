using NUnit.Framework;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Linq;
using System.Windows.Documents;
using System.Windows.Input;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioView
{
    [TestFixture, RequiresSTA]
    public class StudioViewModelTests
    {
        StudioViewModel viewModel;
        string propertyChangedText;

        [SetUp]
        public void Setup()
        {
            viewModel = new StudioViewModel();
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
            viewModel.AddFormatMacroTask(EFormatType.Bold);

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
            Assert.That(viewModel.SupportedSpecialKeys.Count == 25);
        }

        [Test]
        public void TextInputChangesCaratPosition()
        {
            GivenLinesOfTextAreAddedToCurrentRichText();

            Assert.That(viewModel.CaretPosition.CompareTo(viewModel.CurrentRichText.ContentEnd) == 0);
        }

        [TestCase(ESpecialKey.LeftArrow)]
        [TestCase(ESpecialKey.RightArrow)]
        [TestCase(ESpecialKey.ShiftLeftArrow)]
        [TestCase(ESpecialKey.ShiftRightArrow)]
        public void LeftRightSingleCharacterNavigationKeysChangeCaretLocation(ESpecialKey specialKey)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(specialKey);
            viewModel.RunMacro();

            if (specialKey == ESpecialKey.LeftArrow || specialKey == ESpecialKey.ShiftLeftArrow)
            {
                Assert.That(viewModel.CaretPosition.CompareTo(startPosition.GetPositionAtOffset(-1)) == 0);
            }
            else
            {
                Assert.That(viewModel.CaretPosition.CompareTo(startPosition.GetPositionAtOffset(1)) == 0);
            }
        }

        [TestCase(ESpecialKey.UpArrow)]
        [TestCase(ESpecialKey.DownArrow)]
        [TestCase(ESpecialKey.ShiftUpArrow)]
        [TestCase(ESpecialKey.ShiftDownArrow)]
        public void UpDownLineNavigationChangesCaretLocation(ESpecialKey specialKey)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(specialKey);
            viewModel.RunMacro();

            if (specialKey == ESpecialKey.UpArrow || specialKey == ESpecialKey.ShiftUpArrow)
            {
                Assert.That(viewModel.CaretPosition.CompareTo(
                    viewModel.CurrentRichText.Blocks.ElementAt(3).ContentStart.GetPositionAtOffset(14)) == 0);
            }
            else
            {
                Assert.That(viewModel.CaretPosition.CompareTo(
                    viewModel.CurrentRichText.Blocks.ElementAt(5).ContentStart.GetPositionAtOffset(14)) == 0);
            }
        }

        [TestCase(ESpecialKey.Home)]
        [TestCase(ESpecialKey.End)]
        [TestCase(ESpecialKey.ShiftHome)]
        [TestCase(ESpecialKey.ShiftEnd)]
        public void HomeEndNavigationChangesCaretLocation(ESpecialKey specialKey)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(specialKey);
            viewModel.RunMacro();

            if (specialKey == ESpecialKey.Home || specialKey == ESpecialKey.ShiftHome)
            {
                Assert.That(viewModel.CaretPosition.CompareTo(
                    viewModel.CurrentRichText.Blocks.ElementAt(4).ContentStart) == 0);
            }
            else
            {
                Assert.That(viewModel.CaretPosition.CompareTo(
                    viewModel.CurrentRichText.Blocks.ElementAt(4).ContentEnd) == 0);
            }
        }

        [TestCase(ESpecialKey.ControlLeftArrow)]
        [TestCase(ESpecialKey.ControlRightArrow)]
        public void ControlLeftRightArrowsChangeCaretLocationToNextOrPreviousWord(ESpecialKey specialKey)
        {
            GivenLinesOfTextAreAddedToCurrentRichText();
            var startPosition = GivenCaratPositionIsInTheMiddleOfTheDocument();

            viewModel.AddSpecialKeyMacroTask(specialKey);
            viewModel.RunMacro();

            if (specialKey == ESpecialKey.ControlLeftArrow)
            {
                Assert.That(viewModel.CaretPosition.CompareTo(
                    viewModel.CurrentRichText.Blocks.ElementAt(4).ContentStart.GetPositionAtOffset(8)) == 0);
            }
            else
            {
                Assert.That(viewModel.CaretPosition.CompareTo(
                    viewModel.CurrentRichText.Blocks.ElementAt(4).ContentStart.GetPositionAtOffset(16)) == 0);
            }
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
            viewModel.AddFormatMacroTask(EFormatType.Italic);
            viewModel.AddSpecialKeyMacroTask(ESpecialKey.ControlRightArrow);
            viewModel.AddTextInputMacroTask("Hello world!");
        }

        private void ViewModel_PropertyChanged(string propertyName)
        {
            propertyChangedText = propertyName;
        }
    }
}
