using NUnit.Framework;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Windows.Documents;
using System.Windows.Input;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioView
{
    [TestFixture]
    public class StudioViewModelTests
    {
        StudioViewModel viewModel;

        [SetUp]
        public void Setup()
        {
            viewModel = new StudioViewModel();
        }

        [Test]
        public void CanCreateViewModel()
        {
            Assert.That(viewModel != null);
        }

        [Test]
        public void CanAddTextInputMacroTask()
        {
            viewModel.AddTextInputMacroTask("text to add");

            Assert.That(viewModel.CurrentTaskList.Count == 1);
            Assert.That(((Run)viewModel.CurrentTaskList[0].Line.Inlines.FirstInline).Text == "text to add");
        }

        [Test]
        public void CanAddSpecialCharacterMacroTask()
        {
            viewModel.AddSpecialKeyMacroTask(Key.Home);
            viewModel.AddSpecialKeyMacroTask(Key.RightCtrl);

            Assert.That(viewModel.CurrentTaskList.Count == 2);
            Assert.That(viewModel.CurrentTaskList[0].KeyStroke == Key.Home);
            Assert.That(viewModel.CurrentTaskList[1].KeyStroke == Key.RightCtrl);
        }

        [Test]
        public void CanAddFormatMacroTask()
        {
            viewModel.AddFormatMacroTask(EFormatType.Bold);

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



        private void GivenSomeBasicTasks()
        {
            viewModel.AddFormatMacroTask(EFormatType.Italic);
            viewModel.AddSpecialKeyMacroTask(Key.LeftCtrl);
            viewModel.AddTextInputMacroTask("Hello world!");
        }
    }
}
