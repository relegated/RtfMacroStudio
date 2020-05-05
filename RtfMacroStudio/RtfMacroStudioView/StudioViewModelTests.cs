using NUnit.Framework;
using RtfMacroStudioViewModel.ViewModel;
using System;
using System.Windows.Documents;

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
    }
}
