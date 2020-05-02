using NUnit.Framework;
using RtfMacroStudioViewModel.ViewModel;
using System;


namespace RtfMacroStudioView
{
    [TestFixture]
    public class StudioViewModelTests
    {
        [Test]
        public void CanCreateViewModel()
        {
            var viewModel = new StudioViewModel();

            Assert.That(viewModel != null);
        }
    }
}
