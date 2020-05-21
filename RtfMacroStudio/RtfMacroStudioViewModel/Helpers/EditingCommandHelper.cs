using RtfMacroStudioViewModel.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;

namespace RtfMacroStudioViewModel.Helpers
{
    public class EditingCommandHelper : IEditingCommandHelper
    {
        public void Backspace(RichTextBox rtb)
        {
            EditingCommands.Backspace.Execute(null, rtb);
        }

        public void Delete(RichTextBox rtb)
        {
            EditingCommands.Delete.Execute(null, rtb);
        }

        public void EnterParagraphBreak(RichTextBox rtb)
        {
            EditingCommands.EnterParagraphBreak.Execute(null, rtb);
        }

        public void MoveDownByLine(RichTextBox rtb)
        {
            EditingCommands.MoveDownByLine.Execute(null, rtb);
        }

        public void MoveLeftByCharacter(RichTextBox rtb)
        {
            EditingCommands.MoveLeftByCharacter.Execute(null, rtb);
        }

        public void MoveLeftByWord(RichTextBox rtb)
        {
            EditingCommands.MoveLeftByWord.Execute(null, rtb);
        }

        public void MoveRightByCharacter(RichTextBox rtb)
        {
            EditingCommands.MoveRightByCharacter.Execute(null, rtb);
        }

        public void MoveRightByWord(RichTextBox rtb)
        {
            EditingCommands.MoveRightByWord.Execute(null, rtb);
        }

        public void MoveToDocumentEnd(RichTextBox rtb)
        {
            EditingCommands.MoveToDocumentEnd.Execute(null, rtb);
        }

        public void MoveToDocumentStart(RichTextBox rtb)
        {
            EditingCommands.MoveToDocumentStart.Execute(null, rtb);
        }

        public void MoveToLineEnd(RichTextBox rtb)
        {
            EditingCommands.MoveToLineEnd.Execute(null, rtb);
        }

        public void MoveToLineStart(RichTextBox rtb)
        {
            EditingCommands.MoveToLineStart.Execute(null, rtb);
        }

        public void MoveUpByLine(RichTextBox rtb)
        {
            EditingCommands.MoveUpByLine.Execute(null, rtb);
        }

        public void SelectDownByLine(RichTextBox rtb)
        {
            EditingCommands.SelectDownByLine.Execute(null, rtb);
        }

        public void SelectLeftByCharacter(RichTextBox rtb)
        {
            EditingCommands.SelectLeftByCharacter.Execute(null, rtb);
        }

        public void SelectLeftByWord(RichTextBox rtb)
        {
            EditingCommands.SelectLeftByWord.Execute(null, rtb);
        }

        public void SelectRightByCharacter(RichTextBox rtb)
        {
            EditingCommands.SelectRightByCharacter.Execute(null, rtb);
        }

        public void SelectRightByWord(RichTextBox rtb)
        {
            EditingCommands.SelectRightByWord.Execute(null, rtb);
        }

        public void SelectToDocumentEnd(RichTextBox rtb)
        {
            EditingCommands.SelectToDocumentEnd.Execute(null, rtb);
        }

        public void SelectToDocumentStart(RichTextBox rtb)
        {
            EditingCommands.SelectToDocumentStart.Execute(null, rtb);
        }

        public void SelectToLineEnd(RichTextBox rtb)
        {
            EditingCommands.SelectToLineEnd.Execute(null, rtb);
        }

        public void SelectToLineStart(RichTextBox rtb)
        {
            EditingCommands.SelectToLineStart.Execute(null, rtb);
        }

        public void SelectUpByLine(RichTextBox rtb)
        {
            EditingCommands.SelectUpByLine.Execute(null, rtb);
        }
    }
}
