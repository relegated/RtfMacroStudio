using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace RtfMacroStudioViewModel.Interfaces
{
    public interface IEditingCommandHelper
    {
        void MoveLeftByCharacter(RichTextBox rtb);
        void MoveRightByCharacter(RichTextBox rtb);
        void MoveLeftByWord(RichTextBox rtb);
        void MoveRightByWord(RichTextBox rtb);
        void SelectLeftByCharacter(RichTextBox rtb);
        void SelectRightByCharacter(RichTextBox rtb);
        void SelectLeftByWord(RichTextBox rtb);
        void SelectRightByWord(RichTextBox rtb);
        void MoveUpByLine(RichTextBox rtb);
        void MoveDownByLine(RichTextBox rtb);
        void SelectUpByLine(RichTextBox rtb);
        void SelectDownByLine(RichTextBox rtb);
        void MoveToLineStart(RichTextBox rtb);
        void MoveToLineEnd(RichTextBox rtb);
        void SelectToLineStart(RichTextBox rtb);
        void SelectToLineEnd(RichTextBox rtb);
        void MoveToDocumentStart(RichTextBox rtb);
        void MoveToDocumentEnd(RichTextBox rtb);
        void SelectToDocumentStart(RichTextBox rtb);
        void SelectToDocumentEnd(RichTextBox rtb);
        void EnterParagraphBreak(RichTextBox rtb);
        void Delete(RichTextBox rtb);
        void Backspace(RichTextBox rtb);
        void AlignCenter(RichTextBox rtb);
        void AlignJustify(RichTextBox rtb);
        void AlignLeft(RichTextBox rtb);
        void AlignRight(RichTextBox rtb);
        void ToggleBold(RichTextBox rtb);
        void ToggleItalic(RichTextBox rtb);
        void ToggleUnderline(RichTextBox rtb);
    }
}
