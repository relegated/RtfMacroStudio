using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace RtfMacroStudioViewModel.Interfaces
{
    public interface IFileHelper
    {
        void LoadFile(RichTextBox richTextBox);
        void SaveFile(RichTextBox richTextBox);
    }
}
