using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfMacroStudioViewModel.Enums
{
    public class Enums
    {
        public enum EMacroTaskType
        {
            Undefined=0,
            Text,
            SpecialKey,
            Format
        }

        public enum EFormatType
        {
            Undefined=0,
            Bold,
            Unbold,
            Italic,
            Unitalic,
            Font,
            Color,
        }
    }
}
