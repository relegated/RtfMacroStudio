using System.Windows.Documents;
using System.Windows.Input;
using static RtfMacroStudioViewModel.Enums.Enums;
using System.Windows.Media;

namespace RtfMacroStudioViewModel.Models
{
    public class MacroTask
    {
        public EMacroTaskType MacroTaskType { get; set; }
        
        public Paragraph Line { get; set; }

        public ESpecialKey SpecialKey { get; set; }
        public EFormatType FormatType { get; set; }
        public Color TextColor { get; set; }
        public FontFamily TextFont { get; set; }
        public double TextSize { get; set; }
    }
}
