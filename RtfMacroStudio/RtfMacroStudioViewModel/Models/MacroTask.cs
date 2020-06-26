using System.Windows.Documents;
using System.Windows.Input;
using static RtfMacroStudioViewModel.Enums.Enums;
using System.Windows.Media;

namespace RtfMacroStudioViewModel.Models
{
    public class MacroTask
    {
        public int Index { get; set; }
        public EMacroTaskType MacroTaskType { get; set; }
        
        public string Line { get; set; }

        public ESpecialKey SpecialKey { get; set; }
        public EFormatType FormatType { get; set; }
        public Color TextColor { get; set; }
        public FontFamily TextFont { get; set; }
        public double TextSize { get; set; }
        public int VarValue { get; set; }
        public int VarIncrementValue { get; set; }
        public string VarName { get; set; }
        public bool VarUsePlaceValue { get; set; }
        public int VarPlaceValue { get; set; }
    }
}
