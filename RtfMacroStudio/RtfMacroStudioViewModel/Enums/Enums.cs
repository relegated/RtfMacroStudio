using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static RtfMacroStudioViewModel.Enums.Enums;

namespace RtfMacroStudioViewModel.Enums
{
    public class Enums
    {
        public enum ERunPresenterOption
        {
            Once,
            NTimes,
            End
        }

        public enum EMacroTaskType
        {
            Undefined=0,
            Text,
            SpecialKey,
            Format,
            Variable
        }

        public enum EFormatType
        {
            Bold,
            Italic,
            Underline,
            Font,
            Color,
            TextSize,
            AlignCenter,
            AlignJustify,
            AlignLeft,
            AlignRight,
        }

        public enum ESpecialKey
        {
            LeftArrow,
            RightArrow,
            UpArrow,
            DownArrow,
            Home,
            End,
            ControlHome,
            ControlEnd,
            ShiftHome,
            ShiftEnd,
            ControlShiftHome,
            ControlShiftEnd,
            Delete,
            Backspace,
            Enter,
            ShiftLeftArrow,
            ShiftRightArrow,
            ShiftUpArrow,
            ShiftDownArrow,
            ControlLeftArrow,
            ControlRightArrow,
            ControlShiftLeftArrow,
            ControlShiftRightArrow,
            Cut,
            Copy,
            Paste,
        }
    }

    public static class RtfEnumStringRetriever
    {
        public static string GetFriendlyString(this EFormatType type)
        {
            switch (type)
            {
                case EFormatType.Bold:
                    return "Bold";
                case EFormatType.Italic:
                    return "Italic";
                case EFormatType.Underline:
                    return "Underline";
                case EFormatType.Font:
                    return "Font";
                case EFormatType.Color:
                    return "Color";
                case EFormatType.TextSize:
                    return "Size";
                case EFormatType.AlignCenter:
                    return "Align Center";
                case EFormatType.AlignJustify:
                    return "Align Justify";
                case EFormatType.AlignLeft:
                    return "Align Left";
                case EFormatType.AlignRight:
                    return "Align Right";
                default:
                    return string.Empty;
            }
        }

        public static string GetFriendlyString(this EMacroTaskType type)
        {
            switch (type)
            {
                case EMacroTaskType.Undefined:
                    return "Undefined";
                case EMacroTaskType.Text:
                    return "Text";
                case EMacroTaskType.SpecialKey:
                    return "Special Key";
                case EMacroTaskType.Format:
                    return "Format";
                case EMacroTaskType.Variable:
                    return "Variable";
                default:
                    return string.Empty;
            }
        }

        public static string GetFriendlyString(this ESpecialKey specialKey)
        {
            switch (specialKey)
            {
                case ESpecialKey.LeftArrow:
                    return "Left Arrow";
                case ESpecialKey.RightArrow:
                    return "Right Arrow";
                case ESpecialKey.UpArrow:
                    return "Up Arrow";
                case ESpecialKey.DownArrow:
                    return "Down Arrow";
                case ESpecialKey.Home:
                    return "Home";
                case ESpecialKey.End:
                    return "End";
                case ESpecialKey.ControlHome:
                    return "Control + Home";
                case ESpecialKey.ControlEnd:
                    return "Control + End";
                case ESpecialKey.ShiftHome:
                    return "Shift + Home";
                case ESpecialKey.ShiftEnd:
                    return "Shift + End";
                case ESpecialKey.ControlShiftHome:
                    return "Control + Shift + Home";
                case ESpecialKey.ControlShiftEnd:
                    return "Control + Shift + End";
                case ESpecialKey.Delete:
                    return "Delete";
                case ESpecialKey.Backspace:
                    return "Backspace";
                case ESpecialKey.Enter:
                    return "Enter";
                case ESpecialKey.ShiftLeftArrow:
                    return "Shift + Left Arrow";
                case ESpecialKey.ShiftRightArrow:
                    return "Shift + Right Arrow";
                case ESpecialKey.ShiftUpArrow:
                    return "Shift + Up Arrow";
                case ESpecialKey.ShiftDownArrow:
                    return "Shift + Down Arrow";
                case ESpecialKey.ControlLeftArrow:
                    return "Control + Left Arrow";
                case ESpecialKey.ControlRightArrow:
                    return "Control + Right Arrow";
                case ESpecialKey.ControlShiftLeftArrow:
                    return "Control + Shift + Left Arrow";
                case ESpecialKey.ControlShiftRightArrow:
                    return "Control + Shift + Right Arrow";
                case ESpecialKey.Copy:
                    return "Copy";
                case ESpecialKey.Cut:
                    return "Cut";
                case ESpecialKey.Paste:
                    return "Paste";
                default:
                    return string.Empty;
            }
        }
    }
}
