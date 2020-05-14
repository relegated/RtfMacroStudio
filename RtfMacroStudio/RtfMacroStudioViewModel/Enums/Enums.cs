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
        public enum EMacroTaskType
        {
            Undefined=0,
            Text,
            SpecialKey,
            Format
        }

        public enum EFormatType
        {
            Bold,
            Unbold,
            Italic,
            Unitalic,
            Font,
            Color,
        }

        public enum ESpecialKey
        {
            LeftArrow,
            RightArrow,
            UpArrow,
            DownArrow,
            Home,
            End,
            Delete,
            Backspace,
            Enter,
            ShiftLeftArrow,
            ShiftRightArrow,
            ShiftUpArrow,
            ShiftDownArrow,
            ControlLeftArrow,
            ControlRightArrow,
            ControlUpArrow,
            ControlDownArrow,
            ControlShiftLeftArrow,
            ControlShiftRightArrow,
            ControlShiftUpArrow,
            ControlShiftDownArrow,
        }
    }

    public static class RtfEnumStringRetriever
    {
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
                case ESpecialKey.ControlUpArrow:
                    return "Control + Up Arrow";
                case ESpecialKey.ControlDownArrow:
                    return "Control + Down Arrow";
                case ESpecialKey.ControlShiftLeftArrow:
                    return "Control + Shift + Left Arrow";
                case ESpecialKey.ControlShiftRightArrow:
                    return "Control + Shift + Right Arrow";
                case ESpecialKey.ControlShiftUpArrow:
                    return "Control + Shift + Up Arrow";
                case ESpecialKey.ControlShiftDownArrow:
                    return "Control + Shift + Down Arrow";
                default:
                    return string.Empty;
            }
        }
    }
}
