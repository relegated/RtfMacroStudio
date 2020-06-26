using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfMacroStudioViewModel.Models
{
    public class Variable
    {
        public string Name { get; set; }
        public int Value { get; set; }
        public int IncrementByValue { get; set; }
        public int PlaceValuesToFill { get; set; }
        public bool UsePlaceValues { get; set; }
    }
}
