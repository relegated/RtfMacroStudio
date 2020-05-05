using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace RtfMacroStudioViewModel.Models
{
    public class MacroTask
    {
        //will be enum later
        public int MacroTaskType { get; set; }
        
        public Paragraph Line { get; set; }
    }
}
