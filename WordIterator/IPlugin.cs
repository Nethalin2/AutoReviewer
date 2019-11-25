using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;


namespace WordIterator
{
    interface IPlugin
    {
        void ToggleRun(bool shouldPluginRun);
        void ToggleComment(bool shouldPluginAddComments);
        void ToggleHighlight(bool shouldPluginHighlightText);
        void ToggleReport(bool shouldPluginAddToReport);
        void Run(Document document);
    }
}
