using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Questions
{
    class PrintErrorParagraphs
    {
        private XWPFParagraph para = null;
        public PrintErrorParagraphs(XWPFParagraph p)
        {
            para = p;
        }
        public int PrintErrorID()
        {
            return 0;
        }
    }
}



