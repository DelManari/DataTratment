using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportData
{
    class elements
    {
        public int id;
        public int val;
        public Color c;
        public elements(int id , int val,Color c)
        {
            this.id = id;
            this.val = val;
            this.c = c;
        }
    }
}
