using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace tool
{
    class Point
    {
        public int id { get; set; }
        public string name { get; set; }
        public string type { get; set; }
        public int x { get; set; }
        public int y { get; set; }
        public Image image;
        private List<int> _canGos;

        public Point()
        {
            x = 0;
            y = 0;
        }

        public Point(int id,  string name, string type = "")
        {
            this.id = id;
            this.name = name;
            this.type = type;
        }

        public void setXY(int x,int y)
        {
            this.x = x;
            this.y = y;
        }
    }
}
