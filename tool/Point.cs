using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace tool
{
    class Point
    {
        public int id { get; set; }
        public string name { get; set; }

        //Type= 0: point
        //type=1: new point
        public int type { get; set; }
        public int t { get; set; }
        public int x { get; set; }
        public int y { get; set; }
        public string latitude { get; set; }
        public string longtitude { get; set; }
        public Image image;
        public List<int> listRelative = new List<int>();

        //public static string path = @"C:\Users\giapn\Desktop";
        public static string path = @"C:\Users\hoang\Desktop\publish";



        public Point()
        {
            x = 0;
            y = 0;
            longtitude = "";
            latitude = "";
            type = 18;
        }

        public Point(int id,  string name, int type = 0)
        {
            this.id = id;
            this.name = name;
            this.type = type;
            longtitude = "";
            latitude = "";
            type = 18;
        }

        public void setXY(int x,int y)
        {
            this.x = x;
            this.y = y;
        }

        public void setGeo(string longtitude, string latitude)
        {
            this.longtitude = longtitude;
            this.latitude = latitude;
        }

        public void genImage()
        {
            //load image
            if(this.image == null)
            {
                Image imgPoint = new Image();
                Uri imaPointUri;
                if (this.type == 0)
                {
                    imaPointUri = new Uri(path + @"\point.png");
                }
                else
                {
                    imaPointUri = new Uri(path + @"\point_blue.png");
                }

                BitmapImage imagePointBitmap = new BitmapImage(imaPointUri);
                imgPoint.Source = imagePointBitmap;
                this.image = imgPoint;
            }
            else
            {
                Uri imaPointUri;
                if (this.type == 0)
                {
                    imaPointUri = new Uri(path + @"\point.png");
                }
                else
                {
                    imaPointUri = new Uri(path + @"\point_blue.png");
                }
                this.image.Source = new BitmapImage(imaPointUri);
            }

            
        }

        

        public void setRed()
        {
            Uri imaPointUri;
            imaPointUri = new Uri(path + @"\point_red.png");
            this.image.Source =  new BitmapImage(imaPointUri);
        }

        public void setYellow()
        {
            Uri imaPointUri;
            imaPointUri = new Uri(path + @"\point_yellow.png");
            this.image.Source = new BitmapImage(imaPointUri);
        }

        public int getDistance(Point p)
        {
            double result = Math.Sqrt(Math.Pow(this.x - p.x, 2) + Math.Pow(this.y - p.y, 2));
            return (int)result;
        }
    }
}
