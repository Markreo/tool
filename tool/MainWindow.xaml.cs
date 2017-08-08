using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;


namespace tool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        int _width = 0;
        int _height = 0;
        int _sheet;
        int _type;
        int _number;


        Point _currentPoint;

        List<Point> _listPoint { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            _listPoint = new List<Point>();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFile = new OpenFileDialog();
            oFile.Title = "Choose file map (Image)";
            oFile.Filter = "Image files|*.png;*.jpg";
            DialogResult result = oFile.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK || File.Exists(oFile.FileName))
            {
                ImageBrush ib = new ImageBrush();
                ib.ImageSource = new BitmapImage(new Uri(oFile.FileName, UriKind.Absolute));
                imgMap.Background = ib;
                imgMap.Width = ib.ImageSource.Width;
                imgMap.Height = ib.ImageSource.Height;
                _width = (int)imgMap.Width;
                _height = (int)imgMap.Height;
            }
        }


        //load excell file
        private void btnLoadPoint_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFile = new OpenFileDialog();
            oFile.Title = "Choose file to load (Excel)";
            oFile.Filter = "Excel Files|*.xls;*.xlsx";
            DialogResult result = oFile.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK || File.Exists(oFile.FileName))
            {
                loadExcell le = new loadExcell(oFile.FileName);
                le.ShowDialog();
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                _sheet = int.Parse(le._sheet);
                Console.WriteLine(le._type);
                _type = int.Parse(le._type);
                _number = int.Parse(le._number);

                //read file excell

                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(oFile.FileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[_sheet];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int id = 0;

                for (int i = 1; i <= rowCount; i++)
                { 
                    if (xlRange.Cells[i, _number] != null && xlRange.Cells[i, _number].Value2 != null)
                    { 
                        if (xlRange.Cells[i, _type] != null && xlRange.Cells[i, _type].Value2 != null)
                        {
                            _listPoint.Add(new Point(id++, xlRange.Cells[i, _number].Value2.ToString(), xlRange.Cells[i, _type].Value2.ToString()));
                        }
                        else
                        {
                            _listPoint.Add(new Point(id++, xlRange.Cells[i, _number].Value2.ToString()));
                        }
                                                  
                    }
                  
                }



                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.MessageBox.Show("get " + _listPoint.Count.ToString() + " point!");
                foreach(Point point in _listPoint)
                {
                    cbPoints.Items.Add(point);
                }

                if (_listPoint.Count != 0)
                { 
                    _currentPoint = _listPoint[0];
                    cbPoints.SelectedItem = _currentPoint;
                    txtX.Text = _currentPoint.x.ToString();
                    txtY.Text = _currentPoint.y.ToString();
                }


            }
            else
            {
                System.Windows.MessageBox.Show("load file excell fail!");
            }
        }

        private void imgMap_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            //Point p = Mouse.GetPosition(canvas);
            var relativePosition = e.GetPosition(imgMap);
            //var point = PointToScreen(relativePosition);
            pointer.Content = "X:" + (int)relativePosition.X + "; Y:" + (int)relativePosition.Y;
        }

        private void imgMap_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            pointer.Content = "";
        }

        private void imgMap_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var relativePosition = e.GetPosition(imgMap);
            if (cNewPoint.IsChecked == true)
            {
                _currentPoint = new Point(_listPoint.Count + 1, "new Point " + (_listPoint.Count + 1).ToString());
                _currentPoint.setXY((int)relativePosition.X, (int)relativePosition.Y);
                _listPoint.Add(_currentPoint);
                txtX.Text = _currentPoint.x.ToString();
                txtY.Text = _currentPoint.y.ToString();

                     //load image
                Image imgPoint = new Image();
                Uri imaPointUri = new Uri(@"C:\Users\giapn\Desktop\point_blue.png");
                BitmapImage imagePointBitmap = new BitmapImage(imaPointUri);
                imgPoint.Source = imagePointBitmap;

                Canvas.SetTop(imgPoint, _currentPoint.y - 4);
                Canvas.SetLeft(imgPoint, _currentPoint.x - 4);

                _currentPoint.image = imgPoint;
                cbPoints.Items.Add(_currentPoint);
                cbPoints.SelectedItem = _currentPoint;
                this.imgMap.Children.Add(imgPoint);
            }
            else
            {
                if (_listPoint.Count != 0)
                {
                    _listPoint.Find(x => x.id == _currentPoint.id).setXY((int)relativePosition.X, (int)relativePosition.Y);


                    //load image
                    Image imgPoint = new Image();
                    Uri imaPointUri = new Uri(@"C:\Users\giapn\Desktop\point.png");
                    BitmapImage imagePointBitmap = new BitmapImage(imaPointUri);
                    imgPoint.Source = imagePointBitmap;

                    Canvas.SetTop(imgPoint, _currentPoint.y - 4);
                    Canvas.SetLeft(imgPoint, _currentPoint.x - 4);

                    _currentPoint.image = imgPoint;
                    this.imgMap.Children.Add(imgPoint);

                    cbPoints.SelectedIndex++;
                }
            }

        }

        private void cbPoints_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _currentPoint = (Point)cbPoints.SelectedItem;
            txtX.Text = _currentPoint.x.ToString();
            txtY.Text = _currentPoint.y.ToString();
        }

        private void btnDel_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPoint != null)
            {
                this.imgMap.Children.Remove(_currentPoint.image);
                _currentPoint.x = 0;
                _currentPoint.y = 0;
                txtX.Text = "0";
                txtY.Text = "0";
            }
        }
    }
}
