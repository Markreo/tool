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
using System.Drawing;


namespace tool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        int _width = 0;
        int _height = 0;
        int ratio = 100;
        int _sheet;
        int _longtitude;
        string oldLongtitude = "";
        int _latitude;
        string oldLatitude = "";
        int _type;
        int _number;
        Line lineX;
        Line lineY;
        bool isClickPoint = false;

        double geo = 1;

        Point _currentPoint;

        List<Point> _listPoint { get; set; }
       

        public MainWindow()
        {
            InitializeComponent();
            _listPoint = new List<Point>();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Input input = new Input();
            //DialogResult resultInput = input.ShowDialog();
            //if(resultInput == System.Windows.Forms.DialogResult.OK)
            //{
            //    Point.path = input.path;
            //}
            

            OpenFileDialog oFile = new OpenFileDialog();
            oFile.Title = "Choose file map (Image)";
            oFile.Filter = "Image files|*.png;*.jpg";
            DialogResult result = oFile.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK || File.Exists(oFile.FileName))
            {
                ImageBrush ib = new ImageBrush();
                ib.ImageSource = new BitmapImage(new Uri(oFile.FileName, UriKind.Relative));
                System.Drawing.Image i = System.Drawing.Image.FromFile(oFile.FileName);

                //System.Windows.MessageBox.Show(i.Width.ToString() + " " + i.Height.ToString());
                imgMap.Background = ib;
                imgMap.Width = i.Width;
                imgMap.Height = i.Height;
                _width = (int)imgMap.Width;
                _height = (int)imgMap.Height;


                lineX = new Line();
                lineX.Stroke = System.Windows.Media.Brushes.OrangeRed;
                lineY = new Line();
                lineY.Stroke = System.Windows.Media.Brushes.OrangeRed;

                lineX.X1 = 0;
                lineX.X2 = _width;
                lineX.Y1 = 50;
             
                lineY.Y1 = 0;
                lineY.Y2 = _height;

                lineX.Visibility = Visibility.Hidden;
                lineY.Visibility = Visibility.Hidden;


                //lineX.StrokeThickness = 1;
                //lineX.StrokeThickness = 1;
                imgMap.Children.Add(lineX);
                imgMap.Children.Add(lineY);
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
                _longtitude = int.Parse(le._long);
                _latitude = int.Parse(le._la);

                //read file excell

                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(oFile.FileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[_sheet];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int id = 0;

                for (int i = 2; i <= rowCount; i++)
                { 
                    if (xlRange.Cells[i, _number] != null && xlRange.Cells[i, _number].Value2 != null)
                    {
                        Point p = new Point(id++, xlRange.Cells[i, _number].Value2.ToString(), 0);
                        if (xlRange.Cells[i, _longtitude] != null && xlRange.Cells[i, _longtitude].Value2 != null)
                        {
                            oldLatitude = xlRange.Cells[i, _latitude].Value2.ToString();
                            oldLongtitude = xlRange.Cells[i, _longtitude].Value2.ToString();
                        }
                        p.setGeo(oldLongtitude, oldLatitude);
                        p.t = 1;
                        p.genImage();
                        p.image.MouseLeftButtonDown += new MouseButtonEventHandler((s, eve) => point_MouseDown(s, eve, p));
                        _listPoint.Add(p);
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
                    txtName.Text = _currentPoint.name;
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
            pointer.Content = "X:" + (int)relativePosition.X*ratio/100 + "; Y:" + (int)relativePosition.Y*ratio/100;

            

            lineX.Y1 = (int)relativePosition.Y;
            lineX.Y2 = (int)relativePosition.Y;

            lineY.X1 = (int)relativePosition.X;
            lineY.X2 = (int)relativePosition.X;
        }

        private void imgMap_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            pointer.Content = "";
            lineX.Visibility = Visibility.Hidden;
            lineY.Visibility = Visibility.Hidden;
        }

        

        
        private void imgMap_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (cRelative.IsChecked == false)
            {
                if (isClickPoint == false)
                {
                    var relativePosition = e.GetPosition(imgMap);
                    //set new point
                    if (cNewPoint.IsChecked == true)
                    {
                        Point p = new Point(_listPoint.Count + 1, "new Point " + (_listPoint.Count + 1).ToString(), 1);
                        p.setXY((int)(relativePosition.X/(ratio/100.0)), (int)(relativePosition.Y/(ratio/100.0)));
                        _listPoint.Add(p);
                        txtX.Text = p.x.ToString();
                        txtY.Text = p.y.ToString();
                        txtName.Text = p.name;

                        p.genImage();
                        p.image.MouseLeftButtonDown += new MouseButtonEventHandler((s, eve) => point_MouseDown(s, eve, p));
                        System.Windows.Controls.Image imgPoint = p.image;


                        //add image to canvas

                        Canvas.SetTop(imgPoint, relativePosition.Y - 8);
                        Canvas.SetLeft(imgPoint, relativePosition.X - 8);

                        cbPoints.Items.Add(p);
                        cbPoints.SelectedItem = p;
                        _currentPoint = p;
                        this.imgMap.Children.Add(imgPoint);
                    }
                    else //set X Y
                    {
                        if (_listPoint.Count != 0)
                        {
                            //todo: check here
                            _currentPoint.setXY((int)(relativePosition.X/(ratio/100.0)), (int)(relativePosition.Y/(ratio/100.0)));


                            //load image
                            System.Windows.Controls.Image imgPoint = _currentPoint.image;

                            Canvas.SetTop(imgPoint, relativePosition.Y - 8);
                            Canvas.SetLeft(imgPoint, relativePosition.X - 8);
                            if(this.imgMap.Children.IndexOf(imgPoint) == -1)
                            {
                                this.imgMap.Children.Add(imgPoint);
                                cbPoints.SelectedIndex++;
                            }

                            
                        }
                    }
                }
                
            }
            else
            {
                Console.WriteLine("check point");
            }

            isClickPoint = false;
        }

        private void cbPoints_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbPoints.SelectedItem != null)
            {
                _currentPoint = (Point)cbPoints.SelectedItem;
                txtX.Text = _currentPoint.x.ToString();
                txtY.Text = _currentPoint.y.ToString();
                txtName.Text = _currentPoint.name;
                if (cRelative.IsChecked == true)
                {
                    viewRelative();
                }
            }
            else
            {
                cbPoints.SelectedIndex = _listPoint.Count-1;
            }
            

        }

        private void btnDel_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPoint != null)
            {
                Point delPoint = _currentPoint;
                Console.WriteLine("del id" + delPoint.id);
                cbPoints.SelectedIndex++;
                this.imgMap.Children.Remove(delPoint.image);
                _listPoint.Remove(delPoint);
                Console.WriteLine("size: " + _listPoint.Count.ToString());
                cbPoints.Items.Remove(delPoint);
            }
        }


        private void viewRelative()
        {
            if (_currentPoint != null)
            {
                foreach (Point point in _listPoint)
                {
                    if (_currentPoint.listRelative.Exists(x => (x == point.id)))
                    {
                        Console.WriteLine(_currentPoint.listRelative.Find(x => (x == point.id)).ToString());
                        point.setYellow();
                    }
                    else
                    {
                        if (point.id != _currentPoint.id)
                        {
                            point.setRed();
                        }
                        else
                        {
                            point.genImage();
                        }
                    }
                }
            }
            else
            {
                System.Windows.MessageBox.Show("no current point");
            }
        }

        private void cRelative_Checked(object sender, RoutedEventArgs e)
        {
            viewRelative();
        }


        private void point_MouseDown(object sender, System.Windows.Input.MouseEventArgs e, Point p)//sender = image
        {
            isClickPoint = true;
            if(cRelative.IsChecked == true)
            {
                if(_currentPoint.listRelative.Exists(x => (x == p.id)))
                {
                    p.setRed();
                    _currentPoint.listRelative.Remove(p.id);
                }
                else
                {
                    p.setYellow();
                    _currentPoint.listRelative.Add(p.id);
                }
            }
            else
            {
                cbPoints.SelectedItem = p;
            }
            
        }

        private void cRelative_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (Point point in _listPoint)
            {
                if (point.id != _currentPoint.id)
                {
                    point.genImage();
                }
            }
        }

        private void btnLocation_Click(object sender, RoutedEventArgs e)
        {
            if(_listPoint.Count != 0)
            {
                SaveFileDialog sFile = new SaveFileDialog();
                sFile.Title = "Choose file to save (Excel)";
                sFile.Filter = "Excel Files (*.xls)|*.xls";
                DialogResult result = sFile.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    Input input = new Input();
                    DialogResult result2 = input.ShowDialog();
                    if (result2 == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            geo = Double.Parse(input.path);
                            System.Windows.MessageBox.Show(geo.ToString());
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("fail!");
                        }
                    }

                    //System.Windows.MessageBox.Show(sFile.FileName);

                    Excel.Application xlApp = new Excel.Application();

                    if (xlApp == null)
                    {
                        System.Windows.MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }


                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    Excel.Worksheet xlWorkSheet2;
                    object misValue = System.Reflection.Missing.Value;

                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    

                    //sheet 1
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    //xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.cre
                    xlWorkSheet.Name = "floy matrix";


                    xlWorkSheet.Cells[1, 1] = "Order";
                    xlWorkSheet.Cells[1, 2] = "Name";
                    xlWorkSheet.Cells[1, 3] = "Lat";
                    xlWorkSheet.Cells[1, 4] = "Long";
                    xlWorkSheet.Cells[1, 5] = "Type";
                    xlWorkSheet.Cells[1, 6] = "X";
                    xlWorkSheet.Cells[1, 7] = "Y";
                    for (int i = 2; i < _listPoint.Count + 2; i++)
                    {
                        xlWorkSheet.Cells[i, 1] = (i - 1).ToString();
                        xlWorkSheet.Cells[i, 2] = _listPoint[i - 2].name;
                        xlWorkSheet.Cells[i, 3] = _listPoint[i - 2].latitude;
                        xlWorkSheet.Cells[i, 4] = _listPoint[i - 2].longtitude;
                        xlWorkSheet.Cells[i, 5] = _listPoint[i - 2].type;
                        xlWorkSheet.Cells[i, 6] = _listPoint[i - 2].x.ToString();
                        xlWorkSheet.Cells[i, 7] = _listPoint[i - 2].y.ToString();
                    }

                    //sheet 1
                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.Add(misValue, xlWorkSheet);
                    xlWorkSheet2.Name = "map";
                    //header
                    for (int i = 2; i< _listPoint.Count + 2; i++)
                    {
                        xlWorkSheet2.Cells[1, i] = _listPoint[i - 2].name;
                    }
                    //raw
                    for (int row = 2; row < _listPoint.Count + 2; row++)
                    {
                        Point rPoint = _listPoint[row-2];
                        xlWorkSheet2.Cells[row, 1] = rPoint.name;
                        for (int i = 0; i < _listPoint.Count; i++)
                        {
                            Point iPoint = _listPoint[i];
                            if (rPoint.listRelative.Exists(x => (x == iPoint.id))){
                                xlWorkSheet2.Cells[row, i + 2] = (int)(rPoint.getDistance(_listPoint[i]) * geo);
                            }
                            else
                            {
                                xlWorkSheet2.Cells[row, i + 2] = "0";
                            }
                        }
                        pointer.Content = "Done " + (row - 2).ToString() + "/" + (_listPoint.Count - 2).ToString();
                    }



                    xlWorkBook.SaveAs(sFile.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkSheet2);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);

                    System.Windows.MessageBox.Show("Excel file created!");

                }
            }
            else //count ==0
            {
                System.Windows.MessageBox.Show("no data");
            }
        }

        private void btnIn_Click(object sender, RoutedEventArgs e)
        {
            ratio += 10;
            imgMap.Width = ratio * _width / 100;
            imgMap.Height = ratio * _height / 100;
            lRatio.Content = ratio.ToString() + "%";

            foreach(Point p in _listPoint)
            {
                System.Windows.Controls.Image imgPoint = p.image;
                Canvas.SetTop(imgPoint, (p.y * ratio / 100) - 4);
                Canvas.SetLeft(imgPoint, (p.x * ratio / 100) - 4);
            }
           
        }

        private void btnOut_Click(object sender, RoutedEventArgs e)
        {
            ratio -= 10;
            imgMap.Width = ratio * _width / 100;
            imgMap.Height = ratio * _height / 100;
            lRatio.Content = ratio.ToString() + "%";
            foreach (Point p in _listPoint)
            {
                System.Windows.Controls.Image imgPoint = p.image;
                Canvas.SetTop(imgPoint, (p.y * ratio / 100) - 4);
                Canvas.SetLeft(imgPoint, (p.x * ratio / 100) - 4);
            }
        }


        private void groupBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.W)
            {
                cbPoints.SelectedIndex--;
            }
            if (e.Key == Key.S)
            {
                cbPoints.SelectedIndex++;
            }
            if(e.Key == Key.Delete || e.Key == Key.Back)
            {
                btnDel_Click(null, null);
            }
            if(e.Key == Key.LeftCtrl)
            {
                lineX.Visibility = Visibility.Visible;
                lineY.Visibility = Visibility.Visible;
            }
        }

        private void txtName_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if(_currentPoint != null)
                {
                    _currentPoint.name = txtName.Text;
                    cbPoints.SelectedIndex++;
                }
            }
        }

        private void btnLoadMaxtrix_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFile = new OpenFileDialog();
            DialogResult result = oFile.ShowDialog();
            if(result == System.Windows.Forms.DialogResult.OK && File.Exists(oFile.FileName))
            {
                //read file excell

                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(oFile.FileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int id = 0;
                _listPoint = new List<Point>();
                while (imgMap.Children.Count != 0)
                {
                    imgMap.Children.Remove(imgMap.Children[0]);
                }
                

                imgMap.Children.Add(lineX);
                imgMap.Children.Add(lineY);

                for (int i = 2; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                    {
                        int t = 18;
                        if (xlRange.Cells[i, 5].Value2.ToString() == "1")
                        {
                            t = 1;
                        }
                        Point p = new Point(id++, xlRange.Cells[i, 2].Value2.ToString(), 0);
                        p.setGeo(xlRange.Cells[i, 4].Value2.ToString(), xlRange.Cells[i, 3].Value2.ToString());
                        
                        try
                        {
                            p.setXY(int.Parse(xlRange.Cells[i, 6].Value2.ToString()), int.Parse(xlRange.Cells[i, 7].Value2.ToString()));
                            p.t = int.Parse(xlRange.Cells[i, 5].Value2.ToString());
                        } catch(Exception ex)
                        {
                            p.t = 18;
                        }
                        p.genImage();
                        p.image.MouseLeftButtonDown += new MouseButtonEventHandler((s, eve) => point_MouseDown(s, eve, p));

                        
                        System.Windows.Controls.Image imgPoint = p.image;


                        //add image to canvas

                        Canvas.SetTop(imgPoint, p.y - 4);
                        Canvas.SetLeft(imgPoint, p.x - 4);

                        cbPoints.Items.Add(p);
                        this.imgMap.Children.Add(imgPoint);


                        _listPoint.Add(p);
                    }

                }


                

                Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
                //Excel.Range xlRange2 = xlWorksheet2.get_Range(xlWorksheet2.Cells[2,2], );
                int rang = _listPoint.Count + 2;
                for(int row=2;row< rang; row++)
                {
                    Point pRow = _listPoint[row - 2];
                    for (int col = 2; col < rang; col++)
                    {
                        if( xlWorksheet2.Cells[row, col].Value2 != 0)
                        {
                            pRow.listRelative.Add(_listPoint[col - 2].id);
                            Console.WriteLine("Add: " + xlWorksheet2.Cells[row, col].Value2.ToString());
                        }
                    }
                    Console.WriteLine("-----------------Done " + (row - 2).ToString() + "/" + (rang - 2).ToString());
                    
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
        

                if (_listPoint.Count != 0)
                {
                    _currentPoint = _listPoint[0];
                    cbPoints.SelectedItem = _currentPoint;
                    txtX.Text = _currentPoint.x.ToString();
                    txtY.Text = _currentPoint.y.ToString();
                    txtName.Text = _currentPoint.name;
                }
            }
        }

        private void imgMap_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            
        }

        private void txtX_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (_currentPoint != null)
                {
                    try
                    {
                        _currentPoint.x = int.Parse(txtX.Text);
                        _currentPoint.y = int.Parse(txtY.Text);
                        System.Windows.Controls.Image imgPoint = _currentPoint.image;
                        Canvas.SetTop(imgPoint, _currentPoint.y - 8);
                        Canvas.SetLeft(imgPoint, _currentPoint.x - 8);
                    } catch(Exception ex)
                    {

                    }
                }
            }
        }

        private void groupBox_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl)
            {
                lineX.Visibility = Visibility.Hidden;
                lineY.Visibility = Visibility.Hidden;
            }
        }
    }
}

