using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using LiveCharts;
using LiveCharts.Wpf;
using System.IO;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Collections.ObjectModel;
using Accord.Statistics.Models.Regression.Fitting;
using Accord.Math;
using Accord.Math.Optimization;
using System.Globalization;
using System.Threading;

namespace OPV_Helper
{


    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        

        ReadClass input;
        OLS slope;
       
        Stream SrPath;
        IList<XlsInputData> xlsdata = new List<XlsInputData>();
        ObservableCollection<double[,]> DeviceGraphs = new ObservableCollection<double[,]>();
        public ObservableCollection<DeviceManagerP> DeviceNames = new ObservableCollection<DeviceManagerP>();
        int counter = 0;
        
        public MainWindow()
        {

        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();

        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            
           
            double LightIntensity;
            double AAvalue;
            Microsoft.Win32.OpenFileDialog path = new Microsoft.Win32.OpenFileDialog();
            path.InitialDirectory = "c:\\";
            path.DefaultExt = ".txt";
            path.Filter = "txt Files (*.txt)|*.txt|All files (*.*)|*.*";

            Nullable<bool> result = path.ShowDialog();

            SrPath = path.OpenFile();

            string stringpath = path.ToString();
            string PathName = System.IO.Path.GetFileNameWithoutExtension(stringpath);
            
            
            DeviceNames.Add(new DeviceManagerP(){ DeviceTittle=PathName});

            DeviceManager1.ItemsSource = DeviceNames;
            input = new ReadClass(SrPath);
            slope = new OLS(input.get_Data());
            
           
            double[] Volts = input.get_Voltage();
            

           
            double[,] DeviceInfo = input.get_Data();
            DeviceGraphs.Add((DeviceInfo));
            string[] partnumber = PathName.Split('_');
           // DeviceGraphs.Add(new PlotClass(input, I_V_chart, P_V_chart));
            string partofpartnumber = partnumber[0].Substring(2);
            string Part = partofpartnumber;
            try
            {
                 AAvalue = double.Parse(AATextbox.Text);
            }
            catch (FormatException)
            {
                AAvalue = 0.00005;
            }
            double Voc = input.get_Voc();
            try
            {
                LightIntensity = double.Parse(Irradiance.Text);
            }
              
            catch(FormatException)
            {
                LightIntensity = 1000;
            }
          //  double PCE1 = (((input.get_Isc() * Voc * input.get_FF()) / 1000) / AAvalue) * 100;

            xlsdata.Add(new XlsInputData
            {
                partofline = PathName.Substring(1, 1),
                Line = PathName.Substring(0, 1),
                Part = partofpartnumber,
                Pattern = partnumber[1],
                AA = AAvalue,
                Voc = Voc,
                Isc = input.get_Isc(),
                Vmp = input.get_Vmax(),
                Pmp = input.get_maxPower(),
                Imp = input.get_Imax(),
                FF = input.get_FF(),
                Rseries = input.get_Rseries(),
                // Rshunt = input.get_Rshunt1(),
                Rshunt =input.get_Rshunt(),
                PCE = (((input.get_Isc() * input.get_Voc() * input.get_FF()) / LightIntensity) / AAvalue) * 100

            });

            var XlsData = new XlsInputData
            {
                partofline = PathName.Substring(1, 1),
                Line = PathName.Substring(0, 1),
                Part = partofpartnumber,
                Pattern = partnumber[1],
                AA = AAvalue,
                Voc = Voc,
                Isc = input.get_Isc(),
                Vmp = input.get_Vmax(),
                Pmp = input.get_maxPower(),
                Imp = input.get_Imax(),
                FF = input.get_FF(),
                Rseries = input.get_Rseries(),
                //  Rshunt = input.get_Rshunt1(),
                Rshunt = input.get_Rshunt(),
                PCE = (((input.get_Isc() * input.get_Voc() * input.get_FF()) / LightIntensity) / AAvalue) * 100

            };
            XLgrid.Items.Add(XlsData);
            SrPath.Close();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void MyButton_Click(object sender, RoutedEventArgs e)
        {
            PlotClass Plot_IV_Curve = new PlotClass(input, I_V_chart, P_V_chart);
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {

        }

        private void Tools_Clicked(object sender, RoutedEventArgs e)
        {

        }
        private void YBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            dt = ToDataTable(xlsdata);
            GenerateExcel(dt);


            MessageBox.Show("Exported DataGrid to Excel file created");
        }

        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "FOM_Number";
            workSheet.Cells[1, "B"] = "Line";
            workSheet.Cells[1, "C"] = "PartOfLine";
            workSheet.Cells[1, "D"] = "Part";
            workSheet.Cells[1, "E"] = "Pattern";
            workSheet.Cells[1, "F"] = "A.A.";
            workSheet.Cells[1, "G"] = "Isc";
            workSheet.Cells[1, "H"] = "Voc";
            workSheet.Cells[1, "I"] = "FF";
            workSheet.Cells[1, "J"] = "Rseries";
            workSheet.Cells[1, "K"] = "Rshunt";
            workSheet.Cells[1, "L"] = "Imp";
            workSheet.Cells[1, "M"] = "Vmp";
            workSheet.Cells[1, "N"] = "Pmp";
            workSheet.Cells[1, "O"] = "PCE %";

            

        }

        private void GenerateExcel(System.Data.DataTable DtIN)
        {
            try
            {
                var excel = new Microsoft.Office.Interop.Excel.Application();
                excel.DisplayAlerts = true;
                excel.Visible = true;
                var workBook = excel.Workbooks.Add(Type.Missing);
                var workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                workSheet.Name = "FOM_DataAnalysis";
                System.Data.DataTable tempDt = DtIN;


                workSheet.Cells.Font.Size = 11;
                int rowcount = 1;
                for (int i = 1; i <= tempDt.Columns.Count; i++) //taking care of Headers.  
                {
                    workSheet.Cells[1, i] = tempDt.Columns[i - 1].ColumnName;
                }
                foreach (System.Data.DataRow row in tempDt.Rows) //taking care of each Row  
                {
                    rowcount += 1;
                    for (int i = 0; i < tempDt.Columns.Count; i++) //taking care of each column  
                    {
                        workSheet.Cells[rowcount, i + 1] = row[i].ToString();
                    }
                }
                var cellRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowcount, tempDt.Columns.Count]];
                cellRange.EntireColumn.AutoFit();
                xlsdata.Clear();
            }
            catch (Exception)
            {
                throw;
            }
        }
        public System.Data.DataTable ToDataTable<T>(IList<T> XlsData)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            System.Data.DataTable dt = new System.Data.DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                dt.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in XlsData)
            {
                DataRow row = dt.NewRow();
                foreach (PropertyDescriptor pdt in properties)
                {
                    row[pdt.Name] = pdt.GetValue(item) ?? DBNull.Value;
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        private void AATextbox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void MenuItem_Click_6(object sender, RoutedEventArgs e)
        {



            var viewbox = new Viewbox();

            var ParentPanelCollection = (I_V_chart.Parent as System.Windows.Controls.Grid).Children as UIElementCollection;
            ParentPanelCollection.Remove(I_V_chart);


            viewbox.Child = I_V_chart;
            //viewbox.Child = P_V_chart;
            viewbox.Measure(I_V_chart.RenderSize);
            viewbox.Arrange(new Rect(new System.Windows.Point(0, 0), I_V_chart.RenderSize));
            I_V_chart.Update(true, true);
            viewbox.UpdateLayout();
            SaveToPng(I_V_chart, "chart.png");
            viewbox.Child = null;
            ParentPanelCollection.Add(I_V_chart);
        }


        private void SaveToPng(FrameworkElement visual, string fileName)
        {
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            EncodeVisual(visual, fileName, encoder);
        }
        private static void EncodeVisual(FrameworkElement visual, string fileName, BitmapEncoder encoder)
        {
            RenderTargetBitmap bitmap = new RenderTargetBitmap((int)visual.ActualWidth, (int)visual.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            BitmapFrame frame = BitmapFrame.Create(bitmap);
            encoder.Frames.Add(frame);

            // encoder.Save(stream);

            // System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
            using (Stream stream = new FileStream(@"C:\Users\OET Intern\Desktop\OPV_Simulator\OPV_Simulator\bin\Debug\chart.png", FileMode.Create))
            {
                encoder.Save(stream);
            }

        }

        private void DeviceSelection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            PlotClass Plot_IV_Curve = new PlotClass(input, I_V_chart, P_V_chart);
        }

        private void btnShowSelectedItem_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnSelectLast_Click(object sender, RoutedEventArgs e)
        {
            DeviceManager1.SelectedIndex = DeviceManager1.SelectedIndex + 1;
        }

        private void btnSelectNext_Click(object sender, RoutedEventArgs e)
        {
            DeviceManager1.SelectedIndex = DeviceManager1.SelectedIndex - 1;
        }


        public class DeviceManagerP
        {
            public string DeviceTittle { get; set; }
        }

        private void DeviceManager1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            counter++;
           
            DeviceManager1.SelectedIndex = DeviceManager1.SelectedIndex + counter;

        }

        private void btnEnter_Click(object sender, RoutedEventArgs e)
        {
            int counter = DeviceManager1.SelectedIndex;
            double[,] selecteditem =  DeviceGraphs[counter];

           // ListPlotter selectedgraph = new ListPlotter(selecteditem, I_V_chart,P_V_chart);
            ListPlotter selectedgraph = new ListPlotter(selecteditem, I_V_chart, P_V_chart);
        }

        private void MenuItem_clickcklick(object sender, RoutedEventArgs e)
        {

        }

        private void Irradiance_TextChanged(object sender, TextChangedEventArgs e)
        {
           
        }
    }
    
}








    


        

      
