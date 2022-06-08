using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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

namespace OPV_Helper
{
    class XlsInputData
    {   public string FOM_Number { get; set; }
        public string Line { get; set; }
        public string partofline { get; set; }
        public string Part { get; set; }
        public string Pattern { get; set; }
        public double AA { get; set; }
        public double Isc { get; set; }
        public double Voc { get; set; }
        public double FF { get; set; }
        public double Rseries { get; set; }
        public double Rshunt { get; set; }
        public double Vmp { get; set; }
        public double Imp { get; set; }
        public double Pmp { get; set; }
        public double PCE { get; set; }
    }
}


