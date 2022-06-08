using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;
using LiveCharts;
using LiveCharts.Wpf;
using Accord.Statistics.Models.Regression.Fitting;
using Accord.Math;
using Accord.Math.Optimization;


namespace OPV_Helper
{
    class ReadClass
    {
        private string[] header;
        private double[,] data;
        private int nLines;
        private int nColumns;
        double I_values;
        double V_values;
        double[] Vmax;
        double[] Power;
        double[] I;
        double[] V;
        double I_max;
        double V_max;
        double Isc;
        double Voc;
        double Maxpower;
        double FF;
        double PCE;
        double Rshunt;
        double Rseries;
        int Icounter = 0;
        int Vcounter = 0;
        double P;
        double Rshunt1;
     
        

        double[,] RShuntmatrix;

        public ReadClass(Stream input)
        {
                string aux;
                string[] pieces;
                StreamReader sr = new StreamReader(input);
        
                if (!sr.EndOfStream) { sr.ReadLine(); } //skips first line

                aux = sr.ReadLine();
                nLines = 0;
                header = aux.Split('\t');


                nColumns = header.Length;
                while (sr.ReadLine() != null)
                {
                    if (aux.Length > 0) nLines++;
                }
                nLines = nLines - 1;
                data = new double[nLines, nColumns];

                Vmax = new double[nLines];
                Power = new double[nLines];
                I = new double[nLines];
                V = new double[nLines];
                sr.BaseStream.Seek(0, 0);
                if (!sr.EndOfStream) { sr.ReadLine(); }

                sr.ReadLine();

                for (int i = 0; i < nLines; i++)
                {
                    aux = sr.ReadLine();
                    pieces = aux.Split('\t');

                    for (int j = 0; j < nColumns; j++)
                    {
                        if (j == 1)
                        {
                            I_values = double.Parse(pieces[1]);
                            Isc = double.Parse(pieces[1]);
                            I_values = I_values * (-1);
                            data[i, j] = I_values;
                            I[i] = I_values;
                            if (Isc < 0)
                            {
                                Vcounter++;
                            }
                        }
                        else if (j == 0)
                        {

                            I_values = double.Parse(pieces[1]);
                            I_values = I_values * (-1);
                            V_values = double.Parse(pieces[0]);
                            P = V_values * I_values;
                            Power[i] = P;
                            data[i, j] = V_values;
                            V[i] = V_values;
                            if (Power.Max() == P)
                            {
                                V_max = V_values;
                            }
                            if (V_values < 0)
                            {
                                Icounter++;
                            }

                        }
                    }
                }

            /*
              double[] TangentCollection = new double[55];
              double[] IafterIscAvg = new double[55];
              double[] Vavg = new double[55];

             RShuntmatrix = get_Data();
             int tcounter = 0;

              for (int i = 101; i <= 154; i++)
              {
                  TangentCollection[tcounter] = (-1 / ((RShuntmatrix[i, 1] - RShuntmatrix[100, 1]) / (RShuntmatrix[i, 0] - RShuntmatrix[100, 0])));
                  tcounter++;
              }

              tcounter = 0;
              for (int i = 101; i < 154; i++) {
                  IafterIscAvg[tcounter] = RShuntmatrix[i, 1];
                  Vavg[tcounter] = RShuntmatrix[i, 0];
                      }
              double III=IafterIscAvg.Average();
              double VVV = Vavg.Average();
              Rshunt1 = (-1 / ((Isc - III)/(RShuntmatrix[100,0]-VVV)));*/
            RShuntmatrix = get_Data();
            Rshunt1 = (-1 / ((RShuntmatrix[130,1] - RShuntmatrix[100,1]) / (RShuntmatrix[130, 0] -RShuntmatrix[100,0] )));
            sr.Close();
            
            } 

        public string[] get_Header()
        {
            return header;
        }

        public double[,] get_Data()
        {

            return data;
        }

        public int get_nLines()
        {
            return nLines;
        }

        public int get_nColumns()
        {
            return nColumns;
        }

        public double[] get_Current()
        {
            return I;
        }

        public double get_Imax()
        {
            I_max = Power.Max() / V_max;
            return I_max;
        }

        public double[] get_Voltage()
        {
            return V;
        }

        public double get_Vmax()
        {

            return V_max;
        }
        public double[] Get_Power()
        {
            return Power;
        }
        public double get_maxPower()
        {
            Maxpower = Power.Max();
            return Maxpower;
        }
        public double get_Isc()
        {

            Isc = I[Icounter];
            return Isc;
        }
        public double get_Voc()
        {
            Vcounter = Vcounter - 1;
            Voc = data[Vcounter,0];
            return Voc;
        }
        public double get_FF()
        {
            FF = (V_max * I_max) / (Voc * Isc);
            return FF;
        }
        public double get_Rshunt()
        {
            double Rch = V_max / I_max;
            double Pmaxideal = Voc * Isc;
            Rshunt= -1/((I[128]-Isc) / (V[128]-V[100])) ;
            return Rshunt;
        }
        public double get_Rseries()
        {
            double Rch = V_max / I_max;
            double Pmaxideal = Voc * Isc;
            Rseries = -1 / ((I[174] - I[169]) / (Voc - V[169]));
            return Rseries;
        }
        public double get_PCE()
        {
            PCE=(Voc* Isc* FF)/1000;
            return PCE;
        }
        public double get_Rshunt1() 
        {
            
            return Rshunt1;
        }
}           
    }



