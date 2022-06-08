using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OPV_Helper
{
    class OLS
    {
        double[] X = new double[44];
        double[] Y = new double[44];
        double Xavg;
        double Yavg;
        double SumXY=0;
        double SumXXavg=0;
        double slope ;
        double xyavg;
        double yxavg;
        double[] Xsquared = new double[44];
        double[] Ysquared = new double[44];
        double Xsquaredavg;
        double Ysquaredavg;
        double sumX;
        double sumY;
        double sumXYproduct;
        double sumXsquared;
        

        public OLS(double[,] datainput)
        {

            int counter = 0;

            for (int i = 100; i < 144; i++)
            {
                X[counter] = datainput[i, 0];
                sumX += X[counter];
                Y[counter] = datainput[i, 1];
                sumY += Y[counter];
                counter++;

            }
            Xavg = X.Average();
            Yavg = Y.Average();

            counter = 0;
            for (int i = 100; i < 144; i++)
            {
               // slope = ((X[counter] - Xavg) * (Y[counter] * Yavg))/ ((Math.Pow(X[counter], 2) - Math.Pow(Xavg, 2)));
                SumXY += ((X[counter] - Xavg) * (Y[counter] - Yavg));
                SumXXavg += (Math.Pow(X[counter], 2))- Math.Pow(Xavg, 2);
                Xsquared[counter] = Math.Pow(datainput[i, 0], 2);
                sumXsquared += Xsquared[counter];
                Ysquared[counter] = Math.Pow(datainput[i, 1], 2);
                xyavg += X[counter] * Y[counter];
                //slope =-1/(SumXY/SumXXavg);
                counter++;
            }
            sumXYproduct = (sumX * sumY) / 44;
            Xsquaredavg = Xsquared.Average();
            Ysquaredavg = Ysquared.Average();
            xyavg = xyavg/44 ;
            yxavg = Xavg * Yavg;
            // double sqrofslope = Math.Sqrt((Xsquaredavg - Math.Pow(Xavg, 2))) * Math.Sqrt((Ysquaredavg - Math.Pow(Yavg, 2)));
            //double sqrofslope = xyavg / Xsquaredavg;
             slope = ((sumXYproduct)-(SumXY)) / (sumXsquared - ( (Math.Pow(sumX, 2)/40)));
             slope = 1 / slope;
            //slope = 10/sqrofslope;
        }
       public double get_slope()
        {
            return slope;
        }

    }
}
