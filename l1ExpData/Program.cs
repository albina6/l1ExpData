using System;
//using ;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
//using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace l1ExpData
{
    class Program
    {
        //const int n = 77, p = 10, x = 3, y = 2;
        
        static double [] Average(double[,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            var returnAverage = new double[y];
            double sum = 0;
            for (int i=0; i < y; i++)
            {
                for(int k = 0; k < x; k++)
                {
                    sum += matrix[k,i];
                }
                returnAverage[i] = sum / x;
                sum = 0;
            }
            return returnAverage;
        }

        static double [,] Cov(double[] averege, double[,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            var returnAverage = new double[y,y];
            double sum = 0;
            for (int i = 0; i < y; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    for (int k = 0; k < x; k++)
                    {
                        sum += (matrix[k, i] - averege[i]) * (matrix[k, j] - averege[ j]);
                    }
                    returnAverage[i, j] = sum / x;
                    sum = 0;
                }
            }
            return returnAverage;
        }

        

        static double[,] StandartMatrix(double [] dispersion , double[] averege, double[,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            var returnAverage = new double[x, y];
            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                        returnAverage[i, j] = (matrix[i,j]-averege[j])/dispersion[j];
                }
            }
            return returnAverage;
        }

        static double[,] CorrelMatrix(double [,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            var returnAverage = new double[y, y];
            double sum = 0;
            for (int i = 0; i < y; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    for (int k = 0; k < x; k++)
                    {
                        sum += matrix[k, i] * matrix[k, j];
                    }
                    returnAverage[i, j] = sum/x;
                    sum = 0;
                }
            }
            return returnAverage;
        }

        static double [] Dispersion(double [] averege,double[,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            var returnAverage = new double[ y];
            double sum = 0;
            for (int j = 0; j < y; j++)
            {
                for (int k = 0; k < x; k++)
                {
                    sum += (matrix[k, j]-averege[j])* (matrix[k, j] - averege[j]);
                }
                returnAverage[j] =sum / x;
                sum = 0;
            }
            return returnAverage;
        }

        static double [,] Significance(double compare ,double[,] correlMatrix)
        {
            int y = correlMatrix.GetLength(1);
            var signifMatrix = new double[y, y];
            double nsqr = Math.Sqrt(y * y - 2);
            double correl;
            double tcalculated;
            for (int j = 0; j < y; j++)
            {
                for (int k = 0; k < y; k++)
                {
                    correl =Math.Abs( correlMatrix[j, k]);
                    tcalculated = (correl * nsqr) / Math.Sqrt(1 - correl * correl);
                    if(tcalculated - compare < 0)
                        signifMatrix[j,k] =0 ;
                    else
                        signifMatrix[j, k] = 1;
                }
            }
            return signifMatrix;
        }

        static double [,] Transposition(double [,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            var transpMatrix = new double[y,x ];
            for(int i = 0; i < x; i++)
            {
                for (int k = 0; k < y; k++)
                {
                    transpMatrix[k, i] = matrix[i, k];
                }
            }
            return transpMatrix;
        }

        static double [,] MultipleMatrix(double [,] firstMatrix,double [,] secondMatrix)
        {
            int n;
            if (firstMatrix.GetLength(1) != secondMatrix.GetLength(0))
            {
                Console.WriteLine("I can't multiply this matrixs");
                Console.Read();
                var inv = new double[,] { { 1 }, { 1 } };
                return inv;
            }
            else
            {
                n = firstMatrix.GetLength(1);
                int m = firstMatrix.GetLength(0), k = secondMatrix.GetLength(1);
                var multipleMatrix = new double[m,k];
                double sum = 0;
                for (int i = 0; i < m; i++)
                {
                    for (int j = 0; j < k; j++)
                    {
                        for(int s = 0; s < n; s++)
                        {
                            sum += firstMatrix[i, s] * secondMatrix[s, j];
                        }
                        multipleMatrix[i, j] = sum;
                        sum = 0;
                    }
                }
                return multipleMatrix;
            }  
        }

        static double [,] MultipleMatrix(double a,double [,] matrix)
        {
            int x = matrix.GetLength(0), y = matrix.GetLength(1);
            for(int i = 0; i < x; i++)
            {
                for (int k = 0; k < y; k++)
                {
                    matrix[i,k] *= a;
                }
            }
            return matrix;
        }
        ////////////////////

        //void RegressionAnalysis(int columnY,double [,] x)
        //{
        //    int xRow = x.GetLength(0), xColumn = x.GetLength(1);
        //    var y = new double [1, xRow];
        //    for (int i = 0; i < xRow; i++)
        //    {
        //        y[1, i] = x[columnY, i];
        //        x[columnY, i] = 1;
        //    }
        //}
        //////////////////////////////////////

        static double [,] InversMatrix(double [,] xMatrix)
        {
            if (xMatrix.GetLength(0) != xMatrix.GetLength(1))
            {
                Console.WriteLine("This matrix don't have invers matrix");
                Console.Read();
                var inv = new double[,] { { 1},{ 1 } };
                return inv ;
            }
            else
            {
                int n = xMatrix.GetLength(0);
                var inversMatrix = new double[n, n];
                for (int i = 0; i < n; i++)
                {
                    inversMatrix[i, i] = 1;
                }

                for (int i = 0, k = 0; i < n; k++)
                {
                    i = k;
                    var a = xMatrix[i, k];
                    if (a != 0)
                    {
                        for (int j = 0; j < n; j++)
                        {
                                inversMatrix[i, j] /= a;
                                xMatrix[i, j] /= a;
                        }
                        i++;
                        if (i < n)
                        {
                            a = xMatrix[i, k];
                            if (a != 0)
                            {
                                for (int s = 0; s < n; s++)
                                {
                                    inversMatrix[i, s] -= inversMatrix[k, s] * a;
                                }
                                for (int j = 0; j < n; j++)//////////
                                {
                                    xMatrix[i, j] -= xMatrix[k, j] * a;
                                }
                            }
                            else
                                continue;
                        }
                    }
                    else
                    {
                        Console.Read();//////////////////////////
                    }
                   
                    
                }
                for (int k = n-1;k>0 ; k--)
                {
                    for (int i=0;i<k;i++)
                    {
                        var a = xMatrix[i, k];
                        for (int s = 0; s < n; s++)
                        {
                            inversMatrix[i, s] -= inversMatrix[k, s] * a;//////////
                        }
                        for (int j = n - 1; j >= k; j--)
                        {
                            xMatrix[i, j] -= xMatrix[k, j] * a;
                        }
                    }
                       
                }
                return inversMatrix;

            }
        }

        static void Main(string[] args)
        {
            int n = 77, p = 9, startx = 3, starty = 108;//108
            double compare = 1.9839715;

            //int n = 2, p = 2, startx = 3, starty = 2;
            //double compare = 2.3646243;
            Excel ex = new Excel(@"D:\pro\6sem\компОбрЭкспДан\DataL1.xls", 1);
            var read = ex.ReadRange(startx,starty,startx+n,starty+p);
            int lenghtX = read.GetLength(0);
            int lenghtY = read.GetLength(1);
            ////del ws 2


            ex.CreateNewSheet();

            int start = 1;
            ex.SelectWorksheet(2);
            ex.WriteToCellString(start, 1, "Average:");
            start++;
            var average = Average(read);
            ex.WriteRange(start, 2, 1 + average.Length, average);
            start += 2;

            //в идеале можно сократить DRY
            ex.WriteToCellString(start, 1, "Dispersion:");
            start++;
            var dispersion = Dispersion(average, read);
            ex.WriteRange(start, 2, 1 + dispersion.Length, dispersion);
            start += 2;

            ex.WriteToCellString(start, 1, "Sqrt Dispersion:");
            start++;
            var sqrtDispersion = Dispersion(average, read);
            for (int k = 0; k < sqrtDispersion.Length; k++)
            {
                sqrtDispersion[ k] = Math.Sqrt(sqrtDispersion[ k]);
            }

            ex.WriteRange(start, 2, 1 + sqrtDispersion.Length, sqrtDispersion);
            start += 2;

            ex.WriteToCellString(start, 1, "Cov matrix:");
            start++;
            var covMatrix = Cov(average, read);
            ex.WriteRange(start, 2, start + covMatrix.GetLength(0) - 1, 1 + covMatrix.GetLength(1), covMatrix);
            start += 1 + covMatrix.GetLength(1);

            ex.WriteToCellString(start, 1, "Standart matrix:");
            start++;
            var standartMatrix = StandartMatrix(sqrtDispersion, average, read);
            ex.WriteRange(start, 2, start + standartMatrix.GetLength(0) - 1, 1 + standartMatrix.GetLength(1), standartMatrix);
            start += lenghtX + 1;

            ex.WriteToCellString(start, 1, "Average column standart Matrix:");
            start++;
            var averageStand = Average(standartMatrix);
            ex.WriteRange(start, 2, 1 + averageStand.Length, averageStand);
            start += 2;

            ex.WriteToCellString(start, 1, "Correl matrix:");
            start++;
            var correlMatrix = CorrelMatrix(standartMatrix);
            ex.WriteRange(start, 2, start + correlMatrix.GetLength(0) - 1, 1 + correlMatrix.GetLength(1), correlMatrix);
            start += lenghtY + 1;

            ex.WriteToCellString(start, 1, "Significance");
            start++;
            var signifMatrix = Significance(compare, correlMatrix);
            ex.WriteRange(start, 2, start + lenghtY - 1, 1 + signifMatrix.GetLength(1), signifMatrix);
            start += lenghtY + 1;
            // ex.Save();


           // var invees = InversMatrix(read);


            Console.Read();
            ex.Close();
        }
    }
}
