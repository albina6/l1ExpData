﻿using System;
namespace l1ExpData
{
    class Program
    {
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

        static double Average(double[] array)
        {
            int x = array.Length;
            double average=0; 
            for (int k = 0; k < x; k++)
            {
                average+= array[k];
            }
            average /=x;
            return average;
        }

        static double [] CovY(double [] y,double [,] matrix)
        {
            int matrixRow = matrix.GetLength(0), matrixColumn = matrix.GetLength(1);
            var average = Average(matrix);
            var averageY = Average(y);

            var returnAverage = new double[ matrixColumn];
            double sum = 0;
            for (int j = 0; j < matrixColumn; j++)
            {
                for (int k = 0; k < matrixRow; k++)
                {
                    sum += (y[k] - averageY) * (matrix[k, j] - average[j]);
                }
                returnAverage[j] = sum / matrixRow;
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

        static int WriteAndStartChange(int start, string name,Excel ex, double[,] matrix)
        {
            ex.WriteToCellString(start, 1, name);
            start++;
            ex.WriteRange(start, 2, start + matrix.GetLength(0) - 1, 1 + matrix.GetLength(1), matrix);
            start += matrix.GetLength(0) + 1;
            return start;
        }

        static int WriteAndStartChange(int start, string name, Excel ex, double[] array)
        {
            ex.WriteToCellString(start, 1, name);
            start++;
            ex.WriteRange(start, 2, 1 + array.Length, array);
            start ++;
            return start;
        }

        static int WriteAndStartChange(int start, string name, Excel ex, double x)
        {
            ex.WriteToCellString(start, 1, name);
            start++;
            ex.WriteToCellString(start, 2, x);
            start++;
            return start;
        }

        static double [] MultipleMatrix(double [,] first,double [] second)
        {
            if (first.GetLength(1) != second.Length)
            {
                Console.WriteLine("I can't multiply this matrixs");
                Console.Read();
                var inv = new double[] {  1  };
                return inv;
            }
            else
            {
                int n = second.Length, m = first.GetLength(0);
                var a =new double[m];
                double sum = 0;
                for(int i = 0; i < m; i++)
                {
                    for (int r = 0; r < n; r++)
                    {
                        sum += first[i, r] * second[r];
                    }
                    a[i] = sum;
                    sum = 0;
                }
                return a;
            }

        }

        static double [,] MultipleMatrix(double [,] firstMatrix,double [,] secondMatrix)
        {
            int m;
            if (firstMatrix.GetLength(1) != secondMatrix.GetLength(0))
            {
                Console.WriteLine("I can't multiply this matrixs");
                Console.Read();
                var inv = new double[,] { { 1 }, { 1 } };
                return inv;
            }
            else
            {
                m = firstMatrix.GetLength(1);
                int p = firstMatrix.GetLength(0), n = secondMatrix.GetLength(1);
                var multipleMatrix = new double[p,n];
                double sum = 0;
                for (int j = 0; j < n; j++)
                {
                    for (int i = 0; i < p; i++)
                    {
                        for (int r = 0; r < m; r++)
                        {
                            sum += firstMatrix[i, r] * secondMatrix[r, j];
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
            var returnMatrix =new double[x,y];
            for(int i = 0; i < x; i++)
            {
                for (int k = 0; k < y; k++)
                {
                    returnMatrix[i,k] = matrix[i,k]*a;
                }
            }
            return returnMatrix;
        }

        static double [,] RegressionAnalysisMatrixX(int number, double[,] matrix)
        {
            int xRow = matrix.GetLength(0), xColumn = matrix.GetLength(1);
            int count = 0;
            var y = new double[xRow];
            for (int i = 0; i < xRow; i++)
            {
                y[i] = matrix[i, number];
            }
            var covY = CovY(y, matrix);
            var matX = new double[xRow, xColumn];
            for (int i = 0; i < xColumn;count++, i++)
            {
                if ((Math.Abs(covY[i]) > 0.3)&& i!= number)
                {
                    for(int k = 0; k < xRow; k++)
                    {
                        matX[k,count] = matrix[ k,i];
                    }
                }
                else
                    count--;
            }
            count++;
            var x =new double[xRow, count];
            for (int i = 0; i < xRow; i++)
            {
                for (int k = 0; k < count-1; k++)
                {
                    x[i, k] = matX[i, k];
                }
            }
            for (int i = 0,k=count -1; i < xRow; i++)
            {
                x[i, k] = 1;
            }
            return x;
        }

        static double [,] InversMatrix( double [,] xMatrix)
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
                var matX =new double[n, n];
                for (int i = 0; i < n; i++)
                {
                    for(int k = 0; k < n; k++)
                    {
                        matX[i, k] = xMatrix[i, k];
                    }
                }
                var inversMatrix = new double[n, n];
                for (int i = 0; i < n; i++)
                {
                    inversMatrix[i, i] = 1;
                }
                for (int k = 0; k < n; k++)
                {
                    var a = matX[k, k];
                    if (a != 0)
                    {
                        for (int j = 0; j < n; j++)
                        {
                                inversMatrix[k, j] /= a;
                                matX[k, j] /= a;
                        }
                        if (k < n-1)
                        {
                            for(int p = k+1; p < n; p++)
                            {
                                a = matX[p, k];
                                if (a != 0)
                                {
                                    for (int s = 0; s < n; s++)
                                    {
                                        inversMatrix[p, s] -= inversMatrix[k, s] * a;
                                    }
                                    for (int j = 0; j < n; j++)
                                    {
                                        matX[p, j] -= matX[k, j] * a;
                                    }
                                }
                                else
                                    continue;
                            }
                            
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
                        var a = matX[i, k];
                        for (int s = 0; s < n; s++)
                        {
                            inversMatrix[i, s] -= inversMatrix[k, s] * a;
                        }
                        for (int j = n - 1; j >= k; j--)
                        {
                            matX[i, j] -= matX[k, j] * a;
                        }
                    }
                       
                }
                return inversMatrix;

            }
        }

        static double[] CoefficientLinRegression(double[] y, double[,] xMatrix)
        {
            var xTransposition = Transposition(xMatrix);
            var xMultiple = MultipleMatrix(xTransposition, xMatrix);
            var xInvers = InversMatrix(xMultiple);
            var xMulXT = MultipleMatrix(xInvers, xTransposition);
            var a = MultipleMatrix(xMulXT, y);
            return a;
        }

        static double CoeffDeterm(double yAverage, double[] y, double[]yCalculation)
        {
            int yLengh = y.Length;
            double numerator=0, denominator=0;
            for(int i = 0; i < yLengh; i++)
            {
                numerator += (yCalculation[i] - yAverage) * (yCalculation[i] - yAverage);
                denominator += (y[i] - yAverage) * (y[i] - yAverage);
            }
            return numerator / denominator;
        }

        static void Main(string[] args)
        {
            int n = 77, p = 9, startx = 3, starty = 108;//106
            double compare = 1.9839715;

            //int n = 4, p = 2, startx = 3, starty = 2;
            //double compare = 2.3646243;
            Excel ex = new Excel(@"D:\pro\6sem\компОбрЭкспДан\DataL1.xls", 1);
            var read = ex.ReadRange(startx,starty,startx+n,starty+p);
            int lenghtX = read.GetLength(0);
            int lenghtY = read.GetLength(1);

            ex.CreateNewSheet();

            int start = 1;
            ex.SelectWorksheet(2);

            var average = Average(read);
            start = WriteAndStartChange(start, "Average:", ex, average);

            //в идеале можно сократить DRY
            var dispersion = Dispersion(average, read);
            start = WriteAndStartChange(start, "Dispertion", ex,dispersion );

            var sqrtDispersion = Dispersion(average, read);
            for (int k = 0; k < sqrtDispersion.Length; k++)
            {
                sqrtDispersion[ k] = Math.Sqrt(sqrtDispersion[ k]);
            }
            start = WriteAndStartChange(start, "SQRT Dispersion", ex, sqrtDispersion);

            var covMatrix = Cov(average, read);
            start = WriteAndStartChange(start, "COV matrix:", ex, covMatrix);

            var standartMatrix = StandartMatrix(sqrtDispersion, average, read);
            start = WriteAndStartChange(start, "Standart matrix:", ex, standartMatrix);

            var averageStand = Average(standartMatrix);
            start = WriteAndStartChange(start, "Average column standart Matrix:", ex,averageStand );

            var correlMatrix = CorrelMatrix(standartMatrix);
            start = WriteAndStartChange(start, "CORREL matrix", ex, correlMatrix);

            var signifMatrix = Significance(compare, correlMatrix);
            start = WriteAndStartChange(start, "Significance", ex, signifMatrix);

            // ex.Save();

            int number = 1;
            var y = new double[lenghtX];
            for (int i = 0; i <lenghtX ; i++)
            {
                y[i] = read[i,number];
            }

            var matrixX = RegressionAnalysisMatrixX(number, read);
            start = WriteAndStartChange(start, "matrix x", ex, matrixX);

            var a = CoefficientLinRegression(y, matrixX);
            start = WriteAndStartChange(start, "coefficient A", ex, a);

            start = WriteAndStartChange(start, "y", ex, y);

            var yCalculate = MultipleMatrix(matrixX, a);
            start = WriteAndStartChange(start, "Calculate y", ex, yCalculate);

            var yAverage = Average(y);
            start = WriteAndStartChange(start, "y average", ex, yAverage);

            var yCalculateAverage = Average(yCalculate);
            start = WriteAndStartChange(start, "y calculate average", ex, yCalculateAverage);

            var coeffDeterm = CoeffDeterm(yAverage,y,yCalculate);
            start = WriteAndStartChange(start, "coefficient of determination", ex, coeffDeterm);

            Console.Read();
            ex.Close();
        }
    }
}
