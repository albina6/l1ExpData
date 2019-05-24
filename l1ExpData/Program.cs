using System;
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
                        returnAverage[i, j] = (matrix[i,j]-averege[j])/Math.Sqrt( dispersion[j]);
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

        static int WriteAndStartChange(int start, string name, Excel ex, int[] array)
        {
            ex.WriteToCellString(start, 1, name);
            start++;
            ex.WriteRange(start, 2, 1 + array.Length, array);
            start++;
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
            return (numerator / denominator);
        }

        //static double Block(double[,] aMatrix,int n)
        //{
        //    double sum=0 ;
        //    for (int j = 1; j < n; j++)
        //    {
        //        for (int i = 0; i < j - 1; i++)
        //        {
        //            sum += 2 * aMatrix[i, j] * aMatrix[i, j];
        //        }
        //    }
        //    return Math.Sqrt(sum) / n;
        //}
        static double[] MaxElem(double [,]aMatrix,int n)
        {
            var maxElem = new double[] {0,0,0 };
            for(int i = 0; i < n; i++)
            {
                for (int j = 0; j < i; j++)
                {
                    if ( Math.Abs(aMatrix[j, i]) > maxElem[0])
                    {
                        maxElem[0] = Math.Abs(aMatrix[j, i]);//не нравится переделай потом
                        maxElem[1] = j;
                        maxElem[2] = i;
                    }
                }
            }
            return maxElem;
        }

        static void ChangeMatrix(double [,] aMatrix,double[,]tMatrix, double[] maxElem)
        {
            var n = aMatrix.GetLength(0);
            double eps = 0.01;
            int p = (int)maxElem[1], q = (int)maxElem[2];
            double y=(aMatrix[p,p]-aMatrix[q,q])/ 2;
            double x,s,c;
            if (Math.Abs(y) < eps)
                x = -1;
            else
                x=-(Math.Sign(y) * aMatrix[p, q] / (Math.Sqrt(aMatrix[p, q] * aMatrix[p, q] + y * y)));
            s = x / (Math.Sqrt(2 * (1 + Math.Sqrt(1 - x * x))));
            c = Math.Sqrt(1 - s * s); 
            for (int i = 0; i < n; i++)
            {
                if ((i != p) && (i != q))
                {
                    double z1 = aMatrix[i, p], z2 = aMatrix[i, q];
                    aMatrix[q, i] = z1 * s + z2 * c;
                    aMatrix[i, q] = aMatrix[q, i];
                    aMatrix[i, p] = z1 * c - z2 * s;
                    aMatrix[p, i] = aMatrix[i, p];
                }
            }
            double z5 = s * s, z6 = c * c, z7 = s * c, 
                v1 = aMatrix[p, p], v2 = aMatrix[p, q], v3 = aMatrix[q, q];
            aMatrix[p, p] = v1 * z6 + v3 * z5 - 2 * v2 * z7;
            aMatrix[q, q] = v1 * z5 + v3 * z6 + 2 * v2 * z7;
            aMatrix[p, q] = (v1 - v3) * z7 + v2 * (z6 - z5);
            aMatrix[q, p] = aMatrix[p, q];
            
            for (int i = 0; i < n; i++)
            {
                double z3, z4;
                z3 = tMatrix[i, p];
                z4 = tMatrix[i, q];
                tMatrix[i, q] = z3 * s + z4 * c;
                tMatrix[i, p] = z3 * c - z4 * s;
            }
        }

        static void Jacobi( double[,] aMatrix,double [,] tMatrix, double eps,double hi2)
        {
            if(aMatrix.GetLength(0)!=aMatrix.GetLength(1))
            {
                Console.WriteLine("This matrix is invalid Jacobi");
                Console.Read();
            }
            else
            {
                int n = aMatrix.GetLength(0);
                double a0;
                double sum = 0;
                for (int j = 1; j < n; j++)
                {
                    for (int i = 0; i < j ; i++)
                    {
                        sum +=  aMatrix[i, j] * aMatrix[i, j];
                    } 
                }
                double d = 2 * sum * n * n;
                if (d<= hi2)
                {
                    Console.WriteLine("correl matrix like E-matrix");
                    Console.Read();
                }
                else
                {
                    a0 = Math.Sqrt(sum * 2) / n;
                    var ai = a0;
                    var maxElem = MaxElem(aMatrix, n);
                    double epsel = 0.0001;
                    while (maxElem[0] > eps * a0)
                    {
                        if (maxElem[0] - ai > epsel)
                            ChangeMatrix(aMatrix, tMatrix, maxElem);
                        ai = ai / (n * n);
                        maxElem = MaxElem(aMatrix, n);
                    }
                }

                
            }
        }

        static void SortT(double [,] matrixT, int [] index)
        {
            if( (matrixT.GetLength(0) != matrixT.GetLength(1)) && (matrixT.GetLength(1) != index.Length))
            {
                Console.WriteLine("this matrix invalid");
                Console.Read();
            }
            else
            {
                int n = matrixT.GetLength(0);

                for (int i=0;i<n;i++)
                {
                    if (i != index[i]&& i<index[i])
                    {
                        double a;
                        int indexColumn = index[i];
                       // int
                        for (int j = 0; j < n; j++)
                        {
                            a = matrixT[ j,i];
                            matrixT[j, i] = matrixT[j, indexColumn];
                            matrixT[j, indexColumn] = a;
                        }
                        //index[indexColumn] = i;
                       // index[i] = indexColumn;
                    }
                    
                }
            }
        }

        static int [] SortIncrease(double [] array)
        {
            int n = array.Length;
            var index = new int[n];
            for (int i = 0; i < n; i++)
            {
                index[i] = i;
            }
            for (int i = 0; i < n; i++)
            {
                double max = -10000;
                int k =i;
                for(int j=i; j < n; j++)
                {
                    if (array[j] > max)
                    {
                        max =Math.Abs( array[j]);
                        k = j;
                    }
                }
                if (k != i)
                {
                    double a = array[i];
                    int b = index[i];
                    array[i] = max;
                    index[i] = k;
                    array[k] = array[i];
                    index[k] = i;
                }
                
            }
            return index;
        }

        static double [,] MGKMatrix(double [,] cMatrix, double[,] xMatrix)
        {
            int x = xMatrix.GetLength(0), y = xMatrix.GetLength(1);
            var mgkMatrix = new double[x, y];
            if ((cMatrix.GetLength(0)!=cMatrix.GetLength(1)) && (cMatrix.GetLength(0) != y))
            {
                Console.WriteLine("This matrix Invalid MGK Matrix");
                Console.Read();
            }
            else
            {
                for (int j = 0; j < y; j++)
                {
                    for (int i = 0; i < x; i++)
                    {
                        double sum = 0;
                        for (int k = 0; k < y; k++)
                        {
                            sum += cMatrix[ k,j] * xMatrix[ i,k];
                        }
                        mgkMatrix[i, j] = sum;
                    }

                }
            }
            return mgkMatrix;
        }

        static double SumArray(double [] array)
        {
            int n = array.Length;
            double sum = 0;
            for (int i = 0; i < n; i++)
            {
                sum += array[i];
            }
            return sum;
        }

        static void Main(string[] args)
        {
            int n = 77, p = 9, startx = 3, starty = 108;//106
            //double compare = 1.9839715;
            int start = 1;

            //int n = 3, p = 3, startx = 3, starty = 2;
            //int n = 4, p = 2, startx = 3, starty = 2;
            //double compare = 2.3646243;
            Excel ex = new Excel(@"D:\pro\6sem\компОбрЭкспДан\DataL1.xls", 1);
            //var read = ex.ReadRange(startx,starty,startx+n,starty+p);
            //int lenghtX = read.GetLength(0);
            //int lenghtY = read.GetLength(1);

            //n = 3;
            //p = 3;
            //startx = 3;
            //starty = 2;
            
            var read = ex.ReadRange(startx, starty, startx + n, starty + p);
            int x = read.GetLength(0), y = read.GetLength(1);
            var aMatrix = new double[x, y];
            for (int i = 0; i < x; i++)
            {
                for(int j = 0; j < y; j++)
                {
                    aMatrix[i, j] = read[i, j];
                }
            }
            ex.CreateNewSheet();        
            ex.SelectWorksheet(2);

            ////var average = Average(read);
            ////start = WriteAndStartChange(start, "Average:", ex, average);

            //////в идеале можно сократить DRY
            ////var dispersion = Dispersion(average, read);
            ////start = WriteAndStartChange(start, "Dispertion", ex, dispersion);

            ////var sqrtDispersion = Dispersion(average, read);
            ////for (int k = 0; k < sqrtDispersion.Length; k++)
            ////{
            ////    sqrtDispersion[k] = Math.Sqrt(sqrtDispersion[k]);
            ////}
            ////start = WriteAndStartChange(start, "SQRT Dispersion", ex, sqrtDispersion);

            ////var covMatrix = Cov(average, read);
            ////start = WriteAndStartChange(start, "COV matrix:", ex, covMatrix);

            ////var standartMatrix = StandartMatrix(sqrtDispersion, average, read);
            ////start = WriteAndStartChange(start, "Standart matrix:", ex, standartMatrix);

            ////var averageStand = Average(standartMatrix);
            ////start = WriteAndStartChange(start, "Average column standart Matrix:", ex, averageStand);

            ////var correlMatrix = CorrelMatrix(standartMatrix);
            ////start = WriteAndStartChange(start, "CORREL matrix", ex, correlMatrix);

            ////var signifMatrix = Significance(compare, correlMatrix);
            ////start = WriteAndStartChange(start, "Significance", ex, signifMatrix);

            ////// ex.Save();

            ////int number = 1;
            ////var y = new double[lenghtX];
            ////for (int i = 0; i < lenghtX; i++)
            ////{
            ////    y[i] = read[i, number];
            ////}

            ////var matrixX = RegressionAnalysisMatrixX(number, read);
            ////start = WriteAndStartChange(start, "matrix x", ex, matrixX);

            ////var a = CoefficientLinRegression(y, matrixX);
            ////start = WriteAndStartChange(start, "coefficient A", ex, a);

            ////start = WriteAndStartChange(start, "y", ex, y);

            ////var yCalculate = MultipleMatrix(matrixX, a);
            ////start = WriteAndStartChange(start, "Calculate y", ex, yCalculate);

            ////var yAverage = Average(y);
            ////start = WriteAndStartChange(start, "y average", ex, yAverage);

            ////var yCalculateAverage = Average(yCalculate);
            ////start = WriteAndStartChange(start, "y calculate average", ex, yCalculateAverage);

            ////var coeffDeterm = CoeffDeterm(yAverage, y, yCalculate);
            ////start = WriteAndStartChange(start, "coefficient of determination", ex, coeffDeterm);


            //////start = WriteAndStartChange(start, "my data ", ex, standartMatrix);
            //////Jacobi(ex, start, standartMatrix, 0.005);

            /////////////////////////////////

            ///////////
            var average = Average(aMatrix);
            start = WriteAndStartChange(start, "Average:", ex, average);

            //в идеале можно сократить DRY
            var dispersion = Dispersion(average, aMatrix);
            start = WriteAndStartChange(start, "Dispertion", ex, dispersion);

            var sqrtDispersion = Dispersion(average, aMatrix);
            for (int k = 0; k < sqrtDispersion.Length; k++)
            {
                sqrtDispersion[k] = Math.Sqrt(sqrtDispersion[k]);
            }
            start = WriteAndStartChange(start, "SQRT Dispersion", ex, sqrtDispersion);

            var covMatrix = Cov(average, aMatrix);
            start = WriteAndStartChange(start, "COV matrix:", ex, covMatrix);

            var standartMatrix = StandartMatrix(dispersion, average, aMatrix);
            start = WriteAndStartChange(start, "Standart matrix:", ex, standartMatrix);

            covMatrix = Cov(average, standartMatrix);
            start = WriteAndStartChange(start, "COV matrix:", ex, covMatrix);

            var averageStand = Average(standartMatrix);
            start = WriteAndStartChange(start, "Average column standart Matrix:", ex, averageStand);

            var correlMatrix = CorrelMatrix(standartMatrix);
            start = WriteAndStartChange(start, "CORREL matrix", ex, correlMatrix);
            ///////

            int lenghtX =correlMatrix.GetLength(0);
            var xMatrix = new double[lenghtX, lenghtX];
            for(int i = 0; i < lenghtX; i++)
            {
                for(int j = 0; j < lenghtX; j++)
                {
                    xMatrix[i, j] = correlMatrix[i, j];
                }
            }

            //////////////
            
            var tMatrix = new double[lenghtX, lenghtX];
            for (int i = 0; i < lenghtX; i++)
            {
                tMatrix[i, i] = 1;
            }

            start = WriteAndStartChange(start, "test ", ex, xMatrix);
            double hi2 = 7.84;
            Jacobi(xMatrix, tMatrix, 0.01, hi2);

            var lambda = new double[lenghtX];
            for (int i = 0; i < lenghtX; i++)
                lambda[i] = xMatrix[i, i];
            start = WriteAndStartChange(start, "Eigenvector", ex, lambda);

            var index = SortIncrease(lambda);
            start = WriteAndStartChange(start, "lambda ", ex, index);

            start = WriteAndStartChange(start, "tmatrix ", ex, tMatrix);

            SortT(tMatrix, index);
            start = WriteAndStartChange(start, "tmatrix sort", ex, tMatrix);

            var mgkMatrix = MGKMatrix(tMatrix, correlMatrix);
            start = WriteAndStartChange(start, "mgkMatrix", ex,mgkMatrix);

            var averageX = Average(correlMatrix);
            var averageMGK = Average(mgkMatrix);
            start = WriteAndStartChange(start, "averageX", ex, averageX);
            start = WriteAndStartChange(start, "averageMGK", ex, averageMGK);

            var dispersionMGK = Dispersion(averageMGK,mgkMatrix);
            var dispersionXMatrix = Dispersion(averageX, correlMatrix);
            start = WriteAndStartChange(start, "DispertionMGK", ex, dispersionMGK);
            start = WriteAndStartChange(start, "DispertionX", ex, dispersionXMatrix);

            double sumDispersionMGK = SumArray(dispersionMGK);
            double sumDispersionX = SumArray(dispersionXMatrix);

            start = WriteAndStartChange(start, "sumDispersionMGK", ex, sumDispersionMGK);
            start = WriteAndStartChange(start, "sumDispersionX", ex, sumDispersionX);

            Console.Read();
            ex.Close();
        }
    }
}
