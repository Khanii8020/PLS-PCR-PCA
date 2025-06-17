using System;
using System.Collections.Generic;
using System.IO;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using ClosedXML.Excel;
using System.Linq;

namespace PCA_PLS_PCR
{
    public static class PCA_PCR_PLS
    {
        public static Matrix<double> LoadMatrixFromExcel(string filePath, int rowStart, int rowEnd, int colStart, int colEnd, int sheetIndex)
        {
            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.ElementAt(sheetIndex); // Sheet index is 0-based
                int rows = rowEnd - rowStart + 1;
                int cols = colEnd - colStart + 1;

                var matrix = DenseMatrix.Create(rows, cols, (i, j) =>
                {
                    var cell = worksheet.Cell(rowStart + i, colStart + j);
                    var cellText = cell.GetValue<string>();

                    if (double.TryParse(cellText, out double val))
                        return val;
                    else
                        return 0.0; // default fallback for non-numeric cells
                });

                return matrix;
            }
        }

        public static void SaveMatrixToExcel(string filePath, double[,] matrix, string sheetName, int startRow, int startCol)
        {
            using (var workbook = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Contains(sheetName)
                    ? workbook.Worksheet(sheetName)
                    : workbook.AddWorksheet(sheetName);

                int rowCount = matrix.GetLength(0);
                int colCount = matrix.GetLength(1);

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        worksheet.Cell(startRow + i, startCol + j).Value = matrix[i, j];
                    }
                }

                workbook.SaveAs(filePath);
            }
        }

        public static void SavePressDataToExcel(List<double> pressValues, string filePath, string sheetName, int startRow, int colIndex)
        {
            using (var workbook = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Contains(sheetName)
                    ? workbook.Worksheet(sheetName)
                    : workbook.AddWorksheet(sheetName);

                worksheet.Cell(startRow - 1, colIndex).Value = "n";
                worksheet.Cell(startRow - 1, colIndex + 1).Value = "PRESS";

                for (int i = 0; i < pressValues.Count; i++)
                {
                    worksheet.Cell(startRow + i, colIndex).Value = i + 1;
                    worksheet.Cell(startRow + i, colIndex + 1).Value = pressValues[i];
                }

                workbook.SaveAs(filePath);
            }
        }

        public static Matrix<double> MeanCenter(Matrix<double> data)
        {
            var mean = data.ColumnSums() / data.RowCount;
            return data.MapIndexed((i, j, val) => val - mean[j]);
        }

        public static (Matrix<double> Scores, Matrix<double> Loadings) PCA(Matrix<double> X, int nComponents)
        {
            var svd = X.Svd(computeVectors: true);
            var U = svd.U.SubMatrix(0, X.RowCount, 0, nComponents);
            var S = DenseMatrix.OfDiagonalVector(svd.S.SubVector(0, nComponents));
            var Vt = svd.VT.SubMatrix(0, nComponents, 0, X.ColumnCount);
            var T = U.Multiply(S);
            var P = Vt;
            return (T, P);
        }

        public static double ComputePRESS(Matrix<double> X, Matrix<double> Y, int n)
        {
            double press = 0.0;
            for (int i = 0; i < X.RowCount; i++)
            {
                var X_train = X.RemoveRow(i);
                var Y_train = Y.RemoveRow(i);

                var pca = PCA(X_train, n);
                var T = pca.Scores;
                
                var b = (T.TransposeThisAndMultiply(T)).Inverse()
                         .Multiply(T.TransposeThisAndMultiply(Y_train));

                var x_i = X.Row(i);
                var t_i = x_i * pca.Loadings.Transpose(); // 1×n
                var y_pred = t_i * b;                     // 1×2
                var y_true = Y.Row(i);                    // 1×2

                press += (y_true - y_pred).PointwisePower(2).Sum(); // 2x1
            }
            return press;
        }

        public static int FindOptimalNFromPress(List<double> press)
        {
            int n = press.Count;

            // نقاط اولیه و پایانی برای رسم خط
            var x1 = 1;
            var y1 = press[0];
            var x2 = n;
            var y2 = press[n - 1];

            double maxDistance = double.MinValue;
            int optimalN = 1;

            for (int i = 1; i < n - 1; i++)
            {
                // مختصات نقطه فعلی
                double x0 = i + 1;
                double y0 = press[i];

                // فاصله عمودی نقطه از خط بین (x1,y1) و (x2,y2)
                double numerator = Math.Abs((y2 - y1) * x0 - (x2 - x1) * y0 + x2 * y1 - y2 * x1);
                double denominator = Math.Sqrt(Math.Pow(y2 - y1, 2) + Math.Pow(x2 - x1, 2));
                double distance = numerator / denominator;

                if (distance > maxDistance)
                {
                    maxDistance = distance;
                    optimalN = i + 1; // چون اندیس از 0 شروع می‌شود
                }
            }

            return optimalN;
        }

        public static Matrix<double> Denoise(Matrix<double> Xc, Matrix<double> Y_pred, Matrix<double> Y_org)
        {
            var error = Y_org - Y_pred;
            double errorScale = error.L2Norm() / (Y_org.RowCount * Y_org.ColumnCount);
            return Xc.Map(x => x * (1 - errorScale));
        }

        public static void Start(string filePath, int rowXStart, int rowXEnd,
                                 int colStartX, int colEndX,
                                 int rowYStart, int rowYEnd,
                                 int colStartY, int colEndY,int startSheet, int outSheet)
        {
            var X = LoadMatrixFromExcel(filePath, rowXStart, rowXEnd, colStartX, colEndX,startSheet).Transpose();
            var Y = LoadMatrixFromExcel(filePath, rowYStart, rowYEnd, colStartY, colEndY, startSheet);

            var Xc = MeanCenter(X);
            //var Yc = MeanCenter(Y);
            var Yc = Y;

            int kMax = Math.Min(Xc.RowCount, Xc.ColumnCount)-1;
           
            var pressList = new List<double>();
            for (int k = 1; k < kMax; k++)
                pressList.Add(ComputePRESS(Xc, Yc, k));

            int n_opt = FindOptimalNFromPress(pressList);

            var pcaResult = PCA(Xc, n_opt);
            var T = pcaResult.Scores;
            var P = pcaResult.Loadings;

            var b = (T.TransposeThisAndMultiply(T)).Inverse()
           .Multiply(T.TransposeThisAndMultiply(Yc));

            var Y_hat = T.Multiply(b); // Tx = b.Y ==> PCR

            var X_denoised = Denoise(Xc, Y_hat, Yc);

            SaveMatrixToExcel(filePath, T.ToArray(), "Tx_scores(with n_optimize)", 1, 1);
            SaveMatrixToExcel(filePath, P.ToArray(), "Px_loadings(with n_optimize)", 1, 1);
            SaveMatrixToExcel(filePath, Y_hat.ToArray(), "Y_hat(Ypredict = b . TX)", 1, 1);
            SaveMatrixToExcel(filePath, X_denoised.Transpose().ToArray(), "X_denoised", 1, 1);
           
            SavePressDataToExcel(pressList, filePath, "PRESS_values", 2, 1);
        }
       
    }
}
