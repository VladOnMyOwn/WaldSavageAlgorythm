// Copyright © 2021 Gorokhov Vladislav Dmitrievich. All rights reserved.
// Contacts: <gorohov2017vladislav@yandex.ru>
// License: https://www.gnu.org/licenses/old-licenses/gpl-2.0.html

using System;
using System.Collections.Generic;

namespace Wsav_alg
{
    public class WSav_algorythm {

        public static void FindWoptimal(double[,] matrix, out int[] WOstrategies, out double Wsum, out double[] Wi) {
            Wi = new double[matrix.GetLength(0)]; 
            double min;
            for (int i = 0; i < matrix.GetLength(0); ++i) { 
                min = double.MaxValue;
                for (int j = 0; j < matrix.GetLength(1); ++j) {
                    if (matrix[i, j] < min) min = matrix[i, j];
                }
                Wi[i] = min;
            }
            Wsum = double.MinValue; 
            for (int i = 0; i < Wi.Length; ++i) {
                if (Wi[i] > Wsum) Wsum = Wi[i];
            }
            List<int> strategiesList = new List<int>();
            for (int i = 0; i < Wi.Length; ++i) {
                if (Wi[i] == Wsum) strategiesList.Add(i);
            }
            WOstrategies = strategiesList.ToArray(); 
        }
        
        public static void CreateRiskMatrix(double[,] matrix, out double[,] Rmatrix) {
            Rmatrix = new double[matrix.GetLength(0), matrix.GetLength(1)]; 
            double max;
            for (int j = 0; j < matrix.GetLength(1); ++j) {
                max = double.MinValue;
                for (int i = 0; i < matrix.GetLength(0); ++i) { 
                    if (matrix[i, j] > max) {
                        max = matrix[i, j];
                    }
                }
                for (int i = 0; i < Rmatrix.GetLength(0); ++i) { 
                    Rmatrix[i, j] = Math.Round(max - matrix[i, j], 1);
                }
            }
        }

        public static void FindSavOptimal(double[,] Rmatrix, out int[] SavOstrategies, out double Sav_sum, out double[] Savi) {
            Savi = new double[Rmatrix.GetLength(0)];
            double max;
            for (int i = 0; i < Rmatrix.GetLength(0); ++i) { 
                max = double.MinValue;
                for (int j = 0; j < Rmatrix.GetLength(1); ++j) {
                    if (Rmatrix[i, j] > max) {
                        max = Rmatrix[i, j];
                    }
                }
                Savi[i] = max;
            }
            Sav_sum = double.MaxValue; 
            for (int i = 0; i < Savi.Length; ++i) {
                if (Savi[i] < Sav_sum) Sav_sum = Savi[i];
            }
            List<int> strategiesList = new List<int>();
            for (int i = 0; i < Savi.Length; ++i) { 
                if (Savi[i] == Sav_sum) strategiesList.Add(i);
            }
            SavOstrategies = strategiesList.ToArray();
        }

        public static bool IsIntersectionNotEmpty(int[] WOstrategies, int[] SavOstrategies, out int[] WSavIntersection) {
            bool f = false;
            List<int> strategiesList = new List<int>();
            for (int i = 0; i < WOstrategies.Length; ++i) {
                for (int j = 0; j < SavOstrategies.Length; ++j) {
                    if (WOstrategies[i] == SavOstrategies[j]) {
                        f = true;
                        strategiesList.Add(WOstrategies[i]);
                    }
                }
            }
            WSavIntersection = strategiesList.ToArray();
            if (f)
                return true;
            else
                return false;
        }

    }
}