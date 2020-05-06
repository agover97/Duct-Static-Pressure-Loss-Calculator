using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DuctPressureCalc
{
    
    public partial class Form1 : Form
    {

        public List<List<string>> array_list = new List<List<string>>();
        static int numberOfRows = 0;
        static int numberOfCols;
        static int numRowIn2;
        static int numColIn2;
        public double[,] inputArray;
        public string[,] inputArray2;
        static int ductRun = 0;
        static int fitRun = 0;

        /// <summary>
        ///start
        /// </summary>
        /// 

        const double PI = 3.14159265358979;
        const double MEANTEMP = 59.0;                           //Farenheit
        const double PSIA = 14.7;
        const double DENSITYDRYAIR = 0.077;                 //Density of Dry Air (lb/ft^3)
        const double KINEMATICVISCOSITY1 = 1.565;   //Kinematic Viscosity (ft^2/s*10^-4)
        const double KINEMATICVISCOSITY2 = 1.573;   //Kinematic Viscosity (ft^2/s*10^-4)

        double cfm;                                                                 //F column
                                                                                    //Duct Dimensions
        double width;                                                               //G column
        double height;                                                          //H column
        double diameter;                                                    //inputArray[3]
        double ductLength;                                                  //H column and length from inputArray[4]
        double staticPressure;                                              //inputArray[5]
        double pressureDrop;                                                //inputArray[6]
                                                                            //Materials
        double roughness;                                                       //K column
                                                                                //Pressure Loss Data
        double fittingCoefficient;                                  //M column
        double componentPressureDrop;                               //N column

        bool checking;

        string ashraeFittingNumber;
        string description;
        string ductElement;
        int endRow = 45;



        /* --------- ARRAYS OF FIXED DATA --------- */
        /* --------- ARRAYS OF FIXED DATA --------- */
        public static double[,] SD4_1 = {
    {-1,10,15,20,30,45,60,90,120,150,180},
    {0.10,0.05,0.05,0.05,0.05,0.07,0.08,0.19,0.29,0.37,0.43},
    {0.17,0.05,0.04,0.04,0.04,0.06,0.07,0.18,0.28,0.36,0.42},
    {0.25,0.05,0.04,0.04,0.04,0.06,0.07,0.17,0.27,0.35,0.41},
    {0.50,0.05,0.05,0.05,0.05,0.06,0.06,0.12,0.18,0.24,0.26},
    {1.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00},
    {2.00,0.44,0.52,0.76,1.28,1.32,1.32,1.28,1.24,1.20,1.20},
    {4.00,2.56,3.52,4.80,7.36,9.76,10.88,10.24,10.08,9.92,9.92},
    {10.00,21.00,28.00,38.00,59.00,76.00,80.00,83.00,84.00,83.00,83.00},
    {16.00,53.76,74.24,97.28,153.60,215.04,225.28,225.28,225.28,225.28,225.28}
};

        public static double[,] SD4_2 = {
    {-1,10,15,20,30,45,60,90,120,150,180},
{0.1,0.05,0.05,0.05,0.05,0.07,0.08,0.19,0.29,0.37,0.43},
{0.17,0.05,0.05,0.04,0.04,0.06,0.07,0.18,0.28,0.36,0.42},
{0.25,0.06,0.05,0.05,0.04,0.06,0.07,0.17,0.27,0.35,0.41},
{0.5,0.06,0.07,0.07,0.05,0.06,0.06,0.12,0.18,0.24,0.26},
{1,0,0,0,0,0,0,0,0,0,0},
{2,0.6,0.84,1,1.2,1.32,1.32,1.32,1.28,1.24,1.2},
{4,4,5.76,7.2,8.32,9.28,9.92,10.24,10.24,10.24,10.24},
{10,30,50,53,64,75,84,89,91,91,88},
{16,76.8,138.24,135.68,166.4,197.12,225.28,243.2,250.88,250.88,238.08}
};

        public static double[,] SD5_10_Cb = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.65,0.24,0,0,0,0,0,0,0},
{0.2,2.98,0.65,0.33,0.24,0.18,0,0,0,0},
{0.3,7.36,1.56,0.65,0.39,0.29,0.24,0.2,0,0},
{0.4,13.78,2.98,1.2,0.65,0.43,0.33,0.27,0.24,0.21},
{0.5,22.24,4.92,1.98,1.04,0.65,0.47,0.36,0.3,0.26},
{0.6,32.73,7.36,2.98,1.56,0.96,0.65,0.49,0.39,0.33},
{0.7,45.26,10.32,4.21,2.21,1.34,0.9,0.65,0.51,0.42},
{0.8,59.82,13.78,5.67,2.98,1.8,1.2,0.86,0.65,0.52},
{0.9,76.41,17.75,7.36,3.88,2.35,1.56,1.11,0.83,0.65}
};

        public static double[,] SD5_10_Cs = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.13,0.16,0,0,0,0,0,0,0},
{0.2,0.2,0.13,0.15,0.16,0.28,0,0,0,0},
{0.3,0.9,0.13,0.13,0.14,0.15,0.16,0.2,0,0},
{0.4,2.88,0.2,0.14,0.13,0.14,0.15,0.15,0.16,0.34},
{0.5,6.25,0.37,0.17,0.14,0.13,0.14,0.14,0.15,0.15},
{0.6,11.88,0.9,0.2,0.13,0.14,0.13,0.14,0.14,0.15},
{0.7,18.62,1.71,0.33,0.18,0.16,0.14,0.13,0.15,0.14},
{0.8,26.88,2.88,0.5,0.2,0.15,0.14,0.13,0.13,0.14},
{1.00001,36.45,4.46,0.9,0.3,0.19,0.16,0.15,0.14,0.13}
};

        public static double[,] SR4_1 = {
    {-1,10,15,20,30,45,60,90,120,150,180},
{0.1,0.05,0.05,0.05,0.05,0.07,0.08,0.19,0.29,0.37,0.43},
{0.17,0.05,0.04,0.04,0.04,0.05,0.07,0.18,0.28,0.36,0.42},
{0.25,0.05,0.04,0.04,0.04,0.06,0.07,0.17,0.27,0.35,0.41},
{0.5,0.06,0.05,0.05,0.05,0.06,0.07,0.14,0.2,0.26,0.27},
{1,0,0,0,0,0,0,0,0,0,1},
{2,0.56,0.52,0.6,0.96,1.4,1.48,1.52,1.48,1.44,1.4},
{4,2.72,3.04,3.52,6.72,9.6,10.88,11.2,11.04,10.72,10.56},
{10,24,26,36,53,69,82,93,93,92,91},
{16,66.56,69.12,102.4,143.36,181.76,220.16,256,253.44,250.88,250.88}
};

        public static double[,] SR4_3 = {
    {-1,10,15,20,30,45,60,90,120,150,180},
{0.1,0.05,0.05,0.05,0.05,0.07,0.08,0.19,0.29,0.37,0.43},
{0.17,0.05,0.05,0.05,0.04,0.06,0.07,0.18,0.28,0.36,0.42},
{0.25,0.06,0.05,0.05,0.04,0.06,0.07,0.17,0.27,0.35,0.41},
{0.5,0.06,0.07,0.07,0.05,0.06,0.06,0.12,0.18,0.24,0.26},
{1,0,0,0,0,0,0,0,0,0,0},
{2,0.6,0.84,1,1.2,1.32,1.32,1.32,1.28,1.24,1.2},
{4,4,5.76,7.2,8.32,9.28,9.92,10.24,10.24,10.24,10.24},
{10,30,50,53,64,75,84,89,91,91,88},
{16,76.8,138.24,135.68,166.4,197.12,225.28,243.2,250.88,250.88,238.08}
};

        public static double[,] SR5_13_Cb = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.73,0.34,0.32,0.34,0.35,0.37,0.38,0.39,0.4},
{0.2,3.1,0.73,0.41,0.34,0.32,0.32,0.33,0.34,0.35},
{0.3,7.59,1.65,0.73,0.47,0.37,0.34,0.32,0.32,0.32},
{0.4,14.2,3.1,1.28,0.73,0.51,0.41,0.36,0.34,0.32},
{0.5,22.92,5.08,2.07,1.12,0.73,0.54,0.44,0.38,0.35},
{0.6,33.76,7.59,3.1,1.65,1.03,0.73,0.56,0.47,0.41},
{0.7,46.71,10.63,4.36,2.31,1.42,0.98,0.73,0.58,0.49},
{0.8,61.79,14.2,5.86,3.1,1.9,1.28,0.94,0.73,0.6},
{0.9,78.98,18.29,7.59,4.02,2.46,1.65,1.19,0.91,0.73}
};

        public static double[,] SR5_13_Cs = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.04,0,0,0,0,0,0,0,0},
{0.2,0.98,0.04,0,0,0,0,0,0,0},
{0.3,3.48,0.31,0.04,0,0,0,0,0,0},
{0.4,7.55,0.98,0.18,0.04,0,0,0,0,0},
{0.5,13.18,2.03,0.49,0.13,0.04,0,0,0,0},
{0.6,20.38,3.48,0.98,0.31,0.1,0.04,0,0,0},
{0.7,29.15,5.32,1.64,0.6,0.23,0.09,0.04,0,0},
{0.8,39.48,7.55,2.47,0.98,0.42,0.18,0.08,0.04,0},
{0.9,51.37,10.17,3.48,1.46,0.67,0.31,0.15,0.07,0.04}
};

        public static double[,] CR3_1 = {
    {-1,0.25,0.5,0.75,1,1.5,2,3,4,5,6,8},
{0,3.45,3.15,2.97,2.75,2.51,2.39,2.39,2.51,2.63,2.71,2.75},
{0.5,1.53,1.38,1.29,1.18,1.06,1,1,1.06,1.12,1.16,1.18},
{0.75,0.57,0.52,0.48,0.44,0.4,0.39,0.39,0.4,0.42,0.43,0.44},
{1,0.27,0.25,0.23,0.21,0.19,0.18,0.18,0.19,0.2,0.21,0.21},
{1.5,0.22,0.2,0.19,0.17,0.15,0.14,0.14,0.15,0.16,0.17,0.17},
{2,0.2,0.18,0.16,0.15,0.14,0.13,0.13,0.14,0.14,0.15,0.15}
};

        public static double[,] CR3_1_angle_factor = {

};

        public static double[,] SR4_2 = {
    {-1,10,15,20,30,45,60,90,120,150,180},
{0.063,78.58,131.19,134.83,164.94,194.91,222.29,239.75,247.42,247.5,235.44},
{0.1,30.86,49.5,53.04,63.54,74.45,83.42,88.38,90.33,90.33,87.7},
{0.25,4.07,5.76,7.23,8.34,9.29,9.92,10.24,10.24,10.24,10.24},
{0.5,0.62,0.85,0.99,1.2,1.32,1.32,1.32,1.28,1.24,1.2},
{0.8,0.15,0.21,0.25,0.3,0.33,0.33,0.33,0.32,0.31,0.29},
{1.2,0.02,0.02,0.02,0.02,0.02,0.02,0.04,0.06,0.08,0.08},
{2,0.06,0.07,0.07,0.05,0.06,0.06,0.12,0.18,0.24,0.26},
{4,0.06,0.05,0.05,0.04,0.06,0.07,0.17,0.27,0.35,0.4},
{6,0.05,0.05,0.05,0.04,0.06,0.07,0.18,0.28,0.36,0.42},
{10,0.05,0.05,0.05,0.05,0.07,0.08,0.19,0.29,0.37,0.43},
{100,0.05,0.05,0.05,0.05,0.07,0.09,0.2,0.3,0.38,0.45}
};

        public static double[,] SD5_9_Cb = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,1.2,0.62,0.8,1.28,1.99,2.92,4.07,5.44,7.02},
{0.2,4.1,1.2,0.72,0.62,0.66,0.8,1.01,1.28,1.6},
{0.3,8.99,2.4,1.2,0.81,0.66,0.62,0.64,0.7,0.8},
{0.4,15.89,4.1,1.94,1.2,0.88,0.72,0.64,0.62,0.63},
{0.5,24.8,6.29,2.91,1.74,1.2,0.92,0.77,0.68,0.63},
{0.6,35.73,8.99,4.1,2.4,1.62,1.2,0.96,0.81,0.72},
{0.7,48.67,12.19,5.51,3.19,2.12,1.55,1.2,0.99,0.85},
{0.8,63.63,15.89,7.14,4.1,2.7,1.94,1.49,1.2,1.01},
{0.9,80.6,20.1,8.99,5.13,3.36,2.4,1.83,1.46,1.2}
};

        public static double[,] SD5_9_Cs = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.13,0.16,0,0,0,0,0,0,0},
{0.2,0.2,0.13,0.15,0.16,0.28,0,0,0,0},
{0.3,0.9,0.13,0.13,0.14,0.15,0.16,0.2,0,0},
{0.4,2.88,0.2,0.14,0.13,0.14,0.15,0.15,0.16,0.34},
{0.5,6.25,0.37,0.17,0.14,0.13,0.14,0.14,0.15,0.15},
{0.6,11.88,0.9,0.2,0.13,0.14,0.13,0.14,0.14,0.15},
{0.7,18.62,1.71,0.33,0.18,0.16,0.14,0.13,0.15,0.14},
{0.8,26.88,2.88,0.5,0.2,0.15,0.14,0.13,0.13,0.14},
{0.9,36.45,4.46,0.9,0.3,0.19,0.16,0.15,0.14,0.13}
};

        public static double[,] SD5_18_Cb = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,1.15,0.32,0.01,0.01,0,0,0,0,0},
{0.2,3.11,1.11,0.54,0.31,0.01,0.01,0,0,0},
{0.3,6.93,1.64,1.05,0.65,0.39,0.27,0,0,0},
{0.4,13.56,2.96,1.32,1.09,0.76,0.52,0.38,0.29,0},
{0.5,22.94,5.09,2.12,1.27,1.1,0.84,0.62,0.48,0.38},
{0.6,32.89,7.24,2.98,1.68,0.84,1.09,0.87,0.68,0.53},
{0.7,45.43,10.14,4.29,2.29,1.49,1.17,1.1,0.92,0.73},
{0.8,59.84,13.68,5.66,2.98,1.94,1.37,1.14,1.1,0.95},
{0.9,75.45,17.64,7.05,3.91,2.41,1.68,1.3,1.13,1.09}
};

        public static double[,] SD5_18_Cs = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,1.15,0.32,0.01,0.01,0,0,0,0,0},
{0.2,3.11,1.11,0.54,0.31,0.01,0.01,0,0,0},
{0.3,6.93,1.64,1.05,0.65,0.39,0.27,0,0,0},
{0.4,13.56,2.96,1.32,1.09,0.76,0.52,0.38,0.29,0},
{0.5,22.94,5.09,2.12,1.27,1.1,0.84,0.62,0.48,0.38},
{0.6,32.89,7.24,2.98,1.68,0.84,1.09,0.87,0.68,0.53},
{0.7,45.43,10.14,4.29,2.29,1.49,1.17,1.1,0.92,0.73},
{0.8,59.84,13.68,5.66,2.98,1.94,1.37,1.14,1.1,0.95},
{0.9,75.45,17.64,7.05,3.91,2.41,1.68,1.3,1.13,1.09}
};

        public static double[,] SR5_12_Cb = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,1.01,1.2,5.33,13.68,26.37,43.45,64.94,90.85,121.19},
{0.2,3.19,1,0.79,1.66,3.63,6.77,11.1,16.63,23.37},
{0.3,7.34,1.97,0.99,0.65,0.74,1.26,2.22,3.65,5.56},
{0.4,14.39,3.33,1.65,1.01,0.7,0.62,0.78,1.18,1.82},
{0.5,23.87,5.1,2.39,1.47,1,0.74,0.62,0.66,0.84},
{0.6,36.09,7.43,3.28,1.99,1.37,0.99,0.76,0.64,0.62},
{0.7,50.62,10.28,4.35,2.55,1.75,1.29,0.98,0.78,0.66},
{0.8,71.87,14.59,5.94,3.37,2.27,1.66,1.28,1.02,0.83},
{0.9,90.37,18.42,7.38,4.08,2.7,1.98,1.53,1.22,0.99}
};

        public static double[,] SR5_12_Cs = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.04,0,0,0,0,0,0,0,0},
{0.2,0.98,0.04,0,0,0,0,0,0,0},
{0.3,3.39,0.3,0.04,0,0,0,0,0,0},
{0.4,7.55,0.98,0.18,0.04,0,0,0,0,0},
{0.5,13.18,2.03,0.49,0.13,0.04,0,0,0,0},
{0.6,20.38,3.48,0.98,0.31,0.1,0.04,0,0,0},
{0.7,29.29,5.35,1.65,0.61,0.23,0.06,0.04,0,0},
{0.8,36.54,6.92,2.23,0.87,0.37,0.15,0.07,0.03,0},
{0.9,50.99,10.09,3.45,1.44,0.66,0.31,0.15,0.07,0.04}
};

        public static double[,] SR5_11_Cb = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,1.58,0.94,0.83,0.79,0.77,0.76,0.76,0.76,0.75},
{0.2,4.2,1.58,1.1,0.94,0.87,0.83,0.8,0.79,0.78},
{0.3,8.63,2.67,1.58,1.2,1.03,0.94,0.88,0.85,0.83},
{0.4,14.85,4.2,2.25,1.58,1.27,1.1,1,0.94,0.9},
{0.5,22.87,6.19,3.13,2.07,1.58,1.32,1.16,1.06,0.99},
{0.6,32.68,8.63,4.2,2.67,1.96,1.58,1.35,1.2,1.1},
{0.7,44.3,11.51,5.48,3.38,2.41,1.89,1.58,1.38,1.24},
{0.8,57.71,14.85,6.95,4.2,2.94,2.25,1.84,1.58,1.4},
{0.9,72.92,18.63,8.63,5.14,3.53,2.67,2.14,1.81,1.58}
};

        public static double[,] SR5_11_Cs = {
    {-1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9},
{0.1,0.04,0,0,0,0,0,0,0,0},
{0.2,0.98,0.04,0,0,0,0,0,0,0},
{0.3,3.48,0.31,0.04,0,0,0,0,0,0},
{0.4,7.55,0.98,0.18,0.04,0,0,0,0,0},
{0.5,13.18,2.03,0.49,0.13,0.04,0,0,0,0},
{0.6,20.38,3.48,0.98,0.31,0.1,0.04,0,0,0},
{0.7,29.15,5.32,1.64,0.6,0.23,0.09,0.04,0,0},
{0.8,39.48,7.55,2.47,0.98,0.42,0.18,0.08,0.04,0},
{0.9,51.37,10.17,3.48,1.46,0.67,0.31,0.15,0.07,0.04}
};

        public static double[,] CD3_4 = {
            { -1, 3,   4,   5,   6,   100 },
{ 1, 0.2, 0.17 ,   0.14 ,   0.11   , 0.11 },
{ 1, 0.2, 0.17 ,   0.14 ,   0.11   , 0.11 }


};

        public static double[,] CD3_2 = { 
            { -1, 3, 4, 5, 6, 7, 100 }, 
            { 1, 0.44, 0.37, 0.3, 0.25, 0.24, 0.24 } ,
            { 1, 0.44, 0.37, 0.3, 0.25, 0.24, 0.24 }


};



        public static double[,] SR5_14 = {
           {-1,0.5, 1 },
            {1, 0.3, 1 },
            {1, 0.3, 1 }


};

        public static double[,] CD3_1 = {
            {-1, 3.0, 3.9, 4.9, 5.9, 7.1, 7.9, 9.1, 100.0 },
            {1, 0.30,    0.21,    0.16,    0.14,    0.12,    0.11,    0.11,    0.11 },
            {1, 0.30,    0.21,    0.16,    0.14,    0.12,    0.11,    0.11,    0.11 }


};
        public static double[,] CR3_1_K =
        {
            {-1, 0,   20,  30 , 45 , 60 , 75, 90 , 110, 130, 150, 180 },
            {1, 0 ,  0.31 ,   0.45  ,  0.6, 0.78  ,  0.9 ,1 ,  1.13   , 1.2 ,1.28 ,   1.4 },
            {1, 0 ,  0.31 ,   0.45  ,  0.6, 0.78  ,  0.9 ,1 ,  1.13   , 1.2 ,1.28 ,   1.4 }

        };
        //All ASHRAE fitting numbers								//L column
        List<string> ashraeFittingNumbers = new List<string>() {
        "SD4-1", "SD4-2", "SD5-10", "SR4-1", "SR5-13", "SR5-14", "CD3-1", "CR3-1", "CR3-16",
        "SR4-3", "SR4-2", "SD5-9", "SD5-18", "SR5-12", "SR5-11", "CD3-4", "CD3-2"
    };
        //All descriptions													//E column
        List<string> descriptions = new List<string>() {
        "Air Terminal Device (FPU/ACB)","Fire/Smoke Damper","Louver", "Air Terminal Device (VAV)", "Air Terminal (Lab) Valve",
        "Air Terminal (Lab) Valve", "Reheat Coil Hydronic (2 Row)", "Reheat Coil (Electric)", "Sound Attenuator", "Airflow Station",
        "Airflow Station", "Manual Volume Damper", "Automatic Control Damper", "Fire Damper", "Smoke Damper", "UVGI Lamps",
        "Flex Duct (5ft.)", "Flex Elbow", "Flex Connector", "Wire-Mesh Screen", "Diffuser/ Grille"
    };
        //All duct elements													//D column
        List<string> ductElements = new List<string>() {
         "Fitting", "Component", "Flex Duct"
    };
        //All roughness															//K column & Common Tables column C (to match with column B material descriptions)
        List<string> roughnesses = new List<string>() {
         "0.0001", "0.00015", "0.0002", "0.003", "0.004", "0.00038", "0.0005", "0.003", "0.005", "0.01", "0.007", "0.003", "0.01"
    };
        //All materials															//J column & Common Tables column B
        List<string> materials = new List<string>() {
        "Uncolated carbon steel, clean", "PVC plastic pipe", "Aluminum", "Galvanized steel, longitudinal seams, 4ft joins",
        "Galvanized steel, continuously rolled, spiral seams, 10ft joints", "Galvanized steel, spiral seam, with 1,2, and 3 ribs 12ft joints",
        "Galvanized steel, longitudinal seams, with 2.5ft joints", "Galvanized steel, spiral, corrugated, 10ft joints",
        "Fibrous glass duct, air side with facing material", "Fibrous glass duct, air side spray coated",
        "Flexible duct, metallic (when fully extended)", "Flexible duct, utilized in conjuction with the pressure loss correction factor as indicated in 2009 ASHRAE", "Concrete"
    };
        //All common fitting coefficients						//Common tables
        List<string> fittingCoef = new List<string>() {
        "Elbow-90°-Mitered w/ vanes (Single thickness- 3-1/4\"Sp)", "Elbow-90°-Mitered w/ vanes (Single thickness- 3-1/4\"Sp)",
        "Elbow-90°- 1.5 Radius", "Elbow-90°- 1.0 Radius", "Elbow-45°- 1.5 Radius", "Type \"B\" & \"C\" Fire Damper",
        "Round Volume Damper (Open)", "Round Volume Damper (Open)", "Abrubt exit to atmosphere"
    };
        //All component pressure drops							//N column & Common Tables
        List<string> pressureDrops = new List<string>() {
        "0.1","0.15","0.12", "0.25", "0.6", "0.3", "0.15", "0.05", "0.1", "0.05", "0.03",
        "0.08", "0.1", "0.05", "0.1", "0.05", " ", " ", " ", "0.03", "0.1"
    };

        /* --------- 2D ARRAY HOLDING INPUTS --------- */
        //Note: empty values are NULL
        //static int numberOfRows = 7;
        // static int numberOfCols;
        // double[,] inputArray = new double[numberOfRows, numberOfCols];

        /* --------- SINGLE-DIM ARRAY FOR FINAL RESULTS --------- */
        //23 attributes in table so 23 columns
        string[] results = new string[23];


        ///////////////////
        ///end
        ///


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {
            //fitting pieces
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
                fitRun = 1;
            }
            else
            {
                return;
            }
            button1.BackColor = Color.Green;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            numRowIn2 = xlRange.Rows.Count;
            numColIn2 = xlRange.Columns.Count;
            string[,] stringArray = new string[numRowIn2, numColIn2];
            inputArray2 = new string[numRowIn2 - 1, numColIn2];

            for (int i = 1; i <= numRowIn2; i++)
            {
                for (int j = 1; j <= numColIn2; j++)
                {





                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {

                        //dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                        //List<string> temp_list = new List<string>();
                        stringArray[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();


                    }
                }
            }
           /* for (int i = 1; i <= numRowIn2; i++)
            {
                if (stringArray[i - 1, 6] != null)
                {
                    if (stringArray[i - 1, 6].Length > 7)
                    {
                        stringArray[i - 1, 6] = stringArray[i - 1, 6].Remove(7);
                    }
                }


            }*/

            string temp = String.Copy(stringArray[0, 0]);
            Console.WriteLine(temp);
            if (stringArray[3, 3] == null)
            {
                Console.WriteLine("hello");
            }

            for (int i = 1; i <= numRowIn2 - 2; i++)
            {
                for (int j = 1; j <= numColIn2; j++)
                {
                    string ttt = stringArray[i + 1, j - 1];
                    inputArray2[i - 1, j - 1] = stringArray[i + 1, j - 1];
                    if ((stringArray[i + 1, j - 1] == null))
                    {
                        inputArray2[i - 1, j - 1] = "empty";
                    }
                    //Console.WriteLine(inputArray2[i - 1, j - 1]);
                }
            }

            /* int rowCount = xlRange.Rows.Count;
             int colCount = xlRange.Columns.Count;

             // dt.Column = numberOfCols;  
             dataGridView1.ColumnCount = colCount;
             dataGridView1.RowCount = rowCount;

             for (int i = 1; i <= rowCount; i++)
             {
                 for (int j = 1; j <= colCount; j++)
                 {


                     //write the value to the Grid  


                     if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                     {
                         dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();

                     }
                     // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

                     //add useful things here!     
                 }
             }
 */
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



        }

        private void button2_Click(object sender, EventArgs e)
        {
            //regular pieces
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
                ductRun = 1;
            }
            else
            {
                return;
            }
            button2.BackColor = Color.Green;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            numberOfRows = xlRange.Rows.Count;
            numberOfCols = xlRange.Columns.Count;

            // dt.Column = numberOfCols;  
            //dataGridView1.ColumnCount = numberOfCols;
            //dataGridView1.RowCount = numberOfRows;
            string[,] stringArray = new string[numberOfRows, numberOfCols];
            inputArray = new double[numberOfRows-1,numberOfCols];

           /* if (xlRange.Cells[1, 1].Value2.ToString() != "Static Pressure Duct Schedule")
            {
                Console.WriteLine("Wrong file");
            }*/

            for (int i = 1; i <= numberOfRows; i++)
            {
                for (int j = 1; j <= numberOfCols; j++)
                {


   


                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {

                        //dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                        //List<string> temp_list = new List<string>();
                        stringArray[i-1,j-1] = xlRange.Cells[i, j].Value2.ToString();
                        

                    }     
                }
            }
            for (int i = 1; i <= numberOfRows; i++)
            {
                if(stringArray[i - 1, 6] != null)
                {
                    if(stringArray[i - 1, 6].Length > 7)
                    {
                        stringArray[i - 1, 6] = stringArray[i - 1, 6].Remove(7);
                    }
                }
          
               
            }
            
            string temp = String.Copy(stringArray[0, 0]);
            Console.WriteLine(temp);
            if (stringArray[3,3] == null)
            {
                Console.WriteLine("hello");
            }

            for (int i = 1; i <= numberOfRows-2; i++)
            {
                for (int j = 1; j <= numberOfCols; j++)
                {
                    inputArray[i-1,j-1] = Convert.ToDouble(stringArray[i + 1, j - 1]);
                    if ((stringArray[i + 1, j - 1] == null))
                    {
                        inputArray[i - 1, j - 1] = -1;
                    }
                }
            }

            /*for (int i = 1; i <= numberOfRows-2; i++)
            {
                for (int j = 1; j <= numberOfCols; j++)
                {
                    //dataGridView1.Rows[i - 1].Cells[j - 1].Value = stringArray[i - 1, j - 1];
                    dataGridView1.Rows[i - 1].Cells[j - 1].Value = inputArray[i - 1, j - 1];
                }
            }*/


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

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.componentSelect = "-1";
            Form3 f3 = new Form3(this);
            DialogResult ds = f3.ShowDialog();
            double loss = 0;
            
            if (this.componentSelect == "1")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Air Terminal Device (FPU/ACB)";
                loss = 0.1;
            }
            if (this.componentSelect == "2")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Fire/Smoke Damper";
                loss = 0.15;
            }
            if (this.componentSelect == "3")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Louver";
                loss = 0.12;
            }
            if (this.componentSelect == "4")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Air Terminal Device (VAV)";
                loss = 0.25;
            }
            if (this.componentSelect == "5")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Air Terminal (Lab) Valve";
                loss = 0.6;
            }
            if (this.componentSelect == "6")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Air Terminal (Lab) Valve- LP";
                loss = 0.3;
            }
            if (this.componentSelect == "7")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Airflow Station";
                loss = 0.3;
            }
            if (this.componentSelect == "8")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Reheat Coil Hydronic (2 Row)";
                loss = 0.15;
            }
            if (this.componentSelect == "9")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Reheat Coil (Electric)";
                loss = 0.05;
            }
            if (this.componentSelect == "10")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Sound Attenuator";
                loss = 0.1;
            }
            if (this.componentSelect == "11")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Humidifier";
                loss = 0.05;
            }
            if (this.componentSelect == "12")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Airflow Station";
                loss = 0.03;
            }
            if (this.componentSelect == "13")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Manual Volume Damper";
                loss = 0.08;
            }
            if (this.componentSelect == "14")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Automatic Control Damper";
                loss = 0.1;
            }
            if (this.componentSelect == "15")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Fire Damper";
                loss = 0.05;
            }
            if (this.componentSelect == "16")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Smoke Damper";
                loss = 0.1;
            }
            if (this.componentSelect == "17")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "UVGI Lamps";
                loss = 0.05;
            }
            if (this.componentSelect == "18")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Flex Duct (5 ft.)";
                loss = 0.0;
            }
            if (this.componentSelect == "19")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Flex Elbow";
                loss = 0.0;
            }
            if (this.componentSelect == "20")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Flex Connector";
                loss = 0.0;
            }
            if (this.componentSelect == "21")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Wire-Mesh Screen";
                loss = 0.03;
            }
            if (this.componentSelect == "22")
            {
                dataGridView1.Rows[endRow].Cells[2].Value = "Diffuser/ Grille";
                loss = 0.1;
            }
            if (this.componentSelect != "-1")

            {
                dataGridView1.Rows[endRow + 1].Cells[21].Value = Convert.ToDouble(dataGridView1.Rows[endRow].Cells[21].Value) + loss ;
                dataGridView1.Rows[endRow + 1].Cells[20].Value = "Total Loss:";
                dataGridView1.Rows[endRow].Cells[0].Value = endRow;
                dataGridView1.Rows[endRow].Cells[1].Value = "Component";
                //dataGridView1.Rows[endRow].Cells[2].Value = inputArray2[fitRow, 0];   //Description
                dataGridView1.Rows[endRow].Cells[3].Value = "N/A";   //CFM
                dataGridView1.Rows[endRow].Cells[4].Value = "N/A";   //width
                dataGridView1.Rows[endRow].Cells[5].Value = "N/A";   //height/diameter
                dataGridView1.Rows[endRow].Cells[6].Value = "N/A";   //duct length
                dataGridView1.Rows[endRow].Cells[7].Value = "N/A";   //material
                dataGridView1.Rows[endRow].Cells[8].Value = "N/A";   //roughness
                dataGridView1.Rows[endRow].Cells[9].Value = "N/A";   //ashrae number
                dataGridView1.Rows[endRow].Cells[10].Value = "N/A";   //fitting coefficient
                dataGridView1.Rows[endRow].Cells[11].Value = "N/A";   //component pressure loss
                dataGridView1.Rows[endRow].Cells[12].Value = "N/A";   //solve for diameter
                dataGridView1.Rows[endRow].Cells[13].Value = "N/A";   //velocity
                dataGridView1.Rows[endRow].Cells[14].Value = "N/A";   //velocity pressure
                dataGridView1.Rows[endRow].Cells[15].Value = "N/A";    //"Solve for Reynolds No. (Re)";
                dataGridView1.Rows[endRow].Cells[16].Value = "N/A";   //"Flow Type";
                dataGridView1.Rows[endRow].Cells[17].Value = "N/A";   //"Friction Factor (f) turbulent"
                dataGridView1.Rows[endRow].Cells[18].Value = "N/A";   //"Friction Factor (f) laminar"
                dataGridView1.Rows[endRow].Cells[19].Value = "N/A";   //"Friction Factor(fd)"
                dataGridView1.Rows[endRow].Cells[20].Value = "N/A";   //"Pressure Loss(in. of H20/100')"
                dataGridView1.Rows[endRow].Cells[21].Value = loss;   //"Total Pressure Loss(in. of H20)"
                endRow++;

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex == 7 && e.RowIndex >0)
            {
                this.radioSelect = "-1";
                Form2 f2 = new Form2(this);
                DialogResult ds = f2.ShowDialog();
                double roughness = 8.88;
                if(this.radioSelect == "1")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Uncolated carbon steel, clean";
                    roughness = 0.0001;
                }
                else if(this.radioSelect == "2")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "PVC plastic pipe";
                    roughness = 0.00015;
                }
                else if (this.radioSelect == "3")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Aluminum";
                    roughness = 0.0002;
                }
                else if (this.radioSelect == "4")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Galvanized steel, longitudinal seams, 4ft joins";
                    roughness = 0.003;
                }
                else if (this.radioSelect == "5")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Galvanized steel, continuously rolled, spiral seams, 10ft joints";
                    roughness = 0.004;
                }
                else if (this.radioSelect == "6")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Galvanized steel, spiral seam, with 1,2, and 3 ribs 12ft joints";
                    roughness = 0.00038;
                }
                else if (this.radioSelect == "7")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Galvanized steel, longitudinal seams, with 2.5ft joints";
                    roughness = 0.0005;
                }
                else if (this.radioSelect == "8")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Galvanized steel, spiral, corrugated, 10ft joints";
                    roughness = 0.003;
                }
                else if (this.radioSelect == "9")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Fibrous glass duct, air side with facing material";
                    roughness = 0.005;
                }
                else if (this.radioSelect == "10")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Fibrous glass duct, air side spray coated";
                    roughness = 0.01;
                }
                else if (this.radioSelect == "11")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Flexible duct, metallic (when fully extended)";
                    roughness = 0.007;
                }
                else if (this.radioSelect == "12")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Flexible duct, utilized in conjuction with the pressure loss correction factor as indicated in 2009 ASHRAE";
                    roughness = 0.003;
                }
                else if (this.radioSelect == "13")
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Concrete";
                    roughness = 0.01;
                }
                
                if (this.radioSelect == "-1")
                {

                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells[8].Value = roughness;
                    double diameter = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[12].Value);
                    double reynoldsNumber = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[15].Value);
                    string flowType = dataGridView1.Rows[e.RowIndex].Cells[16].Value.ToString();
                    double velocity = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[13].Value);
                    double velocityPressure = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[14].Value);
                    string ductElement = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    double fittingCoefficient = 0;
                    double ductLength = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[6].Value);
                    double componentPressureDrop = 0;
                    double frictionFactorTurbulent = calculateFrictionFactorTurbulent(diameter, reynoldsNumber, 1, roughness);
                    double frictionFactorLaminar = 64.0 / reynoldsNumber;
                    double frictionFactor = calculateFrictionFactor(flowType, frictionFactorLaminar, frictionFactorTurbulent);
                    double pressureLoss = calculatePressureLoss(ductElement, frictionFactor, diameter, velocity);
                    double totalPressureLoss = calculateTotalPressureLoss(ductElement, fittingCoefficient, velocityPressure, ductLength, componentPressureDrop, pressureLoss);
                    dataGridView1.Rows[e.RowIndex].Cells[17].Value = frictionFactorTurbulent;
                    dataGridView1.Rows[e.RowIndex].Cells[18].Value = frictionFactorLaminar;
                    dataGridView1.Rows[e.RowIndex].Cells[19].Value = frictionFactor;
                    dataGridView1.Rows[e.RowIndex].Cells[20].Value = pressureLoss;
                    dataGridView1.Rows[e.RowIndex].Cells[21].Value = totalPressureLoss;
                }
            }
            
        }

        public string radioSelect
        {
            get { return label1.Text; }
            set { label1.Text = value; }
        }

        public string componentSelect
        {
            get { return label2.Text; }
            set { label2.Text = value; }
        }

        

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                int delRow = Convert.ToInt32(this.textBox1.Text);
                Console.WriteLine(delRow);
                for (int i = delRow; i < endRow + 2; i++)
                {
                    for (int j = 1; j < 22; j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Value = dataGridView1.Rows[i + 1].Cells[j].Value;
                        dataGridView1.Rows[i].Cells[j].Style.BackColor = dataGridView1.Rows[i + 1].Cells[j].Style.BackColor;
                    }
                }
                endRow--;
                for (int j = 0; j < 21; j++)
                {
                    dataGridView1.Rows[endRow].Cells[j].Value = "";
                    dataGridView1.Rows[endRow].Cells[j].Style.BackColor = Color.White;
                }
                double totalLoss = 0;
                for (int i = 1; i < endRow; i++)
                {
                    totalLoss += Convert.ToDouble(dataGridView1.Rows[i].Cells[21].Value);
                }

                dataGridView1.Rows[endRow].Cells[21].Value = totalLoss;
                dataGridView1.Rows[endRow].Cells[20].Value = "Total Loss:";
                
            }
            catch (System.FormatException)
            {
                this.textBox1.Text = "invalid input";
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if(ductRun == 0)
            {
                numberOfRows = 2;
                numberOfCols = 22;
                inputArray = new double[numberOfRows - 1, numberOfCols];
            }

            if (fitRun == 0)
            {
                numRowIn2 = 2;
                numColIn2 = 22;
                inputArray2 = new string[numRowIn2 - 1, numColIn2];
            }

            numberOfRows -= 1;
            string[,] outputArray = new string[numberOfRows+2, 22];
            dataGridView1.ColumnCount = 22;
            dataGridView1.RowCount = numberOfRows+numRowIn2+4;
            outputArray[0, 0] = "Identifier";
            outputArray[0, 1] = "Duct Element";
            outputArray[0, 2] = "Description";
            outputArray[0, 3] = "CFM";
            outputArray[0, 4] = "Width(zero if round), (Inches)";
            outputArray[0, 5] = "Height or Diameter (Inches)";
            outputArray[0, 6] = "Duct Length (Feet)";
            outputArray[0, 7] = "Material";
            outputArray[0, 8] = "Roughness";
            outputArray[0, 9] = "ASHRAE Fitting No.";
            outputArray[0, 10] = "Fitting Coefficient(C_o)";
            outputArray[0, 11] = "Component Pressure drop(in. of H2O)";
            outputArray[0, 12] = "Solve for Diameter(Dh)";
            outputArray[0, 13] = "Velocity(FPM)";
            outputArray[0, 14] = "Vel. Press.(in. of H20)";
            outputArray[0, 15] = "Solve for Reynolds No. (Re)";
            outputArray[0, 16] = "Flow Type";
            outputArray[0, 17] = "Friction Factor (f) turbulent";
            outputArray[0, 18] = "Friction Factor (f) laminar";
            outputArray[0, 19] = "Friction Factor(fd)";
            outputArray[0, 20] = "Pressure Loss(in. of H20/100')";
            outputArray[0, 21] = "Total Pressure Loss(in. of H20)";
            



            //Set all, calculate all, save all in result array, then iterate to next row
            for (int j = 0; j < numberOfRows; j++)
            {
                Console.WriteLine(j);

                //SET PARAMETERS FROM INPUT
                setDuctScheduleInputParameters(j);
                setAshraeFittingNumber(j);
                //setDescriptionAndDuctElement(ashraeFittingNumber, ashraeFittingNumbers, descriptions, ductElements);
                setRoughness(j);                            //column K from material
                setFittingCoefficient(j);                   //column M
                setComponentPressureDrop(j);                //column N from description

                //CALCULATE PARAMETERS
                double roughness = 0.0005;
                double diameter = calculateDiameter(width, height, j);
                double velocity = calculateVelocity(cfm, width, height, diameter);
                double velocityPressure = calculateVelocityPressure(velocity);
                double reynoldsNumber = calculateReynoldsNumber(diameter, velocity);
                string flowType = calculateFlowType(reynoldsNumber);
                double frictionFactorTurbulent = calculateFrictionFactorTurbulent(diameter, reynoldsNumber, 1, roughness);
                double frictionFactorLaminar = 64.0 / reynoldsNumber;
                double frictionFactor = calculateFrictionFactor(flowType, frictionFactorLaminar, frictionFactorTurbulent);
                double pressureLoss = calculatePressureLoss(ductElement, frictionFactor, diameter, velocity);
                double totalPressureLoss = calculateTotalPressureLoss(ductElement, fittingCoefficient, velocityPressure, ductLength, componentPressureDrop, pressureLoss);
                


                outputArray[j + 1, 0] = (j + 1).ToString();
                outputArray[j + 1, 1] = "Duct";
                outputArray[j + 1, 2] = "N/A";
                outputArray[j + 1, 3] = inputArray[j, 0].ToString();
                outputArray[j + 1, 4] = inputArray[j, 1].ToString();
                outputArray[j + 1, 5] = inputArray[j, 2].ToString();     //height or diamete
                if (inputArray[j, 1] == -1)
                {
                    outputArray[j + 1, 5] = inputArray[j, 3].ToString();
                }
                outputArray[j + 1, 6] = inputArray[j, 4].ToString();
                outputArray[j + 1, 7] = "Galvanized Steel";
                outputArray[j + 1, 8] = roughness.ToString();
                outputArray[j + 1, 9] = "N/A";
                outputArray[j + 1, 10] = "N/A";
                outputArray[j + 1, 11] = "N/A";
                outputArray[j + 1, 12] = Math.Round(diameter, 2).ToString();
                outputArray[j + 1, 13] = Math.Round(velocity, 0).ToString();
                outputArray[j + 1, 14] = Math.Round(velocityPressure, 2).ToString();
                outputArray[j + 1, 15] = Math.Round(reynoldsNumber, 0).ToString();
                outputArray[j + 1, 16] = flowType;
                outputArray[j + 1, 17] = Math.Round(frictionFactorTurbulent, 3).ToString();
                outputArray[j + 1, 18] = Math.Round(frictionFactorLaminar, 3).ToString();
                outputArray[j + 1, 19] = Math.Round(frictionFactor, 3).ToString();
                outputArray[j + 1, 20] = Math.Round(pressureLoss, 3).ToString();
                outputArray[j + 1, 21] = Math.Round(totalPressureLoss, 3).ToString();
                dataGridView1.Rows[j+1].Cells[13].Value = Math.Round(velocity, 0);   //velocity






            }
            for (int i = 1; i <= numberOfRows; i++)
            {
                for (int j = 1; j <= 22; j++)
                {
                    if (j == 14 && i != 1)
                    {
                        continue;
                    }
                    //dataGridView1.Rows[i - 1].Cells[j - 1].Value = stringArray[i - 1, j - 1];
                    string temp = outputArray[i - 1, j - 1];
                    dataGridView1.Rows[i - 1].Cells[j - 1].Value = temp;
                }
            }

            //creating output for fitting pieces
            int fitRow = 0;
            for (int i = numberOfRows; i < numberOfRows + numRowIn2 -2; i++)
            {
                double width = setWidth(fitRow, inputArray2[fitRow, 2]);
                double cfm = setCFM(fitRow, inputArray2[fitRow, 2]);
                double height = setHeight(fitRow, inputArray2[fitRow, 2]);
                if(height == -1)
                {
                    height = setDiameter(fitRow, inputArray2[fitRow, 2]);
                }
                

                double diameter;
                if (width > 0)
                {
                    diameter = (double)4 * width * height / (2 * (width + height));
                    //=IF(G8>0,4*G8*H8/((G8+H8)*2),H8)
                }
                else
                {
                    diameter = height;
                }
                //string description = getDescription(inputArray2[fitRow, 2]); 
                if (inputArray2[fitRow, 2] == "SD4-1")
                {
                    height = Convert.ToDouble(inputArray2[fitRow, 10]);
                    diameter = Convert.ToDouble(inputArray2[fitRow, 10]);
                }
                double velocity = calculateVelocity(cfm,width,height,diameter);
                double velocity_pressure = calculateVelocityPressure(velocity);
                double reynolds = calculateReynoldsNumber(diameter, velocity);
                string flowType = calculateFlowType(reynolds);
                double frictionFactorTurbulent = calculateFrictionFactorTurbulent(diameter, reynolds, 1, 0.0005);
                double frictionFactorLaminar = 64.0 / reynolds;
                double frictionFactor = calculateFrictionFactor(flowType, frictionFactorLaminar, frictionFactorTurbulent);
                double pressureLoss = calculatePressureLoss("Fitting", frictionFactor, diameter, velocity);
                double fittingCoef = getStaticCoefficient(fitRow, inputArray2[fitRow, 2]);
                double totalPressureLoss = calculateTotalPressureLoss("Fitting", fittingCoef, velocity_pressure, 1, componentPressureDrop, pressureLoss);
                
               

                dataGridView1.Rows[i].Cells[0].Value = i;                   //List number
                dataGridView1.Rows[i].Cells[1].Value = "Fitting";           //Hard coded as Fitting
                dataGridView1.Rows[i].Cells[2].Value = inputArray2[fitRow, 0];   //Description
                dataGridView1.Rows[i].Cells[3].Value = cfm;   //CFM
                dataGridView1.Rows[i].Cells[4].Value = width;   //width
                dataGridView1.Rows[i].Cells[5].Value = height;   //height/diameter
                dataGridView1.Rows[i].Cells[6].Value = 1;   //duct length
                dataGridView1.Rows[i].Cells[7].Value = "Galvanized Steel";   //material
                dataGridView1.Rows[i].Cells[8].Value = 0.0005;   //roughness
                dataGridView1.Rows[i].Cells[9].Value = inputArray2[fitRow,2];   //ashrae number
                dataGridView1.Rows[i].Cells[10].Value = Math.Round(fittingCoef, 2);   //fitting coefficient
                dataGridView1.Rows[i].Cells[11].Value = "N/A";   //component pressure loss
                dataGridView1.Rows[i].Cells[12].Value = Math.Round(diameter, 2);   //solve for diameter
                dataGridView1.Rows[i].Cells[13].Value = Math.Round(velocity, 0);   //velocity
                dataGridView1.Rows[i].Cells[14].Value = Math.Round(velocity_pressure, 2);   //velocity pressure
                dataGridView1.Rows[i].Cells[15].Value = Math.Round(reynolds, 0);   //"Solve for Reynolds No. (Re)";
                dataGridView1.Rows[i].Cells[16].Value = flowType;   //"Flow Type";
                dataGridView1.Rows[i].Cells[17].Value = Math.Round(frictionFactorTurbulent, 3);   //"Friction Factor (f) turbulent"
                dataGridView1.Rows[i].Cells[18].Value = Math.Round(frictionFactorLaminar, 3);   //"Friction Factor (f) laminar"
                dataGridView1.Rows[i].Cells[19].Value = Math.Round(frictionFactor, 3);   //"Friction Factor(fd)"
                dataGridView1.Rows[i].Cells[20].Value = "N/A";   //"Pressure Loss(in. of H20/100')"
                dataGridView1.Rows[i].Cells[21].Value = Math.Round(totalPressureLoss, 3);   //"Total Pressure Loss(in. of H20)"



                fitRow++;
            }

            for(int i =1; i< numberOfRows + numRowIn2 - 2; i++)
            {
                if(dataGridView1.Rows[i].Cells[4].Value.ToString() == "-1")
                {
                    dataGridView1.Rows[i].Cells[4].Value = 0;
                }
            }

            for (int i = 1; i < numberOfRows + numRowIn2 - 2; i++)
            {
                if(dataGridView1.Rows[i].Cells[13].Value.GetType() == 1.0.GetType())
                {
                    if (Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value) > 600)
                    {
                        dataGridView1.Rows[i].Cells[13].Style.BackColor = Color.FromArgb(255, 153, 153);
                    }
                    if (Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value) > 899)
                    {
                        dataGridView1.Rows[i].Cells[13].Style.BackColor = Color.Red;
                    }
                }
               
            }

            //add total row

            //numberOfRows + numRowIn2 - 2
            double totalLoss = 0;
            endRow = numberOfRows + numRowIn2 - 2;
            for (int i =1; i< numberOfRows + numRowIn2 - 2; i++)
            {
                totalLoss += Convert.ToDouble(dataGridView1.Rows[i].Cells[21].Value);
            }

            dataGridView1.Rows[numberOfRows + numRowIn2 - 2].Cells[21].Value = totalLoss;
            dataGridView1.Rows[numberOfRows + numRowIn2 - 2].Cells[20].Value = "Total Loss:";



            /*   for (int i = 0; i <= 22; i++)
               {

                   Console.WriteLine("Console.WriteLine(\"dataGridView1.Rows[\"+i+\"].Cells[" + i + "].Value = \"\";\");");
               }*/


            //Convert results to strings
            //int[] intarray = { 1, 2, 3, 4, 5 };
            //string[] result = intarray.Select(x=>x.ToString()).ToArray();
        }

        /*
	* ----------------------------------
	* SETTER METHODS
	* ----------------------------------
	*/

        
        static public double linearInterpolation(double x, double x0, double x1, double y0, double y1)
        {
            if ((x1 - x0) == 0)
            {
                return (y0 + y1) / 2;
            }
            return y0 + (x - x0) * (y1 - y0) / (x1 - x0);
        }

        static public double bilinearInterpolation(double q11, double q12, double q21, double q22, double x1, double x2, double y1, double y2, double x, double y)
        {
            /*double x2x1, y2y1, x2x, y2y, yy1, xx1;
            x2x1 = x2 - x1;
            y2y1 = y2 - y1;
            x2x = x2 - x;
            y2y = y2 - y;
            yy1 = y - y1;
            xx1 = x - x1;
            return 1.0 / (x2x1 * y2y1) * (
                q11 * x2x * y2y +
                q21 * xx1 * y2y +
                q12 * x2x * yy1 +
                q22 * xx1 * yy1
                );*/

            double R1 = ((x2 - x) / (x2 - x1)) * q11 + ((x - x1) / (x2 - x1)) * q21;

            double R2 = ((x2 - x) / (x2 - x1)) * q12 + ((x - x1) / (x2 - x1)) * q22;
            return ((y2 - y) / (y2 - y1)) * R1 + ((y - y1) / (y2 - y1)) * R2;
        }
        public static double interpolateData(double x, double y, String aFN)
        {
            double[,] coArray;
            switch (aFN)
            {
                case "SD4_1":
                    coArray = SD4_1;
                    break;
                case "SD4_2":
                    coArray = SD4_2;
                    break;
                case "SD5_10_Cs":
                    coArray = SD5_10_Cs;
                    break;
                case "SD5_10_Cb":
                    coArray = SD5_10_Cb;
                    break;
                case "SR4_1":
                    coArray = SR4_1;
                    break;
                case "SR5_13_Cs":
                    coArray = SR5_13_Cs;
                    break;
                case "SR5_13_Cb":
                    coArray = SR5_13_Cb;
                    break;
                case "SR5_14":
                    coArray = SR5_14;
                    break;
                case "CD3_1":
                    coArray = CD3_1;
                    break;
                case "CR3_1":
                    coArray = CR3_1;
                    break;
                case "SR4_3":
                    coArray = SR4_3;
                    break;
                case "SR4_2":
                    coArray = SR4_2;
                    break;
                case "SD5_9_Cs":
                    coArray = SD5_9_Cs;
                    break;
                case "SD5_9_Cb":
                    coArray = SD5_9_Cb;
                    break;
                case "SD5_18_Cs":
                    coArray = SD5_18_Cs;
                    break;
                case "SD5_18_Cb":
                    coArray = SD5_18_Cb;
                    break;
                case "SR5_12_Cs":
                    coArray = SR5_12_Cs;
                    break;
                case "SR5_12_Cb":
                    coArray = SR5_12_Cb;
                    break;
                case "SR5_11_Cs":
                    coArray = SR5_11_Cs;
                    break;
                case "SR5_11_Cb":
                    coArray = SR5_11_Cb;
                    break;
                case "CD3_4":
                    coArray = CD3_4;
                    break;
                case "CD3_2":
                    coArray = CD3_2;
                    break;
                case "CR3_1_K":
                    coArray = CR3_1_K;
                    break;


                default:
                    return -7;
                    break;


            }
            double result = -6;
            double in1 = x;
            double in2 = y;
            int low_1 = -1;
            int high_1 = -1;
            int low_2 = -1;
            int high_2 = -1;

            for (int i = 1; i < coArray.GetLength(1) - 1; i++)
            {
                if (coArray[0, i] == in1)
                {
                    low_1 = i;
                    high_1 = -1;

                    break;

                }
                if (coArray[0, i] < in1 && coArray[0, i + 1] > in1)
                {
                    low_1 = i;
                    high_1 = i + 1;
                    break;
                }
                if (coArray[0, i + 1] == in1)
                {
                    low_1 = i + 1;
                    high_1 = -1;

                    break;
                }
            }


            if (coArray[0, 1] > in1)
            {
                low_1 = 1;
                high_1 = -1;
            }
            if (coArray[0, coArray.GetLength(1) - 1] < in1)
            {
                low_1 = coArray.GetLength(1) - 1;
                high_1 = -1;
            }

            for (int i = 1; i < coArray.GetLength(0) -1; i++)
            {
                if (coArray[i, 0] == in2)
                {
                    low_2 = i;
                    high_2 = -1;
                    break;
                }
                if (coArray[i, 0] < in2 && coArray[i + 1, 0] > in2)
                {
                    low_2 = i;
                    high_2 = i + 1;
                    break;
                }
                if (coArray[i + 1, 0] == in2)
                {
                    low_2 = i + 1;
                    high_2 = -1;
                    break;
                }
            }


            if (coArray[1, 0] > in2)
            {
                low_2 = 1;
                high_2 = -1;
            }
            if (coArray[coArray.GetLength(0) - 1, 0] < in2)
            {
                low_2 = coArray.GetLength(0) - 1;
                high_2 = -1;
            }

            if (low_1 == -1 && high_1 == -1 && low_2 == -1 && high_2 == -1)
            {
                result = -5;
            }

            else if (high_1 == -1 && high_2 == -1)
            {
                result = coArray[low_2, low_1];
            }

            else if (high_1 == -1)
            {
                result = linearInterpolation(in2, coArray[low_2, 0], coArray[high_2, 0], coArray[low_2, low_1], coArray[high_2, low_1]);
            }
            else if (high_2 == -1)
            {
                result = linearInterpolation(in1, coArray[0, low_1], coArray[0, high_1], coArray[low_2, low_1], coArray[low_2, high_1]);
            }
            else
            {
                result = bilinearInterpolation(coArray[low_2, low_1], coArray[high_2, low_1], coArray[low_2, high_1], coArray[high_2, high_1], coArray[0, low_1], coArray[0, high_1], coArray[low_2, 0], coArray[high_2, 0], in1, in2);
            }
            return result;
        }
        public double getStaticCoefficient(int row, String aFN)
        {
            // switch case using ashraeFittingNumber to call the right method

            //string aFN = ashraeFittingNumber;

            double D;
            double D0;
            double D1;
            double theta;
            double H1;
            double W1;
            double Dc;
            double Ds;
            double Db;
            double Qb;
            double Qc;
            double Qs;
            double H0;
            double W;
            double H;
            double L;
            double Hs;
            double Ws;
            double Hb;
            double Wb;
            double Wc;
            double Db1;
            double Db2;
            double Qb1;
            double R;
            double y_axis;
            double x_axis;
            double x_axis_1;
            double y_axis_1;
            double W0;
            double Cp;
            double K;


            double Cb;
            double Cs;
            double C0 = -1;      //what the method returns

            switch (aFN)
            {
                case "SD4-1":
                    D0 = Convert.ToDouble(inputArray2[row, 10]);
                    D1 = Convert.ToDouble(inputArray2[row, 11]);
                    theta = Convert.ToDouble(inputArray2[row, 12]);
                    y_axis = Math.Pow((D0 / 2), 2) / Math.Pow((D1 / 2), 2);
                    C0 = interpolateData(theta, y_axis, "SD4_1");
                    break;
                case "SD4-2":
                    theta = Convert.ToDouble(inputArray2[row, 12]);
                    D0 = Convert.ToDouble(inputArray2[row, 13]);
                    H1 = Convert.ToDouble(inputArray2[row, 14]);
                    W1 = Convert.ToDouble(inputArray2[row, 15]);
                    y_axis = Math.PI * Math.Pow((D0 / 2), 2) / (W1 * H1);
                    C0 = interpolateData(theta, y_axis, "SD4_2");
                    break;
                case "SD5-10":
                    Dc = Convert.ToDouble(inputArray2[row, 19]);
                    Ds = Convert.ToDouble(inputArray2[row, 20]);
                    Db = Convert.ToDouble(inputArray2[row, 21]);
                    Qb = Convert.ToDouble(inputArray2[row, 29]);
                    Qc = Convert.ToDouble(inputArray2[row, 28]);
                    Qs = (Qc - Qb);

                    //Using Table Cb
                    x_axis = Qb / Qc;
                    y_axis = Math.Pow((Db / 2), 2) / Math.Pow((Dc / 2), 2);
                    Cb = interpolateData(x_axis, y_axis, "SD5_10_Cb");

                    //Using Table Cs
                    x_axis_1 = Qs / Qc;
                    y_axis_1 = Math.Pow((Ds / 2), 2) / Math.Pow((Dc / 2), 2);
                    Cs = interpolateData(x_axis_1, y_axis_1, "SD5_10_Cs");

                    //Evaluate C0 according to the branch factor
                    if (Cb == 0 && Cs == 0)
                    {
                        C0 = 0;
                    }
                    else if (inputArray2[row, 30] == "Yes" || inputArray2[row, 30] == "yes")
                    {
                        C0 = Cb;
                    }
                    else
                    {
                        C0 = Cs;
                    }

                    break;
                case "SR4-1":
                    H0 = Convert.ToDouble(inputArray2[row, 14]);
                    W = Convert.ToDouble(inputArray2[row, 15]);
                    H1 = Convert.ToDouble(inputArray2[row, 16]);
                    L = Convert.ToDouble(inputArray2[row, 18]);
                    theta = 2 * Math.Atan(Math.Abs((H0 - H1) / (2 * L))) * (180 / Math.PI);
                    y_axis = H0 / H1;
                    C0 = interpolateData(theta, y_axis, "SR4_1");
                    break;
                case "SR5-13":
                    H = Convert.ToDouble(inputArray2[row, 22]);
                    W = Convert.ToDouble(inputArray2[row, 23]);
                    Hs = Convert.ToDouble(inputArray2[row, 24]);
                    Ws = Convert.ToDouble(inputArray2[row, 25]);
                    Hb = Convert.ToDouble(inputArray2[row, 26]);
                    Wb = Convert.ToDouble(inputArray2[row, 27]);
                    Qc = Convert.ToDouble(inputArray2[row, 28]);
                    Qb = Convert.ToDouble(inputArray2[row, 29]);

                    //Using Table Cb
                    x_axis = Qb / Qc;
                    y_axis = (Hb * Wb) / (H * W);
                    Cb = interpolateData(x_axis, y_axis, "SR5_13_Cb");

                    //Using Table Cs
                    x_axis_1 = (Qc-Qb) / Qc;
                    y_axis_1 = (Hs * Ws) / (H * W);
                    Cs = interpolateData(x_axis, y_axis, "SR5_13_Cs");

                    //Evaluate C0 according to the branch factor
                    if (Cb == 0 && Cs == 0)
                    {
                        C0 = 0;
                    }
                    else if (inputArray2[row, 30] == "Yes" || inputArray2[row, 30] == "yes")
                    {
                        C0 = Cb;
                    }
                    else
                    {
                        C0 = Cs;
                    }

                    break;
                case "SR5-14":
                    H = Convert.ToDouble(inputArray2[row, 36]);
                    Wb = Convert.ToDouble(inputArray2[row, 37]);
                    Wc = Convert.ToDouble(inputArray2[row, 38]);
                    C0 = 1.4 * (H * Wb) / (H * Wc) - 0.4;
                    break;
                case "CD3-1":
                    D = Convert.ToDouble(inputArray2[row, 5]);
                    C0 = interpolateData(D, 10, "CD3_1");
                    break;
                case "CR3-1":
                    W = Convert.ToDouble(inputArray2[row, 6]);
                    H = Convert.ToDouble(inputArray2[row, 7]);
                    R = Convert.ToDouble(inputArray2[row, 8]);
                    theta = Convert.ToDouble(inputArray2[row, 9]);

                    //Table Cp
                    x_axis = H / W;
                    y_axis = R / W;
 
                    Cp = interpolateData(x_axis, y_axis, "CR3_1");
                    //Table K
                    K = interpolateData(theta, 10, "CR3_1_K");
                    //Evaluate C0
                    C0 = Cp * K;
                    

                    break;
                case "CR3-16":
                    C0 = 0.41;
                    break;
                case "SR4-3":
                    D1 = Convert.ToDouble(inputArray2[row, 13]);
                    H0 = Convert.ToDouble(inputArray2[row, 14]);
                    W0 = Convert.ToDouble(inputArray2[row, 15]);
                    theta = Convert.ToDouble(inputArray2[row, 12]);
                    y_axis = (H0 * W0) / (Math.PI * Math.Pow((D1 / 2), 2));
                    C0 = interpolateData(theta, y_axis, "SR4_3"); 
                    /*if (y_axis > 0.5 && y_axis < 1)
                    {
                        C0 = 0;
                    }
                    else
                    {
                        C0 = interpolateData(theta, y_axis, "SR4_3");
                    }*/
                    break;
                case "SR4-2":
                    //theta = Convert.ToDouble(inputArray2[row, 12]);
                    //2*ATAN(ABS(B252-B254)/(2*B256))*(180/PI())
                    
                    H0 = Convert.ToDouble(inputArray2[row, 14]);
                    W0 = Convert.ToDouble(inputArray2[row, 15]);
                    H1 = Convert.ToDouble(inputArray2[row, 16]);
                    W1 = Convert.ToDouble(inputArray2[row, 17]);
                    L = Convert.ToDouble(inputArray2[row, 18]);
                    double theta1 = 2 * Math.Atan(Math.Abs(H0 - H1) / (2 * L)) * (180 / Math.PI);
                    double theta2 = 2 * Math.Atan(Math.Abs(W0 - W1) / (2 * L)) * (180 / Math.PI);
                    theta = Math.Max(theta1, theta2);
                    y_axis = (W1 * H1) / (H0 * W0);
                    C0 = interpolateData(theta, y_axis, "SR4_2");
                    break;
                case "SD5-9":
                    Dc = Convert.ToDouble(inputArray2[row, 19]);
                    Ds = Convert.ToDouble(inputArray2[row, 20]);
                    Db = Convert.ToDouble(inputArray2[row, 21]);
                    Qb = Convert.ToDouble(inputArray2[row, 29]);
                    Qc = Convert.ToDouble(inputArray2[row, 28]);
                    Qs = (Qc - Qb);

                    //Using Table Cb
                    x_axis = Qb / Qc;
                    y_axis = Math.Pow((Db / 2), 2) / Math.Pow((Dc / 2), 2);
                    Cb = interpolateData(x_axis, y_axis, "SD5_9_Cb");

                    //Using Table Cs
                    x_axis_1 = Qs / Qc;
                    y_axis_1 = Math.Pow((Ds / 2), 2) / Math.Pow((Dc / 2), 2);
                    Cs = interpolateData(x_axis_1, y_axis_1, "SD5_9_Cs");

                    //Evaluate C0 according to the branch factor
                    if (Cb == 0 && Cs == 0)
                    {
                        C0 = 0;
                    }
                    else if (inputArray2[row, 30] == "Yes" || inputArray2[row, 30] == "yes")
                    {
                        C0 = Cb;
                    }
                    else
                    {
                        C0 = Cs;
                    }

                    break;
                case "SD5-18":
                    Dc = Convert.ToDouble(inputArray2[row, 31]);
                    Db1 = Convert.ToDouble(inputArray2[row, 32]);
                    Db2 = Convert.ToDouble(inputArray2[row, 33]);
                    Qb1 = Convert.ToDouble(inputArray2[row, 34]);
                    Qc = Convert.ToDouble(inputArray2[row, 35]);

                    //Using Table Cb
                    x_axis = Qb1 / Qc;
                    y_axis = (Math.Pow((Db1 / 2), 2)*Math.PI) / (Math.Pow((Dc / 2), 2) * Math.PI);
                    Cb = interpolateData(x_axis, y_axis, "SD5_18_Cb");

                    //Using Table Cs
                    x_axis_1 = 1 - (Qb1 / Qc);
                    y_axis_1 = (Math.Pow((Db2 / 2), 2) * Math.PI) / (Math.Pow((Dc / 2), 2) * Math.PI);
                    Cs = interpolateData(x_axis_1, y_axis_1, "SD5_18_Cs");

                    //Evaluate C0 according to the branch factor
                    if (Cb == 0 && Cs == 0)
                    {
                        C0 = 0;
                    }
                    else if (inputArray2[row, 30] == "Yes" || inputArray2[row, 30] == "yes")
                    {
                        C0 = Cb;
                    }
                    else
                    {
                        C0 = Cs;
                    }

                    break;
                case "SR5-12":
                    H = Convert.ToDouble(inputArray2[row, 22]);
                    W = Convert.ToDouble(inputArray2[row, 23]);
                    Hs = Convert.ToDouble(inputArray2[row, 24]);
                    Ws = Convert.ToDouble(inputArray2[row, 25]);
                    Db = Convert.ToDouble(inputArray2[row, 21]);
                    Qc = Convert.ToDouble(inputArray2[row, 28]);
                    Qb = Convert.ToDouble(inputArray2[row, 29]);

                    //Using Table Cb
                    x_axis = Qb / Qc;
                    y_axis = Math.Pow((Db / 2), 2) / (H * W) * Math.PI;
                    Cb = interpolateData(x_axis, y_axis, "SR5_12_Cb");

                    //Using Table Cs
                    x_axis_1 = (Qc - Qb) / Qc;
                    y_axis_1 = (Hs * Ws) / (H * W);
                    Cs = interpolateData(x_axis_1, y_axis_1, "SR5_12_Cs");

                    //Evaluate C0 according to the branch factor
                    if (Cb == 0 && Cs == 0)
                    {
                        C0 = 0;
                    }
                    else if (inputArray2[row, 30] == "Yes" || inputArray2[row, 30] == "yes")
                    {
                        C0 = Cb;
                    }
                    else
                    {
                        C0 = Cs;
                    }

                    break;
                case "SR5-11":
                    H = Convert.ToDouble(inputArray2[row, 22]);
                    W = Convert.ToDouble(inputArray2[row, 23]);
                    Hs = Convert.ToDouble(inputArray2[row, 24]);
                    Ws = Convert.ToDouble(inputArray2[row, 25]);
                    Db = Convert.ToDouble(inputArray2[row, 21]);
                    Qc = Convert.ToDouble(inputArray2[row, 28]);
                    Qb = Convert.ToDouble(inputArray2[row, 29]);

                    //Using Table Cb
                    x_axis = Qb / Qc;
                    y_axis = Math.Pow((Db / 2), 2) / (H * W) * Math.PI;
                    Cb = interpolateData(x_axis, y_axis, "SR5_11_Cb");

                    //Using Table Cs
                    x_axis_1 = (Qc - Qb) / Qc;
                    y_axis_1 = (Hs * Ws) / (H * W);
                    Cs = interpolateData(x_axis_1, y_axis_1, "SR5_11_Cs");

                    //Evaluate C0 according to the branch factor
                    if (Cb == 0 && Cs == 0)
                    {
                        C0 = 0;
                    }
                    else if (inputArray2[row, 30] == "Yes" || inputArray2[row, 30] == "yes")
                    {
                        C0 = Cb;
                    }
                    else
                    {
                        C0 = Cs;
                    }

                    break;
                case "CD3-4":
                    D = Convert.ToDouble(inputArray2[row, 5]);
                    

                    C0 = interpolateData(D, 1, "CD3_4");
                    break;
                case "CD3-2":
                    D = Convert.ToDouble(inputArray2[row, 5]);
                   
                    C0 =  interpolateData(D, 1, "CD3_2");
                    break;

                default:
                    //ERROR - throw an exception
                    C0 = -1;
                    break;
            }

            return C0;

        }

        public void setAshraeFittingNumber(int row)
        {
            //set ashraeFittingNumber to the number imported
            //ashrae fitting number is column 3
            //ashraeFittingNumber = inputArray[row, 3];

        }

        public String getDescription(string aFN)
        {
            //description
            //=IFERROR( IF( AND (D8="Duct",G8=0),"Round Duct",IF( AND(D8="Duct",G8>1),"Rectangular Duct",IF(D8="Component","-",IF(D8="Fitting",INDEX('Fitting Name'!$B$218:
            description = "";
            for (int i = 0; i<ashraeFittingNumbers.Count; i++)
            {
                if (ashraeFittingNumbers[i] == aFN)
                {
                    description = descriptions[i];
                }
            }
            return description;
        }

        public bool checkWHD(int row)
        {
            //true <= width and height are set
            //false <= only diameter is set
            if (inputArray[row, 3] == -1)
            {
                if (inputArray[row, 1] != -1 && inputArray[row, 2] != -1)
                {
                    checking = true;
                }
                else
                {
                    //throw an error
                }
            }
            else
            {
                if (inputArray[row, 1] == -1 && inputArray[row, 2] == -1)
                {
                    checking = false;
                }
                else
                {
                    //throw an error
                }
            }

            return checking;
        }

        public void setDuctScheduleInputParameters(int row)
        {
            cfm = Convert.ToDouble(inputArray[row, 0]);
            width = Convert.ToDouble(inputArray[row, 1]);
            height = Convert.ToDouble(inputArray[row, 2]);
            diameter = Convert.ToDouble(inputArray[row, 3]);
            ductLength = Convert.ToDouble(inputArray[row, 4]);
            staticPressure = Convert.ToDouble(inputArray[row, 5]);
        }

        public void setRoughness(int row)
        {
            //match from lists using material chosen *?
            //=INDEX('Common Tables'!$C$3:$C$16,MATCH(J8,'Common Tables'!$B$3:$B$16,0))

        }

        public void setFittingCoefficient(int row)
        {
            //match depending on pressure drop value - lists
            //=IF(D8="Duct",INDEX('Pressure Calc- Ducts'!$M$7:$M$314,MATCH(C8,'Pressure Calc- Ducts'!$C$7:$C$314,0)),IF(D8="Fitting",INDEX('Pressure Calc- Fittings'!$M$7:$M$314,MATCH(C8,'Pressure Calc- Fittings'!$C$7:$C$314,0)),))
        }

        public void setComponentPressureDrop(int row)
        {
            //from input table - duct schedule
            //=IFERROR(IF(D8="Component",INDEX('Common Tables'!$I$4:$I$24,MATCH(E8,'Common Tables'!$H$4:$H$24,0)),),"")
        }


        /*
        * ----------------------------------
        * CALCULATE METHODS
        * ----------------------------------
        */

        public double calculateDiameter(double width, double height, int row)
        {
            double diameter;
            if (width > 0)
            {
                diameter = (double)4 * width * height / (2 * (width + height));
                //=IF(G8>0,4*G8*H8/((G8+H8)*2),H8)
            }
            else
            {
                diameter = inputArray[row, 3];
            }
            return diameter;
        }



        public double calculateVelocity(double cfm, double width, double height, double diameter)
        {
            double velocity;
            if (width > 0)
            {
                velocity = (double)cfm / (width * height / 144);
                //=IFERROR(IF(G8>0,F8/(G8*H8/144), F8/(PI()*(H8/12)^2/4)),0)
            }
            else
            {
                velocity = (double)cfm / (PI * Math.Pow(diameter / 12, 2) / 4);
            }
            return velocity;
        }

        //input: calculateVelocity result
        public double calculateVelocityPressure(double velocity)
        {
            double velPressure = DENSITYDRYAIR * Math.Pow((velocity / 1097), 2);
            //=$L$3*(P8/1097)^2

            return velPressure;
        }

        /* ----------------------------------
        * FLUID PROPERTIES
        * RN inputs are results of functions: calculateDiameter, calculateVelocity
        * FlowType inputs: RN
        * -----------------------------------
        */
        public static double calculateReynoldsNumber(double diameter, double velocity)
        {
            double rn = diameter * velocity / (KINEMATICVISCOSITY2 * 72 * 0.001);
            //=O8*P8/($N$3*72*0.001)

            return rn;
        }

        //input: RN result
        public String calculateFlowType(double reynoldsNumber)
        {
            String flowType;
            if (reynoldsNumber < 2300)
            {
                flowType = "Laminar";
                //=IF(R8<2300,"Laminar",IF(AND(R8>=2300,R8<=4000),"Transient","Turbulent"))
            }
            else if (reynoldsNumber >= 2300 && reynoldsNumber < 4000)
            {
                flowType = "Transient";
            }
            else
            {
                flowType = "Turbulent";
            }

            return flowType;
        }

        //input: RN result
        public double calculateFrictionFactorTurbulent(double diameter, double reynoldsNumber, double frictionFactorTurbulent, double roughness)
        {
            //=IFERROR(
            //IF(ISNUMBER(T8), (-0.5/LOG10(12*K8/(3.7*O8)+2.51/(R8*T8^0.5)))^2, 1),
            //"")
            
            double dd = Math.Pow(-0.5 / Math.Log10(12 * roughness / (3.7 * diameter) + (2.51 / (reynoldsNumber * Math.Sqrt(frictionFactorTurbulent)))), 2);
            if (Math.Abs(frictionFactorTurbulent - dd) > 0.000001)
            {
                dd = calculateFrictionFactorTurbulent(diameter,reynoldsNumber,dd, roughness);
            }
            return dd;
            return 0;
        }

        //input: RN result
        public double calculateFrictionFactorLaminar(double reynoldsNumber)
        {
            //=IFERROR(IF(ISNUMBER(T8), (-0.5/LOG10(12*K8/(3.7*O8)+2.51/(R8*T8^0.5)))^2, 1),"")
            double frictionFactor = 64 / reynoldsNumber;
            return 0;
        }

        //input: RN result
        public double calculateFrictionFactor(string flowType, double frictionFactorLaminar, double frictionFactorTurbulent)
        {
            //=IF(S8="Laminar",U8,T8)
            if (flowType == "Laminar")
            {
                return frictionFactorLaminar;
            }
            else
            {
                return frictionFactorTurbulent;
            }

        }

        //inputs: diameter and velocity calculated
        public double calculatePressureLoss(string ductElement, double frictionFactor, double diameter, double velocity)
        {
            //=IF(D8="FITTING", "-",
            //IF(D8="COMPONENT", "-", 12*V8*100*$L$3*(P8/1097)^2/O8))
            //Duct element = D8
            //FrictionFactor = V8
            //Velocity = P8
            //Density of dry air (constant) = L3
            //Diameter = O8
            if (ductElement == "Fitting")
            {
                return -1;
            }
            else
            {
                return 12 * frictionFactor * 100 * 0.077 * Math.Pow((velocity/1097), 2) / diameter;
            }

        }

        //input parameters from results
        public double calculateTotalPressureLoss(string ductElement, double fittingCoefficient, double velocityPressure, double ductLength, double compPressureDrop, double pressureLoss)
        {
            //Duct element = D8
            //Fitting coefficient = M8
            //Velocity pressure = Q8
            //Duct length = I8
            //Component pressure drop = N8
            //Pressure loss = W8 (previous method)
            //Diameter = O8
            //IF($D8="FITTING", $M8*$Q8*$I8,
            //IF($D8="COMPONENT", $N8*$I8,
            //IF($D8="FLEX DUCT", I8/100*W8*(1+58*(1-0.85)*2.7813^(-0.125*O8)),I8*W8/100)
            //)
            //)
            double pressure;
            if (ductElement == "Fitting")
            {
                pressure = fittingCoefficient * velocityPressure * ductLength;
            }
            else
            {
                if (ductElement == "Component")
                {
                    pressure = compPressureDrop * ductLength;
                }
                else if (ductElement == "Flex Duct")
                {
                    pressure = ductLength * pressureLoss / 100 * (1 + 58 * (1 - 0.85) * Math.Pow(2.7813, -0.125 * diameter));
                }
                else
                {
                    pressure = ductLength * pressureLoss / 100;
                }
            }

            return pressure;
        }
        //Equation for Width
        public double setWidth(int row, string aFN)
        {

            //string[,] inputArray2 = new string[7, 7];

            double width = -1;

            if (aFN == "SR4-1")
            {
                width = Convert.ToDouble(inputArray2[row, 15]);
            }
            else if (aFN == "SR5-13" && inputArray2[row, 30] == "No")
            {
                width = Convert.ToDouble(inputArray2[row, 25]);
            }
            else if (aFN == "SR5-13" && inputArray2[row, 30] != "No")
            {
                width = Convert.ToDouble(inputArray2[row, 27]);
            }
            else if (aFN == "SR5-14")
            {
                width = Convert.ToDouble(inputArray2[row, 37]);
            }
            else if (aFN == "CR3-1")
            {
                width = Convert.ToDouble(inputArray2[row, 6]);
            }
            else if (aFN == "CR3-16")
            {
                width = Convert.ToDouble(inputArray2[row, 6]);
            }
            else if (aFN == "SR4-3")
            {
                width = Convert.ToDouble(inputArray2[row, 15]);
            }
            else if (aFN == "SR4-2")
            {
                width = Convert.ToDouble(inputArray2[row, 15]);
            }
            else if (aFN == "SR5-12" && inputArray2[row, 30] == "No")
            {
                width = Convert.ToDouble(inputArray2[row, 25]);
            }
            else if (aFN == "SR5-11" && inputArray2[row, 30] == "No")
            {
                width = Convert.ToDouble(inputArray2[row, 25]);
            }

            return width;

        }


        //Equation for Height
        public double setHeight(int row, string aFN)
        {

            //string[,] inputArray2 = new string[7, 7];

            double height = -1;

            if (aFN == "SR4-1")
            {
                height = Convert.ToDouble(inputArray2[row, 14]);
            }
            else if (aFN == "SR5-13" && inputArray2[row, 30] == "No")
            {
                height = Convert.ToDouble(inputArray2[row, 24]);
            }
            else if (aFN == "SR5-13" && inputArray2[row, 30] != "No")
            {
                height = Convert.ToDouble(inputArray2[row, 26]);
            }
            else if (aFN == "SR5-14")
            {
                height = Convert.ToDouble(inputArray2[row, 36]);
            }
            else if (aFN == "CR3-1")
            {
                height = Convert.ToDouble(inputArray2[row, 7]);
            }
            else if (aFN == "CR3-16")
            {
                height = Convert.ToDouble(inputArray2[row, 7]);
            }
            else if (aFN == "SR4-3")
            {
                height = Convert.ToDouble(inputArray2[row, 14]);
            }
            else if (aFN == "SR4-2")
            {
                height = Convert.ToDouble(inputArray2[row, 14]);
            }
            else if (aFN == "SR5-12" && inputArray2[row, 30] == "No")
            {
                height = Convert.ToDouble(inputArray2[row, 24]);
            }
            else if (aFN == "SR5-11" && inputArray2[row, 30] == "No")
            {
                height = Convert.ToDouble(inputArray2[row, 24]);
            }

            return height;

        }


        //Equation for Diameter(not solve for diameter)
        public double setDiameter(int row, string aFN)
        {

            //string[,] inputArray2 = new string[7, 7];

            double diameter = -1;

            if (aFN == "SR4-1")
            {
                diameter = Convert.ToDouble(inputArray2[row, 10]);
            }
            else if (aFN == "SD4-2")
            {
                diameter = Convert.ToDouble(inputArray2[row, 13]);
            }
            else if (aFN == "SD5-10" && inputArray2[row, 30] == "No")
            {
                diameter = Convert.ToDouble(inputArray2[row, 20]);
            }
            else if (aFN == "SD5-10" && inputArray2[row, 30] != "No")
            {
                diameter = Convert.ToDouble(inputArray2[row, 21]);
            }
            else if (aFN == "CD3-1")
            {
                diameter = Convert.ToDouble(inputArray2[row, 5]);
            }
            else if (aFN == "SD5-9" && inputArray2[row, 30] == "No")
            {
                diameter = Convert.ToDouble(inputArray2[row, 20]);
            }
            else if (aFN == "SD5-9" && inputArray2[row, 30] != "No")
            {
                diameter = Convert.ToDouble(inputArray2[row, 21]);
            }
            else if (aFN == "SD5-18")
            {
                diameter = Convert.ToDouble(inputArray2[row, 32]);
            }
            else if (aFN == "SR5-12" && inputArray2[row, 30] == "Yes")
            {
                diameter = Convert.ToDouble(inputArray2[row, 21]);
            }
            else if (aFN == "SR5-11" && inputArray2[row, 30] == "Yes")
            {
                diameter = Convert.ToDouble(inputArray2[row, 21]);
            }
            else if (aFN == "CD3-4")
            {
                diameter = Convert.ToDouble(inputArray2[row, 5]);
            }
            else if (aFN == "CD3-2")
            {
                diameter = Convert.ToDouble(inputArray2[row, 5]);
            }

            return diameter;

        }

        //Equation for CFM
        public double setCFM(int row, string aFN)
        {

            double cfm = -1;

            //string[,] inputArray2 = new string[7, 7];

            if (aFN == "SD4-1" || aFN == "SD4-2" || aFN == "SR4-1" || aFN == "CD3-1" || aFN == "CR3-1" || aFN == "CR3-16" || aFN == "SR4-3" || aFN == "SR4-2")
            {
                cfm = Convert.ToDouble(inputArray2[row, 3]);
            }
            else if ((aFN == "SD5-10" || aFN == "SR5-13" || aFN == "SD5-9") && (inputArray2[row, 30] == "No"))
            {
                cfm = Convert.ToDouble(inputArray2[row, 28]) - Convert.ToDouble(inputArray2[row, 29]);
            }
            else if ((aFN == "SD5-10" || aFN == "SR5-13" || aFN == "SD5-9") && (inputArray2[row, 30] != "No"))
            {
                cfm = Convert.ToDouble(inputArray2[row, 29]);
            }
            else if (aFN == "SD5-18")
            {
                cfm = Convert.ToDouble(inputArray2[row, 34]);
            }
            else if (aFN == "SR5-14")
            {
                cfm = Convert.ToDouble(inputArray2[row, 3]) / 2;
            }
            else if (aFN == "SR5-12" && inputArray2[row, 30] == "Yes")
            {
                cfm = Convert.ToDouble(inputArray2[row, 29]);
            }
            else if (aFN == "SR5-12" && inputArray2[row, 30] == "No")
            {
                cfm = Convert.ToDouble(inputArray2[row, 28]) - Convert.ToDouble(inputArray2[row, 29]);
            }
            else if (aFN == "SR5-11" && inputArray2[row, 30] == "Yes")
            {
                cfm = Convert.ToDouble(inputArray2[row, 29]);
            }
            else if (aFN == "SR5-11" && inputArray2[row, 30] == "No")
            {
                cfm = Convert.ToDouble(inputArray2[row, 28]) - Convert.ToDouble(inputArray2[row, 29]);
            }
            else if (aFN == "CD3-4")
            {
                cfm = Convert.ToDouble(inputArray2[row, 3]);
            }
            else if (aFN == "CD3-2")
            {
                cfm = Convert.ToDouble(inputArray2[row, 3]);
            }

            return cfm;

        }

        
    }
}
