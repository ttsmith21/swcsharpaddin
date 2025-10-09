using System;

namespace NM.StepClassifierAddin.Utils
{
    internal static class Math3D
    {
        public static double[] Normalize(double[] v)
        {
            double n = Math.Sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]);
            if (n <= 0) return new[] { 0.0, 0.0, 0.0 };
            return new[] { v[0] / n, v[1] / n, v[2] / n };
        }
        public static double Dot(double[] a, double[] b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        public static double[] Sub(double[] a, double[] b) => new[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] };
        public static double[] Mul(double[] a, double s) => new[] { a[0] * s, a[1] * s, a[2] * s };
        public static double[] Add(double[] a, double[] b) => new[] { a[0] + b[0], a[1] + b[1], a[2] + b[2] };
        public static double Abs(double x) => Math.Abs(x);
        public static double Clamp(double x, double lo, double hi) => x < lo ? lo : (x > hi ? hi : x);
    }
}