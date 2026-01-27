using System;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Geometry
{
    public sealed class BoundingBoxExtractor
    {
        private const double M_TO_IN = 39.37007874015748;

        public sealed class BoundingBoxInfo
        {
            public double LengthInches { get; set; }
            public double WidthInches { get; set; }
            public double HeightInches { get; set; }
            public double[] MinPoint { get; set; }
            public double[] MaxPoint { get; set; }
        }

        public BoundingBoxInfo GetFaceBoundingBox(IFace2 face)
        {
            if (face == null) return null;

            try
            {
                var uv = face.GetUVBounds() as double[];
                if (uv == null || uv.Length < 4) return null;
                var surf = face.GetSurface() as ISurface;
                if (surf == null) return null;

                double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

                const int samples = 10;
                double uStep = (uv[1] - uv[0]) / samples;
                double vStep = (uv[3] - uv[2]) / samples;

                for (int i = 0; i <= samples; i++)
                {
                    for (int j = 0; j <= samples; j++)
                    {
                        double u = uv[0] + (i * uStep);
                        double v = uv[2] + (j * vStep);
                        var pt = surf.Evaluate(u, v, 0, 0) as double[];
                        if (pt != null && pt.Length >= 3)
                        {
                            if (pt[0] < minX) minX = pt[0]; if (pt[0] > maxX) maxX = pt[0];
                            if (pt[1] < minY) minY = pt[1]; if (pt[1] > maxY) maxY = pt[1];
                            if (pt[2] < minZ) minZ = pt[2]; if (pt[2] > maxZ) maxZ = pt[2];
                        }
                    }
                }

                return new BoundingBoxInfo
                {
                    LengthInches = (maxX - minX) * M_TO_IN,
                    WidthInches = (maxY - minY) * M_TO_IN,
                    HeightInches = (maxZ - minZ) * M_TO_IN,
                    MinPoint = new[] { minX * M_TO_IN, minY * M_TO_IN, minZ * M_TO_IN },
                    MaxPoint = new[] { maxX * M_TO_IN, maxY * M_TO_IN, maxZ * M_TO_IN }
                };
            }
            catch { return null; }
        }

        public (double length, double width) GetBlankSize(IModelDoc2 model)
        {
            var fa = new FaceAnalyzer();
            var face = fa.GetProcessingFace(model);
            if (face == null) return (0, 0);

            var box = GetFaceBoundingBox(face);
            if (box == null) return (0, 0);

            double length = box.LengthInches, width = box.WidthInches;
            if (width > length)
            {
                var tmp = length; length = width; width = tmp;
            }
            return (length, width);
        }
    }
}
