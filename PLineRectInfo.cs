using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
using static ArchitecturalWindows.Globals;

namespace ArchitecturalWindows
{
    class PLineRectInfo

    {
        public Point3d Origin;
        public Vector3d XAxis;
        public Vector3d YAxis;
        public double XLength;
        public double YLength;
        Point3d[] Points;

        public PLineRectInfo(Polyline rect)
        {
            Points = Get3DPolylinePoints(rect);

            // Extract relevant points and vectors
            Origin = Points[0];
            Point3d xVectPt = Points[1];
            Point3d yVectPt = Points[3];

            Vector3d xVect = xVectPt - Origin;
            Vector3d yVect = yVectPt - Origin;

            XLength = xVect.Length;
            YLength = yVect.Length;

            XAxis = xVect.GetNormal();
            YAxis = yVect.GetNormal();
        }
    }
}
