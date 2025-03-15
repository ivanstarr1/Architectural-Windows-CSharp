using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using static Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Autodesk.AutoCAD.Geometry;
namespace ArchitecturalWindows
{
    public static class Globals
    {
        public static Point3d[] Get3DPolylinePoints(Polyline polyline)
        {
            //Point3d[] pointsArray = new Point3d[];
            //Database db = HostApplicationServices.WorkingDatabase;
            //Transaction tr = db.TransactionManager.StartTransaction()
            //Polyline polyline = trans.GetObject(result.ObjectId, OpenMode.ForRead) as Polyline;
            // Create a list to store points
            List<Point3d> points = new List<Point3d>();

            // Get all the vertices of the polyline
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                Point3d point = polyline.GetPoint3dAt(i);
                points.Add(point);
            }

            // Convert list to array if needed
            Point3d[] pointsArray = points.ToArray();

            return pointsArray;
        }

        public static void SetXRecord(Transaction tr, DBDictionary dict, string key, ResultBuffer ResBuf)
        {
            Xrecord Xrec = new Xrecord();
            Xrec.Data = ResBuf;
            dict.SetAt(key, Xrec);
            tr.AddNewlyCreatedDBObject(Xrec, true);
            //tr.Commit();
        }
        public static void SetXRecordReal(Transaction tr, DBDictionary dict, string key, double Value)
        {
            SetXRecord(tr, dict, key, new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataReal, Value)));
        }

 
        public static void SetXRecordText(Transaction tr, DBDictionary dict, string key, String Value)
        {
            SetXRecord(tr, dict, key, new ResultBuffer(new TypedValue((int)DxfCode.Text, Value)));
        }
        public static void SetXRecordInt(Transaction tr, DBDictionary dict, string key, int Value)
        {
            SetXRecord(tr, dict, key, new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataInteger32, Value)));
        }



        public static void SetXRecordHandle(Transaction tr, DBDictionary dict, string key, Handle Value)
        {
            SetXRecord(tr, dict, key, new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataHandle, Value)));
        }

        public static void SetXRecordPoint3d(Transaction tr, DBDictionary dict, string key, Point3d Value)
        {
            SetXRecord(tr, dict, key,new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataReal, Value.X),
                new TypedValue((int)DxfCode.ExtendedDataReal, Value.Y),
                new TypedValue((int)DxfCode.ExtendedDataReal, Value.Z)
                ));
        }

        public static void SetXRecordVector3d(Transaction tr, DBDictionary dict, string key, Vector3d Value)
        {
            SetXRecord(tr, dict, key, (new ResultBuffer(
                 new TypedValue((int)DxfCode.ExtendedDataReal, Value.X),
                 new TypedValue((int)DxfCode.ExtendedDataReal, Value.Y),
                 new TypedValue((int)DxfCode.ExtendedDataReal, Value.Z)
                 )));
        }


        public static ResultBuffer GetXRecord(Transaction tr, DBDictionary dict, string key)
        {
            //if (!dict.Contains(key)) return 0.0; // Return default if key doesn't exist

            ObjectId xrecId = dict.GetAt(key);
            Xrecord xrec = tr.GetObject(xrecId, OpenMode.ForRead) as Xrecord;

            return xrec.Data;
        }

        public static double GetXRecordReal(Transaction tr, DBDictionary dict, string key)
        {
            ResultBuffer ResBuf = GetXRecord(tr, dict, key);
            TypedValue[] TypedValues = ResBuf.AsArray();
            return (double)TypedValues[0].Value;
        }
        public static string GetXRecordText(Transaction tr, DBDictionary dict, string key)
        {
            ResultBuffer ResBuf = GetXRecord(tr, dict, key);
            return (string)ResBuf.AsArray()[0].Value;
        }
        public static int GetXRecordInt(Transaction tr, DBDictionary dict, string key)
        {
            ResultBuffer ResBuf = GetXRecord(tr, dict, key);
            return (int)ResBuf.AsArray()[0].Value;
            //if (!dict.Contains(key)) return 0;

            //ObjectId xrecId = dict.GetAt(key);
            //Xrecord xrec = tr.GetObject(xrecId, OpenMode.ForRead) as Xrecord;

            //if (xrec != null)
            //{
            //    using (ResultBuffer rb = xrec.Data)
            //    {
            //        TypedValue[] values = rb.AsArray();
            //        if (values.Length > 0 && values[0].TypeCode == (int)DxfCode.ExtendedDataInteger16)
            //            return (int)values[0].Value;
            //    }
            //}
            //return 0; // Default value
        }

        public static Handle GetXRecordHandle(Transaction tr, DBDictionary dict, string key)
        {
            ResultBuffer ResBuf = GetXRecord(tr, dict, key);
             string HexString = (string)ResBuf.AsArray()[0].Value;
            long ln = Convert.ToInt64(HexString, 16);

            // Not create a Handle from the long integer

            Handle hn = new Handle(ln);
            return hn;
        }
        public static Point3d GetXRecordPoint3d(Transaction tr, DBDictionary dict, string key)
        {
            ResultBuffer ResBuf = GetXRecord(tr, dict, key);
            TypedValue[] values = ResBuf.AsArray();
            double x = (double)values[0].Value;
            double y = (double)values[1].Value;
            double z = (double)values[2].Value;
            return new Point3d(x, y, z);
        }



        public static Vector3d GetXRecordVector3d(Transaction tr, DBDictionary dict, string key)
        {
            ResultBuffer ResBuf = GetXRecord(tr, dict, key);
            var values = ResBuf.AsArray();
            double x = (double)values[0].Value;
            double y = (double)values[1].Value;
            double z = (double)values[2].Value;
            return new Vector3d(x, y, z);
        }


        public static void SetUCS(Point3d origin, Vector3d xAxis, Vector3d yAxis)
        {
            var zAxis = xAxis.CrossProduct(yAxis);
            Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem =
                Matrix3d.AlignCoordinateSystem(
                    Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                    origin, xAxis.GetNormal(), yAxis.GetNormal(), zAxis.GetNormal());
        }



        //public static Handle GetXRecordHandle(Transaction tr, DBDictionary dict, string key)
        //{
        //    ResultBuffer ResBuf = GetXRecord(tr, dict, key);
        //    return (Handle)ResBuf.AsArray()[0].Value;
        //}
    }
}
