using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using static ArchitecturalWindows.Globals;

namespace ArchitecturalWindows
{
    public class WindowObject
    {
        public string Id;
        public Point3d Origin;
        public double XLength;
        public double YLength;
        public Vector3d XAxis;
        public Vector3d YAxis;
        public int Handedness;
        public Handle FrameEntHandle;
        public Handle GlassEntHandle;
        public double WallThickness;
        public Handle WallEntHandle;

        public ObjectId WindowEntityObjectId(Handle WindowEntHandle)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            return db.GetObjectId(false, FrameEntHandle, 0);
        }
        public ObjectId FrameEntObjectId()
        {
            return WindowEntityObjectId(FrameEntHandle);
        }
        public ObjectId GlassEntObjectId()
        {
            return WindowEntityObjectId(GlassEntHandle);
        }
        public Solid3d WindowEntity(Transaction tr, Handle WindowEntHandle, OpenMode Mode)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            ObjectId FrameEntObjectId = db.GetObjectId(false, WindowEntHandle, 0);
            return tr.GetObject(FrameEntObjectId, Mode) as Solid3d;
        }
        public Solid3d FrameEnt(Transaction tr, OpenMode Mode)
        {
            return WindowEntity(tr, FrameEntHandle, Mode);
        }
        public Solid3d GlassEnt(Transaction tr, OpenMode Mode)
        {
            return WindowEntity(tr, GlassEntHandle, Mode);
        }

        public WindowObject(Transaction tr, DBDictionary dict)
        {
            Id = GetXRecordText(tr, dict, "Id");
            Origin = GetXRecordPoint3d(tr, dict, "Origin");
            XLength = GetXRecordReal(tr, dict, "XLength");
            YLength = GetXRecordReal(tr, dict, "YLength");
            WallThickness = GetXRecordReal(tr, dict, "WallThickness");
            XAxis = GetXRecordVector3d(tr, dict, "XAxis");
            YAxis = GetXRecordVector3d(tr, dict, "YAxis");
            Handedness = GetXRecordInt(tr, dict, "Handedness");
            FrameEntHandle = GetXRecordHandle(tr, dict, "FrameEntHandle");
            GlassEntHandle = GetXRecordHandle(tr, dict, "GlassEntHandle");
            WallEntHandle = GetXRecordHandle(tr, dict, "WallEntHandle");
        }

        public Point3d[] GetWindowPoints()
        {
            Point3d Point1 = Origin.Add(XAxis.MultiplyBy(XLength));
            Point3d Point2 = Point1.Add(YAxis.MultiplyBy(YLength));
            Point3d Point3 = Origin.Add(YAxis.MultiplyBy(YLength));
            return [Origin, Point1, Point2, Point3];
        }

        public void ToDictionary(DBDictionary dict)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            //ObjectId windowDictId;
            //DBDictionary windowDict;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                SetXRecordText(tr, dict, "Id", Id);
                SetXRecordPoint3d(tr, dict, "Origin", Origin);
                SetXRecordReal(tr, dict, "XLength", XLength);
                SetXRecordReal(tr, dict, "YLength", YLength);
                SetXRecordVector3d(tr, dict, "XAxis", XAxis);
                SetXRecordVector3d(tr, dict, "YAxis", YAxis);
                SetXRecordInt(tr, dict, "Handedness", Handedness);
                SetXRecordHandle(tr, dict, "FrameEntHandle", FrameEntHandle);
                SetXRecordHandle(tr, dict, "GlassEntHandle", GlassEntHandle);
                SetXRecordHandle(tr, dict, "WallEntHandle", WallEntHandle);
            }
        }

 
    }
}
