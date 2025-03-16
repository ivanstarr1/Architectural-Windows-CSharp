using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArchitecturalWindows
{
    //public class RectangleJig : EntityJig
    //{
    //    private Point3d firstPoint;
    //    private Point3d secondPoint;

    //    public RectangleJig(Point3d basePoint)
    //        : base(new Polyline())
    //    {
    //        firstPoint = basePoint;
    //        secondPoint = basePoint;
    //    }

    //    protected override bool Update()
    //    {
    //        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
    //        Editor ed = doc.Editor;

    //        Polyline pl = Entity as Polyline;
    //        if (pl == null) return false;


    //        if (pl.NumberOfVertices > 0)
    //        {
    //            pl.SetPointAt(0, new Point2d(firstPoint.X, firstPoint.Y));
    //            pl.SetPointAt(1, new Point2d(secondPoint.X, firstPoint.Y));
    //            pl.SetPointAt(2, new Point2d(secondPoint.X, secondPoint.Y));
    //            pl.SetPointAt(3, new Point2d(firstPoint.X, secondPoint.Y));
    //            pl.SetPointAt(4, new Point2d(firstPoint.X, firstPoint.Y));
    //        }
    //        else
    //        {
    //            pl.AddVertexAt(0, new Point2d(firstPoint.X, firstPoint.Y), 0, 0, 0);
    //            pl.AddVertexAt(1, new Point2d(secondPoint.X, firstPoint.Y), 0, 0, 0);
    //            pl.AddVertexAt(2, new Point2d(secondPoint.X, secondPoint.Y), 0, 0, 0);
    //            pl.AddVertexAt(3, new Point2d(firstPoint.X, secondPoint.Y), 0, 0, 0);
    //            pl.AddVertexAt(4, new Point2d(firstPoint.X, firstPoint.Y), 0, 0, 0);
    //            ed.WriteMessage("Y");
    //        }


    //        return true;
    //    }

    //    protected override SamplerStatus Sampler(JigPrompts prompts)
    //    {
    //        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
    //        Editor ed = doc.Editor;

    //        JigPromptPointOptions pointOpts = new JigPromptPointOptions("\nSelect opposite corner:");
    //        pointOpts.UserInputControls = UserInputControls.Accept3dCoordinates;

    //        PromptPointResult pointResult = prompts.AcquirePoint(pointOpts);
    //        if (pointResult.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

    //        Point3d newPoint = pointResult.Value;
    //        if (newPoint == secondPoint) return SamplerStatus.NoChange;

    //        ed.WriteMessage("X");
    //        secondPoint = newPoint;
    //        return SamplerStatus.OK;
    //    }

    //    public static void CreateRectangle()
    //    {
    //        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
    //        Editor ed = doc.Editor;

    //        PromptPointOptions ppo = new PromptPointOptions("\nSelect first corner:");
    //        PromptPointResult ppr = ed.GetPoint(ppo);
    //        if (ppr.Status != PromptStatus.OK) return;

    //        RectangleJig jig = new RectangleJig(ppr.Value);
    //        if (ed.Drag(jig).Status == PromptStatus.OK)
    //        {
    //            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
    //            {
    //                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
    //                Polyline rect = jig.Entity as Polyline;
    //                rect.Closed = true;
    //                btr.AppendEntity(rect);
    //                tr.AddNewlyCreatedDBObject(rect, true);
    //                tr.Commit();

    //            }
    //        }
    //    }


    //}

    public class RectangleJig : EntityJig
    {
        private Point3d firstPoint;
        private Point3d secondPoint;
        private Polyline poly;

        public RectangleJig(Point3d basePoint)
            : base(new Polyline())
        {
            firstPoint = basePoint;
            secondPoint = basePoint;
            poly = Entity as Polyline;
        }

        protected override bool Update()
        {
            if (poly == null) return false;

            // Get the current UCS matrix at runtime
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Matrix3d ucsMatrix = ed.CurrentUserCoordinateSystem;

            // Transform points from WCS to UCS
            Point3d pt1 = firstPoint.TransformBy(ucsMatrix.Inverse());
            Point3d pt2 = secondPoint.TransformBy(ucsMatrix.Inverse());

            // Define rectangle vertices in UCS
            Point2d p1 = new Point2d(pt1.X, pt1.Y);
            Point2d p2 = new Point2d(pt2.X, pt1.Y);
            Point2d p3 = new Point2d(pt2.X, pt2.Y);
            Point2d p4 = new Point2d(pt1.X, pt2.Y);

            // Define polyline (2D)
            if (poly.NumberOfVertices == 0)
            {
                poly.AddVertexAt(0, p1, 0, 0, 0);
                poly.AddVertexAt(1, p2, 0, 0, 0);
                poly.AddVertexAt(2, p3, 0, 0, 0);
                poly.AddVertexAt(3, p4, 0, 0, 0);
                poly.AddVertexAt(4, p1, 0, 0, 0); // Close the rectangle
            }
            else
            {
                poly.SetPointAt(0, p1);
                poly.SetPointAt(1, p2);
                poly.SetPointAt(2, p3);
                poly.SetPointAt(3, p4);
                poly.SetPointAt(4, p1);
            }

            return true;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions pointOpts = new JigPromptPointOptions("\nSelect opposite corner:");
            pointOpts.UserInputControls = UserInputControls.Accept3dCoordinates;

            PromptPointResult pointResult = prompts.AcquirePoint(pointOpts);
            if (pointResult.Status != PromptStatus.OK) return SamplerStatus.Cancel;

            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Matrix3d ucsMatrix = ed.CurrentUserCoordinateSystem; // Get UCS dynamically

            Point3d newPoint = pointResult.Value.TransformBy(ucsMatrix.Inverse()); // Convert to UCS
            if (newPoint == secondPoint) return SamplerStatus.NoChange;

            secondPoint = newPoint;
            return SamplerStatus.OK;
        }

        public static void CreateRectangle()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptPointOptions ppo = new PromptPointOptions("\nSelect first corner:");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            RectangleJig jig = new RectangleJig(ppr.Value);

            if (ed.Drag(jig).Status == PromptStatus.OK)
            {
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                    Polyline rect = jig.Entity as Polyline;
                    rect.Closed = true;
                    btr.AppendEntity(rect);
                    tr.AddNewlyCreatedDBObject(rect, true);
                    tr.Commit();
                }
            }
        }
    }

    public class Commands
    {
        [CommandMethod("MyRectangle")]
        public void CreateCustomRectangle()
        {
            RectangleJig.CreateRectangle();
        }


    }

}
