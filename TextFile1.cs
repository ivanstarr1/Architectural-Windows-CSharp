using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
//using static Autodesk.AutoCAD.ApplicationServices.Core.Application;
//using Autodesk.AutoCAD.ApplicationServices.Core;

namespace WindowManagement



{ public static class SnapMode 
    {
        private static int TempSnapModeSave;
		//public static AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

		public static void ToggleSnapsOff()
        {
	            //var AcadApp = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
	            //Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            TempSnapModeSave = (int)Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("OSMODE");
			Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", 0);
        }

        public static void ToggleSnapsOn()
        {
			//Document doc = Application.DocumentManager.MdiActiveDocument;
			Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", TempSnapModeSave);
        }
}

public static class ErrorHandling
{
    public static void LoadStandardErrorRoutine()
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        ed.WriteMessage("\nError routine loaded.");
    }

    public static void UnloadStandardErrorRoutine()
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        ed.WriteMessage("\nError routine unloaded.");
    }

    public static void HandleError(string message)
    {
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
        ed.WriteMessage($"\nError: {message}");
        SnapMode.ToggleSnapsOn();
    }
}

public static class VectorOperations
{
    public static Vector3d VAdd(Vector3d v1, Vector3d v2) => v1 + v2;
    public static Vector3d VSubtr(Vector3d v1, Vector3d v2) => v1 - v2;
    public static Vector3d VCross(Vector3d v1, Vector3d v2) => v1.CrossProduct(v2);
}

public static class GeometryUtilities
{
    public static Matrix3d GetECSTransformMatrix(Entity ent)
    {
            //if (ent == null) return Matrix3d.Identity;
            //Vector3d normal = ent.Ecs;
            //CoordinateSystem3d ecs = new CoordinateSystem3d(Point3d.Origin, normal);
            //return Matrix3d.AlignCoordinateSystem(
            //    ecs.Origin, ecs.Xaxis, ecs.Yaxis, ecs.Zaxis,
            //    Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis
            return ent.Ecs;
       
    }

    public static List<Point2d> Get2DECSPointsFromWLPline(Polyline polyline)
    {
        List<Point2d> points = new List<Point2d>();
        if (polyline == null || polyline.NumberOfVertices < 2)
            return points;

        Matrix3d ecsMatrix = GetECSTransformMatrix(polyline);
        Matrix3d inverseECSMatrix = ecsMatrix.Inverse();

        for (int i = 0; i < polyline.NumberOfVertices; i++)
        {
            Point3d worldPoint = polyline.GetPoint3dAt(i);
            Point3d ecsPoint = worldPoint.TransformBy(inverseECSMatrix);
            points.Add(new Point2d(ecsPoint.X, ecsPoint.Y));
        }
        return points;
    }
}

public static class SolidOperations
{
    public static Solid3d SpatialConflictSolid(Solid3d solidA, Solid3d solidB)
    {
        if (solidA == null || solidB == null)
            return null;
        try
        {
            Solid3d intersection = solidA.Clone() as Solid3d;
            intersection.BooleanOperation(BooleanOperationType.BoolIntersect, solidB.Clone() as Solid3d);
            return intersection.IsNull ? null : intersection;
        }
        catch
        {
            return null;
        }
    }
}

public class WindowCommands
{
    [CommandMethod("ManageWindows")]
    public void ManageWindows()
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        bool running = true;
        while (running)
        {
            PromptKeywordOptions pko = new PromptKeywordOptions("\nChoose an option [Add/Modify/Delete/Exit]: ");
            pko.Keywords.Add("Add");
            pko.Keywords.Add("Modify");
            pko.Keywords.Add("Delete");
            pko.Keywords.Add("Exit");
            pko.AllowNone = false;
            PromptResult result = ed.GetKeywords(pko);
            if (result.Status == PromptStatus.OK)
            {
                switch (result.StringResult)
                {
                    case "Add":
                        AddWindow(ed);
                        break;
                    case "Modify":
                        ModifyWindow(ed);
                        break;
                    case "Delete":
                        DeleteWindow(ed);
                        break;
                    case "Exit":
                        ed.WriteMessage("\nExiting window management.");
                        running = false;
                        break;
                }
            }
            else
            {
                ed.WriteMessage("\nCommand canceled.");
                running = false;
            }
        }
    }

    private void AddWindow(Editor ed)
    {
        PromptPointOptions ppo = new PromptPointOptions("\nSpecify window insertion point: ");
        PromptPointResult ppr = ed.GetPoint(ppo);
        if (ppr.Status == PromptStatus.OK)
            ed.WriteMessage($"\nWindow added at {ppr.Value}");
        else
            ed.WriteMessage("\nInsertion canceled.");
    }

    private void ModifyWindow(Editor ed)
    {
        ed.WriteMessage("\nModify window logic goes here.");
    }

    private void DeleteWindow(Editor ed)
    {
        ed.WriteMessage("\nDelete window logic goes here.");
    }
}

}

