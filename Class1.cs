using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using static ArchitecturalWindows.Globals;
// using static ArchitecturalWindows.UCSCommands;

using System.Windows.Input;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Windows.Controls;
using System.Windows;
using Autodesk.AutoCAD.GraphicsSystem;
using System.Net.Http.Headers;


[assembly: CommandClass(typeof(ArchitecturalWindows.Class1))]

namespace ArchitecturalWindows
{
    public class Class1
    {
        private static Object? tempSnapModeSave;
        private static string WINDOWFRAMELAYER = "A-Glazing-Frames";
        private static string WINDOWGLASSLAYER = "A-Glazing-Glass";
        private static string WALLLAYER = "Walls";
        private static double WINDOWFRAMETHICKNESS = 7;
        private static double GLASSTHICKNESS = 0.5;
        private static double MAXWALLTHICKNESS = 40;
        private static ObjectId WINDOWDICTIONARYDICTIONARYID = GetWindowDictionaryDictionaryObjectId();
        //private static ObjectId previousUCS = ObjectId.Null;
        private static ObjectId PlineTestEntityObjectId = ObjectId.Null;


        [CommandMethod("MyRectangle")]
        public void CreateCustomRectangle()
        {
            RectangleJig.CreateRectangle();
        }


        [CommandMethod("MyRectangle2")]
        public void RunRectangleCommand()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Call the built-in RECTANGLE command
            ed.Document.SendStringToExecute("_RECTANGLE ", true, false, false);
            ed.WriteMessage("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXxxxxxxxxxxxxxxx");
        }

        [CommandMethod("TestXXX")]
        public void TestXXX()
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            DBDictionary WindowDict = new DBDictionary();
            DBDictionary WindowDictionaryDictionary = tr.GetObject(WINDOWDICTIONARYDICTIONARYID, OpenMode.ForWrite) as DBDictionary;
            string newUniqueID = Guid.NewGuid().ToString("N").Substring(0, 8);
            WindowDictionaryDictionary.SetAt(newUniqueID, WindowDict);

            tr.AddNewlyCreatedDBObject(WindowDict, true);
            SetXRecordText(tr, WindowDict, "Id", newUniqueID);
            tr.Commit();

            tr = db.TransactionManager.StartTransaction();

            tr.Commit();

        }

        [CommandMethod("TestYYY")]

        public void TestYYY()
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            //WindowObject WO = new WindowObject(tr, windowDict);
            Solid3d WindowHole = CreateBoxAtCoordinateSystem(1, 1, 1, new Point3d(1, 1, 1), new Vector3d(1, 0, 0), new Vector3d(0, 1, 0));

            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            btr.AppendEntity(WindowHole);
            tr.AddNewlyCreatedDBObject(WindowHole, true);
            tr.Commit();
            //ObjectId WallEntObjId = db.GetObjectId(false, WO.WallEntHandle, 0);
            //Solid3d Wall = tr.GetObject(WallEntObjId, OpenMode.ForWrite) as Solid3d;
            //tr = db.TransactionManager.StartTransaction();
            //Wall.BooleanOperation(BooleanOperationType.BoolSubtract, WindowHole);
            //tr.Commit();

        }


        [CommandMethod("SpatialConflict")]
        public static void SpatialConflict()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo1 = new PromptEntityOptions("\nSelect 3DSolid 1:");
            PromptEntityOptions peo2 = new PromptEntityOptions("\nSelect 3DSolid 2:");
            peo1.SetRejectMessage("\nInvalid selection. Select a 3DSolid entity");
            peo1.AddAllowedClass(typeof(Solid3d), true);
            peo2.SetRejectMessage("\nInvalid selection. Select a 3DSolid entity");
            peo2.AddAllowedClass(typeof(Solid3d), true);
            PromptEntityResult per1 = ed.GetEntity(peo1);
            if (per1.Status != PromptStatus.OK) return;
            PromptEntityResult per2 = ed.GetEntity(peo2);
            if (per2.Status != PromptStatus.OK) return;

            ObjectId? SPConf = SpatialConflictSolid(per1.ObjectId, per2.ObjectId);

        }


        [CommandMethod("TestZ")]
        public void TestZ()
        {
            if (PlineTestEntityObjectId == ObjectId.Null)
                return;

            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Polyline TestPline;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                TestPline = tr.GetObject(PlineTestEntityObjectId, OpenMode.ForRead) as Polyline;
            }
            CreateNewWindowFromRectInWall(TestPline);
        }


        [CommandMethod("Window")]
        public void Window()
        {
            EnsureLayerExists(WINDOWFRAMELAYER);
            EnsureLayerExists(WINDOWGLASSLAYER);
            ToggleSnapsOff();
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
                            AddWindow();
                            break;
                        case "Modify":
                            ModifyWindow();
                            break;
                        case "Delete":
                            DeleteWindow();
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

        public static void ToggleSnapsOff()
        {
            // Save current OSMODE value
            tempSnapModeSave = Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("OSMODE");
            // Set OSMODE to 0 (turn off object snaps)
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", 0);
        }


        public static void ToggleSnapsOn()
        {
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", tempSnapModeSave);
        }

        public static void EnsureLayerExists(string windowFrameLayer)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Check if the layer exists
                if (!layerTable.Has(windowFrameLayer))
                {
                    // Open the LayerTable for write
                    layerTable.UpgradeOpen();

                    // Create a new layer
                    LayerTableRecord newLayer = new LayerTableRecord
                    {
                        Name = windowFrameLayer,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3) // Green (ACI color 3)
                    };

                    // Add the new layer to the LayerTable
                    layerTable.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                tr.Commit();
            }
        }



        public void AddWindow()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ToggleSnapsOn();

            PromptEntityOptions peo = new PromptEntityOptions("\nPick a polyline rectangle: ");
            peo.SetRejectMessage("\nEntity is not a polyline.");
            peo.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Polyline plRect = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
                    PlineTestEntityObjectId = plRect.ObjectId;

                    if (plRect != null)
                    {
                        CreateNewWindowFromRectInWall(plRect);
                        ed.UpdateScreen();
                        PromptKeywordOptions pko = new PromptKeywordOptions("\nDelete rectangle? [Yes/No] <Yes>: ");
                        pko.Keywords.Add("Yes");
                        pko.Keywords.Add("No");
                        pko.AllowNone = true;

                        PromptResult pr = ed.GetKeywords(pko);
                        if (pr.Status == PromptStatus.OK || pr.Status == PromptStatus.None)
                        {
                            if (pr.StringResult == "Yes" || pr.Status == PromptStatus.None)
                            {
                                plRect.UpgradeOpen();
                                plRect.Erase();
                            }
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nEntity is not a polyline.");
                    }
                    tr.Commit();
                }
            }
            else
            {
                ed.WriteMessage("\nNo entity selected.");
            }
        }
        //else if (ppr.Status == PromptStatus.OK)
        //{
        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
        //        tr.Commit();
        //    }

        //    ed.Document.SendStringToExecute("_RECTANGLE PAUSE PAUSE ", true, false, false);

        //    ObjectId lastEntityId = GetLastCreatedModelSpaceEntity();
        //    Polyline lastEntity;
        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        lastEntity = tr.GetObject(lastEntityId, OpenMode.ForRead) as Polyline;
        //        tr.Commit();
        //    }

        //    CreateNewWindowFromRectInWall(lastEntity);

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        lastEntity.Erase();

        //        tr.Commit();
        //    }



    



    

        private void ModifyWindow(Editor ed)
        {
            ed.WriteMessage("\nModify window logic goes here.");
        }


        [CommandMethod("ShowForm")]
        public void ShowForm()
        {
            TestForm frmTest = new TestForm();
            frmTest.Show();
        }


        public DBDictionary ChildDictionary(DBDictionary ParentDict, string ChildDictName)
        {
            // Get the current document and database
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            DBDictionary childValue;

            // Assuming we are searching within the named objects dictionary for the key
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // Open the named objects dictionary for read
                //DBDictionary namedObjectsDict = trans.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;

                childValue = (DBDictionary)trans.GetObject(ParentDict.GetAt(ChildDictName), OpenMode.ForRead);


                // Commit the transaction
                trans.Commit();
            }
            return childValue;
        }


        private static ObjectId GetNextEntity(ObjectId currentId, Transaction tr, Database db)
        {
            // Get the current entity and its owner (e.g., ModelSpace or PaperSpace)
            Entity currentEntity = tr.GetObject(currentId, OpenMode.ForRead) as Entity;
            if (currentEntity == null) return ObjectId.Null;

            BlockTableRecord owner = tr.GetObject(currentEntity.OwnerId, OpenMode.ForRead) as BlockTableRecord;
            if (owner == null) return ObjectId.Null;

            // Find the next entity in the block table record (same as entnext in LISP)
            bool found = false;
            foreach (ObjectId id in owner)
            {
                if (found) return id; // Return the next entity
                if (id == currentId) found = true;
            }

            return ObjectId.Null; // No next entity found
        }


        private static ObjectId SkipAttributes(BlockReference blkRef, Transaction tr, Database db)
        {
            foreach (ObjectId attId in blkRef.AttributeCollection)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef != null && attRef.Id != ObjectId.Null)
                    return attRef.Id;
            }

            return ObjectId.Null;
        }


        // Helper function to create a marker point
        private static ObjectId CreateMarker(Point3d position, Transaction tr, Database db)
        {
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            DBPoint marker = new DBPoint(position);
            ObjectId markerId = btr.AppendEntity(marker);
            tr.AddNewlyCreatedDBObject(marker, true);
            return markerId;
        }


        public static string UniqueID()
        {
            return Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        public static double Get3DSolidVolume(Solid3d solid3d)
        {
             return solid3d.MassProperties.Volume;
        }

        public static List<Point3d> GetPlinePoints(ObjectId plineId)
        {
            List<Point3d> vertices = new List<Point3d>();

            if (plineId == ObjectId.Null)
                throw new System.ArgumentNullException("Invalid Polyline ObjectId.");

            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Polyline pline = tr.GetObject(plineId, OpenMode.ForRead) as Polyline;
                if (pline != null)
                {
                    // Get the polyline's Entity Coordinate System (ECS) to transform points to World Coordinate System (WCS)
                    Matrix3d ecsToWcs = Matrix3d.Identity;
                    if (!pline.Normal.IsEqualTo(Vector3d.ZAxis))
                        ecsToWcs = Matrix3d.AlignCoordinateSystem(
                            Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                            Point3d.Origin, pline.Normal, pline.Normal.GetPerpendicularVector(), pline.Normal.CrossProduct(pline.Normal.GetPerpendicularVector()));

                    int numVerts = pline.NumberOfVertices;
                    for (int i = 0; i < numVerts; i++)
                    {
                        Point2d pt2d = pline.GetPoint2dAt(i);
                        Point3d pt3d = new Point3d(pt2d.X, pt2d.Y, 0).TransformBy(ecsToWcs); // Convert to WCS
                        vertices.Add(pt3d);
                    }
                }

                tr.Commit();
            }

            return vertices;
        }

        public static void U0()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.Command("_UCS", "");
        }

        public static void UM(Point3d pt)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.Command("_UCS", "M", pt);
        }

        public static void U3(Point3d p1, Point3d p2, Point3d p3)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.Command("_UCS");
            ed.Command("3");
            ed.Command(p1);
            ed.Command(p2);
            ed.Command(p3);

        }

        public static void UP()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.Command("_UCS", "P");
        }

        public static ObjectId CreateDictionary()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId dictId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary dict = new DBDictionary();
                dictId = db.NamedObjectsDictionaryId;
                DBDictionary nod = (DBDictionary)tr.GetObject(dictId, OpenMode.ForWrite);
                dictId = nod.SetAt("MyWindowDict", dict);
                tr.AddNewlyCreatedDBObject(dict, true);
                tr.Commit();
            }

            return dictId;
        }


        void SubtractSolids(ObjectId targetSolidId, ObjectId subtractSolidId)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the target solid for write
                Solid3d targetSolid = tr.GetObject(targetSolidId, OpenMode.ForWrite) as Solid3d;
                Solid3d subtractSolid = tr.GetObject(subtractSolidId, OpenMode.ForWrite) as Solid3d;

                if (targetSolid != null && subtractSolid != null)
                {
                    try
                    {
                        targetSolid.BooleanOperation(BooleanOperationType.BoolSubtract, subtractSolid);
                        subtractSolid.Erase(); // Remove the subtracted solid after the operation
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                            $"\nError in SubtractSolids: {ex.Message}"
                        );
                    }
                }

                tr.Commit();
            }
        }


        void AddDictionaryData(Transaction tr, DBDictionary dict, string key, TypedValue typedValue)
        {
            using (ResultBuffer rb = new ResultBuffer())
            {
                rb.Add(typedValue);
                Xrecord xr = new Xrecord();
                xr.Data = rb;
                tr.AddNewlyCreatedDBObject(xr, true);
                dict.SetAt(key, xr);
            }
        }


        public static void SetUCSToWindowCS(WindowObject WindowObj)
        {
            U3(WindowObj.Origin, WindowObj.Origin + WindowObj.XAxis, WindowObj.Origin + WindowObj.YAxis);
        }


        ObjectId GetObjectIdFromHandle(Database db, Handle handle)
        {
            return db.GetObjectId(false, handle, 0);
        }


     
        public static SelectionSet GetAllWalls()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Create a selection filter to select all entities on the specified layer
            TypedValue[] filterValues = { new TypedValue((int)DxfCode.LayerName, WALLLAYER) };
            SelectionFilter filter = new SelectionFilter(filterValues);
            // Prompt user to select objects, but use the filter for walls
            PromptSelectionResult result = ed.SelectAll(filter);

            if (result.Status == PromptStatus.OK)
            {
                SelectionSet selectedWalls = result.Value;
            }
            else
            {
                ed.WriteMessage("\nNo walls found.");
            }
            return result.Value;
        }


        Solid3d CreateBoxWithVertexAtPoint(Polyline plRect, double Depth)
        {
            PLineRectInfo plRectInfo = new PLineRectInfo(plRect);
            Solid3d box = new Solid3d();
            box.CreateBox(Math.Abs(plRectInfo.XLength), Math.Abs(plRectInfo.YLength), Math.Abs(Depth));
            box.TransformBy(Matrix3d.Displacement(new Vector3d(plRectInfo.XLength / 2, plRectInfo.YLength / 2, Depth / 2)));
            box.TransformBy(Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                plRectInfo.Origin, plRectInfo.XAxis, plRectInfo.YAxis, plRectInfo.XAxis.CrossProduct(plRectInfo.YAxis)));
            return box;
        }



        static Solid3d CreateBox(Point3d Origin, Vector3d XAxis, Vector3d YAxis, double X, Double Y, double Z)
        {
            Solid3d box = new Solid3d();
            box.CreateBox(Math.Abs(X), Math.Abs(Y), Math.Abs(Z));
            box.TransformBy(Matrix3d.Displacement(new Vector3d(X / 2, Y / 2, Z / 2)));
            box.TransformBy(Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                Origin, XAxis, YAxis, XAxis.CrossProduct(YAxis)));
            return box;
        }



        static Solid3d CreateBoxAtCoordinateSystem(double X, Double Y, Double Z, Point3d Origin, Vector3d XAxis, Vector3d YAxis)
        {
            Solid3d box = new Solid3d();

            box.CreateBox(Math.Abs(X), Math.Abs(Y), Math.Abs(Z));
            box.TransformBy(Matrix3d.Displacement(new Vector3d(X / 2, Y / 2, Z / 2)));
            box.TransformBy(Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                Origin, XAxis, YAxis, XAxis.CrossProduct(YAxis)));
 
            return box;
        }


        public static ObjectId GetWindowDictionaryDictionaryObjectId()
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            ObjectId WindowDictDictId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary namedObjDict = tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;

                if (namedObjDict.Contains("WindowDictionaryDictionary"))
                {
                    WindowDictDictId = namedObjDict.GetAt("WindowDictionaryDictionary");
                }
                else
                {
                    namedObjDict.UpgradeOpen();
                    DBDictionary newWindowDict = new DBDictionary();
                    namedObjDict.SetAt("WindowDictionaryDictionary", newWindowDict);
                    tr.AddNewlyCreatedDBObject(newWindowDict, true);
                    WindowDictDictId = newWindowDict.ObjectId;

                }

                tr.Commit();
            }
            Transaction tr2 = db.TransactionManager.StartTransaction();
            DBDictionary namedObjDict2 = tr2.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
            WindowDictDictId = namedObjDict2.GetAt("WindowDictionaryDictionary");
            tr2.Commit();

            return WindowDictDictId;
        }


        public static void CreateNewWindowFromRectInWall(Polyline plRect)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;

            Handle ThisWindowsWallEntityHandle = new Handle();
            ObjectId ThisWindowsWallObjectId = new ObjectId();
            DBDictionary newWindowDict = new DBDictionary();
            Solid3d ThisWindowsWall = new Solid3d();
            bool bFoundWall = false;
            Solid3d WindowHole;
            ObjectId WindowHoleObjectId = new ObjectId();
            int Handedness = 0;
            WindowHole = new Solid3d();
            SelectionSet allWalls = GetAllWalls();
            ThisWindowsWall = new Solid3d();
            Solid3d windowBoxPos;
            Solid3d windowBoxNeg;
            ObjectId? windowBoxIntersectPosId;
            ObjectId? windowBoxIntersectNegId;
            PLineRectInfo PI = new PLineRectInfo(plRect);
            foreach (SelectedObject AllWallsSelectedObject in allWalls)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    windowBoxPos = CreateBox(PI.Origin, PI.XAxis, PI.YAxis, PI.XLength, PI.YLength, MAXWALLTHICKNESS);
                    windowBoxNeg = CreateBox(PI.Origin, PI.XAxis, PI.YAxis, PI.XLength, PI.YLength, MAXWALLTHICKNESS * -1);

                    acBlkTblRec.AppendEntity(windowBoxPos);
                    acBlkTblRec.AppendEntity(windowBoxNeg);
                    tr.AddNewlyCreatedDBObject(windowBoxPos, true);
                    tr.AddNewlyCreatedDBObject(windowBoxNeg, true);

                    tr.Commit();
                }

                windowBoxIntersectPosId = SpatialConflictSolid(windowBoxPos.ObjectId, AllWallsSelectedObject.ObjectId);
                windowBoxIntersectNegId = SpatialConflictSolid(windowBoxNeg.ObjectId, AllWallsSelectedObject.ObjectId);

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    windowBoxPos.Erase();
                    windowBoxNeg.Erase();
                    tr.Commit();
                }
            
                if (windowBoxIntersectPosId != ObjectId.Null)
                {
                    WindowHoleObjectId = (ObjectId)windowBoxIntersectPosId;
                    Handedness = 1;
                    bFoundWall = true;
                    ThisWindowsWallObjectId = AllWallsSelectedObject.ObjectId;
                    break;
                }
                if (windowBoxIntersectNegId != ObjectId.Null)
                {
                    WindowHoleObjectId = (ObjectId)windowBoxIntersectNegId;
                    Handedness = -1;
                    bFoundWall = true;
                    ThisWindowsWallObjectId = AllWallsSelectedObject.ObjectId;
                    break;
                }

            }

            if (bFoundWall) // Found a valid wall
            {
                double WindowHoleVolume;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    WindowHole = tr.GetObject(WindowHoleObjectId, OpenMode.ForWrite) as Solid3d;
                    WindowHoleVolume = Get3DSolidVolume(WindowHole);
                    WindowHole.Erase();
                    Solid3d WallObj = tr.GetObject(ThisWindowsWallObjectId, OpenMode.ForRead) as Solid3d;
                    ThisWindowsWallEntityHandle = WallObj.Handle;
                    tr.Commit();
                };
                double rectArea = plRect.Area;
                double wallThickness = WindowHoleVolume / rectArea;
                if (Handedness == -1) PI.Flip();
                WindowObject WO = new WindowObject(PI, ThisWindowsWallEntityHandle, wallThickness);

                DrawWindow(WO);
            }
            else
            {
                Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("No Window Created");
                return;
            }
        }

        //public static ObjectId DrawWindowFromRect(Polyline rect, Handle wallEntityHandle, double wallThickness, int Handedness)
        //{

        //    // Get rectangle's corner points in world coordinates
        //    //List<Point2d> ecsPts2D = Get2DECSPointsFromWLPline(rect);
        //    //Matrix3d entECS = GetECSTransformMatrix(rect);
        //    Point3d[] rectPts = Get3DPolylinePoints(rect);

        //    // Extract relevant points and vectors
        //    Point3d Origin = rectPts[0];
        //    Point3d xVectPt = rectPts[1];
        //    Point3d yVectPt = rectPts[3];

        //    Vector3d xVect = xVectPt - Origin;
        //    Vector3d yVect = yVectPt - Origin;

        //    double XLength = xVect.Length;
        //    double YLength = yVect.Length;

        //    Vector3d XAxis = xVect.GetNormal();
        //    Vector3d YAxis = yVect.GetNormal();
        //    //Vector3d ZAxis = XAxis.CrossProduct(YAxis);

        //    // Call DrawWindow with calculated parameters
        //    return DrawWindow(Origin, XAxis, YAxis, XLength, YLength, wallThickness, wallEntityHandle, Handedness);
        //}


        public static void DeleteWindow()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Prompt user to select an entity
            PromptEntityResult per = ed.GetEntity("\nPick window to delete: ");
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nEntity selection error");
                return;
            }

            // Check if the entity is part of a window
            ObjectId PossWindowDictId = IsEntityPartOfAWindow(per.ObjectId);
            if (PossWindowDictId != ObjectId.Null)
            {
                DeleteWindowAndWindowDictionary(PossWindowDictId, true);
                Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Window deleted");
            }
            else
            {
                Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("No Window found");
            }


        }


        public static ObjectId IsEntityPartOfAWindow(ObjectId entId)
        {
            ObjectId RetVal = ObjectId.Null;
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary WDD = tr.GetObject(WINDOWDICTIONARYDICTIONARYID, OpenMode.ForRead) as DBDictionary;

                foreach (var entry in WDD)
                {
                    ObjectId currWindowDictId = entry.Value;
                    DBDictionary currWindowDict = (DBDictionary)tr.GetObject(currWindowDictId, OpenMode.ForRead);
                    Handle FrameEntHandle = GetXRecordHandle(tr, currWindowDict, "FrameEntHandle");
                    Handle GlassEntHandle = GetXRecordHandle(tr, currWindowDict, "GlassEntHandle");
                    ObjectId FrameEntId = db.GetObjectId(false, FrameEntHandle, 0);
                    ObjectId GlassEntId = db.GetObjectId(false, GlassEntHandle, 0);
                    if ((FrameEntId == entId) || (GlassEntId == entId))
                    {
                        RetVal = currWindowDictId;
                        break;
                    }
                }
                tr.Commit();
            }
            return RetVal;

        }


        public static void DeleteWindowAndWindowDictionary(ObjectId windowDictId, bool sealHole)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            string Id;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary windowDict = tr.GetObject(windowDictId, OpenMode.ForRead) as DBDictionary;
                WindowObject windowObject = new WindowObject(tr, windowDict);
                // Get the Id to erase the window's dictionary in the windowdictionarydictionary
                Id = windowObject.Id;
                // Retrieve stored handles
                Handle frameEntHandle = windowObject.FrameEntHandle;
                Handle glassEntHandle = windowObject.GlassEntHandle;
                // Get their Ids
                ObjectId frameEntId = db.GetObjectId(false, frameEntHandle, 0);
                ObjectId glassEntId = db.GetObjectId(false, glassEntHandle, 0);
                // Get their objects
                Solid3d frameEnt = tr.GetObject(frameEntId, OpenMode.ForWrite) as Solid3d;
                Solid3d glassEnt = tr.GetObject(glassEntId, OpenMode.ForWrite) as Solid3d;
                // Erase them
                frameEnt.Erase();
                frameEnt.RecordGraphicsModified(true);
                glassEnt.Erase();
                glassEnt.RecordGraphicsModified(true);
                tr.Commit();
            }
            if (sealHole)
            {
                SealWindowHole(windowDictId);
            }
            // Now erase erase the window's dictionary in the windowdictionarydictionary
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary WindowDictionaryDictionary = tr.GetObject(WINDOWDICTIONARYDICTIONARYID, OpenMode.ForWrite) as DBDictionary;
                WindowDictionaryDictionary.Remove(Id);
                tr.Commit();
            }
        }

        static void SealWindowHole (ObjectId windowDictId)
        {

            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Solid3d WindowHole;
            WindowObject WO;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary windowDict = tr.GetObject(windowDictId, OpenMode.ForRead) as DBDictionary;
                WO = new WindowObject(tr, windowDict);
                WindowHole = CreateBoxAtCoordinateSystem(WO.XLength, WO.YLength, WO.WallThickness, WO.Origin,
                    WO.XAxis, WO.YAxis);
                ObjectId WindowHoleObjId = WindowHole.ObjectId;
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                btr.AppendEntity(WindowHole);
                tr.AddNewlyCreatedDBObject(WindowHole, true);
                tr.Commit();
            }
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId WallEntObjId = db.GetObjectId(false, WO.WallEntHandle, 0);
                WindowHole = tr.GetObject(WindowHole.ObjectId, OpenMode.ForWrite) as Solid3d;
                Solid3d Wall = tr.GetObject(WallEntObjId, OpenMode.ForWrite) as Solid3d;
                Wall.BooleanOperation(BooleanOperationType.BoolUnite, WindowHole);
                Wall.RecordGraphicsModified(true);
                tr.Commit();
            }

        }
        

        public static void MakeWindowHole(DBDictionary windowDict)
        {

            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Solid3d WindowHole;
            WindowObject WO;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                WO = new WindowObject(tr, windowDict);
                WindowHole = CreateBoxAtCoordinateSystem(WO.XLength, WO.YLength, WO.WallThickness, WO.Origin,
                    WO.XAxis, WO.YAxis);
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                btr.AppendEntity(WindowHole);
                tr.AddNewlyCreatedDBObject(WindowHole, true);
                tr.Commit();
            }
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId WallEntObjId = db.GetObjectId(false, WO.WallEntHandle, 0);
                WindowHole = tr.GetObject( WindowHole.ObjectId, OpenMode.ForWrite)as Solid3d;
                Solid3d Wall = tr.GetObject(WallEntObjId, OpenMode.ForWrite) as Solid3d;
                Wall.BooleanOperation(BooleanOperationType.BoolSubtract, WindowHole);
                Wall.RecordGraphicsModified(true);
                tr.Commit();
            }

        }




        // Draws the window, 
        //public static ObjectId DrawWindow(Point3d WindowOriginPt, Vector3d XAxis, Vector3d YAxis, double XLength, double YLength, double WallThickness, Handle WallEntHandle, int Handedness)
        //{
        //    Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
        //        tr.Commit();
        //    }
        //    Vector3d ZAxis;

        //    Point3d WindowOriginPt2;
        //    Point3d WindowOriginPt1;
        //    Solid3d FrameEnt1;
        //    Solid3d FrameEnt2;
        //    using (Transaction tr2 = db.TransactionManager.StartTransaction())
        //    {

        //        BlockTable bt2 = tr2.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        BlockTableRecord btr2 = tr2.GetObject(bt2[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

        //        ZAxis = XAxis.CrossProduct(YAxis);

        //        //Create Frame entity box1
        //        double OrgZOffset1 = (WallThickness - WINDOWFRAMETHICKNESS) / 2 * Handedness;
        //        WindowOriginPt1 = new Point3d();
        //        WindowOriginPt1 = WindowOriginPt.Add(ZAxis.MultiplyBy(OrgZOffset1));
        //        FrameEnt1 = CreateBoxAtCoordinateSystem(XLength, YLength, WINDOWFRAMETHICKNESS * Handedness, WindowOriginPt1, XAxis, YAxis);
        //        btr2.AppendEntity(FrameEnt1);

        //    //    tr2.Commit();
        //    //}

        //    //using (Transaction tr2 = db.TransactionManager.StartTransaction())
        //    //{
        //        // Create Frame entity box 2
        //        Vector3d FrameOffset = new Vector3d();
        //        FrameOffset = XAxis.MultiplyBy(WINDOWFRAMETHICKNESS) + YAxis.MultiplyBy(WINDOWFRAMETHICKNESS);
        //        WindowOriginPt2 = WindowOriginPt1.Add(FrameOffset);
        //        FrameEnt2 = CreateBoxAtCoordinateSystem(XLength - 2 * WINDOWFRAMETHICKNESS, YLength - 2 * WINDOWFRAMETHICKNESS,
        //            WINDOWFRAMETHICKNESS * Handedness, WindowOriginPt2, XAxis, YAxis);
        //        // Append them to modelspace and add them to the database
        //        //BlockTable bt2 = tr2.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        //BlockTableRecord btr2 = tr2.GetObject(bt2[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord; btr2.AppendEntity(FrameEnt1);
        //        tr2.AddNewlyCreatedDBObject(FrameEnt1, true);
        //        btr2.AppendEntity(FrameEnt2);
        //        tr2.AddNewlyCreatedDBObject(FrameEnt2, true);
        //        // Ccommit the database transaction so then a boolean operation can be performed on them to create the frame object
        //        tr2.Commit();
        //    }
        //    // Thusly -
        //    using (Transaction tr2 = db.TransactionManager.StartTransaction())
        //    {
        //        FrameEnt1 = tr2.GetObject(FrameEnt1.ObjectId, OpenMode.ForWrite) as Solid3d;
        //        FrameEnt2 = tr2.GetObject(FrameEnt2.ObjectId, OpenMode.ForWrite) as Solid3d;

        //        FrameEnt1.BooleanOperation(BooleanOperationType.BoolSubtract, FrameEnt2);
        //        //Change layer -
        //        FrameEnt1.Layer = WINDOWFRAMELAYER;
        //        FrameEnt1.RecordGraphicsModified(true);
        //        tr2.Commit();
        //    }

        //    // Now do the glass:
        //    Solid3d GlassEnt;
        //    using (Transaction tr2 = db.TransactionManager.StartTransaction())
        //    {
        //        Point3d WindowOriginPt3 = new Point3d();
        //        WindowOriginPt3 = WindowOriginPt2.Add(ZAxis.MultiplyBy((WINDOWFRAMETHICKNESS - GLASSTHICKNESS) / (2 * Handedness)));
        //        GlassEnt = CreateBoxAtCoordinateSystem(XLength - 2 * WINDOWFRAMETHICKNESS, YLength - 2 * WINDOWFRAMETHICKNESS,
        //            GLASSTHICKNESS * Handedness, WindowOriginPt3, XAxis, YAxis);
        //        GlassEnt.Layer = WINDOWGLASSLAYER;
        //        GlassEnt.RecordGraphicsModified(true);
        //        BlockTable bt2 = tr2.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        BlockTableRecord btr2 = tr2.GetObject(bt2[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
        //        btr2.AppendEntity(GlassEnt);
        //        tr2.AddNewlyCreatedDBObject(GlassEnt, true);
        //        tr2.Commit();
        //    }

        //    // Create dictionary
        //    DBDictionary WindowDict;
        //    using (Transaction tr2 = db.TransactionManager.StartTransaction())
        //    {
        //        WindowDict = new DBDictionary();
        //        DBDictionary WindowDictionaryDictionary = tr2.GetObject(WINDOWDICTIONARYDICTIONARYID, OpenMode.ForWrite) as DBDictionary;
        //        string newUniqueId = Guid.NewGuid().ToString("N").Substring(0, 8);
        //        WindowDictionaryDictionary.SetAt(newUniqueId, WindowDict);
        //        tr2.AddNewlyCreatedDBObject(WindowDict, true);
        //        SetXRecordPoint3d(tr2, WindowDict, "Origin", WindowOriginPt);
        //        SetXRecordText(tr2, WindowDict, "Id", newUniqueId);
        //        SetXRecordReal(tr2, WindowDict, "XLength", XLength);
        //        SetXRecordReal(tr2, WindowDict, "YLength", YLength);
        //        SetXRecordInt(tr2, WindowDict, "Handedness", Handedness);
        //        SetXRecordHandle(tr2, WindowDict, "FrameEntHandle", FrameEnt1.Handle);
        //        SetXRecordHandle(tr2, WindowDict, "GlassEntHandle", GlassEnt.Handle);
        //        SetXRecordHandle(tr2, WindowDict, "WallEntHandle", WallEntHandle);
        //        SetXRecordReal(tr2, WindowDict, "WallThickness", WallThickness);
        //        SetXRecordVector3d(tr2, WindowDict, "XAxis", XAxis);
        //        SetXRecordVector3d(tr2, WindowDict, "YAxis", YAxis);
        //        tr2.Commit();
        //    }

        //    MakeWindowHole(WindowDict);

        //    return WindowDict.ObjectId;
        //}

        public static ObjectId AdjustWindow(ObjectId windowDictId, Point3d pt)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            WindowObject WO;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary windowDict = tr.GetObject(windowDictId, OpenMode.ForRead) as DBDictionary;
                WO = new WindowObject(tr, windowDict);
                tr.Commit();

            }

            // Find the closest corner to pt
            Point3d[] windowCornerPoints = WO.GetWindowPoints();
            int closestIndex = 0;
            double minDist = pt.DistanceTo(windowCornerPoints[0]);
            for (int i = 1; i < 4; i++)
            {
                double dist = pt.DistanceTo(windowCornerPoints[i]);
                if (dist < minDist)
                {
                    minDist = dist;
                    closestIndex = i;
                }
            }

            Point3d cornerPt = windowCornerPoints[closestIndex];
            Vector3d delta = pt - cornerPt;

            switch (closestIndex)
            {
                case 0:
                    WO.Origin += delta;
                    WO.XLength -= delta.DotProduct(WO.XAxis);
                    WO.YLength -= delta.DotProduct(WO.YAxis);
                    break;
                case 1:
                    WO.Origin += WO.YAxis * delta.DotProduct(WO.YAxis);
                    WO.YLength -= delta.DotProduct(WO.YAxis);
                    WO.XLength += delta.DotProduct(WO.XAxis);
                    break;
                case 2:
                    WO.XLength += delta.DotProduct(WO.XAxis);
                    WO.YLength += delta.DotProduct(WO.YAxis);
                    break;
                case 3:
                    WO.Origin += WO.XAxis * delta.DotProduct(WO.XAxis);
                    WO.XLength -= delta.DotProduct(WO.XAxis);
                    WO.YLength += delta.DotProduct(WO.YAxis);
                    break;
            }
            DeleteWindowAndWindowDictionary(windowDictId, true);
            return DrawWindow(WO);
        }


        public static ObjectId DrawWindow(WindowObject WO)
        {
            return DrawWindow2(WO.Origin, WO.XAxis, WO.YAxis, WO.XLength, WO.YLength, WO.WallThickness, WO.WallEntHandle);
        }


        public static void ModifyWindow()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect window:");
            peo.SetRejectMessage("\nInvalid selection. Select a valid window entity.");
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            ObjectId possWindowEnt = per.ObjectId;
            ObjectId PossWDObjectId = IsEntityPartOfAWindow(possWindowEnt);
            Matrix3d CurrentUCS;
            if (PossWDObjectId != ObjectId.Null)
            {
              
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBDictionary windowDict = (DBDictionary)tr.GetObject(PossWDObjectId, OpenMode.ForRead);
                    WindowObject WO = new WindowObject(tr, windowDict);

                    CurrentUCS = ed.CurrentUserCoordinateSystem;
                    ed.CurrentUserCoordinateSystem = Matrix3d.AlignCoordinateSystem(
                    Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                    WO.Origin, WO.XAxis, WO.YAxis, WO.XAxis.CrossProduct(WO.YAxis));

                    tr.Commit();
                }


                bool continueModifying = true;
                while (continueModifying)
                {
                    PromptPointOptions ppo = new PromptPointOptions("\nEnter point to adjust the corners of this window or eXit [eXit] <eXit>:");
                    ppo.Keywords.Add("Exit");

                    //ppo.AllowNone = true;

                    PromptPointResult ppr = ed.GetPoint(ppo);

                    if (ppr.Status != PromptStatus.OK)
                    {
                        continueModifying = false;
                        break;
                    }
                    Point3d possInputPoint = ppr.Value.TransformBy(ed.CurrentUserCoordinateSystem);

                    PossWDObjectId = AdjustWindow(PossWDObjectId, possInputPoint);
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    ed.CurrentUserCoordinateSystem = CurrentUCS;
                    tr.Commit();
                }
 
                //vpRecord.SetUcs(ucsTable[previousUCS].ObjectId);


            }
            else
            {
                ed.WriteMessage("\nNot part of any window in the window dictionary.");
            }
        }

        public static ObjectId DrawWindow2(Point3d WindowOriginPt, Vector3d XAxis, Vector3d YAxis, double XLength, double YLength, double WallThickness, Handle WallEntHandle)
        {
            Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                tr.Commit();
            }
            Vector3d ZAxis;

            Point3d WindowOriginPt2;
            Point3d WindowOriginPt1;
            Solid3d FrameEnt1;
            Solid3d FrameEnt2;
            using (Transaction tr2 = db.TransactionManager.StartTransaction())
            {

                BlockTable bt2 = tr2.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr2 = tr2.GetObject(bt2[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                ZAxis = XAxis.CrossProduct(YAxis);

                //Create Frame entity box1
                double OrgZOffset1 = (WallThickness - WINDOWFRAMETHICKNESS) / 2;
                WindowOriginPt1 = new Point3d();
                WindowOriginPt1 = WindowOriginPt.Add(ZAxis.MultiplyBy(OrgZOffset1));
                FrameEnt1 = CreateBoxAtCoordinateSystem(XLength, YLength, WINDOWFRAMETHICKNESS, WindowOriginPt1, XAxis, YAxis);
                btr2.AppendEntity(FrameEnt1);

                //    tr2.Commit();
                //}

                //using (Transaction tr2 = db.TransactionManager.StartTransaction())
                //{
                // Create Frame entity box 2
                Vector3d FrameOffset = new Vector3d();
                FrameOffset = XAxis.MultiplyBy(WINDOWFRAMETHICKNESS) + YAxis.MultiplyBy(WINDOWFRAMETHICKNESS);
                WindowOriginPt2 = WindowOriginPt1.Add(FrameOffset);
                FrameEnt2 = CreateBoxAtCoordinateSystem(XLength - 2 * WINDOWFRAMETHICKNESS, YLength - 2 * WINDOWFRAMETHICKNESS,
                    WINDOWFRAMETHICKNESS, WindowOriginPt2, XAxis, YAxis);
                // Append them to modelspace and add them to the database
                //BlockTable bt2 = tr2.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                //BlockTableRecord btr2 = tr2.GetObject(bt2[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord; btr2.AppendEntity(FrameEnt1);
                tr2.AddNewlyCreatedDBObject(FrameEnt1, true);
                btr2.AppendEntity(FrameEnt2);
                tr2.AddNewlyCreatedDBObject(FrameEnt2, true);
                // Ccommit the database transaction so then a boolean operation can be performed on them to create the frame object
                tr2.Commit();
            }
            // Thusly -
            using (Transaction tr2 = db.TransactionManager.StartTransaction())
            {
                FrameEnt1 = tr2.GetObject(FrameEnt1.ObjectId, OpenMode.ForWrite) as Solid3d;
                FrameEnt2 = tr2.GetObject(FrameEnt2.ObjectId, OpenMode.ForWrite) as Solid3d;

                FrameEnt1.BooleanOperation(BooleanOperationType.BoolSubtract, FrameEnt2);
                //Change layer -
                FrameEnt1.Layer = WINDOWFRAMELAYER;
                FrameEnt1.RecordGraphicsModified(true);
                tr2.Commit();
            }

            // Now do the glass:
            Solid3d GlassEnt;
            using (Transaction tr2 = db.TransactionManager.StartTransaction())
            {
                Point3d WindowOriginPt3 = new Point3d();
                WindowOriginPt3 = WindowOriginPt2.Add(ZAxis.MultiplyBy((WINDOWFRAMETHICKNESS - GLASSTHICKNESS) / 2));
                GlassEnt = CreateBoxAtCoordinateSystem(XLength - 2 * WINDOWFRAMETHICKNESS, YLength - 2 * WINDOWFRAMETHICKNESS,
                    GLASSTHICKNESS, WindowOriginPt3, XAxis, YAxis);
                GlassEnt.Layer = WINDOWGLASSLAYER;
                GlassEnt.RecordGraphicsModified(true);
                BlockTable bt2 = tr2.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr2 = tr2.GetObject(bt2[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                btr2.AppendEntity(GlassEnt);
                tr2.AddNewlyCreatedDBObject(GlassEnt, true);
                tr2.Commit();
            }


            // Create dictionary
            DBDictionary WindowDict;
            using (Transaction tr2 = db.TransactionManager.StartTransaction())
            {
                WindowDict = new DBDictionary();
                DBDictionary WindowDictionaryDictionary = tr2.GetObject(WINDOWDICTIONARYDICTIONARYID, OpenMode.ForWrite) as DBDictionary;
                string newUniqueId = Guid.NewGuid().ToString("N").Substring(0, 8);
                WindowDictionaryDictionary.SetAt(newUniqueId, WindowDict);
                tr2.AddNewlyCreatedDBObject(WindowDict, true);
                SetXRecordPoint3d(tr2, WindowDict, "Origin", WindowOriginPt);
                SetXRecordText(tr2, WindowDict, "Id", newUniqueId);
                SetXRecordReal(tr2, WindowDict, "XLength", XLength);
                SetXRecordReal(tr2, WindowDict, "YLength", YLength);
                SetXRecordInt(tr2, WindowDict, "Handedness", 1);
                SetXRecordHandle(tr2, WindowDict, "FrameEntHandle", FrameEnt1.Handle);
                SetXRecordHandle(tr2, WindowDict, "GlassEntHandle", GlassEnt.Handle);
                SetXRecordHandle(tr2, WindowDict, "WallEntHandle", WallEntHandle);
                SetXRecordReal(tr2, WindowDict, "WallThickness", WallThickness);
                SetXRecordVector3d(tr2, WindowDict, "XAxis", XAxis);
                SetXRecordVector3d(tr2, WindowDict, "YAxis", YAxis);
                tr2.Commit();
            }

            MakeWindowHole(WindowDict);

            return WindowDict.ObjectId;
        }

        //public void Test ()
        //{
        //    Database db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
        //        .

        //    }
        //}


        //[CommandMethod("SetNewUCS")]
        //public void SetNewUCS()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        // Get the UCS Table
        //        UcsTable ucsTable = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;

        //        // Store current UCS
        //        //ViewportTableRecord vpRecord = tr.GetObject(ed.CurrentViewportTableRecordId, OpenMode.ForRead) as ViewportTableRecord;
        //        previousUCS = vpRecord.UcsName; // Save the previous UCS

        //        // Create a new UCS definition
        //        UcsTableRecord newUcs = new UcsTableRecord
        //        {
        //            Name = "MyNewUCS",
        //            Origin = new Point3d(10, 10, 0), // Example new origin
        //            XAxis = new Vector3d(1, 0, 0),
        //            YAxis = new Vector3d(0, 1, 0)
        //        };

        //        // Open the UCS table for write
        //        if (!ucsTable.Has(newUcs.Name))
        //        {
        //            ucsTable.UpgradeOpen();
        //            ucsTable.Add(newUcs);
        //            tr.AddNewlyCreatedDBObject(newUcs, true);
        //        }

        //        // Set the new UCS as current
        //        vpRecord.UpgradeOpen();
        //        vpRecord.SetUcs(newUcs.ObjectId);
        //        //vpRecord.UcsPerViewport = true; // Set UCS per viewport

        //        tr.Commit();
        //        ed.WriteMessage("\nNew UCS set as current.");
        //    }
        //}

        //[CommandMethod("RestorePreviousUCS")]
        //public void RestorePreviousUCS()
        //{
        //    if (previousUCS == ObjectId.Null)
        //    {
        //        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nNo previous UCS saved.");
        //        return;
        //    }

        //    Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        ViewportTableRecord vpRecord = tr.GetObject(ed.CurrentViewportTableRecordId, OpenMode.ForWrite) as ViewportTableRecord;

        //        // Restore the previous UCS
        //        vpRecord.SetUcs(previousUCS);
        //        //vpRecord.UcsPerViewport = true;

        //        tr.Commit();
        //        ed.WriteMessage("\nPrevious UCS restored.");
        //    }
        //}
    }


}
