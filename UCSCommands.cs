using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace ArchitecturalWindows
{
    public class UCSCommands
    {
        private static ObjectId previousUcsId = ObjectId.Null; // Store the previous UCS ObjectId
        private static bool wasWorldUCS = true; // Track if the previous UCS was World UCS

     
        public static void SetNewUCS()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get the UCS table
                UcsTable ucsTable = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;

                // Get the current viewport settings
                ObjectId viewportId = db.CurrentViewportTableRecordId;
                ViewportTableRecord vpRecord = tr.GetObject(viewportId, OpenMode.ForWrite) as ViewportTableRecord;
                previousUcsId = vpRecord.UcsName;
                // Store the current UCS
                if (ucsTable.Has(vpRecord.UcsName))
                {
                    previousUcsId = vpRecord.UcsName;
                    wasWorldUCS = false;
                }
                else
                {
                    previousUcsId = ObjectId.Null;
                    wasWorldUCS = true;
                }

                // Define a new UCS
                string ucsName = "MyNewUCS";
                UcsTableRecord newUcs;

                if (ucsTable.Has(ucsName))
                {
                    newUcs = tr.GetObject(ucsTable[ucsName], OpenMode.ForRead) as UcsTableRecord;
                }
                else
                {
                    ucsTable.UpgradeOpen();
                    newUcs = new UcsTableRecord
                    {
                        Name = ucsName,
                        Origin = new Point3d(10, 10, 0), // Example new origin
                        XAxis = new Vector3d(0, 0, 1),
                        YAxis = new Vector3d(0, 1, 0)
                    };
                    ucsTable.Add(newUcs);
                    tr.AddNewlyCreatedDBObject(newUcs, true);
                }

                // Set the new UCS to the viewport
                vpRecord.SetUcs(newUcs.ObjectId);
                vpRecord.UcsSavedWithViewport = true; // Mark UCS as saved with this viewport

                tr.Commit();
                ed.WriteMessage("\nNew UCS set as current.");
            }
        }

        [CommandMethod("RestorePreviousUCS")]
        public void RestorePreviousUCS()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get the UCS table
                UcsTable ucsTable = tr.GetObject(db.UcsTableId, OpenMode.ForRead) as UcsTable;

                // Get the current viewport settings
                ObjectId viewportId = db.CurrentViewportTableRecordId;
                ViewportTableRecord vpRecord = tr.GetObject(viewportId, OpenMode.ForWrite) as ViewportTableRecord;

                if (wasWorldUCS)
                {
                    // Restore to World UCS
                    vpRecord.SetUcs(ObjectId.Null);
                    vpRecord.UcsSavedWithViewport = false;
                }
                else
                {
                    // Restore to the previous UCS
                    if (!previousUcsId.IsNull && ucsTable.Has(previousUcsId))
                    {
                        vpRecord.SetUcs(previousUcsId);
                        vpRecord.UcsSavedWithViewport = true;
                    }
                    else
                    {
                        ed.WriteMessage("\nPrevious UCS was not found in the UCS table.");
                    }
                }

                tr.Commit();
                ed.WriteMessage("\nPrevious UCS restored.");
            }
        }
    }
}
