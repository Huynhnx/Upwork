using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockAttributePrj
{
    class ArxHelper
    {
        public static ObjectId AppendEntity(Entity ent)
        {
            ObjectId objId = ObjectId.Null;
            if (ent == null)
            {
                return objId;
            }
            Document doc = Application.DocumentManager.MdiActiveDocument;
            try
            {
                using (doc.LockDocument())
                {
                    Database db = doc.Database;
                    // start new transaction
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        // open model space block table record
                        BlockTableRecord spaceBlkTblRec = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        // append entity to model space block table record
                        objId = spaceBlkTblRec.AppendEntity(ent);
                        trans.AddNewlyCreatedDBObject(ent, true);
                        // finish
                        trans.Commit();
                    }
                }

            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage(ex.ToString());
            }
            return objId;
        }
        public static void DeleteEntity(ObjectId id)
        {
            if (ObjectId.Null == id || id.IsErased)
                return;
            using (Application.DocumentManager.MdiActiveDocument.LockDocument())
            {
                Database db = Application.DocumentManager.MdiActiveDocument.Database;
                // start new transaction
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    using (DBObject obj = trans.GetObject(id, OpenMode.ForWrite))
                    {
                        if (null != obj && !obj.IsErased)
                        {
                            obj.Erase();
                        }
                    }
                    // finish
                    trans.Commit();
                }
            }
        }
        public static int GetVertexIndexes(Point3d pickedPosition, ObjectId polyId)
        {
            int first = 0;
            int second = 0;

            using (var tran = polyId.Database.TransactionManager.StartOpenCloseTransaction())
            {
                var poly = (Autodesk.AutoCAD.DatabaseServices.Polyline)
                    tran.GetObject(polyId, OpenMode.ForRead);

                var closestPoint = poly.GetClosestPointTo(pickedPosition, false);
                var len = poly.GetDistAtPoint(closestPoint);

                for (int i = 1; i < poly.NumberOfVertices - 1; i++)
                {
                    var pt1 = poly.GetPoint3dAt(i);
                    var l1 = poly.GetDistAtPoint(pt1);

                    var pt2 = poly.GetPoint3dAt(i + 1);
                    var l2 = poly.GetDistAtPoint(pt2);

                    if (len > l1 && len < l2)
                    {
                        first = i;
                        second = i + 1;
                        break;
                    }
                }
                tran.Commit();
            }
            return first;
        }
    }
}
