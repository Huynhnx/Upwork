using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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
        public static double _PI = 3.14159265358979323846;
        public static double _2PI = 6.28318530717958647693;
        public static double _HALFPI = 1.57079632679489661923;
        public static Tolerance Tol = new Tolerance(0.0872222222222222222222, 0.0872222222222222222222);
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
        public static void BuildLine(Line longline,Line shortline)
        {
            Point3d ptex1 = new Point3d();
            Point3d ptex2 = new Point3d();
            ptex1 = longline.GetClosestPointTo(shortline.StartPoint, true);
            ptex2 = longline.GetClosestPointTo(shortline.EndPoint, true);
            Point3dCollection points1 = new Point3dCollection();
            Point3dCollection points2 = new Point3dCollection();
            points1.Add(ptex1);
            points1.Add(shortline.StartPoint);
            points1.Add(shortline.EndPoint);
            points1.Add(ptex2);
            Polyline3d pl1 = new Polyline3d(Poly3dType.SimplePoly, points1, false);
            ArxHelper.AppendEntity(pl1);
            Vector3d dir = (ptex1 - shortline.StartPoint);
            Point3d pt11 = longline.StartPoint + dir.GetNormal()*3;
            Point3d pt21 = longline.EndPoint + dir.GetNormal()*3;
            if (ptex1.DistanceTo(longline.StartPoint)+ptex1.DistanceTo(longline.EndPoint)> longline.Length)
            {
                if (ptex1.DistanceTo(longline.StartPoint)< ptex1.DistanceTo(longline.EndPoint))
                {
                    Point3d pt = new Point3d();
                    pt = ptex1 + (ptex1- longline.EndPoint).GetNormal()*3;
                    Line l = new Line(pt, ptex1);
                    AppendEntity(l);
                    points2.Add(pt);
                    points2.Add(longline.StartPoint);
                    points2.Add(longline.EndPoint);
                    points2.Add(pt21);
                }
                else
                {
                    Point3d pt = new Point3d();
                    pt = ptex2 +(ptex2 - longline.StartPoint).GetNormal()*3;
                    points2.Add(pt);
                    points2.Add(longline.StartPoint);
                    points2.Add(longline.EndPoint);
                    points2.Add(pt21);
                }
            }
            else
            {
                points2.Add(pt11);
                points2.Add(longline.StartPoint);
                points2.Add(longline.EndPoint);
                points2.Add(pt21);
            }
           
            Polyline3d pl2 = new Polyline3d(Poly3dType.SimplePoly, points2, false);
            ArxHelper.AppendEntity(pl2);
        }
        public static bool IsLongestLine(Point3d pt1,Point3d pt2,Point3dCollection pts)
        {
            if (pts.Contains(pt1) && pts.Contains(pt2))
            {
                pts.Remove(pt2);
                pts.Remove(pt1);
            }
            Point3d pt3 = pts[0];
            Point3d pt4 = pts[1];
            if (pt1.DistanceTo(pt2)>pt3.DistanceTo(pt4) &&
                pt1.DistanceTo(pt2)> pt4.DistanceTo(pt1)&&
                pt1.DistanceTo(pt2)> pt2.DistanceTo(pt3)&&
                pt1.DistanceTo(pt2)> pt2.DistanceTo(pt4)&&
                pt1.DistanceTo(pt2) > pt3.DistanceTo(pt1)
                )
            {
                Line l12 = new Line(pt1, pt2);
                Line l34 = new Line(pt3, pt4);
                Point3dCollection lispoint = new Point3dCollection();
                l12.IntersectWith(l34, Intersect.OnBothOperands, lispoint, IntPtr.Zero, IntPtr.Zero);
                if (lispoint.Count==0)
                {
                    return true;
                }
            }
            return false;
        }
        public static bool BuilPolyline(Point3d pt1,Point3d pt2,Point3d closerPoint,Point3d CloserPointOnLine,Point3d FurtherPoint,Point3d FurthurPointOnLine)
        {
            //Build Longer line
           
            Point3d pt7 = pt1 + (CloserPointOnLine-closerPoint).GetNormal() * 3;
            Point3d pt8 = pt2 + (CloserPointOnLine - closerPoint).GetNormal() * 3;
            Point3dCollection verx = new Point3dCollection();
            verx.Add(pt7);
            verx.Add(pt1);
            verx.Add(pt2);
            verx.Add(pt8);
            Polyline3d longer = new Polyline3d(Poly3dType.SimplePoly, verx, false);
           
            //Build Shorter Line
            Xline xline = new Xline();
            xline.BasePoint = closerPoint;
            xline.UnitDir = pt1 -pt2;
            Line l48 = new Line(FurthurPointOnLine, FurtherPoint);
            Point3dCollection intersec = new Point3dCollection();
            xline.IntersectWith(l48, Intersect.OnBothOperands, intersec, IntPtr.Zero, IntPtr.Zero);
            if (intersec.Count < 1)
            {
                return false;
            }
            Xline xl1 = new Xline();
            if (pt1.DistanceTo(FurtherPoint)< pt2.DistanceTo(FurtherPoint))
            {
                xl1.BasePoint = pt1;
            }
            else
            {
                xl1.BasePoint = pt2;
            }
            xl1.UnitDir = pt7 - pt1;
            Xline xl2 = new Xline();
            xl2.BasePoint = FurtherPoint;
            xl2.UnitDir = pt2 - pt1;
            Point3dCollection intersec2 = new Point3dCollection();
            xl1.IntersectWith(xl2, Intersect.OnBothOperands, intersec2, IntPtr.Zero, IntPtr.Zero);

            Point3dCollection verx2 = new Point3dCollection();
            if (intersec2.Count <1)
            {
                return false;
            }
            ArxHelper.AppendEntity(longer);
            if (pt1.DistanceTo(FurtherPoint) < pt2.DistanceTo(FurtherPoint))
            {
                verx2.Add(pt1);
                verx2.Add(intersec2[0]);
                verx2.Add(FurtherPoint);
                verx2.Add(intersec[0]);
                verx2.Add(closerPoint);
                verx2.Add(CloserPointOnLine);
                Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx2, false);
                ArxHelper.AppendEntity(pl);
                return true;
            }
            else
            {
                verx2.Add(pt2);
                verx2.Add(intersec2[0]);
                verx2.Add(FurtherPoint);
                verx2.Add(intersec[0]);
                verx2.Add(closerPoint);
                verx2.Add(CloserPointOnLine);
                Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx2, false);
                ArxHelper.AppendEntity(pl);
                return true;
            }
        }
        public static bool BuildPolylineFrom3Point(Point3d pt1, Point3d pt2, Point3d pt3)
        {
            Vector3d v12 = (pt2 - pt1).GetNormal();
            Vector3d v13 = (pt3 - pt1).GetNormal();
            Vector3d v23 = (pt3 - pt2).GetNormal();
            if (v12.IsPerpendicularTo(v13, ArxHelper.Tol))
            {
                Point3dCollection verx = new Point3dCollection();
                verx.Add(pt3 + (pt2 - pt1).GetNormal() * 3);
                verx.Add(pt3);
                verx.Add(pt1);
                verx.Add(pt2);
                verx.Add(pt2 + (pt3 - pt1).GetNormal() * 3);
                Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, false);
                AppendEntity(pl);
                return true;
            }
            else if (v13.IsPerpendicularTo(v23, ArxHelper.Tol))
            {
                Point3dCollection verx = new Point3dCollection();
                verx.Add(pt1 + (pt2 - pt3).GetNormal() * 3);
                verx.Add(pt1);
                verx.Add(pt3);
                verx.Add(pt2);
                verx.Add(pt2 + (pt1 - pt3).GetNormal() * 3);
                Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, false);
                AppendEntity(pl);
                return true;
            }
            else if (v12.IsPerpendicularTo(v23, ArxHelper.Tol))
            {
                Point3dCollection verx = new Point3dCollection();
                verx.Add(pt3 + (pt1 - pt2).GetNormal() * 3);
                verx.Add(pt3);
                verx.Add(pt2);
                verx.Add(pt1);
                verx.Add(pt1 + (pt3 - pt2).GetNormal() * 3);
                Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, false);
                AppendEntity(pl);
                return true;
            }
            return false;
        }
        public static bool BuildPolylineFrom2Point(Point3d pt1, Point3d pt2)
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;
            PromptPointOptions op = new PromptPointOptions("Pick Direction:");
            op.UseBasePoint = true;
            op.BasePoint = pt1;
            Line line = new Line(pt1, pt2);
            PromptPointResult r = ed.GetPoint(op);
            if (r.Status == PromptStatus.OK)
            {
                Point3dCollection pts = new Point3dCollection();
                Point3d pt = line.GetClosestPointTo(r.Value, true);
                Point3d EndPoint = pt1 + (r.Value - pt).GetNormal() * 3;
                Point3d StartPoint = pt2 + (r.Value - pt).GetNormal() * 3;
                Point3dCollection vertex = new Point3dCollection();
                vertex.Add(EndPoint);
                vertex.Add(pt1);
                vertex.Add(pt2);
                vertex.Add(StartPoint);
                Polyline3d pl3d = new Polyline3d(Poly3dType.SimplePoly, vertex, false);
                ArxHelper.AppendEntity(pl3d);
                vertex.Dispose();
                return true;
            }
            return false;
        }
    }
    public class SegmentInfor
    {
        public int IndexSegment;
        public int PartNumber;
        public Point3d StartPoint;
        public Point3d EndPoint;

    }
}
