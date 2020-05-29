using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Linq;

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
        public static void AddRegAppTableRecord(string regAppName)
        {
            Document doc =  Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {

                RegAppTable rat =  (RegAppTable)tr.GetObject( db.RegAppTableId, OpenMode.ForRead, false);
                if (!rat.Has(regAppName))
                {
                    rat.UpgradeOpen();

                    RegAppTableRecord ratr = new RegAppTableRecord();
                    ratr.Name = regAppName;
                    rat.Add(ratr);
                    tr.AddNewlyCreatedDBObject(ratr, true);
                }
                tr.Commit();
            }
        }
        static public ObjectIdCollection ReadXdata(Entity ent)
        {
            Document doc =  Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ResultBuffer rb = ent.GetXDataForApplication("ListBlock");
            ObjectIdCollection listId = new ObjectIdCollection();
            if (rb == null)
            {
                return null;
            }
            else
            {
                foreach (TypedValue tv in rb)
                {
                    if (tv.TypeCode == (int)DxfCode.SoftPointerId)
                    {
                        ObjectId id = (ObjectId)tv.Value;
                        if (id.IsValid && id.IsErased ==false)
                        {
                            listId.Add(id);
                        }
                    }             
                }
                rb.Dispose();
            }
            return listId;
        }
        public static bool WriteXdata(DBObject ent, ResultBuffer newData)
        {
            bool ret = false;
            if (ent != null)
            {
                         // Name      
                ent.XData = newData;
                ret = true;
            }
            return ret;
        }
        public static void AttachDataToEntity(ObjectId entId, ResultBuffer rb, string AppName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            // now start a transaction 
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Declare an Entity variable named ent.  
                    Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;
                    // test the IsNull property of the ExtensionDictionary of the ent. 
                    if (ent.ExtensionDictionary.IsNull)
                    {
                        // Upgrade the open of the entity. 
                        ent.UpgradeOpen();

                        // Create the ExtensionDictionary
                        ent.CreateExtensionDictionary();
                    }
                    // variable as DBDictionary. 
                    DBDictionary extensionDict = (DBDictionary)trans.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite);
                    if (extensionDict != null)
                    {
                        //  Create a new XRecord. 
                        Xrecord myXrecord = new Xrecord();
                      
                        // Add the ResultBuffer to the Xrecord 
                        myXrecord.Data = rb;
                        // Create the entry in the ExtensionDictionary. 
                        extensionDict.SetAt(AppName, myXrecord);
                        // Tell the transaction about the newly created Xrecord 
                        trans.AddNewlyCreatedDBObject(myXrecord, true);
                    }
                    // all ok, commit it 
                    trans.Commit();
                    trans.Dispose();
                }
                catch (System.Exception ex)
                {
                    trans.Abort();
                }
            }
        }
        public static bool ReadEntityData(ObjectId entId, ref ResultBuffer rb, string AppName)
        {
            bool bRet = false;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Declare an Entity variable named ent. 
                    if (entId.IsValid && entId.IsErased == false)
                    {
                        Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;                    
                        if (ent != null && ent.ExtensionDictionary.IsValid && !ent.ExtensionDictionary.IsErased)
                        {
                            // Declare an Entity variable named ent.                  
                            // variable as DBDictionary.                        
                            DBDictionary extensionDict = trans.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                            if (extensionDict != null && extensionDict.Contains(AppName))
                            {
                                // Check to see if the entry we are going to add is already there. 
                                ObjectId entryId = extensionDict.GetAt(AppName);
                                // Extract the Xrecord. Declare an Xrecord variable. 
                                Xrecord myXrecord = default(Xrecord);
                                // Instantiate the Xrecord variable 
                                myXrecord = (Xrecord)trans.GetObject(entryId, OpenMode.ForRead);
                                if (myXrecord != null && myXrecord.Data != null)
                                {
                                   rb = myXrecord.Data;  
                                    
                                }
                            }
                        }
                    }
                        trans.Commit();
                }
                catch (System.Exception ex)
                {
                        trans.Abort();
                }
            }
            return bRet;
        }

        /// <summary>
        public static int GetVertexIndexes(Point3d pickedPosition, ObjectId polyId)
        {
            int first = 0;
            int second = 0;

            using (var tran = polyId.Database.TransactionManager.StartOpenCloseTransaction())
            {
                var poly = (Polyline)
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
        /// <summary>
        /// Get closet poin with given Point
        /// </summary>
        /// <param name="givenPoint"></param>
        /// <param name="pts"></param>
        /// <returns></returns>
        public static Point3d GetClosetPoint(Point3d givenPoint,Point3dCollection pts)
        {
            double dis = givenPoint.DistanceTo(pts[0]);
            int index = 0;
            for(int i=0;i< pts.Count;i++)
            {
                Point3d pt = pts[i];
                if (dis> pt.DistanceTo(givenPoint))
                {
                    dis = pt.DistanceTo(givenPoint);
                    index = i;
                }
            }
            return pts[index];
        }
        public static bool IsPointOnPolyline(Polyline pl, Point3d pt)
        {
            bool isOn = false;
            for (int i = 0; i < pl.NumberOfVertices; i++)
            {
                Curve3d seg = null;
                SegmentType segType = pl.GetSegmentType(i);
                if (segType == SegmentType.Arc)
                    seg = pl.GetArcSegmentAt(i);
                else if (segType == SegmentType.Line)
                    seg = pl.GetLineSegmentAt(i);
                if (seg != null)
                {
                    isOn = seg.IsOn(pt);
                    if (isOn)
                        break;
                }
            }
            return isOn;
        }
        /// <summary>
        /// Evaluates if the points are clockwise.
        /// </summary>
        /// <param name="p1">First point.</param>
        /// <param name="p2">Second point</param>
        /// <param name="p3">Third point</param>
        /// <returns>True if points are clockwise, False otherwise.</returns>
        public static bool Clockwise(Point3d p1, Point3d p2, Point3d p3)
        {
            return ((p2.X - p1.X) * (p3.Y - p1.Y) - (p2.Y - p1.Y) * (p3.X - p1.X)) < 1e-8;
        }
        public static double CalculateBulge(Point3d ptStart, Point3d corner, Point3d ptEnd)
        {
            Line line = new Line(ptStart, ptEnd);
            Point3d center = new Point3d();       
            Circle cir = new Circle();
            double p = (line.Length + ptStart.DistanceTo(corner) + ptEnd.DistanceTo(corner)) / 2;
            double S = Math.Sqrt((p - line.Length) * (p - ptStart.DistanceTo(corner)) * (p - ptEnd.DistanceTo(corner)) * p);
            cir.Radius = (line.Length * ptStart.DistanceTo(corner) * ptEnd.DistanceTo(corner)) / (4 * S);
            //Find Center
            // midle line 
            Point3d mid = new Point3d((ptEnd.X + ptStart.X) / 2, (ptEnd.Y + ptStart.Y) / 2, (ptEnd.Z + ptStart.Z) / 2);
            Xline xline1 = new Xline();
            xline1.BasePoint = mid;
            xline1.UnitDir = (ptStart - ptEnd).GetPerpendicularVector();

            Point3d mid2 = new Point3d((ptEnd.X + corner.X) / 2, (ptEnd.Y + corner.Y) / 2, (ptEnd.Z + corner.Z) / 2);
            Xline xline2 = new Xline();
            xline2.BasePoint = mid2;
            xline2.UnitDir = (corner - ptEnd).GetPerpendicularVector();
            Point3dCollection pts = new Point3dCollection();
            xline2.IntersectWith(xline1, Intersect.ExtendArgument, pts, IntPtr.Zero, IntPtr.Zero);
            if (pts.Count == 0)
            {
                return 0;
            }
            cir.Center = pts[0];
            center = cir.Center;
            Vector3d v1 = ptStart - cir.Center;
            Vector3d v2 = corner - cir.Center;
            double angle = v1.GetAngleTo(v2);

            double bul = Math.Tan(angle / 4);

            if (Clockwise(ptStart, corner, ptEnd))
            {
                bul = -1 * bul;
            }
            else
            {

            }
            return bul;
        }
        public static void CreateStairFrom3Point(Point3d pt1, Point3d pt2, Point3d pt3,double stairWidth,string Layer )
        {
            Line l12 = new Line(pt1, pt2);
            Point3d ptonL12 = l12.GetClosestPointTo(pt3, false);
            Xline xl1 = new Xline();
            xl1.BasePoint = pt1;
            xl1.UnitDir = pt3 - ptonL12;
            Xline xl3 = new Xline();
            xl3.BasePoint = pt3;
            xl3.UnitDir = pt1 - pt2;
            Point3dCollection pts3 = new Point3dCollection();
            Point3dCollection pts4 = new Point3dCollection();
            xl1.IntersectWith(xl3, Intersect.OnBothOperands, pts3, IntPtr.Zero, IntPtr.Zero);
            Point3d p3 = new Point3d();
            Point3d p4 = new Point3d();
            if (pts3.Count > 0)
            {
                p3 = pts3[0];
            }
            else
            {
                return;
            }
            Xline xl2 = new Xline();
            xl2.BasePoint = pt2;
            xl2.UnitDir = pt3 - ptonL12;
            xl2.IntersectWith(xl3, Intersect.OnBothOperands, pts4, IntPtr.Zero, IntPtr.Zero);
            if (pts4.Count > 0)
            {
                p4 = pts4[0];
            }
            else
            {
                return;
            }
            Point3dCollection verx = new Point3dCollection();
            verx.Add(pt1);
            verx.Add(pt2);
            verx.Add(p4);
            verx.Add(p3);
            Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, true);
            pl.Layer = Layer;
            ArxHelper.AppendEntity(pl);
            for (int i = 1; i < p4.DistanceTo(pt2) / stairWidth; i++)
            {
                Point3d p1 = pt1 + (p4 - pt2).GetNormal() * i * stairWidth;
                Point3d p2 = pt2 + (p4 - pt2).GetNormal() * i * stairWidth;
                Line line = new Line(p1, p2);
                line.Layer = Layer;
                ArxHelper.AppendEntity(line);
            }
            Point3d mid1 = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, (pt1.Z + pt2.Z) / 2);
            Point3d mid2 = new Point3d((p3.X + p4.X) / 2, (p3.Y + p4.Y) / 2, (p3.Z + p4.Z) / 2);
            if (mid1.Z< mid2.Z)
            {
                drawArrow(mid1, mid2, pt1 - pt2, pt1.DistanceTo(pt2), stairWidth);
            }
            else
            {
                drawArrow(mid2, mid1, pt1 - pt2, pt1.DistanceTo(pt2), stairWidth);
            }
           
        }
        public static void CreateStairFrom4Point(Point3d pt1, Point3d pt2, Point3d pt3,Point3d pt4, double stairWidth, string Layer)
        {
            Point3dCollection verx = new Point3dCollection();
            verx.Add(pt1);
            verx.Add(pt2);
            verx.Add(pt3);
            verx.Add(pt4);
            // Draw arrow
            Point3dCollection arrowPoint = new Point3dCollection();
            Point3d mid1= new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, (pt1.Z + pt2.Z) / 2);
            Point3d mid2 = new Point3d((pt3.X + pt4.X) / 2, (pt3.Y + pt4.Y) / 2, (pt3.Z + pt4.Z) / 2);
            if (mid1.Z < mid2.Z)
            {
                drawArrow(mid1, mid2, pt1 - pt2, pt1.DistanceTo(pt2), stairWidth);
            }
            else
            {
                drawArrow(mid2, mid1, pt1 - pt2, pt1.DistanceTo(pt2), stairWidth);
            }
            Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, true);
            pl.Layer = Layer;
            ArxHelper.AppendEntity(pl);
            for (int i = 1; i < pt3.DistanceTo(pt2) / stairWidth; i++)
            {
                Point3d p1 = pt1 + (pt3 - pt2).GetNormal() * i * stairWidth;
                Point3d p2 = pt2 + (pt3 - pt2).GetNormal() * i * stairWidth;
                Line line = new Line(p1, p2);
                line.Layer = Layer;
                ArxHelper.AppendEntity(line);
            }
        }
        public static void drawArrow(Point3d ptStart, Point3d ptEnd,Vector3d vect,double width , double stairwidth)
        {
            Point3d pt = ptEnd + (ptStart - ptEnd).GetNormal()* stairwidth;
            Point3d pt1 = pt + vect.GetNormal() * width / 10;
            Point3d pt2 = pt - vect.GetNormal() * width / 10;
            Line line1 = new Line(pt1, ptEnd);
            Line line2 = new Line(pt2, ptEnd);
            Line line3 = new Line(ptStart, ptEnd);
            AppendEntity(line1);
            AppendEntity(line2);
            AppendEntity(line3);
        }
        public static Point3d FindBigestCorner(Point3d pt1,Point3d pt2, Point3d pt3)
        {
            double anglePt1 = (pt2 - pt1).GetAngleTo(pt3 - pt1);
            double anglePt2 = (pt1 - pt2).GetAngleTo(pt3 - pt2);
            double anglePt3 = (pt1 - pt3).GetAngleTo(pt2 - pt3);
            if (anglePt1 > anglePt2 && anglePt1 > anglePt3)
            {
                return pt1;
            }
            else if (anglePt2 > anglePt1 && anglePt2 > anglePt3)
            {
                return pt2;
            }
            else
                return pt3;
        }
        public static ObjectIdCollection  GetEntitiesOnLayer(string layerName)

        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Build a filter list so that only entities
            // on the specified layer are selected
            TypedValue[] tvs = new TypedValue[1] {new TypedValue( (int)DxfCode.LayerName, layerName)};
            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);
            if (psr.Status == PromptStatus.OK)
                return new ObjectIdCollection( psr.Value.GetObjectIds() );
            else
                return new ObjectIdCollection();
        }
        public static bool IsPolygonInsideOther(ObjectId outplId,ObjectId inplId)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                Polyline plOut = acTrans.GetObject(outplId,OpenMode.ForRead) as Polyline;
                Polyline plIn = acTrans.GetObject(inplId, OpenMode.ForRead) as Polyline;
                for (int i=0;i< plIn.NumberOfVertices;i++)
                {
                    Point3d pt = plIn.GetPoint3dAt(i);
                    if (InsidePolygon(plOut,pt) == false)
                    {
                        acTrans.Abort();
                        return false;
                    }
                }
                acTrans.Abort();
            }
            return true;
        }
        public static bool InsidePolygon(Polyline polygon, Point3d pt)
        {
            int n = polygon.NumberOfVertices;
            double angle = 0;
            Point pt1, pt2;

            for (int i = 0; i < n; i++)
            {
                pt1.X = polygon.GetPoint2dAt(i).X - pt.X;
                pt1.Y = polygon.GetPoint2dAt(i).Y - pt.Y;
                pt2.X = polygon.GetPoint2dAt((i + 1) % n).X - pt.X;
                pt2.Y = polygon.GetPoint2dAt((i + 1) % n).Y - pt.Y;
                angle += Angle2D(pt1.X, pt1.Y, pt2.X, pt2.Y);
            }

            if (Math.Abs(angle) < Math.PI)
                return false;
            else
                return true;
        }
        public struct Point
        {
            public double X, Y;
        };

        /*
           
        */
        /// <summary>
        /// Return the angle between two vectors on a plane
        /// The angle is from vector 1 to vector 2, positive anticlockwise
        /// The result is between -pi -> pi
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        /// <returns></returns>
        public static double Angle2D(double x1, double y1, double x2, double y2)
        {
            double dtheta, theta1, theta2;

            theta1 = Math.Atan2(y1, x1);
            theta2 = Math.Atan2(y2, x2);
            dtheta = theta2 - theta1;
            while (dtheta > Math.PI)
                dtheta -= (Math.PI * 2);
            while (dtheta < -Math.PI)
                dtheta += (Math.PI * 2);
            return (dtheta);
        }
        public static void SortPoint(Polyline pl,Point3dCollection pts)
        {
            foreach (Point3d pt in pts)
            {

            }
        }
    }
    public class SegmentInfor
    {
        public int IndexSegment;
        public int PartNumber;
        public Point3d StartPoint;
        public Point3d EndPoint;

    }
    public class PolygonOrderInfor
    {
        public string LayerName;
        public string Date;
        public int priority;
        public ObjectIdCollection Ids = new ObjectIdCollection();
    }
}
