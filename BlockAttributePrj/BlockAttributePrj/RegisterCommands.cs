using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcadService = Autodesk.AutoCAD.ApplicationServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace BlockAttributePrj
{   
    public class RegisterCommands
    {
        [CommandMethod("ModifyBlock")]
        public void ModifyBlock()
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;
            //Init Keywords/Option
            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
            pKeyOpts.Message = "\nEnter an option ";
            pKeyOpts.Keywords.Add("Select Block");
            pKeyOpts.Keywords.Add("All");
            pKeyOpts.Keywords.Add("By Name");
            pKeyOpts.Keywords.Default = "Select Block";
            pKeyOpts.AllowNone = true;
            //Get KeyWord
            PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);
            //If Get Keyword seccess
            if (pKeyRes.Status == PromptStatus.OK)
            {
                if (pKeyRes.StringResult.Contains("S"))
                {
                    try
                    {
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            PromptSelectionOptions opts = new PromptSelectionOptions();
                            opts.MessageForAdding = "Select entities: ";
                            //Get Block from User
                            PromptSelectionResult selRes = ed.GetSelection(opts);
                            if (selRes.Status != PromptStatus.OK)
                            {
                                return;
                            }

                            if (selRes.Value.Count != 0)
                            {
                                SelectionSet set = selRes.Value;
                                // Init ViewModel
                                ModifyBlockViewModel vm = new ModifyBlockViewModel();
                                vm.OKCmd = new MVVMCore.RelayCommand(OkCmdInvoke);
                                vm.CancelCmd = new MVVMCore.RelayCommand(CancelCmdInvoke);
                                vm.EditAllCmd = new MVVMCore.RelayCommand(EditAllCmdInvoke);

                                foreach (ObjectId id in set.GetObjectIds())
                                {
                                    BlockReference oEnt = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                                    BlockReferenceProperties blkproperties = new BlockReferenceProperties();
                                    blkproperties.BlockName = oEnt.Name;
                                    blkproperties.BlockId = id;
                                    if (oEnt.AttributeCollection.Count > 0)
                                    {

                                        AttributeCollection AttriCollection = oEnt.AttributeCollection;
                                        foreach (ObjectId objId in AttriCollection)
                                        {
                                            if (objId.IsValid)
                                            {
                                                AttributeReference attRef = (AttributeReference)tr.GetObject(objId, OpenMode.ForRead);
                                                BlockAttribute blkAttr = new BlockAttribute();
                                                blkAttr.AtrributeTag = attRef.Tag;
                                                blkAttr.Value = attRef.TextString;
                                                blkAttr.AtrrId = objId;
                                                blkAttr.TextHeight = attRef.Height;

                                                blkAttr.CheckInvisible = attRef.Visible;
                                                blkproperties.blkAttribute.Add(blkAttr);
                                            }

                                        }
                                    }
                                    vm.BlkProperties.Add(blkproperties);
                                }
                                //Init View
                                ModifyBlock view = new ModifyBlock();
                                //Set Default Selected
                                vm.blkSelected = vm.BlkProperties[0];
                                vm.blkAttributes = vm.blkSelected.blkAttribute;
                                vm.blkAttributeSelected = vm.blkAttributes[0];
                                view.DataContext = vm;
                                AcadService.Application.ShowModalWindow(view);
                            }
                            tr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(ex.ToString());
                    }
                }
                else if (pKeyRes.StringResult == "All")
                {
                    PromptSelectionResult selRes = ed.SelectAll();
                    try
                    {
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            if (selRes.Value.Count != 0)
                            {
                                SelectionSet set = selRes.Value;
                                // Init ViewModel
                                ModifyBlockViewModel vm = new ModifyBlockViewModel();
                                // Register Relay Command
                                vm.OKCmd = new MVVMCore.RelayCommand(OkCmdInvoke);
                                vm.CancelCmd = new MVVMCore.RelayCommand(CancelCmdInvoke);
                                vm.EditAllCmd = new MVVMCore.RelayCommand(EditAllCmdInvoke);
                                //Put All Block data to Window
                                foreach (ObjectId id in set.GetObjectIds())
                                {
                                    BlockReference oEnt = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                                    BlockReferenceProperties blkproperties = new BlockReferenceProperties();
                                    blkproperties.BlockName = oEnt.Name;
                                    blkproperties.BlockId = id;
                                    if (oEnt.AttributeCollection.Count > 0)
                                    {

                                        AttributeCollection AttriCollection = oEnt.AttributeCollection;
                                        foreach (ObjectId objId in AttriCollection)
                                        {
                                            if (objId.IsValid)
                                            {
                                                AttributeReference attRef = (AttributeReference)tr.GetObject(objId, OpenMode.ForRead);
                                                BlockAttribute blkAttr = new BlockAttribute();
                                                blkAttr.AtrributeTag = attRef.Tag;
                                                blkAttr.Value = attRef.TextString;
                                                blkAttr.TextHeight = attRef.Height;
                                                blkAttr.AtrrId = objId;
                                                blkAttr.CheckInvisible = attRef.Visible;
                                                blkproperties.blkAttribute.Add(blkAttr);
                                            }

                                        }
                                    }
                                    vm.BlkProperties.Add(blkproperties);
                                }
                                //Init View
                                ModifyBlock view = new ModifyBlock();
                                //Set Default Selected
                                vm.blkSelected = vm.BlkProperties[0];
                                vm.blkAttributes = vm.blkSelected.blkAttribute;
                                vm.blkAttributeSelected = vm.blkAttributes[0];
                                //Set DataContext
                                view.DataContext = vm;
                                //Show Dialog
                                AcadService.Application.ShowModalWindow(view);
                            }
                            tr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(ex.ToString());
                    }
                }
                else if (pKeyRes.StringResult.Contains("B"))
                {
                    PromptSelectionOptions opts = new PromptSelectionOptions();
                    opts.SingleOnly = true;
                    opts.MessageForAdding = "Select entities: ";
                    //Get Block from User
                    PromptSelectionResult selRes = ed.GetSelection(opts);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    PromptSelectionResult selRes2 = ed.SelectAll();
                    //Start Transaction
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        if (selRes.Value.Count != 0 && selRes.Value.Count != 0)
                        {
                            SelectionSet set = selRes.Value;
                            SelectionSet set2 = selRes2.Value;
                            // Init ViewModel
                            ModifyBlockViewModel vm = new ModifyBlockViewModel();
                            vm.OKCmd = new MVVMCore.RelayCommand(OkCmdInvoke);
                            vm.CancelCmd = new MVVMCore.RelayCommand(CancelCmdInvoke);
                            vm.EditAllCmd = new MVVMCore.RelayCommand(EditAllCmdInvoke);
                            ObjectId[] singleId = set.GetObjectIds();
                            if (singleId[0].IsValid == false)
                            {
                                return;
                            }
                            BlockReference SingleEnt = (BlockReference)tr.GetObject(singleId[0], OpenMode.ForWrite);
                            foreach (ObjectId id in set2.GetObjectIds())
                            {
                                BlockReference oEnt = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                                if (SingleEnt.Name != oEnt.Name)
                                {
                                    continue;
                                }
                                BlockReferenceProperties blkproperties = new BlockReferenceProperties();
                                blkproperties.BlockName = oEnt.Name;
                                blkproperties.BlockId = id;
                                if (oEnt.AttributeCollection.Count > 0)
                                {

                                    AttributeCollection AttriCollection = oEnt.AttributeCollection;
                                    foreach (ObjectId objId in AttriCollection)
                                    {
                                        if (objId.IsValid)
                                        {
                                            AttributeReference attRef = (AttributeReference)tr.GetObject(objId, OpenMode.ForRead);
                                            BlockAttribute blkAttr = new BlockAttribute();
                                            blkAttr.AtrributeTag = attRef.Tag;
                                            blkAttr.Value = attRef.TextString;
                                            blkAttr.AtrrId = objId;
                                            blkAttr.CheckInvisible = attRef.Visible;
                                            blkAttr.TextHeight = attRef.Height;
                                            blkproperties.blkAttribute.Add(blkAttr);
                                        }

                                    }
                                }
                                vm.BlkProperties.Add(blkproperties);
                            }
                            //Init View
                            ModifyBlock view = new ModifyBlock();
                            //Set Default Selected
                            vm.blkSelected = vm.BlkProperties[0];
                            vm.blkAttributes = vm.blkSelected.blkAttribute;
                            vm.blkAttributeSelected = vm.blkAttributes[0];
                            view.DataContext = vm;
                            AcadService.Application.ShowModalWindow(view);
                        }
                        tr.Commit();
                    }

                }
            }
            else
            {
                //Can not get keyword
                return;
            }

        }
        void OkCmdInvoke(object obj)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;
            var dlg = obj as ModifyBlock;
            if (dlg != null)
            {
                var vm = dlg.DataContext as ModifyBlockViewModel;
                if (vm != null)
                {
                    if (vm.BlkProperties != null)
                    {
                        foreach (var blkprop in vm.BlkProperties)
                        {
                            if (blkprop.blkAttribute != null)
                            {
                                foreach (BlockAttribute attr in blkprop.blkAttribute)
                                {
                                    using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                                    {

                                        AttributeReference pEnt = tr.GetObject(attr.AtrrId, OpenMode.ForWrite) as AttributeReference;
                                        if (pEnt != null)
                                        {
                                            if (attr.CheckInvisible == false)
                                            {
                                                pEnt.Visible = false;

                                            }
                                            else
                                            {
                                                pEnt.Visible = true;
                                            }
                                            pEnt.TextString = attr.Value;
                                            pEnt.Height = attr.TextHeight;
                                        }
                                        tr.Commit();
                                    }

                                }

                            }

                        }

                    }

                }
            }
            dlg.Close();
        }

        void CancelCmdInvoke(object obj)
        {
            var dlg = obj as ModifyBlock;
            dlg.Close();
        }
        void EditAllCmdInvoke(object obj)
        {
            ModifyBlock modifyView = obj as ModifyBlock;
            if (modifyView != null)
            {
                ModifyBlockViewModel modifyVm = modifyView.DataContext as ModifyBlockViewModel;
                if (modifyView != null)
                {
                    EditAllBlockView view = new EditAllBlockView();
                    EditAllBlockViewModel vm = new EditAllBlockViewModel();
                    vm.OKEditALlCmd = new MVVMCore.RelayCommand(OKEditALlCmdInvoke);
                    vm.CancelEditAllCmd = new MVVMCore.RelayCommand(CancelEditAllCmdInvoke);
                    view.DataContext = vm;
                    view.Owner = modifyView;
                    if (view.ShowDialog() == true)
                    {
                        if (vm.Invisible == true)
                        {
                            foreach (BlockReferenceProperties blk in modifyVm.BlkProperties)
                            {
                                foreach (BlockAttribute attr in blk.blkAttribute)
                                {
                                    attr.CheckInvisible = false;
                                }
                            }
                           
                        }
                        else
                        {
                            foreach (BlockReferenceProperties blk in modifyVm.BlkProperties)
                            {
                                foreach (BlockAttribute attr in blk.blkAttribute)
                                {
                                    attr.CheckInvisible = true;
                                }
                            }
                           
                        }
                        if (vm.TextHeight>0)
                        {
                            foreach (BlockReferenceProperties blk in modifyVm.BlkProperties)
                            {
                                foreach (BlockAttribute attr in blk.blkAttribute)
                                {
                                    attr.TextHeight = vm.TextHeight;
                                }
                            }
                        }
                        
                    }
                }
              
            }
            
        }
        void OKEditALlCmdInvoke(object obj)
        {
            var dlg = obj as EditAllBlockView;
            dlg.DialogResult = true;
            dlg.Close();
        }
        void CancelEditAllCmdInvoke(object obj)
        {
            var dlg = obj as EditAllBlockView;
            dlg.Close();
        }
        [CommandMethod("rotateBlocks")]
        public void rotateBlocks()
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.MessageForAdding = "Select entities: ";
            //Get Block from User
            PromptSelectionResult selRes = ed.GetSelection(opts);
            if (selRes.Status != PromptStatus.OK)
            {
                return;
            }
            if (selRes.Value.Count != 0)
            {
                SelectionSet set = selRes.Value;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in set.GetObjectIds())
                    {
                        BlockReference oEnt = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                        Point3d basePoint = oEnt.Position;
                        oEnt.TransformBy(Matrix3d.Rotation(3.1412, Vector3d.ZAxis, basePoint));
                    }
                    tr.Commit();
                }
            }
        }
        [CommandMethod("CreateArcBetweenSelectedVetices")]

        static public void CreateArcBetweenSelectedVetices()

        {
            Document doc =Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            //Select Polyline
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.SingleOnly = true;
            opts.MessageForAdding = "Select Polyline: ";
            PromptSelectionResult selRes = ed.GetSelection(opts);
            if (selRes.Status != PromptStatus.OK)
            {
                return;
            }
            SelectionSet set = selRes.Value;
            ObjectId[] singleId = set.GetObjectIds();
            if (singleId[0].IsValid == false)
            {
                return;
            }
            ObjectIdCollection idSegment = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Polyline pl = (Polyline)tr.GetObject(singleId[0], OpenMode.ForWrite);
                DBObjectCollection segments = new DBObjectCollection();
                pl.Explode(segments);
                foreach ( Entity seg in segments)
                {
                    ObjectId id= ArxHelper.AppendEntity(seg);
                    if (id.IsValid)
                    {
                        idSegment.Add(id);
                    }
                }
                pl.Visible = false;
                tr.Commit();                
            }
            //Select Vertex
            PromptSelectionOptions opts2 = new PromptSelectionOptions();
            opts2.MessageForAdding = "Select Segments: ";
            PromptSelectionResult selRes2 = ed.GetSelection(opts2);
            if (selRes2.Status != PromptStatus.OK)
            {
                return;
            }
            SelectionSet set2 = selRes2.Value;
            ObjectId[] listSegmentSelected = set2.GetObjectIds();
            if (listSegmentSelected.Length<1)
            {
                return;
            }
            //Get middle point of line selected
            Point3dCollection ptListMiddle = new Point3dCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in listSegmentSelected)
                {
                   
                    Line ent = tr.GetObject(id, OpenMode.ForWrite) as Line;
                    if (ent != null)
                    {
                        ent.ColorIndex = 1;
                        Point3d pts = ent.StartPoint;
                        Point3d pte = ent.EndPoint;
                        Point3d mid = new Point3d((pts.X + pte.X) / 2, (pts.Y + pte.Y) / 2, (pts.Z + pte.Z) / 2);
                        ptListMiddle.Add(mid);
                    }
                    
                }
                tr.Commit();
            }
            //Get vertex 
            ObservableCollection<int> vertexindex = new ObservableCollection<int>(); 
            foreach (Point3d pt in ptListMiddle)
            {
                int seg =ArxHelper.GetVertexIndexes(pt, singleId[0]);
                vertexindex.Add(seg);
            }
            foreach (ObjectId id in idSegment)
            {
                ArxHelper.DeleteEntity(id);
            }
            List<int> listsegment = vertexindex.OrderBy(x => x).ToList();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Polyline pl = (Polyline)tr.GetObject(singleId[0], OpenMode.ForWrite);
                int j = 1;
                if (listsegment.Count > 1)
                {
                    ObservableCollection<SegmentInfor> segmentInfo = new ObservableCollection<SegmentInfor>();
                    SegmentInfor firstseg = new SegmentInfor();
                    firstseg.IndexSegment = listsegment[0];
                    LineSegment3d segment3d = pl.GetLineSegmentAt(listsegment[0]);

                    firstseg.StartPoint = segment3d.StartPoint;
                    firstseg.PartNumber = j;
                    segmentInfo.Add(firstseg);
                    for (int i = 1; i < listsegment.Count; i++)
                    {
                        if (listsegment[i] != listsegment[i - 1] + 1)
                        {
                            j++;
                        }
                        SegmentInfor partseg = new SegmentInfor();
                        partseg.PartNumber = j;
                        partseg.IndexSegment = listsegment[i];
                        LineSegment3d segmentpart3d = pl.GetLineSegmentAt(listsegment[i]);
                        partseg.StartPoint = new Point3d(segmentpart3d.StartPoint.X, segmentpart3d.StartPoint.Y, segmentpart3d.StartPoint.Z);
                        partseg.EndPoint = new Point3d(segmentpart3d.EndPoint.X, segmentpart3d.EndPoint.Y, segmentpart3d.EndPoint.Z);
                        segmentInfo.Add(partseg);
                    }
                    for (int i = 1; i <= j; i++)
                    {
                        List<SegmentInfor> seg = segmentInfo.Where(x => x.PartNumber == i).ToList();
                        if (seg.Count == 0)
                        {
                            pl.Visible = true;
                            return;
                        }
                        Point3d pt = new Point3d();
                        double distance = 0;
                        Point3d ptStart = seg[0].StartPoint;
                        Point3d ptEnd = new Point3d();

                        ptEnd = seg[seg.Count - 1].EndPoint;
                        Line line = new Line(ptStart, ptEnd);
                        if (seg.Count > 1)
                        {
                            //Find fatest point 
                            foreach (SegmentInfor info in seg)
                            {
                                Point3d ptmidle = line.GetClosestPointTo(info.StartPoint, false);
                                double dis = ptmidle.DistanceTo(info.StartPoint);
                                if (dis > distance)
                                {
                                    distance = dis;
                                    pt = info.StartPoint;
                                }
                            }
                            //Remove vertex
                            for (int n = 1; n < seg.Count; n++)
                            {
                                SegmentInfor info = seg[n];
                                if (pl.NumberOfVertices > 2)
                                {
                                    var vex = info.IndexSegment;
                                    if (info.IndexSegment >= pl.NumberOfVertices)
                                    {
                                        continue;
                                    }
                                    if (n == 1)
                                    {
                                        pl.RemoveVertexAt(vex);
                                        if (j > 0)
                                        {
                                            foreach (var inf in segmentInfo)
                                            {
                                                if (inf.IndexSegment > 0)
                                                {
                                                    inf.IndexSegment--;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        pl.RemoveVertexAt(vex - 1);
                                        if (j > 0)
                                        {
                                            foreach (var inf in segmentInfo)
                                            {
                                                if (inf.IndexSegment > 0)
                                                {
                                                    inf.IndexSegment--;
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                            // Calculate arc bulge
                            Point3d closetpoint = line.GetClosestPointTo(pt, false);
                            double fromStart = closetpoint.DistanceTo(ptStart);
                            double bul = (distance / fromStart);
                            int ind = i;
                            if (Clockwise(ptStart, pt, ptEnd))
                            {
                                
                                if (i>=seg.Count)
                                {
                                    ind = seg.Count - 1;
                                }
                                pl.SetBulgeAt(seg[ind].IndexSegment, -bul);
                            }
                            else
                            {
                                if (i >= seg.Count)
                                {
                                    ind = seg.Count - 1;
                                }
                                pl.SetBulgeAt(seg[ind].IndexSegment, bul);
                            }
                            pl.Visible = true;
                        }
                    }
                }

                tr.Commit();
            }


        }

        /// <summary>
        /// Evaluates if the points are clockwise.
        /// </summary>
        /// <param name="p1">First point.</param>
        /// <param name="p2">Second point</param>
        /// <param name="p3">Third point</param>
        /// <returns>True if points are clockwise, False otherwise.</returns>
        private static bool Clockwise(Point3d p1, Point3d p2, Point3d p3)
        {
            return ((p2.X - p1.X) * (p3.Y - p1.Y) - (p2.Y - p1.Y) * (p3.X - p1.X)) < 1e-8;
        }
        [CommandMethod("CreateBuildingLines")]
        public static void CreateBuildingLines()
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;
            TypedValue[] filList = new TypedValue[1] {new TypedValue((int)DxfCode.Start, "INSERT")};

            SelectionFilter filter = new SelectionFilter(filList);

            PromptSelectionOptions opts =new PromptSelectionOptions();

            opts.MessageForAdding = "Select block references: ";
            PromptSelectionResult res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }

            if (res.Value.Count != 0)
            {
                SelectionSet set = res.Value;
                ObjectId[] Ids = set.GetObjectIds();
                //Select 2 block
                if (Ids.Length == 2)
                {
                    ObjectId Id1 = Ids[0];
                    ObjectId Id2 = Ids[1];
                    if (Id1.IsValid && Id2.IsValid)
                    {
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockReference blk1 = tr.GetObject(Id1,OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2,OpenMode.ForRead) as BlockReference;
                            if (blk1 != null && blk2 != null)
                            {
                                Point3d point1 = blk1.Position;
                                Point3d point2 = blk2.Position;
                                PromptAngleOptions dir = new PromptAngleOptions("\nGet Direction");
                                dir.BasePoint = point1;
                                dir.UseBasePoint = true;
                                PromptDoubleResult ptRes2 = ed.GetAngle(dir);
                                double angle = 0;
                                if (ptRes2.Status == PromptStatus.OK)
                                {
                                    angle = ptRes2.Value;
                                    Line line = new Line(point1,point2);
                                    Vector3d mainVect = point2 - point1;
                                    Vector3d xvect = Vector3d.XAxis;
                                    double mainangle = xvect.GetAngleTo(mainVect);
                                    Vector3d vectdir = point2 - point1;
                                    Point3d EndPoint = new Point3d();
                                    Point3d StartPoint = new Point3d();
                                    if (mainangle > angle)
                                    {
                                        EndPoint = point1 + vectdir.GetPerpendicularVector().GetNormal() * (-3);
                                        StartPoint = point2 + vectdir.GetPerpendicularVector().GetNormal() * (-3);
                                    }
                                    else
                                    {
                                        EndPoint = point1 + vectdir.GetPerpendicularVector().GetNormal() * (3);
                                        StartPoint = point2 + vectdir.GetPerpendicularVector().GetNormal() * (3);
                                    }
                                    Point3dCollection vertex = new Point3dCollection();
                                    vertex.Add(EndPoint);
                                    vertex.Add(point1);
                                    vertex.Add(point2);
                                    vertex.Add(StartPoint);
                                    Polyline3d pl3d = new Polyline3d(Poly3dType.SimplePoly, vertex, false);
                                    ArxHelper.AppendEntity(pl3d);
                                    vertex.Dispose();
                                }

                            }
                            tr.Commit();
                        }
                    }
                }
                //Select 4 block
                else if (Ids.Length == 4)
                {
                    ObjectId Id1 = Ids[0];
                    ObjectId Id2 = Ids[1];
                    ObjectId Id3 = Ids[2];
                    ObjectId Id4 = Ids[3];
                    if (Id1.IsValid && Id2.IsValid && Id3.IsValid && Id4.IsValid)
                    {
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockReference blk1 = tr.GetObject(Id1, OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2, OpenMode.ForRead) as BlockReference;
                            BlockReference blk3 = tr.GetObject(Id3, OpenMode.ForRead) as BlockReference;
                            BlockReference blk4 = tr.GetObject(Id4, OpenMode.ForRead) as BlockReference;
                            if (blk1!= null && blk2!= null && blk3!= null && blk4!= null)
                            {
                                Point3d pt1 = blk1.Position;
                                Point3d pt2 = blk2.Position;
                                Point3d pt3 = blk3.Position;
                                Point3d pt4 = blk4.Position;
                                Vector3d v12 = pt2 - pt1;
                                Vector3d v13 = pt3 - pt1;
                                Vector3d v14 = pt4 - pt1;
                                Vector3d v23 = pt3 - pt2;
                                Vector3d v24 = pt4 - pt2;
                                Vector3d v34 = pt4 - pt3;
                                if (v12.IsParallelTo(v34))
                                {
                                    Line line12 = new Line(pt1, pt2);
                                    Line line34 = new Line(pt3, pt4);
                                    Point3d ptex1 = line12.GetClosestPointTo(pt3, true);
                                    Point3d ptex2 = line34.GetClosestPointTo(pt4, true);
                                    Point3dCollection points1 = new Point3dCollection();
                                    points1.Add(ptex1);
                                    points1.Add(pt3);
                                    points1.Add(pt4);
                                    points1.Add(ptex2);
                                    Polyline3d pl1 = new Polyline3d(Poly3dType.SimplePoly,points1,false);
                                    ArxHelper.AppendEntity(pl1);
                                }
                                else if (v13.IsParallelTo(v24))
                                {

                                }
                                else if (v14.IsParallelTo(v23))
                                {

                                }
                            }
                            tr.Commit();
                        }
                    }
                }
            }
        }
    }
}
