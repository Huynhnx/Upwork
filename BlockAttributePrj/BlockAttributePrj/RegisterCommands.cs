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
            Document doc = Application.DocumentManager.MdiActiveDocument;
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
                foreach (Entity seg in segments)
                {
                    ObjectId id = ArxHelper.AppendEntity(seg);
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
            //Get middle point of line selected
            Point3dCollection ptListMiddle = new Point3dCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in listSegmentSelected)
                {
                    Line ent = tr.GetObject(id, OpenMode.ForRead) as Line;
                    if (ent != null)
                    {
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
                int seg = ArxHelper.GetVertexIndexes(pt, singleId[0]);
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
                pl.Visible = true;
                if (listSegmentSelected.Length < 2)
                {
                    return;
                }
                int j = 1;
                if (listsegment.Count > 1)
                {
                    ObservableCollection<SegmentInfor> segmentInfo = new ObservableCollection<SegmentInfor>();
                    SegmentInfor firstseg = new SegmentInfor();
                    firstseg.IndexSegment = listsegment[0];
                    LineSegment3d segment3d = pl.GetLineSegmentAt(listsegment[0]);

                    firstseg.StartPoint = segment3d.StartPoint;
                    firstseg.EndPoint = segment3d.EndPoint;
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
                        partseg.StartPoint = segmentpart3d.StartPoint;
                        partseg.EndPoint = segmentpart3d.EndPoint;
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

                        if (seg.Count > 1)
                        {
                            double bul = 0.0;
                            Point3d center = new Point3d();
                            for (int n = 0; n < seg.Count - 1; n++)
                            {
                                Point3d ptStart = seg[n].StartPoint;
                                Point3d ptEnd = seg[n + 1].EndPoint;

                                Line line = new Line(ptStart, ptEnd);
                                
                                SegmentInfor info = seg[n];
                                Point3d corner = info.EndPoint;
                                Circle cir = new Circle();
                                double p = (line.Length + ptStart.DistanceTo(corner) + ptEnd.DistanceTo(corner)) / 2;
                                double S = Math.Sqrt((p - line.Length) * (p - ptStart.DistanceTo(corner)) * (p - ptEnd.DistanceTo(corner))*p);
                                cir.Radius = (line.Length * ptStart.DistanceTo(corner)* ptEnd.DistanceTo(corner))/(4*S);

                                //Find Center
                                // midle line 
                                Point3d mid = new Point3d((ptEnd.X+ptStart.X)/2, (ptEnd.Y + ptStart.Y) / 2, (ptEnd.Z + ptStart.Z) / 2);
                                Xline xline1 = new Xline();
                                xline1.BasePoint = mid;
                                xline1.UnitDir = (ptStart-ptEnd).GetPerpendicularVector();
                                
                                Point3d mid2 = new Point3d((ptEnd.X + corner.X) / 2, (ptEnd.Y + corner.Y) / 2, (ptEnd.Z + corner.Z) / 2);
                                Xline xline2 = new Xline();
                                xline2.BasePoint = mid2;
                                xline2.UnitDir = (corner-ptEnd).GetPerpendicularVector();
                                Point3dCollection pts = new Point3dCollection();
                                xline2.IntersectWith(xline1, Intersect.ExtendArgument, pts, IntPtr.Zero, IntPtr.Zero);
                                cir.Center = pts[0];
                                center = cir.Center;
                                Vector3d v1 = seg[n].StartPoint - cir.Center;
                                Vector3d v2 = seg[n].EndPoint - cir.Center;
                                double angle = v1.GetAngleTo(v2);
                                
                                bul = Math.Tan(angle/4);
                               
                                if (Clockwise(ptStart, corner, ptEnd))
                                {
                                    bul = -1 * bul;
                                }
                                else
                                {
                                    
                                }
                                pl.SetBulgeAt(seg[n].IndexSegment, bul);
                            }
                            Vector3d v1last = seg[seg.Count - 1].StartPoint - center;
                            Vector3d v2last = seg[seg.Count - 1].EndPoint - center;
                            double anglelast = v1last.GetAngleTo(v2last);
                            bul = -Math.Tan(anglelast / 4);                           
                            pl.SetBulgeAt(seg[seg.Count - 1].IndexSegment, bul);
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
                            BlockReference blk1 = tr.GetObject(Id1, OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2, OpenMode.ForRead) as BlockReference;
                            if (blk1 != null && blk2 != null)
                            {
                                Point3d point1 = new Point3d(blk1.Position.X,blk1.Position.Y,0);
                                Point3d point2 = new Point3d(blk2.Position.X, blk2.Position.Y, 0);
                                ArxHelper.BuildPolylineFrom2Point(point1, point2);
                            }
                            tr.Commit();
                        }
                    }
                }
                else if (Ids.Length == 3)
                {
                    ObjectId Id1 = Ids[0];
                    ObjectId Id2 = Ids[1];
                    ObjectId Id3 = Ids[2];
                    if (Id1.IsValid && Id2.IsValid && Id3.IsValid)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockReference blk1 = tr.GetObject(Id1, OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2, OpenMode.ForRead) as BlockReference;
                            BlockReference blk3 = tr.GetObject(Id3, OpenMode.ForRead) as BlockReference;
                            Point3d pt1 = new Point3d(blk1.Position.X, blk1.Position.Y, 0);
                            Point3d pt2 = new Point3d(blk2.Position.X, blk2.Position.Y, 0);
                            Point3d pt3 = new Point3d(blk3.Position.X, blk3.Position.Y, 0);
                            ArxHelper.BuildPolylineFrom3Point(pt1,pt2,pt3);
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
                            if (blk1 != null && blk2 != null && blk3 != null && blk4 != null)
                            {
                                Point3d pt1 = new Point3d(blk1.Position.X, blk1.Position.Y, 0);
                                Point3d pt2 = new Point3d(blk2.Position.X, blk2.Position.Y, 0);
                                Point3d pt3 = new Point3d(blk3.Position.X, blk3.Position.Y, 0);
                                Point3d pt4 = new Point3d(blk4.Position.X, blk4.Position.Y, 0);
                                Vector3d v12 = (pt2 - pt1).GetNormal();
                                Vector3d v13 = (pt3 - pt1).GetNormal();
                                Vector3d v14 = (pt4 - pt1).GetNormal();
                                Vector3d v23 = (pt3 - pt2).GetNormal();
                                Vector3d v24 = (pt4 - pt2).GetNormal();
                                Vector3d v34 = (pt4 - pt3).GetNormal();
                                double angle1213 = (v12.GetAngleTo(v13) * 180) / Math.PI;
                                double angle2434 = (v24.GetAngleTo(v34) * 180) / Math.PI;
                                double angle2324 = (v24.GetAngleTo(v23) * 180) / Math.PI;
                                double angle1314 = (v13.GetAngleTo(v14) * 180) / Math.PI;
                                double angle1223 = (v12.GetAngleTo(v23) * 180) / Math.PI;
                                double angle1214 = (v12.GetAngleTo(v14) * 180) / Math.PI;

                                //if ((Math.Abs(angle1213 -90)<=5 && Math.Abs(angle2434-90)<=5)
                                //    ||(Math.Abs(angle1223 - 90) <= 5 && Math.Abs(angle1214 - 90) <= 5))
                                if (v12.IsPerpendicularTo(v23, ArxHelper.Tol) && v34.IsPerpendicularTo(v14, ArxHelper.Tol))
                                {
                                    Point3dCollection pts = new Point3dCollection();
                                    pts.Add(pt1);
                                    pts.Add(pt2);
                                    pts.Add(pt3);
                                    pts.Add(pt4);
                                    Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, pts, true);
                                    ArxHelper.AppendEntity(pl);
                                }
                                else if (Math.Abs(angle2324 - 90) <= 5 && Math.Abs(angle1314 - 90) <= 5)
                                {
                                    Point3dCollection pts = new Point3dCollection();
                                    pts.Add(pt1);
                                    pts.Add(pt3);
                                    pts.Add(pt2);
                                    pts.Add(pt4);
                                    Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, pts, true);
                                    ArxHelper.AppendEntity(pl);
                                }
                                else if (v12.IsParallelTo(v34, ArxHelper.Tol))
                                {
                                    Line line12 = new Line(pt1, pt2);
                                    Line line34 = new Line(pt3, pt4);
                                    if (line12.Length > line34.Length)
                                    {
                                        ArxHelper.BuildLine(line12, line34);
                                    }
                                    else
                                    {
                                        ArxHelper.BuildLine(line34, line12);
                                    }


                                }
                                else if (v13.IsParallelTo(v24, ArxHelper.Tol))
                                {
                                    Line line13 = new Line(pt1, pt3);
                                    Line line24 = new Line(pt2, pt4);
                                    if (line13.Length > line24.Length)
                                    {
                                        ArxHelper.BuildLine(line13, line24);
                                    }
                                    else
                                    {
                                        ArxHelper.BuildLine(line24, line13);
                                    }
                                }
                                else if (v14.IsParallelTo(v23, ArxHelper.Tol))
                                {
                                    Line line14 = new Line(pt1, pt4);
                                    Line line23 = new Line(pt2, pt3);
                                    if (line14.Length > line23.Length)
                                    {
                                        ArxHelper.BuildLine(line14, line23);
                                    }
                                    else
                                    {
                                        ArxHelper.BuildLine(line23, line14);
                                    }
                                }
                                else
                                {
                                    TypedValue[] filList1 = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };

                                    SelectionFilter filter1 = new SelectionFilter(filList);

                                    PromptSelectionOptions opts1 = new PromptSelectionOptions();

                                    opts1.MessageForAdding = "Select 2 point to create long line: ";
                                    PromptSelectionResult res1 = ed.GetSelection(opts1, filter1);
                                    if (res1.Status != PromptStatus.OK)
                                    {
                                        return;
                                    }
                                    if (res1.Value.Count == 2)
                                    {
                                        SelectionSet set1 = res1.Value;
                                        ObjectId[] Ids1 = set1.GetObjectIds();
                                        BlockReference blk_1 = tr.GetObject(Ids1[0], OpenMode.ForRead) as BlockReference;
                                        BlockReference blk_2 = tr.GetObject(Ids1[1], OpenMode.ForRead) as BlockReference;
                                        //Point3d pt1
                                        Point3dCollection pts = new Point3dCollection();
                                        pts.Add(pt1);
                                        pts.Add(pt2);
                                        pts.Add(pt3);
                                        pts.Add(pt4);
                                        if (ArxHelper.IsLongestLine(blk_1.Position, blk_2.Position, pts))
                                        {
                                            Line l12 = new Line(blk_1.Position, blk_2.Position);
                                            Point3d pt5 = l12.GetClosestPointTo(pts[0], false);
                                            Point3d pt6 = l12.GetClosestPointTo(pts[1], false);
                                            if (pt5.DistanceTo(pts[0]) > pt6.DistanceTo(pts[1]))
                                            {
                                                ArxHelper.BuilPolyline(blk_1.Position, blk_2.Position, pts[1], pt6, pts[0], pt5);
                                            }
                                            else
                                            {
                                                ArxHelper.BuilPolyline(blk_1.Position, blk_2.Position, pts[0], pt5, pts[1], pt6);
                                            }
                                        }
                                    }
                                }
                            }
                            tr.Commit();
                        }
                    }
                }
                else if (Ids.Length == 5)
                {
                    ObjectId Id1 = Ids[0];
                    ObjectId Id2 = Ids[1];
                    ObjectId Id3 = Ids[2];
                    ObjectId Id4 = Ids[3];
                    ObjectId Id5 = Ids[4];
                    if (Id1.IsValid && Id2.IsValid && Id3.IsValid && Id4.IsValid && Id5.IsValid)
                    {
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockReference blk1 = tr.GetObject(Id1, OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2, OpenMode.ForRead) as BlockReference;
                            BlockReference blk3 = tr.GetObject(Id3, OpenMode.ForRead) as BlockReference;
                            BlockReference blk4 = tr.GetObject(Id4, OpenMode.ForRead) as BlockReference;
                            BlockReference blk5 = tr.GetObject(Id5, OpenMode.ForRead) as BlockReference;
                            if (blk1 != null && blk2 != null && blk3 != null && blk4 != null && blk5 != null)
                            {
                                Point3d pt1 = new Point3d(blk1.Position.X, blk1.Position.Y, 0);
                                Point3d pt2 = new Point3d(blk2.Position.X, blk2.Position.Y, 0);
                                Point3d pt3 = new Point3d(blk3.Position.X, blk3.Position.Y, 0);
                                Point3d pt4 = new Point3d(blk4.Position.X, blk4.Position.Y, 0);
                                Point3d pt5 = new Point3d(blk5.Position.X, blk5.Position.Y, 0);
                                if (ArxHelper.BuildPolylineFrom3Point(pt1,pt2,pt3))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt4, pt5);
                                }
                                
                                else if (ArxHelper.BuildPolylineFrom3Point(pt1,pt2,pt4))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt3, pt5);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt1,pt2,pt5))
                                {

                                    ArxHelper.BuildPolylineFrom2Point(pt4, pt3);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt2,pt3,pt4))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt1, pt5);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt2,pt3,pt5))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt1, pt4);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt3,pt4,pt5))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt1, pt2);
                                }

                            }
                            tr.Commit();
                        }
                    }
                }
            }
        }
        [CommandMethod("TransformToXOY")]
        public static void GetAngle()
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;
            PromptSelectionResult res = ed.SelectAll();
            if (res.Status != PromptStatus.OK)
            {
                return;
            }

            if (res.Value.Count != 0)
            {
                SelectionSet set = res.Value;
                ObjectId[] Ids = set.GetObjectIds();
                //Select 2 block
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in Ids)
                {
                   
                        BlockReference blk1 = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;
                        if (blk1 != null)
                        {
                            Point3d pt = new Point3d(blk1.Position.X, blk1.Position.Y, 0);
                            blk1.TransformBy(Matrix3d.Displacement(pt - blk1.Position));
                        }
                       
                       
                    }
                    tr.Commit();
                }
               
            }
        }
    }
}
