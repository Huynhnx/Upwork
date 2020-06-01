using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadService = Autodesk.AutoCAD.ApplicationServices;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;

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
                        if (vm.TextHeight > 0)
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
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            opts.MessageForAdding = "Select block references: ";
            //Get Block from User
            PromptSelectionResult selRes = ed.GetSelection(opts, filter);
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
                                    //xline1.ColorIndex = 1;
                                    //ArxHelper.AppendEntity(xline1);
                                    //ArxHelper.AppendEntity(xline2);
                                    continue;
                                }
                                cir.Center = pts[0];
                                center = cir.Center;
                                Vector3d v1 = seg[n].StartPoint - cir.Center;
                                Vector3d v2 = seg[n].EndPoint - cir.Center;
                                double angle = v1.GetAngleTo(v2);

                                bul = Math.Tan(angle / 4);

                                if (ArxHelper.Clockwise(ptStart, corner, ptEnd))
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
        [CommandMethod("CreateBuildingLines")]
        public static void CreateBuildingLines()
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();

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
                                Point3d point1 = new Point3d(blk1.Position.X, blk1.Position.Y, 0);
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
                            ArxHelper.BuildPolylineFrom3Point(pt1, pt2, pt3);
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
                                if (ArxHelper.BuildPolylineFrom3Point(pt1, pt2, pt3))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt4, pt5);
                                }

                                else if (ArxHelper.BuildPolylineFrom3Point(pt1, pt2, pt4))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt3, pt5);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt1, pt2, pt5))
                                {

                                    ArxHelper.BuildPolylineFrom2Point(pt4, pt3);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt2, pt3, pt4))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt1, pt5);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt2, pt3, pt5))
                                {
                                    ArxHelper.BuildPolylineFrom2Point(pt1, pt4);
                                }
                                else if (ArxHelper.BuildPolylineFrom3Point(pt3, pt4, pt5))
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
        [CommandMethod("ConnectPntToLine")]
        public static void ConnectPntToLine()
        {
            //Get document, DataBase and Editor
            Document doc = Application.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };

            SelectionFilter filter = new SelectionFilter(filList);

            PromptSelectionOptions opts = new PromptSelectionOptions();

            opts.MessageForAdding = "Select block references: ";
            PromptSelectionResult res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }
            //Select Vertex
            PromptSelectionOptions opts2 = new PromptSelectionOptions();
            opts2.MessageForAdding = "Select PolyLine: ";
            PromptSelectionResult selRes2 = ed.GetSelection(opts2);
            if (selRes2.Status != PromptStatus.OK)
            {
                return;
            }
            SelectionSet set = res.Value;
            ObjectId[] BlockIds = set.GetObjectIds();
            SelectionSet set2 = selRes2.Value;
            ObjectId[] PolyLineIds = set2.GetObjectIds();
            // Case 1: Select 1 Polyline
            if (PolyLineIds.Length == 1 && BlockIds.Length > 1)
            {
                Point3dCollection ListPointBlock = new Point3dCollection();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Polyline pl = (Polyline)tr.GetObject(PolyLineIds[0], OpenMode.ForWrite);
                    foreach (ObjectId id in BlockIds)
                    {
                        if (id.IsErased == true || id.IsValid == false)
                        {
                            continue;
                        }
                        BlockReference blk = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (blk != null)
                        {
                            Point3d pt = new Point3d(blk.Position.X, blk.Position.Y, 0);
                            ListPointBlock.Add(pt);
                        }
                    }
                    Point3d ptStart = ArxHelper.GetClosetPoint(pl.StartPoint, ListPointBlock);
                    Point3d ptEnd = ArxHelper.GetClosetPoint(pl.EndPoint, ListPointBlock);
                    if (ptStart.DistanceTo(pl.StartPoint) < ptEnd.DistanceTo(pl.EndPoint))
                    {
                        Point3d pt = ptStart;
                        ptStart = ptEnd;
                        ptEnd = pt;
                    }
                    Point3dCollection verx = new Point3dCollection();

                    verx.Add(ptStart);
                    ListPointBlock.Remove(ptStart);
                    for (int i = 0; i < ListPointBlock.Count; i++)
                    {
                        Point3d closer = ArxHelper.GetClosetPoint(ptStart, ListPointBlock);
                        verx.Add(closer);
                        ptStart = closer;
                        ListPointBlock.Remove(closer);
                        i = 0;
                    }
                    verx.Add(ListPointBlock[0]);
                    Xline atStart = new Xline();
                    atStart.BasePoint = verx[0];
                    atStart.UnitDir = verx[1] - verx[2];
                    Point3dCollection intersectPts = new Point3dCollection();
                    atStart.IntersectWith(pl, Intersect.OnBothOperands, intersectPts, IntPtr.Zero, IntPtr.Zero);
                    if (intersectPts.Count > 0)
                    {
                        Point3d ptOnPl = ArxHelper.GetClosetPoint(verx[0], intersectPts);
                        verx.Insert(0, ptOnPl);
                    }
                    else
                    {
                        Xline xl = new Xline();
                        xl.BasePoint = verx[0];
                        xl.UnitDir = verx[0] - verx[1];
                        xl.IntersectWith(pl, Intersect.OnBothOperands, intersectPts, IntPtr.Zero, IntPtr.Zero);

                        if (intersectPts.Count > 0)
                        {
                            Point3d ptOnPl = ArxHelper.GetClosetPoint(verx[0], intersectPts);
                            verx.Insert(0, ptOnPl);
                        }
                        ed.WriteMessage("\nCan not Close Polyline");
                    }
                    Xline atEnd = new Xline();
                    atEnd.BasePoint = verx[verx.Count - 1];
                    atEnd.UnitDir = verx[verx.Count - 1] - verx[verx.Count - 2];
                    Point3dCollection intersectPts2 = new Point3dCollection();
                    atEnd.IntersectWith(pl, Intersect.OnBothOperands, intersectPts2, IntPtr.Zero, IntPtr.Zero);
                    if (intersectPts2.Count > 0)
                    {
                        Point3d ptOnPlEnd = ArxHelper.GetClosetPoint(verx[verx.Count - 1], intersectPts2);
                        verx.Add(ptOnPlEnd);
                    }
                    else
                    {
                        ed.WriteMessage("\nCan not Close Polyline");
                    }
                    Polyline3d pl3d = new Polyline3d(Poly3dType.SimplePoly, verx, false);
                    ArxHelper.AppendEntity(pl3d);
                    tr.Commit();
                }
            }

            if (PolyLineIds.Length == 2 && BlockIds.Length > 1)
            {
                // Case 3: Select 2 polyline And Have a block on a Polyline
                Point3dCollection ListPointBlock = new Point3dCollection();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Polyline pl1 = (Polyline)tr.GetObject(PolyLineIds[0], OpenMode.ForWrite);
                    Polyline pl2 = (Polyline)tr.GetObject(PolyLineIds[1], OpenMode.ForWrite);
                    foreach (ObjectId id in BlockIds)
                    {
                        if (id.IsErased == true || id.IsValid == false)
                        {
                            continue;
                        }
                        BlockReference blk = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (blk != null)
                        {
                            Point3d pt = new Point3d(blk.Position.X, blk.Position.Y, 0);
                            ListPointBlock.Add(pt);
                        }
                    }
                    int HaveAPointOnPl = 0;
                    Point3d ptOnPl = new Point3d();
                    //Check On Pl1
                    foreach (Point3d pt in ListPointBlock)
                    {
                        if (ArxHelper.IsPointOnPolyline(pl1, pt) == true)
                        {
                            HaveAPointOnPl = 1;
                            ptOnPl = pt;
                            break;
                        }
                        if (ArxHelper.IsPointOnPolyline(pl2, pt) == true)
                        {
                            HaveAPointOnPl = 2;
                            ptOnPl = pt;
                            break;
                        }
                    }
                    if (HaveAPointOnPl > 0)
                    {
                        Point3dCollection verx = new Point3dCollection();
                        verx.Add(ptOnPl);
                        ListPointBlock.Remove(ptOnPl);
                        for (int i = 0; i < ListPointBlock.Count; i++)
                        {
                            Point3d closer = ArxHelper.GetClosetPoint(ptOnPl, ListPointBlock);
                            verx.Add(closer);
                            ptOnPl = closer;
                            ListPointBlock.Remove(closer);
                            i = 0;
                        }
                        verx.Add(ListPointBlock[0]);
                        if (HaveAPointOnPl == 1)
                        {
                            if (ArxHelper.IsPointOnPolyline(pl2, ListPointBlock[0]) == false)
                            {
                                Line line = new Line(verx[verx.Count - 1], verx[verx.Count - 2]);
                                Point3dCollection ptonpl = new Point3dCollection();
                                line.IntersectWith(pl2, Intersect.ExtendBoth, ptonpl, IntPtr.Zero, IntPtr.Zero);
                                if (ptonpl.Count > 0)
                                {
                                    verx.Add(ArxHelper.GetClosetPoint(verx[verx.Count - 1], ptonpl));
                                }
                            }
                        }

                        Polyline3d pl3d = new Polyline3d(Poly3dType.SimplePoly, verx, false);
                        ArxHelper.AppendEntity(pl3d);
                        tr.Commit();
                    }
                    else
                    {
                        PromptPointOptions op = new PromptPointOptions("Pick Point On Polyline:");
                        PromptPointResult r = ed.GetPoint(op);
                        if (r.Status == PromptStatus.OK)
                        {
                            Point3d ptOnFirstPolyline = r.Value;
                            Point3dCollection verx = new Point3dCollection();
                            Point2dCollection pt2d = new Point2dCollection();

                            verx.Add(ptOnFirstPolyline);
                            pt2d.Add(new Point2d(ptOnFirstPolyline.X, ptOnFirstPolyline.Y));
                            for (int i = 0; i < ListPointBlock.Count; i++)
                            {
                                Point3d closer = ArxHelper.GetClosetPoint(ptOnFirstPolyline, ListPointBlock);
                                pt2d.Add(new Point2d(closer.X, closer.Y));
                                verx.Add(closer);
                                ptOnFirstPolyline = closer;
                                ListPointBlock.Remove(closer);
                                i = 0;
                            }
                            verx.Add(ListPointBlock[0]);
                            pt2d.Add(new Point2d(ListPointBlock[0].X, ListPointBlock[0].Y));
                            //Init Keywords/Option
                            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
                            pKeyOpts.Message = "\nDo you want to create arc at last 3 point? ";
                            pKeyOpts.Keywords.Add("Yes");
                            pKeyOpts.Keywords.Add("No");
                            pKeyOpts.Keywords.Default = "No";
                            pKeyOpts.AllowNone = true;
                            //Get KeyWord
                            PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);
                            //If Get Keyword seccess
                            if (pKeyRes.Status == PromptStatus.OK)
                            {
                                if (pKeyRes.StringResult.Contains("Y"))
                                {
                                    Polyline pl = new Polyline();


                                    Point3d ptStart = verx[verx.Count - 3];
                                    Point3d ptEnd = verx[verx.Count - 1];

                                    Line line = new Line(ptStart, ptEnd);
                                    Point3d center = new Point3d();
                                    Point3d corner = verx[verx.Count - 2];
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
                                        return;
                                    }
                                    cir.Center = pts[0];
                                    center = cir.Center;
                                    Vector3d v1 = ptStart - cir.Center;
                                    Vector3d v2 = corner - cir.Center;
                                    double angle = v1.GetAngleTo(v2);

                                    double bul = Math.Tan(angle / 4);

                                    if (ArxHelper.Clockwise(ptStart, corner, ptEnd))
                                    {
                                        bul = -1 * bul;
                                    }
                                    else
                                    {

                                    }
                                    Point3dCollection listpt1 = new Point3dCollection();
                                    Point3dCollection listpt2 = new Point3dCollection();
                                    cir.IntersectWith(pl1, Intersect.OnBothOperands, listpt1, IntPtr.Zero, IntPtr.Zero);
                                    cir.IntersectWith(pl2, Intersect.OnBothOperands, listpt2, IntPtr.Zero, IntPtr.Zero);
                                    Point3d pt1 = new Point3d(), pt2 = new Point3d();
                                    if (listpt1.Count > 0)
                                    {
                                        pt1 = ArxHelper.GetClosetPoint(ptEnd, listpt1);
                                    }
                                    if (listpt2.Count > 0)
                                    {
                                        pt2 = ArxHelper.GetClosetPoint(ptEnd, listpt2);
                                    }
                                    if (listpt1.Count > 0 && listpt2.Count > 0)
                                    {
                                        if (ptEnd.DistanceTo(pt1) > ptEnd.DistanceTo(pt2))
                                        {
                                            pt2d.Add(new Point2d(pt2.X, pt2.Y));
                                            verx.Add(pt2);
                                        }
                                        else
                                        {
                                            pt2d.Add(new Point2d(pt1.X, pt1.Y));
                                            verx.Add(pt1);
                                        }
                                    }
                                    else if (listpt1.Count > 0 && listpt2.Count == 0)
                                    {
                                        pt2d.Add(new Point2d(pt1.X, pt1.Y));
                                        verx.Add(pt1);
                                    }
                                    else if (listpt2.Count > 0 && listpt1.Count == 0)
                                    {
                                        pt2d.Add(new Point2d(pt2.X, pt2.Y));
                                        verx.Add(pt2);
                                    }
                                    for (int i = 0; i < pt2d.Count; i++)
                                    {
                                        if (i == pt2d.Count - 4)
                                        {
                                            pl.AddVertexAt(i, pt2d[i], bul, 0, 0);
                                        }
                                        else if (i >= pt2d.Count - 3)
                                        {
                                            double bulge = ArxHelper.CalculateBulge(verx[pt2d.Count - 3], verx[pt2d.Count - 2], verx[pt2d.Count - 1]);
                                            pl.AddVertexAt(i, pt2d[i], bulge, 0, 0);
                                        }
                                        else
                                        {
                                            pl.AddVertexAt(i, pt2d[i], 0, 0, 0);
                                        }
                                    }
                                    ArxHelper.AppendEntity(pl);
                                }
                                else
                                {

                                    Xline xl = new Xline();
                                    xl.BasePoint = verx[verx.Count - 1];
                                    xl.UnitDir = verx[verx.Count - 1] - verx[verx.Count - 2];
                                    Point3dCollection intersect = new Point3dCollection();
                                    xl.IntersectWith(pl2, Intersect.OnBothOperands, intersect, IntPtr.Zero, IntPtr.Zero);
                                    Point3d point1 = new Point3d(), point2 = new Point3d();
                                    if (intersect.Count > 0)
                                    {
                                        point1 = ArxHelper.GetClosetPoint(verx[verx.Count - 1], intersect);
                                    }
                                    Point3dCollection intersect2 = new Point3dCollection();
                                    xl.IntersectWith(pl1, Intersect.OnBothOperands, intersect2, IntPtr.Zero, IntPtr.Zero);
                                    if (intersect2.Count > 0)
                                    {
                                        point2 = ArxHelper.GetClosetPoint(verx[verx.Count - 1], intersect2);
                                    }
                                    if (intersect.Count > 0 && intersect2.Count > 0)
                                    {
                                        if (point1.DistanceTo(verx[verx.Count - 1]) < point2.DistanceTo(verx[verx.Count - 1]))
                                        {
                                            verx.Add(point1);
                                        }
                                        else
                                        {
                                            verx.Add(point2);
                                        }
                                    }
                                    else if (intersect.Count > 0)
                                    {
                                        verx.Add(point1);
                                    }
                                    else if (intersect2.Count > 0)
                                    {
                                        verx.Add(point2);
                                    }

                                    Polyline3d pl3d = new Polyline3d(Poly3dType.SimplePoly, verx, false);
                                    ArxHelper.AppendEntity(pl3d);
                                }
                            }

                            tr.Commit();
                        }
                    }

                }
            }
            if (PolyLineIds.Length == 1 && BlockIds.Length == 1)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Polyline pl = (Polyline)tr.GetObject(PolyLineIds[0], OpenMode.ForWrite);
                    BlockReference blk = tr.GetObject(BlockIds[0], OpenMode.ForRead) as BlockReference;
                    Point3d blkptn = new Point3d(blk.Position.X, blk.Position.Y, 0);
                    Point3dCollection listpt = new Point3dCollection();
                    for (int i = 0; i < pl.NumberOfVertices - 2; i = i + 1)
                    {
                        LineSegment3d seg = pl.GetLineSegmentAt(i);
                        PointOnCurve3d pt = seg.GetClosestPointTo(blkptn);
                        listpt.Add(pt.Point);
                    }
                    Point3d pt1 = ArxHelper.GetClosetPoint(blkptn, listpt);
                    Xline xl = new Xline();
                    xl.BasePoint = blkptn;
                    xl.UnitDir = (pt1 - blkptn).GetNormal().RotateBy(Math.PI / 2, Vector3d.ZAxis);
                    Point3dCollection intersect = new Point3dCollection();
                    xl.IntersectWith(pl, Intersect.OnBothOperands, intersect, IntPtr.Zero, IntPtr.Zero);
                    Point3d pt2 = new Point3d();
                    if (intersect.Count > 0)
                    {
                        pt2 = ArxHelper.GetClosetPoint(blkptn, intersect);
                    }
                    Point3dCollection pts = new Point3dCollection();
                    pts.Add(pt1);
                    pts.Add(blkptn);
                    pts.Add(pt2);
                    Polyline3d pl3d = new Polyline3d(Poly3dType.SimplePoly, pts, false);
                    ArxHelper.AppendEntity(pl3d);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("Convert3dPolyToPoly")]
        public static void Convert3dPolyToPoly()
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
                Polyline3d pl = (Polyline3d)tr.GetObject(singleId[0], OpenMode.ForWrite);
                DBObjectCollection segments = new DBObjectCollection();
                pl.Explode(segments);
                Point2dCollection pts = new Point2dCollection();
                foreach (Entity ent in segments)
                {
                    if (ent is Line)
                    {
                        Line line = ent as Line;
                        Point2d ptstart = new Point2d(line.StartPoint.X, line.StartPoint.Y);
                        Point2d ptend = new Point2d(line.EndPoint.X, line.EndPoint.Y);
                        if (pts.Contains(ptstart) == false)
                        {
                            pts.Add(ptstart);
                        }
                        if (pts.Contains(ptend) == false)
                        {
                            pts.Add(ptend);
                        }

                    }
                }
                Polyline poly2d = new Polyline();
                for (int i = 0; i < pts.Count; i++)
                {
                    Point2d pt = pts[i];
                    poly2d.AddVertexAt(i, pt, 0, 0, 0);
                }
                ArxHelper.AppendEntity(poly2d);
                ArxHelper.DeleteEntity(pl.ObjectId);
                tr.Commit();
            }
        }

        [CommandMethod("CreateStairs")]

        public static void CreateStairs()

        {
            double ElevationTol = 0.03;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            opts.MessageForAdding = "Select block references to Create Stair: ";
            PromptSelectionResult res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }
            if (res.Value.Count != 0)
            {
                SelectionSet set = res.Value;
                ObjectId[] Ids = set.GetObjectIds();
                double stairWidth = 0.2;
                PromptDoubleOptions dOpt = new PromptDoubleOptions("Stair Width:");
                dOpt.AllowNone = true;
                dOpt.DefaultValue = stairWidth;
                PromptDoubleResult dr = ed.GetDouble(dOpt);
                if (dr.Status == PromptStatus.OK)
                {
                    stairWidth = dr.Value;
                }
                if (Ids.Length == 4)
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
                                // Get Point
                                Point3d pt1 = blk1.Position;
                                Point3d pt2 = blk2.Position;
                                Point3d pt3 = blk3.Position;
                                Point3d pt4 = blk4.Position;
                                Vector3d v12 = (pt2 - pt1).GetNormal();
                                Vector3d v13 = (pt3 - pt1).GetNormal();
                                Vector3d v14 = (pt4 - pt1).GetNormal();
                                Vector3d v23 = (pt3 - pt2).GetNormal();
                                Vector3d v24 = (pt4 - pt2).GetNormal();
                                Vector3d v34 = (pt4 - pt3).GetNormal();

                                // Check if is a rectang
                                if (v12.IsPerpendicularTo(v13, ArxHelper.Tol) && v34.IsPerpendicularTo(v14, ArxHelper.Tol))
                                {
                                    if (Math.Abs(pt1.Z - pt2.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt2, pt4, pt3, stairWidth, blk1.Layer);
                                    }
                                    else if (Math.Abs(pt1.Z - pt3.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt3, pt4, pt2, stairWidth, blk1.Layer);
                                    }

                                }
                                else if (v12.IsPerpendicularTo(v14, ArxHelper.Tol) && v34.IsPerpendicularTo(v23, ArxHelper.Tol))
                                {
                                    if (Math.Abs(pt1.Z - pt2.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt2, pt3, pt4, stairWidth, blk1.Layer);
                                    }
                                    else if (Math.Abs(pt1.Z - pt4.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt4, pt3, pt2, stairWidth, blk1.Layer);
                                    }

                                }
                                else if (v13.IsPerpendicularTo(v14, ArxHelper.Tol) && v23.IsPerpendicularTo(v24, ArxHelper.Tol))
                                {
                                    if (Math.Abs(pt1.Z - pt3.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt3, pt2, pt4, stairWidth, blk1.Layer);
                                    }
                                    else if (Math.Abs(pt1.Z - pt4.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt4, pt2, pt3, stairWidth, blk1.Layer);
                                    }

                                }
                                else if (v12.IsPerpendicularTo(v13, ArxHelper.Tol) && v34.IsPerpendicularTo(v24, ArxHelper.Tol))
                                {
                                    if (Math.Abs(pt1.Z - pt2.Z) < ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt2, pt4, pt3, stairWidth, blk1.Layer);
                                    }
                                    else if (Math.Abs(pt1.Z - pt3.Z) < 0.001)
                                    {
                                        ArxHelper.CreateStairFrom4Point(pt1, pt3, pt4, pt2, stairWidth, blk1.Layer);
                                    }
                                }
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
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockReference blk1 = tr.GetObject(Id1, OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2, OpenMode.ForRead) as BlockReference;
                            BlockReference blk3 = tr.GetObject(Id3, OpenMode.ForRead) as BlockReference;
                            if (blk1 != null && blk2 != null && blk3 != null)
                            {
                                // Get Point
                                Point3d pt1 = blk1.Position;
                                Point3d pt2 = blk2.Position;
                                Point3d pt3 = blk3.Position;
                                Vector3d v12 = (pt2 - pt1).GetNormal();
                                Vector3d v13 = (pt3 - pt1).GetNormal();
                                Vector3d v23 = (pt3 - pt2).GetNormal();
                                if (v12.IsPerpendicularTo(v13, ArxHelper.Tol))
                                {
                                    Point3dCollection verx = new Point3dCollection();

                                    verx.Add(pt3);
                                    verx.Add(pt1);
                                    verx.Add(pt2);
                                    verx.Add(pt3 + (pt2 - pt1));
                                    //Find top point and Bottom Point
                                    Point3dCollection listPoint1 = new Point3dCollection();
                                    Point3dCollection listPoint2 = new Point3dCollection();
                                    listPoint1.Add(blk1.Position);
                                    if (Math.Abs(blk2.Position.Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(blk2.Position);
                                    }
                                    else
                                    {
                                        listPoint2.Add(blk2.Position);
                                    }
                                    if (Math.Abs(blk3.Position.Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(blk3.Position);
                                    }
                                    else
                                    {
                                        listPoint2.Add(blk3.Position);
                                    }
                                    if (Math.Abs(verx[3].Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(verx[3]);
                                    }
                                    else
                                    {
                                        listPoint2.Add(verx[3]);
                                    }
                                    Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, true);
                                    pl.Layer = blk1.Layer;
                                    ArxHelper.AppendEntity(pl);
                                    Point3d mid1 = new Point3d((listPoint1[0].X + listPoint1[1].X) / 2, (listPoint1[0].Y + listPoint1[1].Y) / 2, (listPoint1[0].Z + listPoint1[1].Z) / 2);
                                    Point3d mid2 = new Point3d((listPoint2[0].X + listPoint2[1].X) / 2, (listPoint2[0].Y + listPoint2[1].Y) / 2, (listPoint2[0].Z + listPoint2[1].Z) / 2);
                                    if (listPoint1[0].Z < listPoint2[0].Z)
                                    {
                                        ArxHelper.drawArrow(mid1, mid2, listPoint1[0] - listPoint1[1], listPoint1[0].DistanceTo(listPoint1[1]), stairWidth);
                                        for (int i = 1; i < mid1.DistanceTo(mid2) / stairWidth; i++)
                                        {
                                            Point3d p1 = listPoint1[0] + (mid2 - mid1).GetNormal() * i * stairWidth;
                                            Point3d p2 = listPoint1[1] + (mid2 - mid1).GetNormal() * i * stairWidth;
                                            Line line = new Line(p1, p2);
                                            line.Layer = blk1.Layer;
                                            ArxHelper.AppendEntity(line);
                                        }
                                    }
                                    else
                                    {
                                        ArxHelper.drawArrow(mid2, mid1, listPoint1[0] - listPoint1[1], listPoint1[0].DistanceTo(listPoint1[1]), stairWidth);
                                        for (int i = 1; i < mid1.DistanceTo(mid2) / stairWidth; i++)
                                        {
                                            Point3d p1 = listPoint2[0] + (mid1 - mid2).GetNormal() * i * stairWidth;
                                            Point3d p2 = listPoint2[1] + (mid1 - mid2).GetNormal() * i * stairWidth;
                                            Line line = new Line(p1, p2);
                                            line.Layer = blk1.Layer;
                                            ArxHelper.AppendEntity(line);
                                        }
                                    }


                                }
                                else if (v13.IsPerpendicularTo(v23, ArxHelper.Tol))
                                {
                                    Point3dCollection verx = new Point3dCollection();

                                    verx.Add(pt1);
                                    verx.Add(pt3);
                                    verx.Add(pt2);
                                    verx.Add(pt1 + (pt2 - pt3));
                                    Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, true);
                                    pl.Layer = blk1.Layer;
                                    ArxHelper.AppendEntity(pl);
                                    //Find top point and Bottom Point
                                    Point3dCollection listPoint1 = new Point3dCollection();
                                    Point3dCollection listPoint2 = new Point3dCollection();
                                    listPoint1.Add(blk1.Position);
                                    if (Math.Abs(blk2.Position.Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(blk2.Position);
                                    }
                                    else
                                    {
                                        listPoint2.Add(blk2.Position);
                                    }
                                    if (Math.Abs(blk3.Position.Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(blk3.Position);
                                    }
                                    else
                                    {
                                        listPoint2.Add(blk3.Position);
                                    }
                                    if (Math.Abs(verx[3].Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(verx[3]);
                                    }
                                    else
                                    {
                                        listPoint2.Add(verx[3]);
                                    }
                                    Point3d mid1 = new Point3d((listPoint1[0].X + listPoint1[1].X) / 2, (listPoint1[0].Y + listPoint1[1].Y) / 2, (listPoint1[0].Z + listPoint1[1].Z) / 2);
                                    Point3d mid2 = new Point3d((listPoint2[0].X + listPoint2[1].X) / 2, (listPoint2[0].Y + listPoint2[1].Y) / 2, (listPoint2[0].Z + listPoint2[1].Z) / 2);
                                    if (listPoint1[0].Z < listPoint2[0].Z)
                                    {
                                        ArxHelper.drawArrow(mid1, mid2, listPoint1[0] - listPoint1[1], listPoint1[0].DistanceTo(listPoint1[1]), stairWidth);
                                        for (int i = 1; i < mid1.DistanceTo(mid2) / stairWidth; i++)
                                        {
                                            Point3d p1 = listPoint1[0] + (mid2 - mid1).GetNormal() * i * stairWidth;
                                            Point3d p2 = listPoint1[1] + (mid2 - mid1).GetNormal() * i * stairWidth;
                                            Line line = new Line(p1, p2);
                                            line.Layer = blk1.Layer;
                                            ArxHelper.AppendEntity(line);
                                        }
                                    }
                                    else
                                    {
                                        ArxHelper.drawArrow(mid2, mid1, listPoint1[0] - listPoint1[1], listPoint1[0].DistanceTo(listPoint1[1]), stairWidth);
                                        for (int i = 1; i < mid1.DistanceTo(mid2) / stairWidth; i++)
                                        {
                                            Point3d p1 = listPoint2[0] + (mid1 - mid2).GetNormal() * i * stairWidth;
                                            Point3d p2 = listPoint2[1] + (mid1 - mid2).GetNormal() * i * stairWidth;
                                            Line line = new Line(p1, p2);
                                            line.Layer = blk1.Layer;
                                            ArxHelper.AppendEntity(line);
                                        }
                                    }

                                }
                                else if (v12.IsPerpendicularTo(v23, ArxHelper.Tol))
                                {
                                    Point3dCollection verx = new Point3dCollection();

                                    verx.Add(pt3);
                                    verx.Add(pt2);
                                    verx.Add(pt1);
                                    verx.Add(pt3 + (pt1 - pt2));
                                    Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, true);
                                    pl.Layer = blk1.Layer;
                                    ArxHelper.AppendEntity(pl);
                                    //Find top point and Bottom Point
                                    Point3dCollection listPoint1 = new Point3dCollection();
                                    Point3dCollection listPoint2 = new Point3dCollection();
                                    listPoint1.Add(blk1.Position);
                                    if (Math.Abs(blk2.Position.Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(blk2.Position);
                                    }
                                    else
                                    {
                                        listPoint2.Add(blk2.Position);
                                    }
                                    if (Math.Abs(blk3.Position.Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(blk3.Position);
                                    }
                                    else
                                    {
                                        listPoint2.Add(blk3.Position);
                                    }
                                    if (Math.Abs(verx[3].Z - blk1.Position.Z) < ElevationTol)
                                    {
                                        listPoint1.Add(verx[3]);
                                    }
                                    else
                                    {
                                        listPoint2.Add(verx[3]);
                                    }
                                    Point3d mid1 = new Point3d((listPoint1[0].X + listPoint1[1].X) / 2, (listPoint1[0].Y + listPoint1[1].Y) / 2, (listPoint1[0].Z + listPoint1[1].Z) / 2);
                                    Point3d mid2 = new Point3d((listPoint2[0].X + listPoint2[1].X) / 2, (listPoint2[0].Y + listPoint2[1].Y) / 2, (listPoint2[0].Z + listPoint2[1].Z) / 2);
                                    if (listPoint1[0].Z < listPoint2[0].Z)
                                    {
                                        ArxHelper.drawArrow(mid1, mid2, listPoint1[0] - listPoint1[1], listPoint1[0].DistanceTo(listPoint1[1]), stairWidth);
                                        for (int i = 1; i < mid1.DistanceTo(mid2) / stairWidth; i++)
                                        {
                                            Point3d p1 = listPoint1[0] + (mid2 - mid1).GetNormal() * i * stairWidth;
                                            Point3d p2 = listPoint1[1] + (mid2 - mid1).GetNormal() * i * stairWidth;
                                            Line line = new Line(p1, p2);
                                            line.Layer = blk1.Layer;
                                            ArxHelper.AppendEntity(line);
                                        }
                                    }
                                    else
                                    {
                                        ArxHelper.drawArrow(mid2, mid1, listPoint1[0] - listPoint1[1], listPoint1[0].DistanceTo(listPoint1[1]), stairWidth);
                                        for (int i = 1; i < mid1.DistanceTo(mid2) / stairWidth; i++)
                                        {
                                            Point3d p1 = listPoint2[0] + (mid1 - mid2).GetNormal() * i * stairWidth;
                                            Point3d p2 = listPoint2[1] + (mid1 - mid2).GetNormal() * i * stairWidth;
                                            Line line = new Line(p1, p2);
                                            line.Layer = blk1.Layer;
                                            ArxHelper.AppendEntity(line);
                                        }
                                    }
                                }
                                else
                                {
                                    if (Math.Abs(pt1.Z - pt2.Z) <= ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom3Point(pt1, pt2, pt3, stairWidth, blk1.Layer);
                                    }
                                    else if (Math.Abs(pt1.Z - pt3.Z) <= ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom3Point(pt1, pt3, pt2, stairWidth, blk1.Layer);
                                    }
                                    else if (Math.Abs(pt3.Z - pt2.Z) <= ElevationTol)
                                    {
                                        ArxHelper.CreateStairFrom3Point(pt3, pt2, pt1, stairWidth, blk1.Layer);
                                    }
                                }
                            }
                            tr.Commit();
                        }
                    }
                }
                else if (Ids.Length == 6)
                {
                    ObjectId Id1 = Ids[0];
                    ObjectId Id2 = Ids[1];
                    ObjectId Id3 = Ids[2];
                    ObjectId Id4 = Ids[3];
                    ObjectId Id5 = Ids[4];
                    ObjectId Id6 = Ids[5];
                    if (Id1.IsValid && Id2.IsValid && Id3.IsValid && Id4.IsValid && Id5.IsValid && Id6.IsValid)
                    {
                        //Start Transaction
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockReference blk1 = tr.GetObject(Id1, OpenMode.ForRead) as BlockReference;
                            BlockReference blk2 = tr.GetObject(Id2, OpenMode.ForRead) as BlockReference;
                            BlockReference blk3 = tr.GetObject(Id3, OpenMode.ForRead) as BlockReference;
                            BlockReference blk4 = tr.GetObject(Id4, OpenMode.ForRead) as BlockReference;
                            BlockReference blk5 = tr.GetObject(Id5, OpenMode.ForRead) as BlockReference;
                            BlockReference blk6 = tr.GetObject(Id6, OpenMode.ForRead) as BlockReference;
                            if (blk1 != null && blk2 != null && blk3 != null && blk4 != null && blk5 != null && blk6 != null)
                            {
                                //Find top point and Bottom Point
                                Point3dCollection listPoint1 = new Point3dCollection();
                                Point3dCollection listPoint2 = new Point3dCollection();
                                listPoint1.Add(blk1.Position);
                                if (Math.Abs(blk2.Position.Z - blk1.Position.Z) <= ElevationTol)
                                {
                                    listPoint1.Add(blk2.Position);
                                }
                                else
                                {
                                    listPoint2.Add(blk2.Position);
                                }
                                if (Math.Abs(blk3.Position.Z - blk1.Position.Z) <= ElevationTol)
                                {
                                    listPoint1.Add(blk3.Position);
                                }
                                else
                                {
                                    listPoint2.Add(blk3.Position);
                                }
                                if (Math.Abs(blk4.Position.Z - blk1.Position.Z) <= ElevationTol)
                                {
                                    listPoint1.Add(blk4.Position);
                                }
                                else
                                {
                                    listPoint2.Add(blk4.Position);
                                }
                                if (Math.Abs(blk5.Position.Z - blk1.Position.Z) <= ElevationTol)
                                {
                                    listPoint1.Add(blk5.Position);
                                }
                                else
                                {
                                    listPoint2.Add(blk5.Position);
                                }
                                if (Math.Abs(blk6.Position.Z - blk1.Position.Z) <= ElevationTol)
                                {
                                    listPoint1.Add(blk6.Position);
                                }
                                else
                                {
                                    listPoint2.Add(blk6.Position);
                                }
                                if (listPoint1.Count != listPoint2.Count)
                                {
                                    tr.Abort();
                                    return;
                                }
                                // Find Bigest Triangle Corner
                                Point3d pt1 = ArxHelper.FindBigestCorner(listPoint1[0], listPoint1[1], listPoint1[2]);
                                Point3d pt5 = ArxHelper.FindBigestCorner(listPoint2[0], listPoint2[1], listPoint2[2]);
                                listPoint1.Remove(pt1);
                                listPoint2.Remove(pt5);
                                // Create Bounding PolyLine
                                Point3dCollection verx = new Point3dCollection();
                                verx.Add(listPoint1[0]);
                                verx.Add(pt1);
                                verx.Add(listPoint1[1]);
                                if (listPoint1[1].DistanceTo(listPoint2[0]) > listPoint1[1].DistanceTo(listPoint2[1]))
                                {
                                    verx.Add(listPoint2[1]);
                                    listPoint2.RemoveAt(1);
                                }
                                else
                                {
                                    verx.Add(listPoint2[0]);
                                    listPoint2.RemoveAt(0);
                                }
                                verx.Add(pt5);
                                verx.Add(listPoint2[0]);
                                Polyline3d pl = new Polyline3d(Poly3dType.SimplePoly, verx, true);
                                pl.Layer = blk1.Layer;
                                ArxHelper.AppendEntity(pl);
                                Line line = new Line(pt1, pt5);
                                line.Layer = blk1.Layer;
                                ArxHelper.AppendEntity(line);
                                int count = (int)(pt1.DistanceTo(pt5) / stairWidth);
                                //Create Stair
                                for (int i = 1; i < count; i++)
                                {
                                    Point3dCollection pts = new Point3dCollection();
                                    pts.Add(verx[0] + (verx[5] - verx[0]).GetNormal() * i * (verx[0].DistanceTo(verx[5]) / count));
                                    pts.Add(pt1 + (pt5 - pt1).GetNormal() * i * stairWidth);
                                    pts.Add(verx[2] + (verx[3] - verx[2]).GetNormal() * i * (verx[2].DistanceTo(verx[3]) / count));
                                    Polyline3d pStep = new Polyline3d(Poly3dType.SimplePoly, pts, false);
                                    pStep.Layer = blk1.Layer;
                                    ArxHelper.AppendEntity(pStep);
                                }
                                Point3d mid1 = new Point3d((verx[0].X + pt1.X) / 2, (verx[0].Y + pt1.Y) / 2, (verx[0].Z + pt1.Z) / 2);
                                Point3d mid2 = new Point3d((pt5.X + verx[5].X) / 2, (pt5.Y + verx[5].Y) / 2, (pt5.Z + verx[5].Z) / 2);
                                if (mid1.Z < mid2.Z)
                                {
                                    ArxHelper.drawArrow(mid1, mid2, verx[0] - pt1, pt1.DistanceTo(verx[0]), stairWidth);
                                }
                                else
                                {
                                    ArxHelper.drawArrow(mid2, mid1, verx[0] - pt1, pt1.DistanceTo(verx[0]), stairWidth);
                                }
                                Point3d mid3 = new Point3d((verx[2].X + pt1.X) / 2, (verx[2].Y + pt1.Y) / 2, (verx[2].Z + pt1.Z) / 2);
                                Point3d mid4 = new Point3d((pt5.X + verx[3].X) / 2, (pt5.Y + verx[3].Y) / 2, (pt5.Z + verx[3].Z) / 2);
                                if (mid3.Z < mid4.Z)
                                {
                                    ArxHelper.drawArrow(mid3, mid4, verx[2] - pt1, pt1.DistanceTo(verx[2]), stairWidth);
                                }
                                else
                                {
                                    ArxHelper.drawArrow(mid4, mid3, verx[2] - pt1, pt1.DistanceTo(verx[2]), stairWidth);
                                }

                            }
                            tr.Commit();
                        }
                    }
                }
            }
        }
        [CommandMethod("projectFrontOfBuilding")]
        public static void projectFrontOfBuilding()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            Point3d point1 = new Point3d(), point2 = new Point3d(), point3 = new Point3d();
            // Create UCS
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the UCS table for read
                UcsTable acUCSTbl;
                acUCSTbl = acTrans.GetObject(acCurDb.UcsTableId,
                                                   OpenMode.ForRead) as UcsTable;
                UcsTableRecord acUCSTblRec;
                // Check to see if the "New_UCS" UCS table record exists
                if (acUCSTbl.Has("New_UCS") == false)
                {
                    acUCSTblRec = new UcsTableRecord();
                    acUCSTblRec.Name = "New_UCS";
                    // Open the UCSTable for write
                    acUCSTbl.UpgradeOpen();
                    // Add the new UCS table record
                    acUCSTbl.Add(acUCSTblRec);
                    acTrans.AddNewlyCreatedDBObject(acUCSTblRec, true);
                }
                else
                {
                    acUCSTblRec = acTrans.GetObject(acUCSTbl["New_UCS"],
                                                    OpenMode.ForWrite) as UcsTableRecord;
                }
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("");
                // Prompt for a point
                pPtOpts.Message = "\nOrigin Point: ";
                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                Point3d pPtOrigin;
                // If a point was entered, then translate it to the current UCS
                if (pPtRes.Status == PromptStatus.OK)
                {
                    pPtOrigin = pPtRes.Value;
                    acUCSTblRec.Origin = pPtOrigin;
                    PromptPointResult ptXdir;
                    PromptPointOptions ptXdirOpt = new PromptPointOptions("");
                    // Prompt for a point
                    ptXdirOpt.Message = "\nX direction: ";
                    ptXdirOpt.UseBasePoint = true;
                    ptXdirOpt.BasePoint = pPtOrigin;
                    ptXdir = acDoc.Editor.GetPoint(ptXdirOpt);
                    if (ptXdir.Status == PromptStatus.OK)
                    {
                        acUCSTblRec.XAxis = (ptXdir.Value - pPtOrigin).GetNormal();
                        PromptPointResult ptYdir;
                        PromptPointOptions ptYdirOpt = new PromptPointOptions("");
                        // Prompt for a point
                        ptYdirOpt.Message = "\nY direction: ";
                        ptYdirOpt.UseBasePoint = true;
                        ptYdirOpt.BasePoint = pPtOrigin;
                        while (true)
                        {
                            ptYdir = acDoc.Editor.GetPoint(ptYdirOpt);
                            if (ptYdir.Status == PromptStatus.OK)
                            {
                                if ((ptYdir.Value - pPtOrigin).IsPerpendicularTo(acUCSTblRec.XAxis))
                                {
                                    acUCSTblRec.YAxis = (ptYdir.Value - pPtOrigin).GetNormal();
                                    break;
                                }
                                ed.WriteMessage("\nY vector is not perpencular to X vector");
                            }
                            else
                            {
                                return;
                            }
                        }
                        point1 = pPtOrigin;
                        point2 = ptXdir.Value;
                        point3 = ptYdir.Value;
                        // Open the active viewport
                        ViewportTableRecord acVportTblRec;
                        acVportTblRec = acTrans.GetObject(acDoc.Editor.ActiveViewportId,
                                                                OpenMode.ForWrite) as ViewportTableRecord;
                        // Display the UCS Icon at the origin of the current viewport
                        acVportTblRec.IconAtOrigin = true;
                        acVportTblRec.IconEnabled = true;
                        // Set the UCS current
                        acVportTblRec.SetUcs(acUCSTblRec.ObjectId);
                        acDoc.Editor.UpdateTiledViewportsFromDatabase();

                    }
                }
                // Save the new objects to the database
                acTrans.Commit();
            }


            // Select Block to Project 
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();

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
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    for (int i = 0; i < Ids.Length; i++)
                    {
                        ObjectId blkId = Ids[i];
                        if (blkId.IsValid)
                        {
                            BlockReference blk = acTrans.GetObject(blkId, OpenMode.ForWrite) as BlockReference;
                            if (blk != null)
                            {
                                //Point3d pt = blk.Position.TransformBy(ed.CurrentUserCoordinateSystem);
                                Point3d pt = blk.Position.TransformBy(ed.CurrentUserCoordinateSystem.Inverse());
                                Plane pl = new Plane(ed.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin, ed.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis);

                                Point3d pointOnPlan = new Point3d(pt.X, pt.Y, 0);
                                pointOnPlan = pointOnPlan.TransformBy(ed.CurrentUserCoordinateSystem);

                                Matrix3d movementMat = Matrix3d.Displacement(pointOnPlan - blk.Position);


                                blk.TransformBy(movementMat);
                            }
                        }
                    }

                    acTrans.Commit();
                }

            }

        }
        [CommandMethod("ConnectLineByExcelOrder")]
        public static void ConnectLineByExcelOrder()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            // Read Excel File
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.RestoreDirectory = false;
            openFileDialog.CheckFileExists = true;
            if (openFileDialog.ShowDialog() == true)
            {
                string outputFileName = openFileDialog.FileName;
                // Check file exist or not
                if (ExcelHelper.IsFileLocked(outputFileName))
                {
                    MessageBox.Show("This File is Lock", "Warnning");
                    return;
                }
                try
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(outputFileName, false))
                    {
                        // Get the SharedStringTablePart. If it does not exist, create a new one.
                        WorkbookPart workbookPart = document.WorkbookPart;
                        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                        string relationshipId = sheets.First().Id.Value;
                        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                        Worksheet workSheet = worksheetPart.Worksheet;
                        SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                        // Get rows data from the sheet
                        IEnumerable<Row> rows = sheetData.Descendants<Row>();
                        // We only accept excel file that has 2 columns (layer name, color and transparency)
                        if (rows.ElementAt(0).Descendants<Cell>().Count() != 4)
                        {
                            MessageBox.Show("File Format is wrong. File should have 3 column", "Warning");
                            return;
                        }

                        foreach (Row row in rows)
                        {
                            // Skip header row
                            if (row == rows.ElementAt(0))
                                continue;
                            if (ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(0)) == string.Empty)
                            {
                                continue;
                            }

                            // Get cell value and paste into new Layer Item
                            string Point = ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(0));
                            // Skip invalid name
                            if (Point == string.Empty)
                            {
                                continue;
                            }
                            string Point1 = ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(1));
                            string Point2 = ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(2));
                            string Point3 = ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(3));

                            if (Point != string.Empty && Point1 != string.Empty)
                            {
                                Point3d p1 = ExcelHelper.ToPoint3d(Point);
                                Point3d p2 = ExcelHelper.ToPoint3d(Point1);

                                Line l = new Line(p1, p2);
                                l.ColorIndex = 1;
                                ArxHelper.AppendEntity(l);
                            }
                            if (Point != string.Empty && Point1 == string.Empty && Point2 != string.Empty && Point3 != string.Empty)
                            {
                                Point3d p1 = ExcelHelper.ToPoint3d(Point);
                                Point3d p2 = ExcelHelper.ToPoint3d(Point2);
                                Point3d p3 = ExcelHelper.ToPoint3d(Point3);

                                Xline xl = new Xline();
                                xl.BasePoint = p2;
                                xl.UnitDir = p3 - p2;
                                Point3d pExtent = xl.GetClosestPointTo(p1, false);
                                Line l = new Line(p1, pExtent);
                                Line l23 = new Line(p2, p3);
                                Line ex = new Line(pExtent, p3);
                                ArxHelper.AppendEntity(l);
                                ArxHelper.AppendEntity(l23);
                                ArxHelper.AppendEntity(ex);

                            }

                        }

                    }
                }
                // Using OpenXml to read the excel file
                catch (System.Exception exc)
                {

                    return;
                }

            }
        }
        [CommandMethod("PrintPossition")]
        public static void PrintPossition()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.SingleOnly = true;

            opts.MessageForAdding = "Select block references to Create Stair: ";
            PromptSelectionResult res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }
            if (res.Value.Count != 0)
            {
                SelectionSet set = res.Value;
                ObjectId[] Ids = set.GetObjectIds();
                using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockReference blk1 = tr.GetObject(Ids[0], OpenMode.ForRead) as BlockReference;
                    ed.WriteMessage("\n Poissition:({0},{1},{2})", blk1.Position.X, blk1.Position.Y, blk1.Position.Z);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("BlockPatternByPolyType")]
        public static void BlockPatternByPolyType()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            // Select Block to Project 
            //TypedValue[] filList = new TypedValue[2] {new TypedValue((int)DxfCode.Start, "LINE"),new TypedValue((int)DxfCode.Start,"POLYLINE") };
            //SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            opts.MessageForAdding = "Select Polyline: ";
            PromptSelectionResult res = ed.GetSelection(opts);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }

            if (res.Value.Count != 0)
            {
                SelectionSet set = res.Value;
                ObjectId[] Ids = set.GetObjectIds();
                ObjectIdCollection SortId = new ObjectIdCollection();
                ObjectIdCollection IdOther = new ObjectIdCollection();
                for (int i = 0; i < Ids.Length; i++)
                {
                    ObjectId id = Ids[i];
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        Entity pl = acTrans.GetObject(id, OpenMode.ForRead) as Entity;
                        if (pl.Layer == "1MIDRON_DOWN")
                        {
                            SortId.Add(id);
                        }
                        else
                        {
                            IdOther.Add(id);
                        }
                        acTrans.Abort();
                    }
                }
                for (int i = 0; i < IdOther.Count; i++)
                {
                    SortId.Add(IdOther[i]);
                }
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    ObjectIdCollection ignoredlist = new ObjectIdCollection();
                    ObjectId idtoSaveId = ObjectId.Null;
                    foreach (ObjectId id in SortId)
                    {
                        if (ignoredlist.Contains(id))
                        {
                            continue;
                        }
                        Polyline pl = acTrans.GetObject(id, OpenMode.ForRead) as Polyline;
                        Line line = acTrans.GetObject(id, OpenMode.ForRead) as Line;
                        ObjectIdCollection listBlock = new ObjectIdCollection();
                        idtoSaveId = id;
                        if (pl != null)
                        {
                            if (pl.Layer == "1MIDRON_DOWN")
                            {
                                double min = 1000;
                                ObjectId closestId = ObjectId.Null;
                                for (int j = 0; j < pl.NumberOfVertices; j++)
                                {
                                    for (int i = 0; i < SortId.Count; i++)
                                    {
                                        ObjectId plId = SortId[i];
                                        if (plId == id)
                                        {
                                            continue;
                                        }
                                        Polyline closestPl = acTrans.GetObject(plId, OpenMode.ForRead) as Polyline;
                                        if (closestPl == null)
                                        {
                                            continue;
                                        }
                                        if (closestPl.Layer != "1MIDRON_UP")
                                        {
                                            continue;
                                        }
                                        for (int k = 0; k < closestPl.NumberOfVertices; k++)
                                        {
                                            double distance = closestPl.GetPoint3dAt(k).DistanceTo(pl.GetPoint3dAt(j));
                                            if (distance < min)
                                            {
                                                min = distance;
                                                closestId = plId;
                                            }
                                        }
                                    }
                                }
                                Polyline clPl = acTrans.GetObject(closestId, OpenMode.ForWrite) as Polyline;
                                var bt = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                                var PointOnclPl = clPl.StartPoint;
                                var pointonPl = pl.GetClosestPointTo(PointOnclPl, true);
                                bool AddPI = false;
                                if (PointOnclPl.DistanceTo(Point3d.Origin) > pointonPl.DistanceTo(Point3d.Origin))
                                {
                                    AddPI = true;
                                }
                                // Check whether the first block exists
                                if (bt.Has("ANOT9106"))
                                {
                                    ObjectId blkRecId = bt["ANOT9106"];
                                    double distance = 3.5;

                                    for (int i = 1; i < clPl.Length / distance; i++)
                                    {
                                        Point3d pt = clPl.GetPointAtDist(i * distance);
                                        int ind = ArxHelper.GetVertexIndexes(pt, closestId);

                                        double angle = (clPl.GetPoint3dAt(ind + 1) - clPl.GetPoint3dAt(ind)).GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);

                                        using (BlockReference acBlkRef = new BlockReference(pt, blkRecId))
                                        {
                                            acBlkRef.ScaleFactors = new Scale3d(0.25, 0.25, 0.25);
                                            if (angle > Math.PI && angle < Math.PI * 2)
                                            {
                                                double rot = Math.PI - angle;
                                                if (AddPI)
                                                {
                                                    acBlkRef.Rotation = rot + Math.PI;
                                                }
                                                else
                                                {
                                                    acBlkRef.Rotation = rot;
                                                }

                                            }
                                            else
                                            {
                                                double rot = 2 * Math.PI - angle;

                                                if (AddPI)
                                                {
                                                    acBlkRef.Rotation = rot + Math.PI;
                                                }
                                                else
                                                {
                                                    acBlkRef.Rotation = rot;
                                                }
                                            }
                                            //else
                                            //{
                                            //    acBlkRef.Rotation = Math.PI + angle;
                                            //}
                                            BlockTableRecord acCurSpaceBlkTblRec;
                                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            listBlock.Add(acCurSpaceBlkTblRec.AppendEntity(acBlkRef));
                                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                                            idtoSaveId = closestId;

                                        }
                                    }
                                    ignoredlist.Add(closestId);
                                }

                            }
                            if (pl.Layer == "1TERASA_UP" || pl.Layer == "1MIDRON_UP")
                            {
                                var bt = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                                // Check whether the first block exists
                                if (bt.Has("ANOT9106"))
                                {
                                    ObjectId blkRecId = bt["ANOT9106"];
                                    double distance = 3.5;
                                    for (int i = 0; i < pl.Length / distance; i++)
                                    {
                                        Point3d pt = pl.GetPointAtDist(i * distance);
                                        int ind = ArxHelper.GetVertexIndexes(pt, id);
                                        double angle = (pl.GetPoint3dAt(ind + 1) - pl.GetPoint3dAt(ind)).GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);
                                        using (BlockReference acBlkRef = new BlockReference(pt, blkRecId))
                                        {
                                            acBlkRef.ScaleFactors = new Scale3d(0.25, 0.25, 0.25);
                                            if (angle > Math.PI && angle < Math.PI * 2)
                                            {
                                                acBlkRef.Rotation = Math.PI - angle;
                                            }
                                            else
                                            {
                                                acBlkRef.Rotation = 2 * Math.PI - angle;
                                            }
                                            //else
                                            //{
                                            //    acBlkRef.Rotation = Math.PI + angle;
                                            //}
                                            BlockTableRecord acCurSpaceBlkTblRec;
                                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            listBlock.Add(acCurSpaceBlkTblRec.AppendEntity(acBlkRef));
                                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                                        }
                                    }
                                }
                            }
                            if (pl.Layer == "M2603")
                            {
                                var bt = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                                // Check whether the first block exists
                                if (bt.Has("M2603"))
                                {
                                    ObjectId blkRecId = bt["M2603"];
                                    double distance = 3.5;
                                    for (int i = 0; i < pl.Length / distance; i++)
                                    {
                                        Point3d pt = pl.GetPointAtDist(i * distance);
                                        int ind = ArxHelper.GetVertexIndexes(pt, id);
                                        double angle = (pl.GetPoint3dAt(ind + 1) - pl.GetPoint3dAt(ind)).GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);
                                        using (BlockReference acBlkRef = new BlockReference(pt, blkRecId))
                                        {
                                            acBlkRef.ScaleFactors = new Scale3d(0.25, 0.25, 0.25);
                                            if (angle > Math.PI && angle < Math.PI * 2)
                                            {
                                                acBlkRef.Rotation = Math.PI - angle;
                                            }
                                            else
                                            {
                                                acBlkRef.Rotation = 2 * Math.PI - angle;
                                            }
                                            //else
                                            //{
                                            //    acBlkRef.Rotation = Math.PI + angle;
                                            //}
                                            BlockTableRecord acCurSpaceBlkTblRec;
                                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            listBlock.Add(acCurSpaceBlkTblRec.AppendEntity(acBlkRef));
                                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                                        }
                                    }
                                }
                            }
                            if (pl.Layer == "1GADER-BARZEL")
                            {
                                var bt = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                                // Check whether the first block exists
                                if (bt.Has("ANOT104"))
                                {
                                    ObjectId blkRecId = bt["ANOT104"];
                                    double distance = 3.5;
                                    for (int i = 0; i < pl.Length / distance; i++)
                                    {
                                        Point3d pt = pl.GetPointAtDist(i * distance);
                                        int ind = ArxHelper.GetVertexIndexes(pt, id);
                                        double angle = (pl.GetPoint3dAt(ind + 1) - pl.GetPoint3dAt(ind)).GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);
                                        using (BlockReference acBlkRef = new BlockReference(pt, blkRecId))
                                        {
                                            acBlkRef.ScaleFactors = new Scale3d(0.25, 0.25, 0.25);
                                            if (angle > Math.PI && angle < Math.PI * 2)
                                            {
                                                acBlkRef.Rotation = Math.PI - angle;
                                            }
                                            else
                                            {
                                                acBlkRef.Rotation = 2 * Math.PI - angle;
                                            }
                                            //else
                                            //{
                                            //    acBlkRef.Rotation = Math.PI + angle;
                                            //}
                                            BlockTableRecord acCurSpaceBlkTblRec;
                                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            listBlock.Add(acCurSpaceBlkTblRec.AppendEntity(acBlkRef));
                                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                                        }
                                    }
                                }
                            }
                            if (pl.Layer == "gush")
                            {
                                var bt = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                                // Check whether the first block exists
                                if (bt.Has("ANOT101"))
                                {
                                    ObjectId blkRecId = bt["ANOT101"];
                                    double distance = 3.5;
                                    for (int i = 1; i < pl.Length / distance; i++)
                                    {
                                        Point3d pt = pl.GetPointAtDist(i * distance);
                                        int ind = ArxHelper.GetVertexIndexes(pt, id);
                                        double angle = (pl.GetPoint3dAt(ind + 1) - pl.GetPoint3dAt(ind)).GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);

                                        using (BlockReference acBlkRef = new BlockReference(pt, blkRecId))
                                        {
                                            acBlkRef.ScaleFactors = new Scale3d(0.25, 0.25, 0.25);
                                            if (angle > Math.PI && angle < Math.PI * 2)
                                            {
                                                double rot = Math.PI - angle;
                                                if (i % 2 == 0)
                                                {
                                                    rot = rot + Math.PI;
                                                }
                                                acBlkRef.Rotation = rot;
                                            }
                                            else
                                            {
                                                double rot = 2 * Math.PI - angle;
                                                if (i % 2 == 0)
                                                {
                                                    rot = rot + Math.PI;
                                                }
                                                acBlkRef.Rotation = rot;
                                            }
                                            //else
                                            //{
                                            //    acBlkRef.Rotation = Math.PI + angle;
                                            //}
                                            BlockTableRecord acCurSpaceBlkTblRec;
                                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            listBlock.Add(acCurSpaceBlkTblRec.AppendEntity(acBlkRef));
                                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                                        }
                                    }
                                }
                            }

                        }
                        if (line != null)
                        {

                            if (line.Layer == "1TERASA_UP")
                            {
                                var bt = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                                // Check whether the first block exists
                                if (bt.Has("ANOT9106"))
                                {
                                    ObjectId blkRecId = bt["ANOT9106"];
                                    double distance = 3.5;
                                    double angle = (line.EndPoint - line.StartPoint).GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);
                                    for (int i = 0; i < line.Length / distance; i++)
                                    {
                                        Point3d pt = line.GetPointAtDist(i * distance);
                                        using (BlockReference acBlkRef = new BlockReference(pt, blkRecId))
                                        {
                                            acBlkRef.ScaleFactors = new Scale3d(0.25, 0.25, 0.25);
                                            if (angle > Math.PI && angle < Math.PI * 2)
                                            {
                                                acBlkRef.Rotation = Math.PI - angle;
                                            }
                                            else
                                            {
                                                acBlkRef.Rotation = 2 * Math.PI - angle;
                                            }
                                            //else
                                            //{
                                            //    acBlkRef.Rotation = Math.PI + angle;
                                            //}
                                            BlockTableRecord acCurSpaceBlkTblRec;
                                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            listBlock.Add(acCurSpaceBlkTblRec.AppendEntity(acBlkRef));
                                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                                        }
                                    }
                                }
                            }

                        }
                        if (listBlock.Count > 0)
                        {
                            ArxHelper.AddRegAppTableRecord("BlockList");
                            ResultBuffer rb = new ResultBuffer();

                            foreach (ObjectId blkId in listBlock)
                            {
                                if (blkId.IsValid)
                                {
                                    rb.Add(new TypedValue((int)DxfCode.SoftPointerId, blkId));
                                }

                            }
                            ArxHelper.AttachDataToEntity(idtoSaveId, rb, "BlockList");
                            idtoSaveId = ObjectId.Null;
                            rb.Dispose();
                        }
                    }
                    acTrans.Commit();
                }
            }
        }
        [CommandMethod("RotateBlockOnPolyline")]
        public static void RotateBlockOnPolyline()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.SingleOnly = true;
            opts.MessageForAdding = "Select Polyline: ";
            PromptSelectionResult res = ed.GetSelection(opts);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }
            SelectionSet set = res.Value;
            ObjectId[] Ids = set.GetObjectIds();
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {


                {
                    ResultBuffer rb = new ResultBuffer();
                    ArxHelper.ReadEntityData(Ids[0], ref rb, "BlockList");
                    ObjectIdCollection listBlock = new ObjectIdCollection();
                    if (rb != null)
                    {
                        foreach (TypedValue tv in rb)
                        {
                            if (tv.TypeCode == (int)DxfCode.SoftPointerId)
                            {
                                ObjectId id = (ObjectId)tv.Value;
                                if (id.IsValid && id.IsErased == false)
                                {
                                    listBlock.Add(id);
                                }
                            }
                        }
                    }
                    foreach (ObjectId id in listBlock)
                    {
                        if (id.IsValid == false || id.IsErased == true)
                        {
                            continue;
                        }
                        BlockReference oEnt = acTrans.GetObject(id, OpenMode.ForWrite) as BlockReference;
                        if (oEnt != null)
                        {
                            Point3d basePoint = oEnt.Position;
                            oEnt.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, basePoint));
                        }

                    }
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("MoveVertices")]
        public static void MoveVertices2C1610()
        {
            double ElevationTol = 4;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
            SelectionFilter filter = new SelectionFilter(filList);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            opts.MessageForAdding = "Select block references: ";
            PromptSelectionResult res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }
            TypedValue[] filListPoly = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "LWPOLYLINE") };
            SelectionFilter filterPoly = new SelectionFilter(filListPoly);
            PromptSelectionOptions optPoly = new PromptSelectionOptions();

            optPoly.MessageForAdding = "Select Polyline: ";
            PromptSelectionResult r = ed.GetSelection(optPoly, filterPoly);
            if (r.Status != PromptStatus.OK)
            {
                return;
            }
            if (res.Value.Count != 0 && r.Value.Count != 0)
            {
                SelectionSet set = res.Value;
                SelectionSet ssPl = r.Value;
                ObjectId[] Ids = set.GetObjectIds();
                ObjectId[] IdPolys = ssPl.GetObjectIds();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in Ids)
                    {
                        if (id.IsValid == false || id.IsErased == true)
                        {
                            continue;
                        }
                        BlockReference blk = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (blk == null)
                        {
                            continue;
                        }
                        if (blk.Layer != "C1610" && blk.Layer != "C1610_0")
                        {
                            continue;
                        }
                        foreach (ObjectId plId in IdPolys)
                        {
                            if (plId.IsValid == false || plId.IsErased == true)
                            {
                                continue;
                            }
                            Polyline pl = tr.GetObject(plId, OpenMode.ForWrite) as Polyline;
                            for (int i = 0; i < pl.NumberOfVertices; i++)
                            {
                                Point3d pt = pl.GetPoint3dAt(i);
                                if (pt.DistanceTo(blk.Position) <= 4)
                                {
                                    pl.SetPointAt(i, new Point2d(blk.Position.X, blk.Position.Y));
                                }
                            }
                        }
                    }
                    tr.Commit();
                }

            }
        }
        [CommandMethod("Taba_Puzzle")]
        public static void Taba_Puzzle()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            // Read Excel File
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.RestoreDirectory = false;
            openFileDialog.CheckFileExists = true;
            if (openFileDialog.ShowDialog() == true)
            {
                string outputFileName = openFileDialog.FileName;
                // Check file exist or not
                if (ExcelHelper.IsFileLocked(outputFileName))
                {
                    MessageBox.Show("This File is Lock", "Warnning");
                    return;
                }
                try
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(outputFileName, false))
                    {
                        // Get the SharedStringTablePart. If it does not exist, create a new one.
                        WorkbookPart workbookPart = document.WorkbookPart;
                        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                        string relationshipId = sheets.First().Id.Value;
                        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                        Worksheet workSheet = worksheetPart.Worksheet;
                        SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                        // Get rows data from the sheet
                        IEnumerable<Row> rows = sheetData.Descendants<Row>();
                        // We only accept excel file that has 2 columns (layer name, color and transparency)
                        if (rows.ElementAt(0).Descendants<Cell>().Count() != 2)
                        {
                            MessageBox.Show("File Format is wrong. File should have 3 column", "Warning");
                            return;
                        }
                        ObservableCollection<PolygonOrderInfor> polygons = new ObservableCollection<PolygonOrderInfor>();
                        int prio = 0;
                        foreach (Row row in rows)
                        {
                            // Skip header row
                            if (row == rows.ElementAt(0))
                                continue;
                            if (ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(0)) == string.Empty)
                            {
                                continue;
                            }

                            // Get cell value and paste into new Layer Item
                            string Layer = ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(0));
                            // Skip invalid name
                            if (Layer == string.Empty)
                            {
                                continue;
                            }
                            string date = ExcelHelper.GetCellValue(document, row.Descendants<Cell>().ElementAt(1));
                            prio++;
                            PolygonOrderInfor plygoninfor = new PolygonOrderInfor();
                            plygoninfor.LayerName = Layer;
                            plygoninfor.Date = date;
                            plygoninfor.priority = prio;
                            plygoninfor.Ids = ArxHelper.GetEntitiesOnLayer(Layer);
                            polygons.Add(plygoninfor);

                        }
                        ObjectIdCollection delId = new ObjectIdCollection();
                        foreach (PolygonOrderInfor plinfor in polygons)
                        {
                            //Get all entity have priority lower than current polyline
                            var checklist = polygons.Where(x => x.priority > plinfor.priority).ToList();
                            foreach (ObjectId id in plinfor.Ids)
                            {
                                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                {
                                    if (id.IsValid == false || id.IsErased == true)
                                    {
                                        acTrans.Abort();
                                        continue;
                                    }
                                    Polyline pline = acTrans.GetObject(id, OpenMode.ForWrite) as Polyline;
                                    if (pline == null)
                                    {
                                        acTrans.Abort();
                                        continue;
                                    }

                                    foreach (PolygonOrderInfor inf in checklist)
                                    {
                                        ObjectIdCollection addIds = new ObjectIdCollection();
                                        if (inf.Ids == null)
                                        {
                                            continue;
                                        }
                                        foreach (ObjectId objid in inf.Ids)
                                        {
                                            if (objid.IsValid == false || objid.IsErased == true)
                                            {
                                                continue;
                                            }
                                            Polyline pltocheck = acTrans.GetObject(objid, OpenMode.ForWrite) as Polyline;
                                            if (pltocheck == null)
                                            {
                                                continue;
                                            }
                                            if (ArxHelper.IsPolygonInsideOther(id, objid))
                                            {
                                                ArxHelper.DeleteEntity(objid);
                                            }
                                            Point3dCollection pts = new Point3dCollection();
                                            pltocheck.IntersectWith(pline, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);
                                            if (pts.Count > 1)
                                            {
                                                bool isvertice = false;
                                                Point3dCollection Vertices = new Point3dCollection();
                                                for (int i =0;i< pltocheck.NumberOfVertices; i++)
                                                {
                                                    foreach (Point3d pt in pts)
                                                    {
                                                        if (pt ==pltocheck.GetPoint3dAt(i))
                                                        {
                                                            isvertice = true;
                                                            break;
                                                        }
                                                    }
                                                    if (isvertice)
                                                    {
                                                        break;
                                                    }
                                                }
                                                DBObjectCollection dbobj = new DBObjectCollection();
                                                //if (isvertice)
                                                //{
                                                //    pltocheck.Explode(dbobj);
                                                //    Entity[] joinlist = null;
                                                //    Line first = null;
                                                //    int i = 0;
                                                //    foreach (DBObject obj in dbobj)
                                                //    {
                                                //        i++;
                                                //        var cur = obj as Line;
                                                        
                                                //        if (cur != null)
                                                //        {
                                                //            Point3dCollection ptintersec = new Point3dCollection();
                                                //            cur.IntersectWith(pline, Intersect.OnBothOperands, ptintersec, IntPtr.Zero, IntPtr.Zero);
                                                //            if (ptintersec.Count==0)
                                                //            {
                                                //                if (ArxHelper.InsidePolygon(pline,cur.StartPoint) && ArxHelper.InsidePolygon(pline, cur.EndPoint))
                                                //                {

                                                //                }
                                                //                else
                                                //                {
                                                //                    if (first == null)
                                                //                    {
                                                //                        first = cur;
                                                //                    }
                                                //                    else
                                                //                    {
                                                //                        first.UpgradeOpen();
                                                //                        first.JoinEntity(cur);
                                                //                    }
                                                                    
                                                //                }
                                                //            }
                                                //        }
                                                //    }
                                                                                                    
                                                //    ArxHelper.AppendEntity(first);
                                                //    // Remove Origin Polyline
                                                //    //ArxHelper.DeleteEntity(objid);
                                                //    delId.Add(objid);
                                                //    inf.Ids.Remove(objid);
                                                //}
                                                //else
                                                {
                                                    dbobj = pltocheck.GetSplitCurves(pts);
                                                    foreach (DBObject obj in dbobj)
                                                    {
                                                        var cur = obj as Polyline;
                                                        if (cur != null)
                                                        {
                                                            Point3d mid = new Point3d();
                                                            // Find a point on polyline to check Inside or outside of Polygon
                                                            if (cur.NumberOfVertices == 2)
                                                            {
                                                                mid = new Point3d((cur.StartPoint.X + cur.EndPoint.X) / 2, (cur.StartPoint.Y + cur.EndPoint.Y) / 2, (cur.StartPoint.Z + cur.EndPoint.Z) / 2);
                                                            }
                                                            else
                                                            {
                                                                mid = cur.GetPoint3dAt(1);
                                                            }
                                                            ResultBuffer read_rb = new ResultBuffer();
                                                            ArxHelper.ReadEntityData(pline.ObjectId, ref read_rb, "OriginId");
                                                            ObjectId OriginId = ObjectId.Null;
                                                            if (read_rb != null)
                                                            {
                                                                foreach (TypedValue tv in read_rb)
                                                                {
                                                                    if (tv.TypeCode == (int)DxfCode.SoftPointerId)
                                                                    {
                                                                        OriginId = (ObjectId)tv.Value;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                            // If Outside, Append to DB
                                                            if (OriginId.IsValid && OriginId.IsErased == false)
                                                            {
                                                                Polyline originpl = acTrans.GetObject(OriginId, OpenMode.ForRead) as Polyline;
                                                                if (ArxHelper.InsidePolygon(originpl, mid) || ArxHelper.IsPointOnPolyline(originpl, mid))
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    ObjectId curid = ArxHelper.AppendEntity(cur);
                                                                    ArxHelper.AddRegAppTableRecord("OriginId");
                                                                    ResultBuffer rb = new ResultBuffer();

                                                                    if (curid.IsValid)
                                                                    {
                                                                        rb.Add(new TypedValue((int)DxfCode.SoftPointerId, objid));
                                                                    }
                                                                    ArxHelper.AttachDataToEntity(curid, rb, "OriginId");
                                                                    addIds.Add(curid);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (ArxHelper.InsidePolygon(pline, mid) || ArxHelper.IsPointOnPolyline(pline, mid))
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    ObjectId curid = ArxHelper.AppendEntity(cur);
                                                                    ArxHelper.AddRegAppTableRecord("OriginId");
                                                                    ResultBuffer rb = new ResultBuffer();

                                                                    if (curid.IsValid)
                                                                    {
                                                                        rb.Add(new TypedValue((int)DxfCode.SoftPointerId, objid));
                                                                    }
                                                                    ArxHelper.AttachDataToEntity(curid, rb, "OriginId");
                                                                    addIds.Add(curid);
                                                                }
                                                            }
                                                        }
                                                    }
                                                    // Remove Origin Polyline
                                                    //ArxHelper.DeleteEntity(objid);
                                                    delId.Add(objid);
                                                    inf.Ids.Remove(objid);
                                                }
                                            }
                                        }
                                        foreach (ObjectId x in addIds)
                                        {
                                            inf.Ids.Add(x);
                                        }
                                    }

                                    acTrans.Commit();
                                }
                            }
                        }
                        foreach (ObjectId id in delId)
                        {
                            ArxHelper.DeleteEntity(id);
                        }
                    }
                }
                // Using OpenXml to read the excel file
                catch (System.Exception exc)
                {

                    return;
                }

            }
        }
       [CommandMethod("PointInPoly")]
        public static void PointInPoly()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.SingleOnly = true;
            opts.SingleOnly = true;
            opts.MessageForAdding = "Select Polyline: ";
            PromptSelectionResult res = ed.GetSelection(opts);
            if (res.Status != PromptStatus.OK)
            {
                return;
            }
            SelectionSet set = res.Value;
            ObjectId[] Ids = set.GetObjectIds();
            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            // Prompt for a point
            pPtOpts.Message = "\nOrigin Point: ";
            pPtRes = acDoc.Editor.GetPoint(pPtOpts);
            Point3d pPtOrigin;
            // If a point was entered, then translate it to the current UCS
            if (pPtRes.Status == PromptStatus.OK)
            {
                pPtOrigin = pPtRes.Value;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    Polyline polygon = acTrans.GetObject(Ids[0],OpenMode.ForRead) as Polyline;
                    if(ArxHelper.IsPointOnPolyline(polygon, pPtOrigin))
                    {
                        ed.WriteMessage("\n Point On Polyline");
                    }
                    Xline xl = new Xline();
                    xl.BasePoint = pPtOrigin;
                    xl.UnitDir = Vector3d.XAxis;
                    Point3dCollection listPt = new Point3dCollection();
                    xl.IntersectWith(polygon,Intersect.OnBothOperands,listPt,IntPtr.Zero,IntPtr.Zero);
                    if (listPt.Count>0)
                    {
                        Point3dCollection pts = new Point3dCollection();
                        foreach (Point3d pt in listPt)
                        {
                            if (pt.X> pPtOrigin.X)
                            {
                                if (ArxHelper.PointIsVertices(pt,polygon))
                                {
                                    pts.Add(pt);
                                }
                                pts.Add(pt);
                            }
                        }
                        if (pts.Count == 0)
                        {
                            ed.WriteMessage("\n Point Outside Polyline");
                        }
                        if (pts.Count % 2 != 0)
                        {
                            ed.WriteMessage("\n Point Inside Polyline");
                        }
                        else
                        {
                            ed.WriteMessage("\n Point Outside Polyline");
                        }
                    }
                    acTrans.Commit();
                }
            }
        }
    }
}
