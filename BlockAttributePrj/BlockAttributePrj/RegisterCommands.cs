using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.BoundaryRepresentation;
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
    public class SegmentInfor
    {
        public int IndexSegment;
        public int PartNumber;
        public Point3d StartPoint;

    }
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
            PromptSelectionOptions opts = new PromptSelectionOptions();
            //            opts.AllowSubSelections = true;
            //            opts.ForceSubSelections = true;
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
                int j = 0;
                if (listsegment.Count>1)
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
                        if (listsegment[i] != listsegment[i-1]+1)
                        {
                            j++;
                        }
                        SegmentInfor partseg = new SegmentInfor();
                        partseg.PartNumber = j;
                        partseg.IndexSegment = listsegment[i];
                        LineSegment3d segmentpart3d = pl.GetLineSegmentAt(listsegment[i]);
                        partseg.StartPoint = segmentpart3d.StartPoint;
                        segmentInfo.Add(partseg);
                    }
                    for (int i=0;i<=j;i++)
                    {
                        List<SegmentInfor> seg = segmentInfo.Where(x => x.PartNumber == i).ToList();
                        Point3d pt = new Point3d();
                        double distance = 0;
                        Point3d ptStart = seg[0].StartPoint;
                        Point3d ptEnd = pl.GetLineSegmentAt(pl.NumberOfVertices - 2).EndPoint;
                        Line line = new Line(ptStart,ptEnd);
                        if (seg.Count>1)
                        {
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
                            for (int n = 1; n < seg.Count-1; n++)
                            {
                                SegmentInfor info = seg[n];
                                if (pl.NumberOfVertices > 2)
                                {
                                    var vex = info.IndexSegment + 1;
                                    if (info.IndexSegment >= pl.NumberOfVertices)
                                    {
                                        continue;
                                    }
                                    pl.RemoveVertexAt(vex);
                                }
                            }
                            
                            // convert points to 2d points
                            var plane = new Plane(Point3d.Origin, Vector3d.ZAxis);
                            var p1 = ptStart.Convert2d(plane);
                            var p2 = pt.Convert2d(plane);
                            var p3 = ptEnd.Convert2d(plane);

                            // compute the bulge of the second segment
                            var angle1 = p1.GetVectorTo(p2).Angle;
                            var angle2 = p2.GetVectorTo(p3).Angle;
                            var bulge = Math.Tan((angle2 - angle1) / 2.0);
                            pl.SetBulgeAt(seg[0].IndexSegment,bulge);
                        }                    
                    }
                }

                tr.Commit();
            }


        }

        [CommandMethod("Test")]
        public void Test()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // prompt for first point
            var options = new PromptPointOptions("\nFirst point: ");
            var result = ed.GetPoint(options);
            if (result.Status != PromptStatus.OK)
                return;
            var pt1 = result.Value;

            // prompt for second point
            options.Message = "\nSecond point: ";
            options.BasePoint = pt1;
            options.UseBasePoint = true;
            result = ed.GetPoint(options);
            if (result.Status != PromptStatus.OK)
                return;
            var pt2 = result.Value;

            // prompt for third point
            options.Message = "\nThird point: ";
            options.BasePoint = pt2;
            result = ed.GetPoint(options);
            if (result.Status != PromptStatus.OK)
                return;
            var pt3 = result.Value;

            // convert points to 2d points
            var plane = new Plane(Point3d.Origin, Vector3d.ZAxis);
            var p1 = pt1.Convert2d(plane);
            var p2 = pt2.Convert2d(plane);
            var p3 = pt3.Convert2d(plane);

            // compute the bulge of the second segment
            var angle1 = p1.GetVectorTo(p2).Angle;
            var angle2 = p2.GetVectorTo(p3).Angle;
            var bulge = Math.Tan((angle2 - angle1) / 2.0);

            // add the polyline to the current space
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                using (var pline = new Polyline())
                {
                    pline.AddVertexAt(0, p1, 0.0, 0.0, 0.0);
                    pline.AddVertexAt(1, p2, bulge, 0.0, 0.0);
                    pline.AddVertexAt(2, p3, 0.0, 0.0, 0.0);
                    pline.TransformBy(ed.CurrentUserCoordinateSystem);
                    curSpace.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);
                }
                tr.Commit();
            }
        }
    }
}
