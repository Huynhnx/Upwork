using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
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
                        using (Transaction tr =db.TransactionManager.StartTransaction())
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

                                foreach (ObjectId id in set.GetObjectIds())
                                {
                                    BlockReference oEnt = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                                    BlockReferenceProperties blkproperties = new BlockReferenceProperties();
                                    blkproperties.BlockName = oEnt.Name;
                                    blkproperties.BlockId = id;
                                    if (oEnt.AttributeCollection.Count>0)
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
                    if (vm.BlkProperties !=null)
                    {
                        foreach(var blkprop in vm.BlkProperties)
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
    }
}
