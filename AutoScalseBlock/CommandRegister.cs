using Autodesk.AutoCAD.ApplicationServices;
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
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Internal;
using Microsoft.Win32;
using System.Windows.Forms;
using System.IO;

namespace AutoScalseBlock
{
    public class CommandRegister
    {
        [CommandMethod("SetScaleFactor")]
        public static void SetScaleFactor()
        {
            Document acDoc = AcadService.Application.DocumentManager.MdiActiveDocument;

            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\OurSettings");

            PromptDoubleResult pr = acDoc.Editor.GetDouble("Get ScaleFactor:");
            if (pr.Status == PromptStatus.OK)
            {
                if (key != null)
                {
                    key.SetValue("Factor", pr.Value);
                    AppLoader.factor = pr.Value;
                }
            }
            key.Close();
        }
        [CommandMethod("LO")]
        public static void TurnOnAllLayer()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    SymbolTable symTable = (SymbolTable)acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForWrite);
                    foreach (ObjectId id in symTable)
                    {
                        LayerTableRecord symbol = (LayerTableRecord)acTrans.GetObject(id, OpenMode.ForWrite);
                        if (symbol.IsOff)
                        {
                            symbol.IsOff = false;
                        }
                        if (symbol.IsHidden)
                        {
                            symbol.IsHidden = false;
                        }
                        if (symbol.IsLocked)
                        {
                            symbol.IsLocked = false;
                        }
                        if (symbol.IsFrozen)
                        {
                            symbol.IsFrozen = false;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    acTrans.Abort();
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("BA")]
        public void ExplodeBock()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection ids = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                BlockTableRecord ms = tr.GetObject(msId, OpenMode.ForWrite) as BlockTableRecord;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                // loop current drawing entities
                foreach (ObjectId objId in btr)
                {
                    if (objId.IsValid == false || objId.IsErased == true)
                    {
                        continue;
                    }
                    BlockRefType type = Helper.GetBlockReferenceType(objId);
                    if (type == BlockRefType.Block)
                    {
                        BlockReference blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                        if (blockRef != null)
                        {
                            ExplodeBlock(blockRef, ms);
                        }
                    }

                }
                tr.Commit();
            }
            //foreach (ObjectId id in ids)
            //{
            //    ExplodeBlock(id);
            //}
           
        }
        static void ExplodeBlock(BlockReference blockRef, BlockTableRecord ms = null)
        {
            
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                ObjectId msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                if (ms == null)
                {
                    ms = tr.GetObject(msId, OpenMode.ForWrite) as BlockTableRecord;
                }
                    if (blockRef != null)
                    {
                        DBObjectCollection toAddColl = new DBObjectCollection();
                        BlockTableRecord blockDef = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        // Create a text for const and  
                        // visible attributes 
                        foreach (ObjectId entId in blockDef)
                        {
                            if (entId.ObjectClass.Name == "AcDbAttributeDefinition")
                            {
                                AttributeDefinition attDef = tr.GetObject(entId, OpenMode.ForRead) as AttributeDefinition;
                                if ((attDef.Constant && !attDef.Invisible))
                                {
                                    DBText text = new DBText();
                                    text.Height = attDef.Height;
                                    text.TextString = attDef.TextString;
                                    text.Position = attDef.Position.TransformBy(blockRef.BlockTransform);
                                    toAddColl.Add(text);
                                }
                            }
                        }
                    // Create a text for non-const  
                    // and visible attributes
                    try
                    {
                        foreach (ObjectId attId in blockRef.AttributeCollection)
                        {
                            AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                            if (attRef.Invisible == false)
                            {
                                DBText text = new DBText();
                                text.Height = attRef.Height;
                                text.TextString = attRef.TextString;
                                text.Position = attRef.Position;
                                toAddColl.Add(text);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                    try
                    {
                        foreach (AttributeReference attRef in blockRef.AttributeCollection)
                        {
                            
                            if (attRef.Invisible == false)
                            {
                                DBText text = new DBText();
                                text.Height = attRef.Height;
                                text.TextString = attRef.TextString;
                                text.Position = attRef.Position;
                                toAddColl.Add(text);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                    // Get the entities from the  
                    // block reference 
                    // Attribute definitions have  
                    // been taken care of.. 
                    // So ignore them 
                    DBObjectCollection entityColl = new DBObjectCollection();
                    try
                    {
                        blockRef.Explode(entityColl);
                    }
                    catch (System.Exception ex)
                    {

                    }
                        
                        foreach (Entity ent in entityColl)
                        {
                        if ((ent is BlockReference))
                        {
                            ExplodeBlock((BlockReference)ent, ms);
                        }                         
                            if (!(ent is AttributeDefinition) && !(ent is BlockReference))
                            {
                                toAddColl.Add(ent);
                            }
                        }
                        // Add the entities to modelspace 
                        foreach (Entity ent in toAddColl)
                        {
                            ms.AppendEntity(ent);
                            tr.AddNewlyCreatedDBObject
                            (ent, true);
                        }
                    // Erase the block reference 
                    try
                    {
                        blockRef.UpgradeOpen();
                    }
                    catch (System.Exception ex)
                    {

                    }
                    try
                    {
                        blockRef.Erase();
                    }
                    catch (System.Exception ex)
                    {

                    }
                   
                    }

                tr.Commit();
            }
        }
        [CommandMethod("XU")]
        static public void ReloadXRefs()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectIdCollection ids = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable table = tr.GetObject(db.BlockTableId,
                                            OpenMode.ForRead) as BlockTable;
                foreach (ObjectId id in table)
                {
                    BlockTableRecord record = tr.GetObject(id,

                                       OpenMode.ForRead) as BlockTableRecord;
                    if (record.IsFromExternalReference)
                    {
                        ids.Add(id);
                    }
                }
                tr.Commit();
            }
            //now relaod the xrefs
            if (ids.Count != 0)
            {
                db.ReloadXrefs(ids);
            }
        }
        [CommandMethod("BD")]
        public void BindXrefs()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectIdCollection xrefCollection = new ObjectIdCollection();
            using (XrefGraph xg = db.GetHostDwgXrefGraph(false))
            {
                int numOfNodes = xg.NumNodes;
                for (int cnt = 0; cnt < xg.NumNodes; cnt++)
                {
                    XrefGraphNode xNode = xg.GetXrefNode(cnt)
                                                        as XrefGraphNode;
                    if (xNode.Database != null)
                    {
                        if (!xNode.Database.Filename.Equals(db.Filename))
                        {
                            if (xNode.XrefStatus == XrefStatus.Resolved)
                            {
                                xrefCollection.Add(xNode.BlockTableRecordId);
                            }
                        }
                    }
                }
            }
            if (xrefCollection.Count != 0)
                db.BindXrefs(xrefCollection, true);
        }
        [CommandMethod("PA")]

        public static void PurgeBlocks()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable table = Tx.GetObject(db.BlockTableId,
                    OpenMode.ForRead) as BlockTable;
                ObjectIdCollection blockIds = new ObjectIdCollection();
                //using do/while loop to purge nested blocks.
                do
                {
                    blockIds.Clear();
                    foreach (ObjectId id in table)
                        blockIds.Add(id);
                    //this function will remove all
                    //blocks which are in use
                    db.Purge(blockIds);
                    foreach (ObjectId id in blockIds)
                    {
                        DBObject obj = Tx.GetObject(id, OpenMode.ForWrite);
                        obj.Erase();
                    }
                }
                while (blockIds.Count != 0);
                //Layer
                SymbolTable symTable = (SymbolTable)Tx.GetObject(db.LayerTableId, OpenMode.ForWrite);
                ObjectIdCollection layerIds = new ObjectIdCollection();

                foreach (ObjectId id in symTable)
                {
                    layerIds.Add(id);
                }
                db.Purge(layerIds);
                foreach (ObjectId id in layerIds)
                {
                    DBObject obj = Tx.GetObject(id, OpenMode.ForWrite);
                    obj.Erase();
                }

                Tx.Commit();
            }

        }
        [CommandMethod("TEST")]
        public void Test()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                // if the bloc table has the block definition
                if (bt.Has("prz_podl"))
                {
                    // create a new block reference
                    var br = new BlockReference(Point3d.Origin, bt["prz_podl"]);

                    // add the block reference to the curentSpace and the transaction
                    var curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    curSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // set the dynamic property value
                    foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
                    {
                        if (prop.PropertyName == "par_l")
                        {
                            prop.Value = 500.0;
                        }
                    }
                }
                // save changes
                tr.Commit();
            } // <- end using: disposing the transaction and all objects opened with it (block table) or added to it (block reference)
        }
        [CommandMethod("CopyAndBindXR")]
        public void CopyAndBindXr()
        {
            FolderBrowserDialog source = new FolderBrowserDialog();
            source.Description = "Select Folder to Copy";
            DialogResult result_source = source.ShowDialog();

            FolderBrowserDialog destination = new FolderBrowserDialog();
            DialogResult result_des = destination.ShowDialog();
            Helper.Copy(source.SelectedPath, destination.SelectedPath);
            DirectoryInfo dirInfo = new DirectoryInfo(destination.SelectedPath);
            Document doc =  AcadApp.DocumentManager.MdiActiveDocument;
           
            Editor ed = doc.Editor;
            foreach (var file in dirInfo.GetFiles())
            {
                if (file.Extension != ".dwg")
                {
                    continue;
                }
                // Create a database and try to load the file

                BindXRefs(file.FullName, file.FullName);
            }
        }
        private void BindXRefs(string strFileIn, string strFileOut)
        {
            Database db = new Database(false, true);
            using (db)
            {
                db.ReadDwgFile(strFileIn, System.IO.FileShare.ReadWrite, false, "");
                db.ResolveXrefs(true, false);

                using (ObjectIdCollection xrefCollection = new ObjectIdCollection())
                {
                    using (XrefGraph xg = db.GetHostDwgXrefGraph(false))
                    {
                        int numOfNodes = xg.NumNodes;
                        for (int cnt = 0; cnt < xg.NumNodes; cnt++)
                        {
                            XrefGraphNode xrefNode = xg.GetXrefNode(cnt) as XrefGraphNode;

                            if (xrefNode.Database != null)
                            {
                                if (!xrefNode.Database.Filename.Equals(db.Filename))
                                {
                                    if (xrefNode.XrefStatus == XrefStatus.Resolved && xrefNode.BlockTableRecordId.IsValid)
                                    {
                                        xrefCollection.Add(xrefNode.BlockTableRecordId);
                                    }
                                }
                            }
                        }
                    }
                    try
                    {
                        db.BindXrefs(xrefCollection, true);

                    }
                    catch
                    {

                    }
                }

                db.SaveAs(strFileOut, true, db.OriginalFileVersion, db.SecurityParameters);
            }
        }
        [CommandMethod("LL")]

        public void LoadLinetypes()

        {

            Document doc =

              AcadApp.DocumentManager.MdiActiveDocument;

            Database db = doc.Database;

            Editor ed = doc.Editor;



            const string filename = "acad.lin";

            try

            {

                string path =

                  HostApplicationServices.Current.FindFile(

                    filename, db, FindFileHint.Default

                  );

                db.LoadLineTypeFile("*", path);

            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)

            {

                if (ex.ErrorStatus == ErrorStatus.FilerError)

                    ed.WriteMessage(

                      "\nCould not find file \"{0}\".",

                      filename

                    );

                else if (ex.ErrorStatus == ErrorStatus.DuplicateRecordName)

                    ed.WriteMessage(

                      "\nCannot load already defined linetypes."

                    );

                else

                    ed.WriteMessage(

                      "\nException: {0}", ex.Message

                    );

            }

        }

    }
}

