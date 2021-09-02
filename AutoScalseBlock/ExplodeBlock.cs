using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.Runtime;



namespace AutoScalseBlock

{

    public class Commands

    {

        [CommandMethod("EB")]

        public void ExplodeBock()

        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection ids = new ObjectIdCollection();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
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
                        ids.Add(objId);
                    }

                }
                tr.Commit();
            }


            using (var tr = db.TransactionManager.StartTransaction())

            {

                // Call our explode function recursively, starting

                // with the top-level block reference

                // (you can pass false as a 4th parameter if you

                // don't want originating entities erased)
                ProgressMeter pm = new ProgressMeter();

                pm.Start("Explode Block...");

                pm.SetLimit(ids.Count);
                foreach (ObjectId id in ids)
                {
                    ExplodeBlock(tr, db, id);
                    pm.MeterProgress();
                }
                pm.Stop();
                tr.Commit();

            }

        }



        private void ExplodeBlock(Transaction tr, Database db, ObjectId id, bool erase = true)

        {
            DBObjectCollection toAddColl = new DBObjectCollection();
            // Open out block reference - only needs to be readable
            // for the explode operation, as it's non-destructive
            if (id.IsErased == true||id.IsNull)
            {
                return;
            }

            if (id.ObjectClass.Name == "AcDbAttributeDefinition")
            {
                AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                //if ((attDef.Constant && !attDef.Invisible))
                {
                    DBText text = new DBText();
                    text.Height = attDef.Height;
                    text.TextString = attDef.TextString;
                    text.Position = attDef.Position;
                    toAddColl.Add(text);
                }
            }
            // Add the entities to modelspace 
            foreach (Entity ent in toAddColl)
            {
                // open model space block table record
                BlockTableRecord spaceBlkTblRec = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                // append entity to model space block table record
                spaceBlkTblRec.AppendEntity(ent);
                tr.AddNewlyCreatedDBObject(ent, true);
            }
            var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

            // We'll collect the BlockReferences created in a collection
            var toExplode = new ObjectIdCollection();
            // Define our handler to capture the nested block references
            ObjectEventHandler handler =
              (s, e) =>
              {
                  if (e.DBObject is BlockReference)
                  {
                      toExplode.Add(e.DBObject.ObjectId);
                  }

              };
            // Add our handler around the explode call, removing it
            // directly afterwards
            db.ObjectAppended += handler;

            br.ExplodeToOwnerSpace();

            db.ObjectAppended -= handler;
            // Go through the results and recurse, exploding the
            // contents
            foreach (ObjectId bid in toExplode)
            {
                ExplodeBlock(tr, db, bid, erase);
            }
            // We might also just let it drop out of scope
            toExplode.Clear();
            // To replicate the explode command, we're delete the
            // original entity
            if (erase)
            {
                br.UpgradeOpen();
                br.Erase();
                br.DowngradeOpen();
            }
        }
    }
}