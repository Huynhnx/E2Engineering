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
using System.IO;

namespace AutoScalseBlock
{
    public enum BlockRefType
    {
        None = 0,
        Xref,
        Block,
    };
    public class Helper
    {
        public static BlockRefType GetBlockReferenceType(ObjectId objId)
        {
            bool bIsBlock = false;
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Entity acadEnt = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                if (acadEnt != null)
                {
                    if (acadEnt is BlockReference)
                    {
                        BlockReference pBlockRef = acadEnt as BlockReference;
                        if (pBlockRef != null)
                        {
                            // GET THE BLOCK TABLE RECORD ID OF THE BLOCK REFERENCE
                            ObjectId btrcd_id_inside = pBlockRef.BlockTableRecord;
                            if (btrcd_id_inside.IsValid == true && btrcd_id_inside.IsErased == false)
                            {
                                BlockTableRecord btrcd_inside = tr.GetObject(btrcd_id_inside, OpenMode.ForRead) as BlockTableRecord;
                                if (btrcd_inside != null)
                                {
                                    // CHECK IF THIS BLOCK TABLE IS FROM EXTERNAL REF
                                    if (btrcd_inside.IsFromExternalReference)
                                    {
                                        return BlockRefType.Xref;
                                    }
                                    else
                                    {
                                        return BlockRefType.Block;
                                    }
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }
            return BlockRefType.None;
        }
        public static void Copy(string sourceDirectory, string targetDirectory)
        {
            var diSource = new DirectoryInfo(sourceDirectory);
            var diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
    }
}
