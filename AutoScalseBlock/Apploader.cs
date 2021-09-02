using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadService = Autodesk.AutoCAD.ApplicationServices;

namespace AutoScalseBlock
{
    public class AppLoader: IExtensionApplication
    {
        static List<Document> RegistedDocuments = new List<Document>();
        #region Ctor/ Destor
        Assembly ResolveHandler(Object Sender, ResolveEventArgs e)
        {            
            return null;
        }
        public AppLoader()
        {
        }

        public void Initialize()
        {
            try
            {
                AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(ResolveHandler);
                
                DocumentManagerReactorRegister();

                Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                DocumentReactorRegister(doc);
            }
            catch (System.Exception ex)
            {
               
            }
        }

        private void callback_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            Document doc = e.Document;
            DocumentReactorRegister(doc);
        }

        public void Terminate()
        {
           
        }

        void OnIdle(object sender, EventArgs e)
        {
            //bFilterProcessing = false;
        }
        #endregion
        #region Reactor Register
        public void DocumentManagerReactorRegister()
        {
            AcadApp.DocumentManager.DocumentCreated += new DocumentCollectionEventHandler(callback_DocumentCreated);
        }

        private void DocumentReactorRegister(Document doc)
        {
            if (RegistedDocuments != null && doc != null && RegistedDocuments.IndexOf(doc) == -1)
            {
                RegistedDocuments.Add(doc);
                doc.CommandWillStart += new CommandEventHandler(WillStartCmd);
                doc.CommandEnded += new CommandEventHandler(EndCommand);       
                doc.Database.ObjectAppended += Database_ObjectAppended;               
            }
        }
        public static double factor = 1;
        public void EndCommand(object sender, CommandEventArgs e)
        {
            
            Document acDoc = AcadService.Application.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;
            
            if (AppendObj.Count > 0)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                   
                    foreach (ObjectId id in AppendObj)
                    {
                        if (id.IsErased == false && id.IsValid == true)
                        {
                            BlockReference blk = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;
                            blk.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(factor);
                        }
                    }
                    tr.Commit();
                }

            }
            AppendObj.Clear();
        }

        public void WillStartCmd(object sender, CommandEventArgs e)
        { //opening the subkey  
            //RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\OurSettings");
            //if (key != null)
            //{
            //    factor = (double)key.GetValue("Factor");
            //}
            AppendObj.Clear();
        }
        #endregion
        #region Database Reactor
        ObjectIdCollection AppendObj = new ObjectIdCollection();
        public void Database_ObjectAppended(object sender, ObjectEventArgs e)
        {
            if (e.DBObject is BlockReference)
            {
                AppendObj.Add(e.DBObject.ObjectId);
            }
        }
        #endregion
    }
}
