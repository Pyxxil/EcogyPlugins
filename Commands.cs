// (C) Copyright 2022 by  
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;

using System.Text.RegularExpressions;
using System.Linq;
using System.IO;
using System;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(Ecogy.Commands))]

namespace Ecogy
{
    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class Commands
    {
        private static readonly string REG_KEY_NAME = "Google Drive";
        private static readonly string REG_KEY_DEPTH = "Google Drive Depth";

        private static readonly RegistryKey dialogs = Registry.CurrentUser.OpenSubKey(
            $@"{HostApplicationServices.Current.UserRegistryProductRootKey}\Profiles\{Application.GetSystemVariable("CPROFILE")}\Dialogs\AllAnavDialogs"
        , true);

        private static readonly Regex rgx = new Regex(@"PlacesOrder\d$", RegexOptions.Compiled);

        public static void GoogleDrive()
        {
            var path = dialogs.GetValue(REG_KEY_NAME).ToString();
            var depth = (int)dialogs.GetValue(REG_KEY_DEPTH);

            Document doc = Application.DocumentManager.MdiActiveDocument;
            var documentPath = doc.Database.OriginalFileName;
            var pathPostfix = "";

            for (int i = 0; i < depth; i++)
            {
                documentPath = Directory.GetParent(documentPath).ToString();
                pathPostfix = new DirectoryInfo(documentPath).Name + @"\" + pathPostfix;
            }

            string directory = $@"{path}\{pathPostfix}";
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var count = 0;
            foreach (var key in dialogs.GetValueNames())
            {
                if (rgx.IsMatch(key)) count++;

                // It's already in there, overwrite it
                if (dialogs.GetValue(key).ToString() == path) break;
            }

            // AutoCad occasionally puts an empty key in
            if (rgx.IsMatch($"PlacesOrder{count}") && !dialogs.GetValueNames().Contains($"PlacesOrder{count}Display")) count--;

            dialogs.SetValue($"PlacesOrder{count}", directory);
            dialogs.SetValue($"PlacesOrder{count}Display", REG_KEY_NAME);
            dialogs.SetValue($"PlacesOrder{count}Ext", "");
        }

        public List<string> GetPlaces()
        {
            var places = new List<string>();

            foreach (var key in dialogs.GetValueNames())
            {
                places.Add(key);
            }

            return places;
        }

        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        [CommandMethod("Ecogy", "AddGoogleDrive", CommandFlags.Modal)]
        public void AddGoogleDrive() // This method can have any name
        {
            // Put your command code here
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var ed = doc.Editor;

                var pathPrompt = new PromptStringOptions("\nWhat's the google drive path? ")
                {
                    AllowSpaces = true
                };

                var pathResponse = ed.GetString(pathPrompt);
                var path = pathResponse.StringResult;

                var attrs = File.GetAttributes(path);
                switch (attrs)
                {
                    case FileAttributes.Directory:
                        if (!Directory.Exists(path))
                        {
                            ed.WriteMessage($"Invalid Path {path}");
                            return;
                        }
                        break;
                    default:
                        if (!File.Exists(path))
                        {
                            ed.WriteMessage($"Invalid Path {path}");
                            return;
                        }
                        break;
                }

                dialogs.SetValue(REG_KEY_NAME, path);

                var depthPrompt = new PromptStringOptions("\nAt what depth? ")
                {
                    AllowSpaces = false
                };

                var depthResponse = ed.GetString(depthPrompt);
                if (!int.TryParse(depthResponse.StringResult, out int depth))
                {
                    ed.WriteMessage($"Invalid number: {depthResponse}");
                    return;
                }

                dialogs.SetValue(REG_KEY_DEPTH, depth);

                AddGoogleDrive();
            }
        }

        static void Import(string path)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
                {
                    foreach (var file in Directory.GetFiles(path))
                    {
                        Import(file);
                    }
                }
                else
                {
                    var db = doc.Database;
                    using (var transaction = doc.Database.TransactionManager.StartTransaction())
                    {
                        DBDictionary nod = (DBDictionary)transaction.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

                        string defDictKey = UnderlayDefinition.GetDictionaryKey(typeof(PdfDefinition));
                        if (!nod.Contains(defDictKey))
                        {
                            using (DBDictionary dict = new DBDictionary())
                            {
                                nod.SetAt(defDictKey, dict);
                                transaction.AddNewlyCreatedDBObject(dict, true);
                            }
                        }

                        ObjectId idPdfDef;
                        DBDictionary pdfDict = (DBDictionary)transaction.GetObject(nod.GetAt(defDictKey), OpenMode.ForWrite);

                        using (PdfDefinition pdfDef = new PdfDefinition())
                        {
                            pdfDef.SourceFileName = path;
                            idPdfDef = pdfDict.SetAt(Path.GetFileNameWithoutExtension(path), pdfDef);
                            transaction.AddNewlyCreatedDBObject(pdfDef, true);
                        }

                        BlockTable bt = (BlockTable)transaction.GetObject(db.BlockTableId, OpenMode.ForRead);

                        BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(
                            bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite
                        );

                        using (PdfReference pdf = new PdfReference())
                        {
                            pdf.DefinitionId = idPdfDef;
                            btr.AppendEntity(pdf);
                            transaction.AddNewlyCreatedDBObject(pdf, true);
                        }

                        transaction.Commit();
                    }
                }
            }
        }

        // Modal Command with localized name
        [CommandMethod("Ecogy", "ImportSpecSheet", CommandFlags.Modal)]
        public void ImportSpecSheet()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var flags = Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.ForceDefaultFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowAnyExtension;

                var ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select Spec Sheet(s)", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "*", "Select Spec Sheet(s)", flags);

                var dr = ofd.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (var file in ofd.GetFilenames())
                    {
                        Import(file);
                    }
                }
            }
        }
    }
}
