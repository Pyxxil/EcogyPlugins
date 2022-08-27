using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using System.Linq;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
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

        private static readonly Regex rgx = new Regex(@"PlacesOrder(\d+)$", RegexOptions.Compiled);
        private static readonly Regex deleteRegex = new Regex(@"^PlacesOrder(\d+)", RegexOptions.Compiled);

        private static int SpecSheetCount = 0;

        public static void GoogleDrive()
        {
            if (dialogs.GetValue(REG_KEY_NAME) == null)
                return;

            var path = dialogs.GetValue(REG_KEY_NAME).ToString();
            var depth = (int)dialogs.GetValue(REG_KEY_DEPTH);

            var doc = Application.DocumentManager.MdiActiveDocument;
            var documentPath = doc.Database.OriginalFileName;
            var pathPostfix = "";

            for (int i = 0; i < depth; i++)
            {
                documentPath = Directory.GetParent(documentPath).ToString();
                pathPostfix = new DirectoryInfo(documentPath).Name + @"\" + pathPostfix;
            }

            var directory = $@"{path}\{pathPostfix}";
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var pos = 0;

            var entries = dialogs
                .GetValueNames()
                .AsEnumerable()
                .Where(value => rgx.IsMatch(value) && dialogs.GetValue(value).ToString().Length != 0)
                .Select(value =>
                {
                    int position = int.Parse(rgx.Match(value).Groups[1].ToString());
                    return (
                        position: pos++,
                        val: dialogs.GetValue($"PlacesOrder{position}").ToString(),
                        display: (dialogs.GetValue($"PlacesOrder{position}Display") ?? "").ToString(),
                        ext: (dialogs.GetValue($"PlacesOrder{position}Ext") ?? "").ToString()
                    );
                })
                .OrderBy(value => value.position)
                .ToList();

            foreach (var entry in dialogs.GetValueNames().AsEnumerable().Where(entry => deleteRegex.IsMatch(entry)))
            {
                dialogs.DeleteValue(entry);
            }

            foreach (var (position, val, display, ext) in entries)
            {
                dialogs.SetValue($"PlacesOrder{position}", val);
                dialogs.SetValue($"PlacesOrder{position}Display", display);
                dialogs.SetValue($"PlacesOrder{position}Ext", ext);
            }

            var match = entries.FindIndex(value => value.display == REG_KEY_NAME);

            var idx = match == -1 ? entries.Count : entries[match].position;
            var end = match == -1 ? entries.Count + 1 : entries.Count;

            dialogs.SetValue($"PlacesOrder{idx}", directory);
            dialogs.SetValue($"PlacesOrder{idx}Display", REG_KEY_NAME);
            dialogs.SetValue($"PlacesOrder{idx}Ext", "");

            dialogs.SetValue($"PlacesOrder{end}", "");
        }

        [CommandMethod("Ecogy", "AddGoogleDrive", CommandFlags.Modal)]
        public void AddGoogleDrive()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var ed = doc.Editor;

                var flags = Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.ForceDefaultFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowFoldersOnly;

                var ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select Google Drive Path", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "*", "Select Spec Sheet(s)", flags);
                var dr = ofd.ShowDialog();

                if (dr != System.Windows.Forms.DialogResult.OK)
                {
                    return;
                }

                var path = ofd.Filename;

                dialogs.SetValue(REG_KEY_NAME, path);

                var depthPrompt = new PromptIntegerOptions("\nAt what depth? ");

                var depthResponse = ed.GetInteger(depthPrompt);
                var depth = depthResponse.Value;

                dialogs.SetValue(REG_KEY_DEPTH, depth);

                GoogleDrive();
            }
        }

        private static int PDFPageCount(string fileName)
        {
            PdfReader reader = new PdfReader(fileName);
            PdfDocument document = new PdfDocument(reader);
            return document.GetNumberOfPages();
        }

        private static readonly double PDF_WIDTH = 8.2677;
        private static readonly double PDF_HEIGHT = 11.6929;
        private static readonly int SHEETS_PER_LINE = 4;

        private static int Min(int a, int b)
        {
            return a <= b ? a : b;
        }

        private static void DoImport(string path, double scale)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            for (int i = 0; i < Min(PDFPageCount(path), 2); i++)
            {
                doc.Editor.Command(
                    $"-PDFATTACH",
                    path,
                    i + 1,
                    new Point2d((SpecSheetCount % SHEETS_PER_LINE) * PDF_WIDTH * scale, (SpecSheetCount / SHEETS_PER_LINE) * -PDF_HEIGHT * scale),
                    scale,
                    0
                );

                SpecSheetCount++;
            }
        }

        private static void Import(string path, double scale)
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                foreach (var file in Directory.GetFiles(path))
                {
                    Import(file, scale);
                }
            }
            else
            {
                DoImport(path, scale);
            }
        }

        [CommandMethod("Ecogy", "ImportSpecSheet", CommandFlags.Modal)]
        public void ImportSpecSheet()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var ed = doc.Editor;
                var scalePrompt = new PromptDoubleOptions("\nAt what scale? ");

                var scaleResponse = ed.GetDouble(scalePrompt);

                if (scaleResponse.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("You must supply a valid scale\n");
                    return;
                }

                var scale = scaleResponse.Value;

                var flags = Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.ForceDefaultFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowAnyExtension;

                var openFileDialog = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select Spec Sheet(s)", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "*", "Select Spec Sheet(s)", flags);

                var result = openFileDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    using (var pm = new ProgressMeter())
                    {
                        var files = openFileDialog.GetFilenames();
                        pm.Start("Filleting Polylines");
                        pm.SetLimit(files.Count());

                        foreach (var file in files)
                        {
                            Import(file, scale);

                            pm.MeterProgress();
                            System.Windows.Forms.Application.DoEvents();
                        }

                        pm.Stop();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
        }

        private static ObjectIdCollection GetPolylineEntities(string layerName = null)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            TypedValue[] filter;

            if (!String.IsNullOrEmpty(layerName))
            {
                filter = new TypedValue[]{
                                   new TypedValue((int)DxfCode.Operator,"<and"),
                                   new TypedValue((int)DxfCode.LayerName,layerName),
                                   new TypedValue((int)DxfCode.Start,"LWPolyline"),
                                   new TypedValue((int)DxfCode.Operator,"and>")
                };
            }
            else
            {
                filter = new TypedValue[] { new TypedValue((int)DxfCode.Start, "LWPolyline") };
            }

            // Build a filter list so that only entities
            // on the specified layer are selected

            var selectionFilter = new SelectionFilter(filter);
            var promptStatusResult = ed.SelectAll(selectionFilter);

            if (promptStatusResult.Status == PromptStatus.OK)
                return
                  new ObjectIdCollection(
                    promptStatusResult.Value.GetObjectIds()
                  );
            else
                return new ObjectIdCollection();
        }


        [CommandMethod("Ecogy", "FilletAll", CommandFlags.Modal)]
        public void FilletAll()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var plineIds = GetPolylineEntities();

            var radiusPrompt = new PromptDoubleOptions("\nWhat radius? ");
            var radiusResponse = ed.GetDouble(radiusPrompt);

            if (radiusResponse.Status != PromptStatus.OK)
            {
                ed.WriteMessage("You must supply a valid radius\n");
                return;
            }

            // Suppress Command Line noise.
            var noMutt = Application.GetSystemVariable("NOMUTT");
            Application.SetSystemVariable("NOMUTT", 1);

            ed.Command("FILLETRAD", 0.5);

            var db = Application.DocumentManager.MdiActiveDocument.Database;

            using (var pm = new ProgressMeter())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                pm.Start("Filleting Polylines");
                pm.SetLimit(plineIds.Count);

                foreach (ObjectId id in plineIds)
                {
                    ed.Command("_.FILLET", "_P", id);
                    pm.MeterProgress();
                    System.Windows.Forms.Application.DoEvents();
                }

                pm.Stop();
                System.Windows.Forms.Application.DoEvents();
            }

            Application.SetSystemVariable("NOMUTT", noMutt);
        }
    }
}
