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
using iText.Kernel.Geom;
using System.Collections.Generic;

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

        private static readonly double X_OFFSET = 0;
        private static readonly double Y_OFFSET = 0;
        private static readonly int SHEETS_PER_LINE = 4;
        private static readonly double PIXELS_PER_INCH = 72;

        private static List<string> Import(string path, double scale)
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                var pdfs = new List<string>();
                foreach (var file in Directory.GetFiles(path))
                {
                    pdfs.AddRange(Import(file, scale));
                }

                return pdfs;
            }
            else
            {
                return new List<string> { path };
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

                var noMutt = Application.GetSystemVariable("NOMUTT");
                Application.SetSystemVariable("NOMUTT", 1);

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
                    var pdfs = new List<string>();
                    var files = openFileDialog.GetFilenames();

                    foreach (var file in files)
                    {
                        pdfs.AddRange(Import(file, scale));
                    }

                    var pages = new List<(string, int)>();
                    foreach (var pdf in pdfs)
                    {
                        PdfReader reader = new PdfReader(pdf);
                        PdfDocument document = new PdfDocument(reader);
                        var pageCount = document.GetNumberOfPages();

                        for (var page = 0; page < pageCount; page++)
                        {
                            pages.Add((pdf, page + 1));
                        }
                    }

                    var y_offsets = new List<double>();
                    for (var i = 0; i < pages.Count; i++)
                    {
                        var x_offset = 0.0;

                        for (var j = 0; j < SHEETS_PER_LINE; j++)
                        {
                            if ((i * SHEETS_PER_LINE + j) >= pages.Count())
                            {
                                break;
                            }

                            var (pdf, page) = pages[i * SHEETS_PER_LINE + j];
                            PdfReader reader = new PdfReader(pdf);
                            PdfDocument document = new PdfDocument(reader);

                            Rectangle rectangle = document.GetPage(page).GetPageSize();

                            if (y_offsets.Count <= j)
                            {
                                y_offsets.Add(0.0);
                            }
                            else
                            {
                                y_offsets[j] += rectangle.GetHeight() / PIXELS_PER_INCH * scale;
                            }

                            doc.Editor.Command(
                                $"-PDFATTACH",
                                pdf,
                                page,
                                new Point2d(
                                    X_OFFSET + x_offset,
                                    Y_OFFSET - y_offsets[j]
                                ),
                                scale,
                                0
                            );

                            x_offset += rectangle.GetWidth() / PIXELS_PER_INCH * scale;
                        }
                    }
                }

                Application.SetSystemVariable("NOMUTT", noMutt);
            }
        }

        private static ObjectIdCollection GetPolylineEntities(string layer = null)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            TypedValue[] filter;

            if (!String.IsNullOrEmpty(layer))
            {
                filter = new TypedValue[]{
                    new TypedValue((int)DxfCode.Operator, "<and"),
                    new TypedValue((int)DxfCode.Start, "LWPolyline"),
                    new TypedValue((int)DxfCode.LayerName, layer),
                    new TypedValue((int)DxfCode.Operator, "and>")
                };
            }
            else
            {
                filter = new TypedValue[] { new TypedValue((int)DxfCode.Start, "LWPolyline") };
            }

            var selectionFilter = new SelectionFilter(filter);
            var promptStatusResult = ed.SelectAll(selectionFilter);

            if (promptStatusResult.Status == PromptStatus.OK)
                return new ObjectIdCollection(promptStatusResult.Value.GetObjectIds());

            return new ObjectIdCollection();
        }


        [CommandMethod("Ecogy", "FilletAll", CommandFlags.Modal)]
        public void FilletAll()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var radiusPrompt = new PromptDoubleOptions("\nWhat radius? ");
            var radiusResponse = ed.GetDouble(radiusPrompt);

            var layerPrompt = new PromptKeywordOptions("\nOn what Layer? ");
            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                foreach (ObjectId layerId in lt)
                {
                    var layer = tr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                    layerPrompt.Keywords.Add(layer.Name);
                }
            }
            layerPrompt.Keywords.Default = "PV-string";

            var lay = ed.GetKeywords(layerPrompt);

            if (radiusResponse.Status != PromptStatus.OK)
            {
                ed.WriteMessage("You must supply a valid radius\n");
                return;
            }
            else if (lay.Status != PromptStatus.OK)
            {
                ed.WriteMessage("You must select one of the layers\n");
                return;
            }

            var noMutt = Application.GetSystemVariable("NOMUTT");
            Application.SetSystemVariable("NOMUTT", 1);

            ed.Command("FILLETRAD", radiusResponse.Value);

            using (var pm = new ProgressMeter())
            {
                pm.Start("Filleting Polylines");

                var collection = GetPolylineEntities(lay.StringResult);
                foreach (ObjectId id in collection)
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
