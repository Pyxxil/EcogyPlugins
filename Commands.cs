using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;

using System.Text.RegularExpressions;
using System.Linq;
using System.IO;
using System;
using Autodesk.AutoCAD.Geometry;

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
                if (dialogs.GetValue(key).ToString() == REG_KEY_NAME) break;
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

                string path;
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    path = ofd.Filename;
                }
                else
                {
                    return;
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

                GoogleDrive();
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

        private static bool Clockwise(Point2d p1, Point2d p2, Point2d p3)
        {
            return ((p2.X - p1.X) * (p3.Y - p1.Y) - (p2.Y - p1.Y) * (p3.X - p1.X)) < 1e-8;
        }

        public static int Fillet(Polyline line, int index, double radius)
        {
            int prev = index == 0 && line.Closed ? line.NumberOfVertices - 1 : index - 1;
            if (line.GetSegmentType(prev) != SegmentType.Line ||
                line.GetSegmentType(index) != SegmentType.Line)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("One or both are not lines\n");
                return 0;
            }

            LineSegment2d seg1 = line.GetLineSegment2dAt(prev);
            LineSegment2d seg2 = line.GetLineSegment2dAt(index);
            Vector2d vec1 = seg1.StartPoint - seg1.EndPoint;
            Vector2d vec2 = seg2.EndPoint - seg2.StartPoint;

            double angle = (Math.PI - vec1.GetAngleTo(vec2)) / 2.0;
            double dist = radius * Math.Tan(angle);
            if (dist == 0.0 || dist > seg1.Length || dist > seg2.Length)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Distance issue\n");
                return 0;
            }

            Point2d pt1 = seg1.EndPoint + vec1.GetNormal() * dist;
            Point2d pt2 = seg2.StartPoint + vec2.GetNormal() * dist;
            double bulge = Math.Tan(angle / 2.0);

            if (Clockwise(seg1.StartPoint, seg1.EndPoint, seg2.EndPoint))
            {
                bulge = -bulge;
            }

            line.AddVertexAt(index, pt1, bulge, 0.0, 0.0);
            line.SetPointAt(index + 1, pt2);

            return 1;
        }

        public static void FilletAll(Polyline pline, double radius)
        {
            int n = pline.Closed ? 0 : 1;
            for (int i = n; i < pline.NumberOfVertices - n; i += 1 + Fillet(pline, i, radius))
            { }
        }

        [CommandMethod("Ecogy", "FilletAll", CommandFlags.Modal)]
        [Obsolete]
        public void FilletAll()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;

            if (doc != null)
            {
                var editor = doc.Editor;

                var valueSelector = new TypedValue[] {
                        new TypedValue(Convert.ToInt32(DxfCode.Operator), "<and"),
                        // new TypedValue(Convert.ToInt32(DxfCode.LayerName), layer.Name),
                        new TypedValue(Convert.ToInt32(DxfCode.Operator), "<or"),
                        new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE"),
                        // new TypedValue(Convert.ToInt32(DxfCode.Start), "LWPOLYLINE"),
                        new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE2D"),
                        // new TypedValue(Convert.ToInt32(DxfCode.Start), "POLYLINE3d"),
                        new TypedValue(Convert.ToInt32(DxfCode.Operator), "or>"),
                        new TypedValue(Convert.ToInt32(DxfCode.Operator), "and>")
                    };

                SelectionFilter selectionFilter = new SelectionFilter(valueSelector);

                var selection = editor.SelectAll(selectionFilter);

                ObjectIdCollection polylineIDCollection;
                if (selection.Status == PromptStatus.OK)
                {
                    polylineIDCollection = new ObjectIdCollection(selection.Value.GetObjectIds());
                }
                else
                {
                    polylineIDCollection = new ObjectIdCollection();
                }

                using (var transcation = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in polylineIDCollection)
                    {
                        using (var obj = id.Open(OpenMode.ForRead))
                        {
                            if (obj is Polyline line)
                            {
                                FilletAll(line, 0.2);
                            }
                            else if (obj is Polyline2d poly2d)
                            {
                                poly2d.UpgradeOpen();

                                using (var poly = new Polyline())
                                {
                                    poly.ConvertFrom(poly2d, true);
                                    FilletAll(poly, 0.2);
                                }
                            }
                        }
                    }

                    transcation.Commit();
                }
            }
        }
    }
}
