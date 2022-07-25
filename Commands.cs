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

        private static int SpecSheetCount = 0;

        public static void GoogleDrive()
        {
            if (dialogs.GetValue(REG_KEY_NAME) == null)
                return;

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

            var directory = $@"{path}\{pathPostfix}";
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var count = 0;
            foreach (var key in dialogs.GetValueNames())
            {
                // It's already in there, overwrite it
                if (dialogs.GetValue(key).ToString() == REG_KEY_NAME) break;

                if (rgx.IsMatch(key)) count++;
            }

            doc.Editor.WriteMessage("Values: {}", string.Join(", ", dialogs.GetValueNames()));

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

                AddGoogleDrive();
            }
        }

        private static int PDFPageCount(string fileName)
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(fileName)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }

        private static void _Import(string path)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            for (int i = 0; i < PDFPageCount(path); i++)
            {
                doc.Editor.Command($"-PDFATTACH", path, 1, new Point2d(SpecSheetCount * 11, 0), 1, 0);
                SpecSheetCount++;
            }
        }

        private static void Import(string path)
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
                _Import(path);
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

                var openFileDialog = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select Spec Sheet(s)", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "*", "Select Spec Sheet(s)", flags);

                var result = openFileDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (var file in openFileDialog.GetFilenames())
                    {
                        Import(file);
                    }
                }
            }
        }
    }
}
