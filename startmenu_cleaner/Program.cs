using System;
using System.IO;
using IWshRuntimeLibrary;
using File = System.IO.File;
using System.Linq;
using System.Collections.Generic;

namespace startmenu_cleaner {
    class Program {
        [MTAThread]
        static void Main (string[ ] args) {
            string[ ] startmenuLocations = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) ,@"Microsoft\Windows\Start Menu\Programs"),
                @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
            };
            Console.WriteLine(Directory.GetParent(@"C:\Users\Tim\Desktop").FullName);
            bool argsCleanup = args.Contains("-c"), argsRemoveUninstall = args.Contains("-u"), argsRemoveBroken = args.Contains("-b");
            if (argsCleanup) {
                foreach (string location in startmenuLocations) {
                    Cleanup(location, argsRemoveUninstall, true);
                }
            }
            if (argsRemoveBroken) {

            }
            Console.ReadLine( );
        }

        [MTAThread]
        static string GetTargetPath (string lnk) {
            if (!File.Exists(lnk)) return null;
            WshShell shell = new WshShell( );
            IWshShortcut link = (IWshShortcut)shell.CreateShortcut(lnk); //Link the interface to our shortcut
            return link.TargetPath;
        }

        static bool IsLink (string file) {
            return Path.GetExtension(file) == ".lnk";
        }

        static void Cleanup (string dir, bool removeUninstall, bool isRoot) {
            if (!Directory.Exists(dir)) return;
            foreach (string subdir in Directory.GetDirectories(dir)) {
                Cleanup(subdir, removeUninstall, false);
            }
            List<string> files = Directory.GetFiles(dir).ToList( );
            if (removeUninstall) {
                for (int i = files.Count - 1; i > -1; i--) {
                    if (Path.GetFileNameWithoutExtension(files[i]).StartsWith("Uninstall")) {
                        File.Delete(files[i]);
                        files.RemoveAt(i);
                    }
                }
            }
            if (files.Count <= 1 && !isRoot) {
                try {
                    string root = Directory.GetParent(dir.EndsWith("\\") ? dir.Substring(0, dir.Length - 1) : dir).FullName;
                    foreach (string file in files) {
                        if (IsLink(file)) {
                            string dest = Path.Combine(root, Path.GetFileName(file));
                            if (File.Exists(dest)) File.Delete(dest);
                            File.Move(file, dest);
                        }
                    }
                    Directory.Delete(dir);
                } catch(Exception e) {
                    Console.WriteLine("Error cleaning up " + dir + ": " + e.Message);
                }
            }
        }

        static void RemoveBroken(string dir) {
            foreach(string subdir in Directory.GetDirectories(dir)) {
                RemoveBroken(dir);
            }
            foreach(string file in Directory.GetFiles(dir)) {
                if (IsLink(file)) {
                    string target = GetTargetPath(file);
                    if(target != null && !File.Exists(target)) {
                        File.Delete(file);
                    }
                }
            }
        }
    }
}
