using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;

namespace NeuExt.PTCGLUtility
{
    internal class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            var action = "CheckPTCGLIsRunning";
#else
            if (args.Length <= 0)
            {
                Environment.Exit(1);
                return;
            }
            var action = args[0];
#endif
            var resets = args.Skip(1).ToArray();
            switch (action.TrimStart('-'))
            {
                case "CheckPTCGLIsRunning":
                    Output(CheckPTCGLIsRunning() ? "1" : "0");
                    break;
                case "CreateDesktopShortcut":
                    if (resets.Length < 2)
                    {
                        Environment.Exit(1);
                        break;
                    }
                    CreateDesktopShortcut(resets[0], resets[1]);
                    break;
                case "DetectPTCGLInstallDirectory":
                    Output(DetectPTCGLInstallDirectory());
                    break;
                case "GetShortcutTarget":
                    if (resets.Length < 1)
                    {
                        Environment.Exit(1);
                        break;
                    }
                    Output(GetShortcutTarget(resets[0]));
                    break;
                case "help":
                    Console.WriteLine("Usage: NeuExt.PTCGLUtility.exe <action>");
                    break;
                case "version":
                    Output("v1.0.0");
                    break;
                default:
                    Environment.Exit(1);
                    break;
            }
#if DEBUG
            Console.ReadLine();
#endif
        }

        static string EncodeURIComponent(string str) => WebUtility.UrlEncode(str).Replace("+", "%20");

        static void Output(string str) => Console.WriteLine(EncodeURIComponent(str));

        static bool CheckPTCGLIsRunning()
        {
            var ps = Process.GetProcessesByName("Pokemon TCG Live");
            return ps.Length > 0;
        }


        static void CreateDesktopShortcut(string name, string target)
        {
            var lnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{name}.lnk");
            if (File.Exists(lnk))
            {
                File.Delete(lnk);
            }
            var shell = new IWshRuntimeLibrary.WshShell();
            var shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(lnk);
            shortcut.Description = name;
            shortcut.TargetPath = target;
            shortcut.WorkingDirectory = Path.GetDirectoryName(target);
            shortcut.Save();
        }

        static string DetectPTCGLInstallDirectory()
        {
            var desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Pokémon TCG Live.lnk");
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "The Pokémon Company International");

            var reg = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64).OpenSubKey(@"Software\The Pokémon Company International\Pokémon Trading Card Game Live");

            if (reg != null)
            {
                path = reg.GetValue("Path")?.ToString() ?? path;
            }

            if (File.Exists(desktopLnk))
            {
                path = Path.GetDirectoryName(GetShortcutTarget(desktopLnk));
            }

            if (!File.Exists(Path.Combine(path, "Pokemon TCG Live.exe")))
            {
                path = Path.Combine(path, "Pokémon Trading Card Game Live");
            }

            return File.Exists(Path.Combine(path, "Pokemon TCG Live.exe")) ? path : "";
        }

        static string GetShortcutTarget(string lnk)
        {
            if (!File.Exists(lnk))
            {
                return "";
            }
            var shell = new IWshRuntimeLibrary.WshShell();
            var shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(lnk);
            return Environment.ExpandEnvironmentVariables(shortcut.TargetPath);
        }
    }
}
