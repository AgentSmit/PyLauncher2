using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PyLauncher2
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            //Определяем версию windows
            OperatingSystem os = Environment.OSVersion;

            if (os.Platform != PlatformID.Win32NT)
            {
                return;
            }

            if (os.Version.Major < 6)
            {
                return;
            }

            //Читаем параметры
            if (File.Exists(configFilename))
            {
                string[] lines = File.ReadAllLines(configFilename);
                foreach (string line in lines)
                {
                    if (line.Length < 1)
                    {
                        continue;
                    }

                    string[] dict = line.Split("=");
                    string key = dict[0];
                    string value = dict[1];

                    switch (key)
                    {
                        case "ConsoleMode":
                            consoleMode = Convert.ToBoolean(Convert.ToInt32(value));
                            break;
                    }
                }
            }

            bool pythonFound = CheckPython();
            if (!pythonFound)
            {
                if (!File.Exists(pyDistFilename))
                {
                    //Загружаем Python
                    DownloadPython(os);
                }

                //Устанавливаем Python
                SetupPython();
            }

            //Обновляем pip
            //string output = RunCommand("python", "-m pip install --upgrade pip");

            CheckPackages();

            LaunchPyScript();
        }

        static readonly string basePath = AppDomain.CurrentDomain.BaseDirectory;
        static readonly string distDir = $"{basePath}dist";
        static readonly string pyDir = $"{basePath }PythonFiles";
        static readonly string pyDistFilename = $"{distDir}\\python-dist.exe";
        static readonly string pyFileName = $"{pyDir}\\main.py";
        static readonly string pycFileName = $"{pyDir}\\main.pyc";
        static readonly string requirementsFilename = $"{pyDir}\\requirements.txt";
        static readonly string configFilename = $"{basePath}config.cfg";

        static bool consoleMode = true;

        static bool CheckPython()
        {
            using (Process cmd = new Process())
            {
                cmd.StartInfo.FileName = "python";
                cmd.StartInfo.Arguments = "-V";
                cmd.StartInfo.UseShellExecute = false;
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                try
                {
                    cmd.Start();
                    string outputStr = cmd.StandardOutput.ReadToEnd().Trim();
                    cmd.WaitForExit();
                    return outputStr != "";
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        static void DownloadPython(OperatingSystem os)
        {
            if (File.Exists(pyDistFilename))
            {
                return;
            }

            string downloadUrl = "https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe"; //For release
            //string downloadUrl = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe"; // For debug

            if (os.Version.Major == 6 && os.Version.Minor == 1)
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    downloadUrl = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe";
                }
                else
                {
                    downloadUrl = "https://www.python.org/ftp/python/3.8.10/python-3.8.10.exe";
                }
            }

            CreateDirectoies();
            if (AllocConsole())
            {
                try
                {
                    Console.WriteLine("Началась загрузка Python");
                    Task.Delay(1500).GetAwaiter().GetResult();
                    WebClient webClient = new WebClient();
                    webClient.DownloadFile(downloadUrl, pyDistFilename);
                }
                catch (WebException)
                {
                    Console.WriteLine("Проверьте соединение с интернет");
                    Task.Delay(5000).GetAwaiter().GetResult();
                    FreeConsole();
                    Environment.Exit(-1);
                }
                Console.WriteLine("Python успешно загружен");
                Task.Delay(3000).GetAwaiter().GetResult();
                FreeConsole();
            }
        }

        static void CreateDirectoies()
        {
            if (!Directory.Exists(distDir))
            {
                Directory.CreateDirectory(distDir);
            }
            if (!Directory.Exists(pyDir))
            {
                Directory.CreateDirectory(pyDir);
            }
        }

        static void SetupPython()
        {
            if (AllocConsole())
            {
                Console.WriteLine("Установка Python");
                Task.Delay(1500).GetAwaiter().GetResult();
                using (Process cmd = new Process())
                {
                    try
                    {
                        cmd.StartInfo.FileName = "cmd";
                        cmd.StartInfo.Arguments = $"/C {pyDistFilename} /quiet InstallAllUsers=1 PrependPath=1 Shortcuts=0";
                        cmd.StartInfo.UseShellExecute = false;
                        cmd.StartInfo.RedirectStandardOutput = true;
                        cmd.StartInfo.CreateNoWindow = false;
                        cmd.Start();
                        cmd.WaitForExit();
                        MessageBox.Show("Python установлен. \nПерезапустите программу!", "PyLauncher 2",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Не удалось установить Python", "PyLauncher 2",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    
                }
                FreeConsole();
                Environment.Exit(0);
            }
        }

        static string RunCommand(string filename, string arguments)
        {
            using (Process cmd = new Process())
            {
                cmd.StartInfo.FileName = filename;
                cmd.StartInfo.Arguments = arguments;
                cmd.StartInfo.UseShellExecute = false;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.Start();
                cmd.WaitForExit();
                string output = cmd.StandardOutput.ReadToEnd();
                return output;
            }
        }

        static void RunInteractive(string filename, string arguments)
        {
            using (Process cmd = new Process())
            {
                cmd.StartInfo.FileName = "cmd";
                cmd.StartInfo.Arguments = $"/C {filename} {arguments}";
                cmd.StartInfo.UseShellExecute = true;
                cmd.StartInfo.CreateNoWindow = false;
                cmd.Start();
                cmd.WaitForExit();
            }
        }

        static void CheckPackages()
        {
            if (File.Exists(requirementsFilename))
            {
                //Проверяем зависимости
                string output = RunCommand("pip", "list --format freeze");
                List<string> installedPackageList = new List<string>();

                foreach (string str in output.Split("\n"))
                {
                    string packageName = str.Split("==")[0];
                    installedPackageList.Add(packageName);
                }

                string reqFile = File.ReadAllText(requirementsFilename);
                List<string> requirementList = new List<string>();
                foreach (string req in reqFile.Split("\n"))
                {
                    requirementList.Add(req.Trim());
                }

                bool found = false;

                foreach (string req in requirementList)
                {
                    foreach (string package in installedPackageList)
                    {
                        if (package.ToUpper() == req.ToUpper())
                        {
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        break;
                    }
                }

                if (!found)
                {
                    RunInteractive("pip", $"install -r {requirementsFilename}");
                }
            }
        }

        static void LaunchPyScript()
        {

            if (File.Exists(pyFileName))
            {
                if (consoleMode)
                {
                    RunInteractive("python", pyFileName);
                }
                else
                {
                    RunCommand("python", pyFileName);
                }
                return;
            }

            if (File.Exists(pycFileName))
            {
                if (consoleMode)
                {
                    RunInteractive("python", pyFileName);
                }
                else
                {
                    RunCommand("python", pyFileName);
                }
                return;
            }
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FreeConsole();
    }
}
