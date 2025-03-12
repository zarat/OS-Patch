using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Management;
using System.Runtime.CompilerServices;

class Program
{
    private const string Token = "test";
    private const string RepositoryUrl = $"https://zarat.lima-city.de/repository.php?token={Token}";
    private const string WingetRepoUrl = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"; // GitHub-URL für Winget

    private static string arch = String.Empty; // x86, x64, arm, etc.
    private static string type = "client"; // server, client

    static async Task Main(string[] args)
    {

        // Winget herunterladen und installieren
        //await InstallWingetAsync();

        ShowBanner();

        CheckOS();

        CheckLicense();

        if(IsServerVersion())
        {
            Console.WriteLine("[debug] IS SERVER");
            type = "server";
        }

        if (Environment.Is64BitOperatingSystem)
        {
            arch = "x64";
        }
        else
        {
            arch = "x86";
        }

        Console.WriteLine("[info] Lade aktuellen Datensatz...");
        var repoPackages = await FetchRepositoryPackagesAsync();
        if (repoPackages.Count == 0)
        {
            Console.WriteLine("[error] Repository konnte nicht heruntergeladen werden");
            return;
        }
        Console.WriteLine($"[info] Manifest-Dateien für {repoPackages.Count} Software-Pakete geladen");

        // \todo
        bool update = false;
        if (args.Count() > 0) { update = true; }

        // Installierte 32 bit programme
        //PSExecute("Get-ItemProperty HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | select DisplayName,DisplayVersion,UninstallString | convertto-csv -NoTypeInformation");
        // installierte 64 bit programme
        //PSExecute("Get-ItemProperty HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | select DisplayName,DisplayVersion,UninstallString | convertto-csv -NoTypeInformation");

        Console.WriteLine($"[info] Suche nach auf dem System installierten Software-Paketen...");
        var installedPackages = GetInstalledPackages();      
        if (installedPackages.Count == 0)
        {
            Console.WriteLine("[info] Keine installierten Pakete gefunden.");
            return;
        }
        //Console.WriteLine($"[info] {installedPackages.Count} installierte Software-Pakete gefunden");
        
        List<Package> managedPackages = new List<Package>();
        List<Package> unmanagedPackages = new List<Package>();
        
        if (installedPackages.Count() > 0)
        {
            //Console.WriteLine("");
            foreach (Package p in installedPackages)
            {
                var repoPkg = repoPackages.FirstOrDefault(r => r.Name.Equals(p.Name, StringComparison.OrdinalIgnoreCase));
                //var repoPkg = repoPackages.FirstOrDefault(r => r.Id.Equals(p.Id, StringComparison.OrdinalIgnoreCase));
                if (repoPkg != null) 
                {
                    //Console.WriteLine(" +++ " + p.Name + " - " + p.Version);
                    managedPackages.Add(p);
                } 
                else
                {
                    //Console.WriteLine(" --- " + p.Name + " - " + p.Version);
                    unmanagedPackages.Add(p);
                }
            }
            //Console.WriteLine("");
        }

        Console.WriteLine($"[info] {managedPackages.Count()} verwaltete und {unmanagedPackages.Count()} nicht-verwaltete installierte Software-Pakete gefunden\n");
        foreach(Package p in managedPackages) { Console.WriteLine("\t+++ " + p.Name + " - " + p.Version); }
        foreach (Package p in unmanagedPackages) { Console.WriteLine("\t--- " + p.Name + " - " + p.Version); }
        Console.WriteLine("");

        Console.WriteLine("[info] Vergleiche installierte und verfügbare Versionen...\n");
        List<string> uptodateprograms = new List<string>();
        bool updatesAvailable = false;
        foreach (var installed in installedPackages)
        {
            var repoPkg = repoPackages.FirstOrDefault(r => r.Name.Equals(installed.Name, StringComparison.OrdinalIgnoreCase));
            if (repoPkg != null)
            {
                var latestRepoVersion = repoPkg.AvailableVersions.OrderByDescending(v => new Version(v)).FirstOrDefault();
                if (latestRepoVersion != null && IsNewerVersion(latestRepoVersion, installed.Version))
                {
                    Console.Write($"{installed.Name} ");
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.Write($"{installed.Version} ");
                    Console.ResetColor();
                    Console.Write("<");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($" {latestRepoVersion}");
                    Console.ResetColor();
                    updatesAvailable = true;   
                    string base_url = "https://raw.githubusercontent.com/microsoft/winget-pkgs/refs/heads/master/manifests/";
                    string firstLetter = repoPkg.Id.Substring(0, 1).ToLower();
                    string id = repoPkg.Id.Replace(".", "/"); // Main.Sub => Main/Sub
                    string version = repoPkg.AvailableVersions[0];
                    string finalURL = base_url + firstLetter + "/" + id + "/" + version + "/" + repoPkg.Id + ".installer.yaml";
                    string yamlContent = String.Empty;
                    try
                    {
                        yamlContent = await DownloadYamlAsync(finalURL);
                    } catch(Exception e) {
                        Console.WriteLine("[debug] error downloading installer");
                    }
                    List<(string Architecture, string InstallerUrl)> installerData = ExtractInstallerData(yamlContent);
                    foreach (var data in installerData)
                    {
                        Console.WriteLine($"\t{data.Architecture}: {data.InstallerUrl}");
                        if(data.Architecture == arch)
                        {
                            //await InstallPackageAsync(data.InstallerUrl);
                        }
                    }     
                }
                else if (latestRepoVersion == null)
                {
                    Console.WriteLine($"[info] Keine Version im Repository für {installed.Name}.");
                }
                else
                {
                    uptodateprograms.Add($"{installed.Name} {installed.Version} >= {latestRepoVersion}");
                }
            }
        }
        Console.WriteLine("");

        if (uptodateprograms.Count() > 0)
        {
            Console.WriteLine("\n[info] Up to Date Software Pakete.");
            Console.ForegroundColor = ConsoleColor.Green;
            foreach (string s in uptodateprograms)
            {
                Console.WriteLine(" * " + s);
            }
            Console.ResetColor();

            //Console.ReadLine();
        }

        if (!updatesAvailable)
        {
            Console.WriteLine("[info] Alle Pakete sind auf dem neuesten Stand.");
        }

        //await InstallPackageAsync("https://raw.githubusercontent.com/microsoft/winget-pkgs/refs/heads/master/manifests/d/Devolutions/RemoteDesktopManager/2024.3.29.0/Devolutions.RemoteDesktopManager.installer.yaml");

        Console.WriteLine("\n[debug] HAFTUNGSHINWEIS: Dieses programm befindet sich in der Testphase! Klicken Sie Enter um zu Beenden.");
        Console.ReadLine();
    }

    public static bool IsServerVersion()
    {
        using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
        {
            foreach (ManagementObject managementObject in searcher.Get())
            {
                // ProductType will be one of:
                // 1: Workstation
                // 2: Domain Controller
                // 3: Server
                uint productType = (uint)managementObject.GetPropertyValue("ProductType");
                return productType != 1;
            }
        }

        return false;
    }

    public static void CheckOS()
    {
        var os = Environment.OSVersion;
        Console.WriteLine("[info] Prüfe Betriebssystem...\n");
        Console.WriteLine("\tPlatform: {0:G}", os.Platform);
        Console.WriteLine("\tVersion String: {0}", os.VersionString);
        Console.WriteLine("\tVersion Information:");
        Console.WriteLine("\t\tMajor: {0}", os.Version.Major);
        Console.WriteLine("\t\tMinor: {0}", os.Version.Minor);
        Console.WriteLine("\tService Pack: '{0}'", os.ServicePack);

        if (Environment.OSVersion.VersionString.Contains("Windows Server"))
        {
            Console.WriteLine("\tOS Type: Windows Server");
            type = "server";
        }
        else
        {
            Console.WriteLine("\tOS Type: Windows Client");
            type = "client";
        }
        Console.WriteLine("");
    }

    public static void CheckLicense()
    {
        try
        {
            Console.WriteLine("[info] Prüfe Lizenzstatus...\n");
            // WMI-Abfrage zur Win32_OperatingSystem-Klasse
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL");

            foreach (ManagementObject obj in searcher.Get())
            {
                Console.WriteLine("\tName: " + obj["Name"]);
                Console.WriteLine("\tLizenzstatus: " + obj["LicenseStatus"]);
                Console.WriteLine("\tTeilprodukt-Key: " + obj["PartialProductKey"]);
                Console.WriteLine("\tBeschreibung: " + obj["Description"]);
                //Console.WriteLine(new string('-', 40));
                Console.WriteLine("");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("[error] Fehler: " + ex.Message);
        }
    }
    #region Production

    // Zeige ASCII Banner
    static void ShowBanner()
    {
        Console.WriteLine(@"        __                                     __         .__                  
_____  |  | __ _____             ___________ _/  |_  ____ |  |__   ___________ 
\__  \ |  |/ //     \    ______  \____ \__  \\   __\/ ___\|  |  \_/ __ \_  __ \
 / __ \|    <|  Y Y  \  /_____/  |  |_> > __ \|  | \  \___|   Y  \  ___/|  | \/
(____  /__|_ \__|_|  /           |   __(____  /__|  \___  >___|  /\___  >__|   
     \/     \/     \/            |__|       \/          \/     \/     \/       ");
        Console.WriteLine("");
    }

    // Download and install update package
    static async Task InstallPackageAsync(string url)
    {
        Console.WriteLine("[info] [download] Starte den Download der neuen Version...");

        string fileName = Path.GetFileName(new Uri(url).LocalPath);
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

        bool downloadSuccess = await DownloadFileAsync(url, filePath);

        if (downloadSuccess)
        {
            // Prüfen, ob Datei wirklich existiert
            if (File.Exists(filePath))
            {
                Console.WriteLine($"[info] [download] Datei erfolgreich gespeichert: {filePath}");

                // TODO: Installation starten
                Console.WriteLine("[info] [install] Installation starten...");

                // Datei löschen nach der Installation
                //File.Delete(filePath);
                //Console.WriteLine("Datei gelöscht.");
            }
            else
            {
                Console.WriteLine("[error] [download] Datei wurde nicht gespeichert!");
            }
        }
        else
        {
            Console.WriteLine("[error] FEHLER: Download fehlgeschlagen!");
        }

        // Hier kannst du die Installation starten
        //Console.WriteLine($"Datei gespeichert unter: {filePath}");
        //File.Delete(filePath);
    }

    static async Task<bool> DownloadFileAsync(string url, string destinationPath)
    {
        try
        {
            using HttpClient client = new HttpClient();
            using HttpResponseMessage response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode();

            await using FileStream fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None);
            await response.Content.CopyToAsync(fileStream);

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Download: {ex.Message}");
            return false;
        }
    }

    static void PSExecute(String command)
    {
        try
        {

            string arguments = " -NoProfile -ExecutionPolicy Bypass -Command \"" + command + "\"";

            var processStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = command,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = processStartInfo };
            process.Start();

            var packageList = new List<Package>();
            bool firstLine = true;

            while (!process.StandardOutput.EndOfStream)
            {
                string line = process.StandardOutput.ReadLine()?.Trim();
                Console.WriteLine(line);
            }

            process.WaitForExit();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Abruf: {ex.Message}");
        }
    }

    static List<Package> GetInstalledPackages()
    {

        // test which packages are available
        /*
        if(IsWingetAvailable()) { 
            Console.WriteLine($"[debug] package managment is available"); 
        }
        else
        {
            Console.WriteLine($"[debug] package managment is not available");
        }
        if (IsGetWingetPackageCmdletAvailable())
        {
            Console.WriteLine($"[debug] powershell package managment is available");
        } else
        {
            Console.WriteLine($"[debug] powershell package managment is not available");
        }
        */

        //type = "server";
            try
            {

                string arguments = String.Empty;
                if (type == "server")
                {
                    // WMI kennt keine source
                    // Name,Version,Fake(Source=akm)
                    arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"Get-WMIObject -Class Win32_Product | Where-Object { -not ($_.Vendor -match 'Microsoft') } | Select-Object Name,Id,Version,@{ Name = 'Source'; Expression = { 'wmi-has-none' } } | ConvertTo-Csv -NoTypeInformation\"";
                    Console.WriteLine($"[info] Wähle WMI für \"{type}\"");
                }
                else
                {
                    // Name,Version,Source
                    arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"Get-WingetPackage | Where-Object { -not ($_.Name -match 'Microsoft') }| select-object Name,Id,InstalledVersion,Source | ConvertTo-Csv -NoTypeInformation\"";
                    Console.WriteLine($"[info] Wähle package manager für \"{type}\"");
                }

                var processStartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = new Process { StartInfo = processStartInfo };
                process.Start();

                var packageList = new List<Package>();
                bool firstLine = true;

                while (!process.StandardOutput.EndOfStream)
                {
                    string line = process.StandardOutput.ReadLine()?.Trim();
                    if (string.IsNullOrWhiteSpace(line) || firstLine)
                    {
                        firstLine = false; // Erste Zeile (CSV-Header) ignorieren
                        continue;
                    }
                    Console.WriteLine("[debug] Read line..");

                    var parts = line.Split(',');
                    if (parts.Length >= 2)
                    {
                        packageList.Add(new Package
                        {
                            Name = parts[0].Trim('"'),
                            Id = parts[1].Trim('"'),
                            Version = parts[2].Trim('"'),
                            Source = parts[3].Trim('"') == "" ? "no-source": parts[3].Trim('"')
                        });
                    }
                }

                process.WaitForExit();
                return packageList;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Abrufen der Pakete: {ex.Message}");
                return new List<Package>();
            }
    }

    static bool IsWingetAvailable()
    {
        try
        {
            // Versuche, winget mit dem Befehl "winget --version" auszuführen
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"winget --version\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = processStartInfo };
            process.Start();

            // Wenn der Befehl erfolgreich ausgeführt wird, existiert winget
            process.WaitForExit();
            return process.ExitCode == 0; // ExitCode 0 bedeutet Erfolg
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Überprüfen von winget: {ex.Message}");
            return false;
        }
    }

    // \todo false positives
    static bool IsGetWingetPackageCmdletAvailable()
    {
        try
        {
            // Überprüfe, ob das Cmdlet "Get-WingetPackage" verfügbar ist
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"Get-Command Get-WingetPackage\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = processStartInfo };
            process.Start();

            // Liest die Ausgabe von Get-Command
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            return process.ExitCode == 0;
            // Wenn der Befehl "Get-WingetPackage" existiert, wird es in der Ausgabe von Get-Command erscheinen
            //return !string.IsNullOrEmpty(output);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Überprüfen des Cmdlets: {ex.Message}");
            return false;
        }
    }

    static void InstallWingetIfNotAvailable()
    {
        try
        {
            // Überprüfen, ob Get-WingetPackage verfügbar ist
            if (!IsGetWingetPackageCmdletAvailable())
            {
                Console.WriteLine("[info] Das Cmdlet 'Get-WingetPackage' ist nicht verfügbar. Versuche, winget zu installieren.");

                // PowerShell-Befehl zum Installieren des Windows Package Managers (winget) über den Microsoft Store
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"Get-AppxPackage Microsoft.DesktopAppInstaller | Add-AppxPackage -ForceReinstall\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = new Process { StartInfo = processStartInfo };
                process.Start();
                process.WaitForExit();

                Console.WriteLine("[info] Installation abgeschlossen, bitte neu starten, um sicherzustellen, dass winget funktioniert.");
            }
            else
            {
                Console.WriteLine("[info] Get-WingetPackage ist bereits verfügbar.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler bei der Installation von winget: {ex.Message}");
        }
    }

    static async Task<List<Package>> FetchRepositoryPackagesAsync()
    {
        try
        {
            using var client = new HttpClient();
            string jsonString = await client.GetStringAsync(RepositoryUrl);

            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            return JsonSerializer.Deserialize<List<Package>>(jsonString, options) ?? new List<Package>();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Abrufen der Repository-Daten: {ex.Message}");
            return new List<Package>();
        }
    }

    static bool IsNewerVersion(string repoVersion, string installedVersion)
    {
        if (Version.TryParse(installedVersion, out var installedVer) && Version.TryParse(repoVersion, out var repoVer))
        {
            return repoVer > installedVer;
        }
        return false;
    }

    private class Package
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Id { get; set; }
        public string Source { get; set; }
        public List<string> AvailableVersions { get; set; }  // Array von Versionen
    }

    static async Task<string> DownloadYamlAsync(string url)
    {
        using HttpClient client = new HttpClient();
        return await client.GetStringAsync(url);
    }

    static List<(string Architecture, string InstallerUrl)> ExtractInstallerData(string yaml)
    {
        List<(string Architecture, string InstallerUrl)> installerData = new List<(string, string)>();

        // Regex um nach Architektur und InstallerUrl zu suchen
        Regex architectureRegex = new Regex(@"Architecture:\s*(\S+)", RegexOptions.IgnoreCase);
        Regex urlRegex = new Regex(@"InstallerUrl:\s*(https?://\S+)", RegexOptions.IgnoreCase);

        string architecture = null;
        foreach (var line in yaml.Split('\n'))
        {
            if (architectureRegex.IsMatch(line))
            {
                architecture = architectureRegex.Match(line).Groups[1].Value;
            }

            if (urlRegex.IsMatch(line))
            {
                string url = urlRegex.Match(line).Groups[1].Value;
                if (architecture != null)
                {
                    installerData.Add((architecture, url));
                    architecture = null; // Architektur für den nächsten Installer zurücksetzen
                }
            }
        }

        return installerData;
    }

    #endregion

    #region Temporary / unused

    // Holt die neueste Version aus den GitHub-Releases
    // unused
    // \todo
    static async Task InstallGithubPackageAsync(string url)
    {
        Console.WriteLine("Starte den Download der neuen Version...");
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), ""); // filename?
        await DownloadFileAsync1(url, filePath);
        // install
        File.Delete(filePath);
    }

    // Holt die neueste Version eines Pakets aus den GitHub-Releases
    static async Task<string> GetLatestGithubDownloadUrlAsync()
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");

        try
        {
            var response = await client.GetStringAsync(WingetRepoUrl);
            var releaseData = JObject.Parse(response);

            // Suche nach dem MSIX-Bundle-Download-Link
            foreach (var asset in releaseData["assets"])
            {
                var browserDownloadUrl = asset["browser_download_url"].ToString();
                if (browserDownloadUrl.Contains(".msixbundle"))
                {
                    return browserDownloadUrl;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Abrufen der Release-Daten: {ex.Message}");
        }

        return null;
    }

    // Lädt die Datei von der angegebenen URL herunter
    static async Task DownloadFileAsync1(string url, string filePath)
    {
        using var client = new HttpClient();
        try
        {
            var response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var content = await response.Content.ReadAsByteArrayAsync();
            await File.WriteAllBytesAsync(filePath, content);
            Console.WriteLine("Download abgeschlossen.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Downloaden der Datei: {ex.Message}");
        }
    }

    // Installiert das MSIX-Paket
    static void InstallMsixPackage(string filePath)
    {
        try
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"Add-AppxPackage -Path \"{filePath}\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = processStartInfo };
            process.Start();
            process.WaitForExit();
            Console.WriteLine("Winget wurde erfolgreich installiert.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler bei der Installation: {ex.Message}");
        }
    }

    #endregion

}
