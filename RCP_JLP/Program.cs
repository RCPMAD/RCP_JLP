using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using OpenMcdf;
using Shellify;
namespace woanware
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Console.WriteLine("JLP STARTED");
            new FormMain();
            Console.WriteLine("JLP STOPPED");
        }
    }
    public class DestListEntry
    {
        public Guid[] Droid { get; set; }
        public Guid[] DroidBirth { get; set; }
        public Uuid Uuid { get; set; }
        public Uuid UuidBirth { get; set; }
        public string NetBiosName { get; set; }
        public string StreamNo { get; set; }
        public DateTime FileTime { get; set; }
        public string Data { get; set; }
        public DestListEntry()
        {
            Droid = new Guid[2];
            DroidBirth = new Guid[2];
            NetBiosName = string.Empty;
            StreamNo = string.Empty;
            Data = string.Empty;
        }
    }
    public class JumpList
    {
        public string Name { get; set; }
        public int Size { get; set; }
        public List<NameValue> Data { get; set; }
        public DestListEntry DestListEntry { get; set; }
        public JumpList()
        {
            Data = new List<NameValue>();
            DestListEntry = new DestListEntry();
        }
    }
    public class JumpListFile
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string AppName { get; set; }
        public List<JumpList> JumpLists { get; set; }
        public List<DestListEntry> DestListEntries { get; set; }
        public long DestListSize { get; set; }
        public bool IsCustom { get; private set; }
        public JumpListFile(bool isCustom)
        {
            IsCustom = isCustom;
            FilePath = string.Empty;
            FileName = string.Empty;
            JumpLists = new List<JumpList>();
            DestListEntries = new List<DestListEntry>();
        }
    }
    public partial class FormMain : Form
    {
        #region Constructor
        public FormMain()
        {
            LoadFiles();
            Close();
        }
        #endregion
        #region Methods
        private void LoadFiles()
        {
            string temp = Environment.CurrentDirectory + "\\temp";
            string myHeaders = "User,Application ID,Application Name,Path,Arguments,Type,Accessed Time,Modified Time,Created Time";
            File.WriteAllText("FM-JLP.csv", myHeaders + Environment.NewLine);
            for (int index1 = 0; index1 <= 50; index1 = index1 + 1)
            {
                string args1 = @"(gwmi -query 'SELECT * FROM Win32_UserAccount WHERE LocalAccount=True').Name.get(" + index1 + ")";
                var startInfo = new ProcessStartInfo();
                startInfo.Arguments = args1;
                startInfo.WorkingDirectory = "c:\\windows\\system32";
                startInfo.RedirectStandardOutput = true;
                startInfo.UseShellExecute = false;
                startInfo.FileName = "powershell.exe";
                startInfo.CreateNoWindow = true;
                Process procHD = Process.Start(startInfo);
                string username = (procHD.StandardOutput.ReadToEnd()).TrimEnd();
                if (username.Contains("Exception")) { index1 = 50; }
                if (index1 != 50)
                {
                    string dirC = "C:\\Users\\" + username + @"\AppData\Roaming\Microsoft\Windows\Recent\CustomDestinations";
                    string dirA = "C:\\Users\\" + username + @"\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations";
                    DirectoryInfo folderB = new DirectoryInfo(dirC);
                    DirectoryInfo folderA = new DirectoryInfo(dirA);
                    DirectoryInfo folderTemp = new DirectoryInfo(temp + "\\" + username);
                    System.IO.DirectoryInfo di = new DirectoryInfo(temp + "\\" + username);
                    if (folderTemp.Exists == true) { di.Delete(true); }
                    folderTemp = new DirectoryInfo(temp + "\\" + username);
                    if (folderTemp.Exists == false) { System.IO.Directory.CreateDirectory(temp + "\\" + username); }
                    folderTemp = new DirectoryInfo(temp + "\\" + username);
                    if (folderB.Exists == true)
                    {
                        foreach (var file1 in folderB.GetFiles("*.*Destinations-ms"))
                        {
                            string file = file1.DirectoryName.ToString() + "\\" + file1.ToString();
                            string basename = Path.GetFileNameFromPath(file);
                            System.IO.File.Copy(file, temp + "\\" + username + "\\" + basename, true);
                        }
                    }
                    if (folderA.Exists == true)
                    {
                        foreach (var file1 in folderA.GetFiles("*.*Destinations-ms"))
                        {
                            string file = file1.DirectoryName.ToString() + "\\" + file1.ToString();
                            string basename = Path.GetFileNameFromPath(file);
                            System.IO.File.Copy(file, temp + "\\" + username + "\\" + basename, true);
                        }
                    }
                }
                if (index1 == 50)
                {
                    foreach (string subdir in System.IO.Directory.GetDirectories(temp))
                    {
                        DirectoryInfo folderTemp = new DirectoryInfo(subdir);
                        foreach (var file1 in folderTemp.GetFiles("*.*Destinations-ms"))
                        {
                            DirectoryInfo dir = new DirectoryInfo(subdir);
                            username = dir.Name.ToString();
                            string file = file1.DirectoryName.ToString() + "\\" + file1.ToString();
                            if (File.Exists(file) == false)
                            {
                                continue;
                            }
                            string extension = System.IO.Path.GetExtension(file);

                            if (extension == ".customDestinations-ms")
                            {
                                List<ShellLinkFile> shellLinkFiles = ShellLinkFile.LoadJumpList(file);
                                if (shellLinkFiles.Count == 0)
                                {
                                    continue;
                                }
                                string appID = System.IO.Path.GetFileNameWithoutExtension(file);
                                string appName = "Unknown";
                                StreamReader reader = new StreamReader(File.OpenRead(Environment.CurrentDirectory + "\\AppID.txt"));
                                List<string> listA = new List<String>();
                                List<string> listB = new List<String>();
                                while (!reader.EndOfStream)
                                {
                                    string line = reader.ReadLine();
                                    if (!String.IsNullOrWhiteSpace(line))
                                    {
                                        string[] values = line.Split(',');
                                        if (values.Length >= 2)
                                        {
                                            listA.Add(values[0]);
                                            listB.Add(values[1]);
                                        }
                                    }
                                }
                                string[] firstlistA = listA.ToArray();
                                string[] firstlistB = listB.ToArray();
                                if (firstlistA.Contains(appID))
                                {
                                    for (int indx = 1; indx <= 367; indx = indx + 1)
                                    {
                                        if (firstlistA.GetValue(indx).ToString() == appID)
                                        {
                                            appName = firstlistB.GetValue(indx).ToString();
                                            indx = 367;
                                        }

                                    }
                                }
                                ListViewItem listItem = new ListViewItem();
                                listItem.Text = System.IO.Path.GetFileName(file);
                                listItem.SubItems.Add(appName);
                                listItem.Tag = file;
                                ParseJumpListC(file, appID, extension, username, appName, shellLinkFiles);
                            }
                            else if (extension == ".automaticDestinations-ms")
                            {
                                string appID = System.IO.Path.GetFileNameWithoutExtension(file);
                                string appName = "Unknown";
                                appID = System.IO.Path.GetFileNameWithoutExtension(file);
                                appName = string.Empty;
                                StreamReader reader = new StreamReader(File.OpenRead(Environment.CurrentDirectory.ToString() + "\\AppID.txt"));
                                List<string> listA = new List<String>();
                                List<string> listB = new List<String>();
                                while (!reader.EndOfStream)
                                {
                                    string line = reader.ReadLine();
                                    if (!String.IsNullOrWhiteSpace(line))
                                    {
                                        string[] values = line.Split(',');
                                        if (values.Length >= 2)
                                        {
                                            listA.Add(values[0]);
                                            listB.Add(values[1]);
                                        }
                                    }
                                }
                                string[] firstlistA = listA.ToArray();
                                string[] firstlistB = listB.ToArray();
                                if (firstlistA.Contains(appID))
                                {
                                    for (int indx = 1; indx <= 367; indx = indx + 1)
                                    {
                                        if (firstlistA.GetValue(indx).ToString() == appID)
                                        {
                                            appName = firstlistB.GetValue(indx).ToString();
                                            indx = 367;
                                        }
                                    }
                                }
                                ListViewItem listItem = new ListViewItem();
                                listItem.Text = System.IO.Path.GetFileName(file);
                                listItem.SubItems.Add(appName);
                                listItem.Tag = file;
                                ParseJumpListA(appID, extension, username, file, appName);
                            }
                            Close();
                        }
                    }
                }
            }
        }
        private void ParseJumpListA(string appid, string extension, string username, string file, string appName)
        {
            JumpListFile jumpListFile = new JumpListFile(false);
            jumpListFile.FilePath = file;
            jumpListFile.FileName = System.IO.Path.GetFileName(file);
            jumpListFile.AppName = appName;
            CompoundFile compoundFile = new CompoundFile(file);
            CFStream cfStream = compoundFile.RootStorage.GetStream("DestList");
            jumpListFile.DestListSize = cfStream.Size;
            List<DestListEntry> destListEntries = ParseDestList(cfStream.GetData());
            jumpListFile.DestListEntries = destListEntries;

            int ii = 1;
            foreach (DestListEntry destListEntry in destListEntries)
            {
                CFStream cfStreamJf = null;
                try
                {
                    cfStreamJf = compoundFile.RootStorage.GetStream(ii.ToString());
                    ShellLinkFile linkFile = ShellLinkFile.Load(cfStreamJf.GetData());
                    string arguments = linkFile.Arguments.ToString();
                    if (arguments == null) { arguments = "None"; }
                    DateTime accesstime_ = linkFile.Header.AccessTime;
                    string accesstime = Convert.ToDateTime(accesstime_).ToString("dd/MM/yyyy HH:mm:ss");
                    string writetime = linkFile.Header.WriteTime.ToString("dd/MM/yyyy HH:mm:ss");
                    string createtime = linkFile.Header.CreationTime.ToString("dd/MM/yyyy HH:mm:ss");
                    string ExeName = linkFile.LinkInfo.LocalBasePath;
                    string myArray = username + "," + appid + "," + appName + "," + ExeName + "," + arguments + "," + extension + "," + accesstime + "," + writetime + "," + createtime;
                    File.AppendAllText("FM-JLP.csv", myArray + Environment.NewLine);
                    if (linkFile.Header.CreationTime == DateTime.MinValue |
                        linkFile.Header.AccessTime == DateTime.MinValue |
                        linkFile.Header.WriteTime == DateTime.MinValue |
                        linkFile.Header.CreationTime == ShellLinkFile.WindowsEpoch |
                        linkFile.Header.AccessTime == ShellLinkFile.WindowsEpoch |
                        linkFile.Header.WriteTime == ShellLinkFile.WindowsEpoch)
                    {
                        continue;
                    }
                    if (linkFile != null)
                    {
                        JumpList jumpList = new JumpList();
                        jumpList.Name = destListEntry.StreamNo;
                        jumpList.Size = cfStreamJf.GetData().Length;
                        jumpList.DestListEntry = destListEntry;
                        jumpListFile.JumpLists.Add(jumpList);
                    }
                }
                catch { Exception ex; }
                ii = ii + 1;
            }
        }
        private void ParseJumpListC(string file, string appid, string extension, string username, string appName, List<ShellLinkFile> lnkFiles)
        {
            JumpListFile jumpListFile = new JumpListFile(true);
            jumpListFile.FilePath = file;
            jumpListFile.FileName = System.IO.Path.GetFileName(file);
            jumpListFile.AppName = appName;
            for (int index = 0; index < lnkFiles.Count; index++)
            {
                ShellLinkFile linkFile = lnkFiles[index];
                JumpList jumpList = new JumpList();
                jumpList.Name = (index + 1).ToString();
                jumpList.Size = 0;
                var Data = lnkFiles.ToArray();
                string arguments = Data[index].Arguments;
                if (arguments == null) { arguments = "None"; }
                string accesstime = Data[index].Header.AccessTime.ToString("dd/MM/yyyy HH:mm:ss");
                string writetime = Data[index].Header.WriteTime.ToString("dd/MM/yyyy HH:mm:ss");
                string createtime = Data[index].Header.CreationTime.ToString("dd/MM/yyyy HH:mm:ss");
                string Name = Data[index].IconLocation;
                string ExeName = Data[index].LinkInfo.LocalBasePath;
                string myArray = username + "," + appid + "," + appName + "," + ExeName + "," + arguments + "," + extension + "," + accesstime + "," + writetime + "," + createtime;
                File.AppendAllText("FM-JLP.csv", myArray + Environment.NewLine);
            }
        }
        private List<DestListEntry> ParseDestList(byte[] data)
        {
            List<DestListEntry> entries = new List<DestListEntry>();
            try
            {
                using (MemoryStream memoryStream = new MemoryStream(data))
                {
                    memoryStream.Seek(32, SeekOrigin.Begin);
                    do
                    {
                        DestListEntry destListEntry = new DestListEntry();
                        memoryStream.Seek(8, SeekOrigin.Current);
                        destListEntry.Droid = new Guid[] { new Guid(StreamReaderHelper.ReadByteArray(memoryStream, 16)), new Guid(StreamReaderHelper.ReadByteArray(memoryStream, 16)) };
                        destListEntry.Uuid = new Uuid(destListEntry.Droid[1].ToString());
                        destListEntry.DroidBirth = new Guid[] { new Guid(StreamReaderHelper.ReadByteArray(memoryStream, 16)), new Guid(StreamReaderHelper.ReadByteArray(memoryStream, 16)) };
                        destListEntry.UuidBirth = new Uuid(destListEntry.DroidBirth[1].ToString());
                        destListEntry.NetBiosName = woanware.Text.ReplaceNulls(StreamReaderHelper.ReadString(memoryStream, 16));
                        destListEntry.StreamNo = StreamReaderHelper.ReadInt64(memoryStream).ToString("X");
                        memoryStream.Seek(4, SeekOrigin.Current);
                        destListEntry.FileTime = StreamReaderHelper.ReadDateTime(memoryStream);
                        memoryStream.Seek(4, SeekOrigin.Current);
                        int stringLength = StreamReaderHelper.ReadInt16(memoryStream);
                        if (stringLength != -1)
                        {
                            destListEntry.Data = StreamReaderHelper.ReadStringUnicode(memoryStream, stringLength * 2);
                        }
                        else
                        {
                            memoryStream.Seek(4, SeekOrigin.Current);
                        }
                        destListEntry.Data = woanware.Text.ReplaceNulls(destListEntry.Data);
                        entries.Add(destListEntry);
                    }
                    while (memoryStream.Position < memoryStream.Length);
                }
            }
            catch (Exception ex)
            {
            }
            return entries;
        }
        #endregion
    }
}