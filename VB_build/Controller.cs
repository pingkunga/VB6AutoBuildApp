using System;
using System.Collections.Generic;
using System.Collections;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;

namespace VB_build
{
    internal class Controller
    {

        public struct FormItem
        {
            public string ref_tmp_files;
            public string filename;
            public string guid;
            public string version;
            public string lcid;
        }

        private Hashtable dlls = new Hashtable();
        private string message = "";
        public string filename = null;
        public bool[] logFile = { false , false, false, false, false, false , false, false, false, false }; //0-9
        private bool check32Bit = true;
        private int time = 0;
        private string[] command;

        private bool buildVBProject(bool type, string build_path, string ref_path, string compat)
        {
            string[] files = System.IO.Directory.GetFiles(build_path);
            string file;
            //Protect Error
            if (files.Length > 0)
            {
                file = files[0];
            }
            else
            {
                return false;
            }

            //Get .vbp file
            foreach (string item in files)
            {
                if (item.LastIndexOf(".vbp") > -1)
                {
                    file = item;
                    break;
                }
            }

            Hashtable text = editVBP(type, file, compat);

            Console.Write("Write Data to VBProject: ");
            System.IO.StreamWriter vbp = new System.IO.StreamWriter(file);
            try
            {
                for (int i = 0; i < text.Count; i++)
                {
                    vbp.WriteLine(text[i].ToString());
                }
                Console.WriteLine("Successful...");
                logFile[5] = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                logFile[5] = false;
                throw ex;
            }
            finally
            {
                vbp.Close();
            }

            // Build VBProject
            makeVBProject(file);

            // Move result file and remove old file
            bool check = moveFile(type, build_path, ref_path);
            if (!check)
            {
                return false;
            }

            return true;
        }
            
        public bool[] processBuild(bool type, string build_path, string ref_path, string version, string compat, string time, string[] _command)
        {   //type [DLL:true] [EXE:FALSE]
            // Create Temp files: log[0-1]
            this.command = _command;
            this.time = int.Parse(time);
            bool check = createTempDLL(ref_path); 
            if (!check)
            {
                return logFile;
            }

            //Check OS: log[2]
            if (System.IO.Directory.Exists(@"C:\Program Files (x86)")) logFile[2] = false;
            else logFile[2] = true;
            if (logFile[2])
            {
                Console.WriteLine("OS-32bits = Yes");
                check32Bit = true;
            }
            else
            {
                Console.WriteLine("OS-32bits = No");
                check32Bit = false;
                logFile[2] = true;
            }

            //Generate name file, name folder and guid file: log[3]
            check = setValueDll(build_path);
            if (!check)
            {
                return logFile;
            }

            //Build VB project: log[4-8]
            check = buildVBProject(type, build_path, ref_path, compat);
            if (!check)
            {
                return logFile;
            }
            
            //Change Version files: log[9]
            check = changeVersion(version, type, ref_path);
            if (check)
            {
                Console.WriteLine("All Success: Successful...");
            }
            else
            {
                Console.WriteLine("All Success: Unsuccessful...");
            }
            return logFile;
        }

        private bool changeVersion(string version, bool type, string ref_path)
        {
            waitting(1, 5f);
            bool check = true;
            Console.Write("Change Version: ");
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();
            cmd.StandardInput.AutoFlush = true;
            string commandText;
            try
            {
                if (type)
                {
                    commandText = "D:\\verpatch.exe \"" + ref_path + "\\" + this.filename + ".dll\" " + version + " /va /pv " + version + " /s product \"" + this.filename + "\" /sc \"(" + DateTime.Now + ")\"";
                }
                else
                {
                    commandText = "D:\\verpatch.exe \"C:\\BuildExe\\" + this.filename + ".exe\" " + version + " /va /pv " + version + " /s product \"" + this.filename + "\" /sc \"(" + DateTime.Now + ")\"";
                }
                cmd.StandardInput.WriteLine(commandText);
                cmd.StandardInput.Flush();
                logFile[9] = true;
            }
            catch (Exception ex)
            {
                logFile[9] = false;
                check = false;
                message = ex.StackTrace;
            }
            finally
            {
                if (logFile[9])
                {
                    Console.WriteLine("Successful...");
                }
                else
                {
                    Console.WriteLine(message);
                }
                waitting(1, 1f);
                cmd.StandardInput.Close();
                cmd.WaitForExit();
            }
            return check;
        }

        private void makeVBProject(string path)
        {
            Console.Write("Build Project: ");
            string vb6;
            Process command = new Process();
            command.StartInfo.FileName = "cmd.exe";
            command.StartInfo.RedirectStandardInput = true;
            command.StartInfo.RedirectStandardOutput = true;
            command.StartInfo.CreateNoWindow = true;
            command.StartInfo.UseShellExecute = false;
            command.Start();
            command.StandardInput.AutoFlush = true;
            try
            {
                if (check32Bit)
                {
                    vb6 = "\"C:\\Program Files\\Microsoft Visual Studio\\VB98\\VB6.EXE\" /Make \"";
                }
                else
                {
                    vb6 = "\"C:\\Program Files (x86)\\Microsoft Visual Studio\\VB98\\VB6.EXE\" /Make \"";
                }
                vb6 += (path + "\"");
                int count = 0, max = this.command.Length;

                if (max != count && this.command[count].IndexOf("/out:") > -1)
                {
                    vb6 += (" /out " + this.command[count].Substring(this.command[count].IndexOf(':') + 1));
                    count++;
                }

                if (max != count && (this.command[count].IndexOf("/d") > -1 || this.command[count].IndexOf("/D") > -1))
                    vb6 += (" " + this.command[count++]);

                command.StandardInput.WriteLine(vb6);
                command.StandardInput.Flush();
                logFile[6] = true;
                Console.WriteLine("Successful...");
            }
            catch (Exception ex)
            {
                logFile[6] = false;
                Console.WriteLine(ex.StackTrace);
            }
            command.StandardInput.Close();
            command.WaitForExit();
        }

        private bool moveFile(bool type, string build_path, string ref_path)
        {
            Console.Write("Check result file from build: ");
            string source = build_path + @"\", destinasion;
            float sec = 0f;
            int time = 0;
            bool check;

            if (this.filename == null)
            {
                Console.WriteLine("Failure... Software have not \"ExeName32\".");
                return false;
            }

            if (type)
            {
                source += (this.filename + ".dll");
                sec = 10f;
                time = 6;
            }
            else
            {
                source += (this.filename + ".exe");
                sec = 60f;
                time = 20;
            }

            check = waitTime(source, sec, this.time);
            if (!check)
            {
                Console.WriteLine("Failure... Time out at Build project. (use time is " + (sec * time) +" sec)");
                return false;
            }
            else
            {
                Console.WriteLine("Successful...");
            }

            Console.Write("Remove Old file: ");
            if (type)
            {
                destinasion = ref_path + "\\" + this.filename + ".dll";
            }
            else
            {
                destinasion = @"C:\BuildExe\" + this.filename + ".exe";
            }
            if (System.IO.File.Exists(destinasion))
            {
                try
                {
                    System.IO.File.Delete(destinasion);
                    logFile[7] = true;
                }
                catch (System.IO.IOException ex)
                {
                    message = ex.StackTrace;
                    logFile[7] = false;
                }
                finally
                {
                    if (logFile[7])
                    {
                        Console.WriteLine("Successful...");
                    }
                    else
                    {
                        Console.WriteLine(message);
                    }
                }
                waitting(1, 3f);
            }
            else
            {
                logFile[7] = true;
                Console.WriteLine("Successful...");
            }

            Console.Write("Move Result file to Target Folder: ");
            try
            {
                System.IO.File.Move(source, destinasion);
                logFile[8] = true;
            }
            catch (System.IO.FileNotFoundException ex)
            {
                message = ex.StackTrace;
                waitting(2, 30f);
                waitTime(source, 15f, 2);
                System.IO.File.Move(source, destinasion);
                logFile[8] = true;
            }
            catch (Exception ex)
            {
                message = ex.StackTrace;
                logFile[8] = false;
            }
            finally
            {
                if (logFile[8])
                {
                    Console.WriteLine("Successful...");
                }
                else
                {
                    Console.WriteLine(message);
                }
            }
            waitting(2, 3f);
            
            return true;
        }

        private Hashtable editVBP(bool type, string file, string compat)
        {
            int count = 0;
            bool chkRef = false, editCompat;
            List<int> index_Ref = new List<int>();
            Hashtable text = new Hashtable();
            string[] vbfile = System.IO.File.ReadAllLines(file);
            string name = "NOT";
            
            try
            {
                if (compat.Equals("-b") && type)
                {
                    editCompat = true;
                    Console.WriteLine("CompatibleMode: Binary Compatible");
                    Console.Write("Edit VBProject: ");
                    for (int i = 0; i < vbfile.Length; i++, count++)
                    {
                        //Get Reference Line
                        if (vbfile[i].IndexOf(".vbp") > -1)
                        {
                            index_Ref.Add(i);
                            chkRef = true;
                        }

                        // Get Name
                        if (vbfile[i].IndexOf("ExeName32=") > -1)
                        {
                            name = vbfile[i].Substring(vbfile[i].IndexOf('\"') + 1);
                            name = name.Substring(0, name.IndexOf('.'));
                            this.filename = name;
                            if (dlls[name.ToUpper()] == null)
                            {
                                editCompat = false;
                            }
                        }
                        else if (vbfile[i].IndexOf("Name=") > -1 && !(vbfile[i].Substring(0, vbfile[i].IndexOf('='))).Equals("VersionCompanyName"))
                        {
                            name = vbfile[i].Substring(vbfile[i].IndexOf('\"') + 1);
                            name = name.Substring(0, name.IndexOf('\"'));
                            this.filename = name;
                            if (dlls[name.ToUpper()] == null)
                            {
                                editCompat = false;
                            }
                        }

                        //Edit Compatible
                        if (vbfile[i].IndexOf("CompatibleMode") > -1 && editCompat)        
                        {
                            string tmp = "CompatibleMode=\"2\"";
                            text.Add(count, tmp);
                            count++;
                            tmp = "CompatibleEXE32=\"" + ((FormItem)dlls[name.ToUpper()]).ref_tmp_files + "\"";
                            text.Add(count, tmp);
                            count++;
                            text.Add(count, "VersionCompatible32=\"" + ((((int)(float.Parse(((FormItem)dlls[name.ToUpper()]).version) * 10)) % 10) + 1) + "\"");
                        }
                        else if (vbfile[i].IndexOf("CompatibleMode") > -1 && !editCompat)        
                        {
                            string tmp = "CompatibleMode=\"0\"";
                            text.Add(count, tmp);
                        }
                        else if (((vbfile[i].IndexOf("CompatibleEXE32=\"") > -1) || (vbfile[i].IndexOf("VersionCompatible32=\"") > -1)))
                        {
                            count--;
                            continue;
                        }
                        else
                        {
                            text.Add(count, vbfile[i]);
                        }
                    }
                }
                else if (compat.Equals("-n") && type)
                {
                    editCompat = false;
                    Console.WriteLine("CompatibleMode: No Compatible");
                    Console.Write("Edit VBProject: ");
                    for (int i = 0; i < vbfile.Length; i++, count++)
                    {
                        //Get Reference Line
                        if (vbfile[i].IndexOf(".vbp") > -1)
                        {
                            index_Ref.Add(i);
                            chkRef = true;
                        }
                        else if (vbfile[i].IndexOf("ExeName32=") > -1)
                        {
                            name = vbfile[i].Substring(vbfile[i].IndexOf('\"') + 1);
                            name = name.Substring(0, name.IndexOf('.'));
                            this.filename = name;
                        }

                        if (vbfile[i].IndexOf("CompatibleMode") > -1 && !editCompat)        
                        {
                            string tmp = "CompatibleMode=\"0\"";
                            text.Add(count, tmp);
                        }
                        else if (((vbfile[i].IndexOf("CompatibleEXE32=\"") > -1) || (vbfile[i].IndexOf("VersionCompatible32=\"") > -1)))
                        {
                            count--;
                            continue;
                        }
                        else
                        {
                            text.Add(count, vbfile[i]);
                        }
                    }
                }
                else
                {
                    editCompat = false;
                    Console.WriteLine("CompatibleMode: No Compatible Only");
                    Console.Write("Edit VBProject: ");
                    for (int i = 0; i < vbfile.Length; i++, count++)
                    {
                        //Get Reference Line
                        if (vbfile[i].IndexOf(".vbp") > -1)
                        {
                            index_Ref.Add(i);
                            chkRef = true;
                        }

                        if (vbfile[i].IndexOf("Title=") > -1)
                        {
                            name = vbfile[i].Substring(vbfile[i].IndexOf('\"') + 1);
                            name = name.Substring(0, name.IndexOf('\"'));
                            this.filename = name;
                        }

                        // Get Name
                        if (vbfile[i].IndexOf("ExeName32=") > -1)
                        {
                            name = "ExeName32=\"" + name + ".exe\"";
                            text.Add(count, name);
                        }
                        else if (vbfile[i].IndexOf("CompatibleMode") > -1 && !editCompat)
                        {
                            string tmp = "CompatibleMode=\"0\"";
                            text.Add(count, tmp);
                        }
                        else if (((vbfile[i].IndexOf("CompatibleEXE32=\"") > -1) || (vbfile[i].IndexOf("VersionCompatible32=\"") > -1)))
                        {
                            count--;
                            continue;
                        }
                        else if (vbfile[i].LastIndexOf("craxdui.dll") > -1)
                        {
                            string tmp;
                            if (check32Bit)
                            {
                                tmp = @"Reference=*\G{BD4B4E53-F7B8-11D0-964D-00A0C9273C2A}#b.0#0#C:\Program Files\Common Files\Crystal Decisions\3.0\bin\craxdui.dll#Crystal Reports ActiveX Designer Library 11.0";
                            }
                            else
                            {
                                tmp = @"Reference=*\G{BD4B4E53-F7B8-11D0-964D-00A0C9273C2A}#b.0#0#C:\Program Files (x86)\Common Files\Crystal Decisions\3.0\bin\craxdui.dll#Crystal Reports ActiveX Designer Library 11.0";
                            }
                            text.Add(count, tmp);
                        }
                        else if (vbfile[i].LastIndexOf("mscomctl.ocx") > -1 || vbfile[i].LastIndexOf(("mscomctl.ocx").ToUpper()) > -1 || vbfile[i].LastIndexOf("mscomctl.OCX") > -1 || vbfile[i].LastIndexOf("MSCOMCTL.ocx") > -1)
                        {
                            string tmp = @"Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; mscomctl.OCX";
                            text.Add(count, tmp);
                        }
                        else
                        {
                            text.Add(count, vbfile[i]);
                        }
                    }
                }

                if (chkRef)
                {
                    foreach (int item in index_Ref)
                    {
                        string line = text[item].ToString();
                        line = line.Substring(line.LastIndexOf('\\') + 1);
                        line = line.Substring(0, line.LastIndexOf('.'));
                        text[item] = "Reference=*\\G" + ((FormItem)dlls[line.ToUpper()]).guid + "#" + ((FormItem)dlls[line.ToUpper()]).version + "#" + ((FormItem)dlls[line.ToUpper()]).lcid + "#" + ((FormItem)dlls[line.ToUpper()]).ref_tmp_files + "#" + ((FormItem)dlls[line.ToUpper()]).filename + ".dll";
                    }
                }
                logFile[4] = true;
            }
            catch (Exception ex)
            {
                logFile[4] = false;
                message = ex.StackTrace;
            }
            finally
            {
                if (logFile[4])
                {
                    Console.WriteLine("Successful...");
                }
                else
                {
                    Console.WriteLine(message);
                }
            }

            return text;
        }

        private bool setValueDll(string build_path)
        {
            bool check = true;
            string[] tempdlls = System.IO.Directory.GetFiles(@"C:\tempFile_vbauto\");
            Console.Write("Generate File DLLs: ");
            try
            {
                Parallel.ForEach(tempdlls, (item) =>
                {
                    if (item.LastIndexOf(".dll") > -1)
                    {
                        FormItem dll = generateGUID(item);
                        dlls.Add(dll.filename.ToUpper(), dll);
                    }
                });
                logFile[3] = true;
            }
            catch (Exception ex)
            {
                message = ex.StackTrace;
                check = false;
                logFile[3] = false;
                throw ex;
            }
            finally
            {
                if (logFile[3])
                {
                    Console.WriteLine("Successful...");
                }
                else
                {
                    Console.WriteLine(message);
                }
            }
            return check;
        }

        private FormItem generateGUID(string file)
        {
            FormItem dll = new FormItem();
            TLI.TypeLibInfo guid = new TLI.TLIApplication().TypeLibInfoFromFile(file);
            dll.ref_tmp_files = file;
            dll.filename = guid.Name;
            dll.guid = guid.GUID.ToString();
            dll.version = guid.MajorVersion + "." + guid.MinorVersion;
            dll.lcid = guid.LCID.ToString();
            logFile[3] = true;
            return dll;
        }

        private bool createTempDLL(string ref_path)
        {
            string destPath = @"C:\tempFile_vbauto\";
            bool check = true;

            //Prepare Temp Folder
            if (System.IO.Directory.Exists(destPath))
            {
                Console.Write("Delete Old Temp File: ");
                try
                {
                    System.IO.Directory.Delete(destPath, true);
                    logFile[0] = true;
                }
                catch (System.IO.IOException ex)
                {
                    message = ex.StackTrace;
                    logFile[0] = false;
                    throw ex;
                }
                finally
                {
                    if (logFile[0])
                    {
                        Console.WriteLine("Successful...");
                    }
                    else
                    {
                        Console.WriteLine(message);
                    }
                }
            }

            Console.Write("Create New Temp File: ");
            System.IO.Directory.CreateDirectory(destPath);
            waitting(1, 0.5f);
            string[] sourceFiles = System.IO.Directory.GetFiles(ref_path);
            if (System.IO.Directory.Exists(destPath))
            {
                logFile[1] = true;
                Console.WriteLine("Successful...");
                foreach (string item in sourceFiles)
                {
                    System.IO.File.Copy(item, (destPath + (item.Substring(item.LastIndexOf('\\') + 1))));
                }
            }
            else
            {
                logFile[1] = false;
                Console.WriteLine("Failure...");
                check = false;
            }
            return check;
        }

        private bool waitTime(string path, float sec, int time)
        {
            bool check = false;
            for (int i = 0; i < time; i++)
            {
                waitting(1, sec);
                if (System.IO.File.Exists(path))
                {
                    check = true;
                    break;
                }
            }
            return check;
        }

        private void waitting(int time, float sec)
        {
            int t = (int)(sec * 1000);
            for (int i = 0; i < time; i++)
            {
                Thread.Sleep(t);
            }
        }
    }
}
