using System;

namespace VB_build
{
    class Program
    {
        public static bool[] log;
        public static string err, starttime = DateTime.Now.ToString(), endtime;
        public static int er = 0;

        public static void Main(string[] args)
        {
            Controller con = new Controller();
            bool check = true;
            if (args[0].Equals("-D"))
            {
                Console.WriteLine("DLL time start: " + starttime);
                check = true;
            }
            else
            {
                Console.WriteLine("EXE time start: " + starttime);
                check = false;
            }
            
            try
            {
                // 0=type 1=tag-build 2=ref-path 3=version 4=compat 5=time
                int count = 6;
                string[] command = new string[args.Length - count];
                for (int i = 0; i < args.Length - count; i++)
                {
                    command[i] = args[(count + i)];
                }
                log = con.processBuild(check, args[1], args[2], args[3], args[4], args[5], command);
            }
            catch (Exception ex)
            {
                log = con.logFile;
                err = ex.Message;
                Console.WriteLine(err);
            }
            finally
            {
                endtime = DateTime.Now.ToString();
                Console.WriteLine("time end: " + endtime);
                er = getLog(args[0], args[3], args[4], args[1]);
                Environment.ExitCode = er;
                Environment.Exit(er);
            }
        }

        private static int getLog(string type, string version, string compat, string filename)
        {
            int error = 0;
            filename = filename.Substring((filename.Substring(0, filename.LastIndexOf('\\'))).LastIndexOf('\\') + 1);
            filename = filename.Substring(0, filename.IndexOf('\\')) + filename.Substring(filename.LastIndexOf('\\') + 1);
            string path = "D:\\log" + type + "_" + version + "_" + compat + "_" + filename + ".txt";
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }
            if ( !log[9] || !log[8] || !log[7] || !log[6] || !log[5] || !log[4] || !log[3] ||!log[2] || !log[1] || !log[0])
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(path);
                if (type.Equals("-D"))
                {
                    file.WriteLine("DLL index: " + filename + " time start: " + starttime);
                }
                else if (type.Equals("-E"))
                {
                    file.WriteLine("EXE index: " + filename + " time start: " + starttime);
                }
                file.WriteLine("Delete Old Temp File: " + checkError(0));
                file.WriteLine("Create New Temp File: " + checkError(1));
                file.WriteLine("OS-32bits: " + checkError(2));
                file.WriteLine("Generate File DLLs: " + checkError(3));
                if (type.Equals("-D"))
                {
                    if (compat.Equals("-b"))
                    {
                        file.WriteLine("CompatibleMode: Binary Compatible");
                    }
                    else
                    {
                        file.WriteLine("CompatibleMode: No Compatible");
                    }
                }
                else
                {
                    file.WriteLine("CompatibleMode: No Compatible Only");
                }
                file.WriteLine("Edit VBProject: " + checkError(4));
                file.WriteLine("Write Data to VBProject: " + checkError(5));
                file.WriteLine("Build Project: " + checkError(6));
                file.WriteLine("Remove Old file: " + checkError(7));
                file.WriteLine("Move Result file to Target Folder: " + checkError(8));
                file.WriteLine("Change Version: " + checkError(9));
                file.WriteLine("All Status: " + checkError(9));

                file.WriteLine("time end: " + endtime);
                Console.WriteLine("time end: " + endtime);
                file.WriteLine("========================================================================================");
                Console.WriteLine("========================================================================================");
                file.Close();
            }

            for (int i = 0; i < log.Length; i++)
            {
                if ((!log[i]))
                {
                    error++;
                }
            }
           return error;
        }

        private static string checkError (int index)
        {
            if (index == 11 && !log[index])
            {
                return "Failure...";
            }
            else if (log[index])
            {
                return "Successful...";
            }
            else if (err != null)
            {
                string str = err;
                err = null;
                return str;
            }
            else
            {
                return "Failure...";
            }
        }
    }
}
