using System;

namespace VB_build
{
    /// <summary>
    /// Configurable filesystem locations used by <see cref="Controller"/>.
    /// These were previously hardcoded absolute paths in the build logic; they are now
    /// externalized so they can be bound from configuration (appsettings.json, environment
    /// variables, etc.) by the console app / Web API hosts. The <see cref="Default"/> values
    /// match the original hardcoded paths exactly, so behavior is unchanged unless a host
    /// explicitly overrides them.
    /// </summary>
    public class BuildOptions
    {
        /// <summary>Folder scanned/recreated for temporary DLL type-library inspection (was "C:\tempFile_vbauto\").</summary>
        public string TempDllFolder { get; set; } = @"C:\tempFile_vbauto\";

        /// <summary>Folder that built EXE output is moved into (was "C:\BuildExe").</summary>
        public string BuildExeFolder { get; set; } = @"C:\BuildExe";

        /// <summary>Folder that per-build log text files are written to (was "D:\log...").</summary>
        public string LogFolder { get; set; } = @"D:\";

        /// <summary>Full path to verpatch.exe used to stamp version info (was "D:\verpatch.exe").</summary>
        public string VerpatchExePath { get; set; } = @"D:\verpatch.exe";

        /// <summary>VB6.EXE path when the OS is 32-bit (was "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE").</summary>
        public string Vb6ExePathOn32BitOs { get; set; } = @"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE";

        /// <summary>VB6.EXE path when the OS is 64-bit (was "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE").</summary>
        public string Vb6ExePathOn64BitOs { get; set; } = @"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE";

        /// <summary>craxdui.dll reference path when the OS is 32-bit.</summary>
        public string CraxduiPathOn32BitOs { get; set; } = @"C:\Program Files\Common Files\Crystal Decisions\3.0\bin\craxdui.dll";

        /// <summary>craxdui.dll reference path when the OS is 64-bit.</summary>
        public string CraxduiPathOn64BitOs { get; set; } = @"C:\Program Files (x86)\Common Files\Crystal Decisions\3.0\bin\craxdui.dll";

        /// <summary>Returns a new <see cref="BuildOptions"/> populated with the original hardcoded defaults.</summary>
        public static BuildOptions Default => new BuildOptions();
    }
}
