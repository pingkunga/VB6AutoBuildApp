using System;

namespace VB_build
{
    /// <summary>
    /// Strongly-typed equivalent of the original CLI arguments
    /// (type, tag-build/build_path, ref-path, version, compat, time, extra VB6 command args).
    /// Used by both the console app and the Web API so the same request shape can be
    /// constructed from argv or from an HTTP JSON payload.
    /// </summary>
    public class BuildRequest
    {
        /// <summary>true = DLL build ("-D"), false = EXE build ("-E").</summary>
        public bool IsDll { get; set; }

        /// <summary>Folder containing the .vbp project to build (was args[1]).</summary>
        public string BuildPath { get; set; }

        /// <summary>Reference folder containing dependency DLLs (was args[2]).</summary>
        public string RefPath { get; set; }

        /// <summary>Version string passed to verpatch (was args[3]).</summary>
        public string Version { get; set; }

        /// <summary>Compatibility mode: "-b" (binary compatible) or "-n" (no compatible) (was args[4]).</summary>
        public string Compat { get; set; }

        /// <summary>Timeout multiplier used while waiting for build output (was args[5]).</summary>
        public string Time { get; set; }

        /// <summary>Additional raw VB6.EXE command-line arguments (was args[6..]).</summary>
        public string[] Command { get; set; } = Array.Empty<string>();
    }
}
