namespace VB_build.Api.Models
{
    /// <summary>
    /// HTTP request body for <c>POST /api/build/trigger</c>. Mirrors the original CLI arguments
    /// (type, tag-build/build_path, ref-path, version, compat, time, extra VB6 command args) so a
    /// Jenkins pipeline can pass the same information it used to pass on the command line.
    /// </summary>
    public class BuildTriggerRequest
    {
        /// <summary>true = DLL build, false = EXE build.</summary>
        public bool IsDll { get; set; }

        public string BuildPath { get; set; } = string.Empty;

        public string RefPath { get; set; } = string.Empty;

        public string Version { get; set; } = string.Empty;

        /// <summary>"-b" (binary compatible) or "-n" (no compatible).</summary>
        public string Compat { get; set; } = string.Empty;

        /// <summary>Timeout multiplier used while waiting for build output.</summary>
        public string Time { get; set; } = string.Empty;

        /// <summary>Additional raw VB6.EXE command-line arguments.</summary>
        public string[] Command { get; set; } = System.Array.Empty<string>();
    }
}
