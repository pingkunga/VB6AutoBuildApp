namespace VB_build
{
    /// <summary>
    /// Result of a single build run: the raw per-step success flags (same order/meaning as
    /// <c>Controller.logFile</c>: 0=Delete Old Temp File .. 9=Change Version), plus a
    /// convenience <see cref="Success"/>/<see cref="ErrorCount"/> summary and any captured error.
    /// </summary>
    public class BuildResult
    {
        public bool[] StepResults { get; set; }

        public int ErrorCount { get; set; }

        public bool Success => ErrorCount == 0;

        public string ErrorMessage { get; set; }
    }
}
