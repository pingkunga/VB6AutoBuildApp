using System;

namespace VB_build
{
    /// <summary>
    /// Default <see cref="IBuildRunner"/> implementation: wraps <see cref="Controller.processBuild"/>
    /// (the same logic the original console app used) with a <see cref="BuildOptions"/> that can
    /// come from configuration.
    /// </summary>
    public class VbBuildRunner : IBuildRunner
    {
        private readonly BuildOptions options;

        public VbBuildRunner(BuildOptions options)
        {
            this.options = options ?? BuildOptions.Default;
        }

        public BuildResult Run(BuildRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            Controller controller = new Controller(this.options);
            string errorMessage = null;
            bool[] log;

            try
            {
                log = controller.processBuild(
                    request.IsDll,
                    request.BuildPath,
                    request.RefPath,
                    request.Version,
                    request.Compat,
                    request.Time,
                    request.Command ?? Array.Empty<string>());
            }
            catch (Exception ex)
            {
                log = controller.logFile;
                errorMessage = ex.Message;
            }

            int errorCount = 0;
            foreach (bool step in log)
            {
                if (!step)
                {
                    errorCount++;
                }
            }

            return new BuildResult
            {
                StepResults = log,
                ErrorCount = errorCount,
                ErrorMessage = errorMessage
            };
        }
    }
}
