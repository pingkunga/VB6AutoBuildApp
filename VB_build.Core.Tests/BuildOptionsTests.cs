using System;
using Xunit;

namespace VB_build.Core.Tests
{
    public class BuildOptionsTests
    {
        [Fact]
        public void Default_MatchesOriginalHardcodedPaths()
        {
            BuildOptions options = BuildOptions.Default;

            Assert.Equal(@"C:\tempFile_vbauto\", options.TempDllFolder);
            Assert.Equal(@"C:\BuildExe", options.BuildExeFolder);
            Assert.Equal(@"D:\", options.LogFolder);
            Assert.Equal(@"D:\verpatch.exe", options.VerpatchExePath);
            Assert.Equal(@"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", options.Vb6ExePathOn32BitOs);
            Assert.Equal(@"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE", options.Vb6ExePathOn64BitOs);
            Assert.Equal(@"C:\Program Files\Common Files\Crystal Decisions\3.0\bin\craxdui.dll", options.CraxduiPathOn32BitOs);
            Assert.Equal(@"C:\Program Files (x86)\Common Files\Crystal Decisions\3.0\bin\craxdui.dll", options.CraxduiPathOn64BitOs);
        }

        [Fact]
        public void CanOverrideIndividualPaths()
        {
            BuildOptions options = new BuildOptions
            {
                VerpatchExePath = @"E:\tools\verpatch.exe"
            };

            Assert.Equal(@"E:\tools\verpatch.exe", options.VerpatchExePath);
            // Other values remain the original defaults.
            Assert.Equal(@"C:\BuildExe", options.BuildExeFolder);
        }
    }

    public class ControllerConstructionTests
    {
        [Fact]
        public void ParameterlessConstructor_UsesDefaultOptions()
        {
            Controller controller = new Controller();

            Assert.NotNull(controller);
            Assert.Equal(10, controller.logFile.Length);
            Assert.All(controller.logFile, step => Assert.False(step));
        }

        [Fact]
        public void Constructor_WithNullOptions_FallsBackToDefault()
        {
            // Should not throw; behavior parity with the original hardcoded-path constructor.
            Controller controller = new Controller(null);

            Assert.NotNull(controller);
        }

        [Fact]
        public void Constructor_WithCustomOptions_DoesNotThrow()
        {
            BuildOptions options = new BuildOptions { LogFolder = @"E:\logs\" };

            Controller controller = new Controller(options);

            Assert.NotNull(controller);
        }
    }

    public class BuildResultTests
    {
        [Fact]
        public void Success_IsTrue_WhenNoErrors()
        {
            BuildResult result = new BuildResult { ErrorCount = 0 };

            Assert.True(result.Success);
        }

        [Fact]
        public void Success_IsFalse_WhenAnyErrors()
        {
            BuildResult result = new BuildResult { ErrorCount = 1 };

            Assert.False(result.Success);
        }
    }
}
