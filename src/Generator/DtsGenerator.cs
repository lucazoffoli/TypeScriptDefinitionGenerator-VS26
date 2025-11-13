using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.TextTemplating.VSHost;
using Microsoft.VisualStudio.Shell;

namespace TypeScriptDefinitionGenerator
{
    [Guid("d1e92907-20ee-4b6f-ba64-142297def4e4")]
    public sealed class DtsGenerator : BaseCodeGeneratorWithSite
    {
        public const string Name = nameof(DtsGenerator);
        public const string Description = "Automatically generates the .d.ts file based on the C#/VB model class.";

        private string originalExt { get; set; }

        public override string GetDefaultExtension()
        {
            if (Options.WebEssentials2015)
            {
                return originalExt + Constants.FileExtension;
            }
            else
            {
                return Constants.FileExtension;
            }
        }

        protected override byte[] GenerateCode(string inputFileName, string inputFileContent)
        {
            ProjectItem item = null;
            originalExt = Path.GetExtension(inputFileName);

            try
            {
                // All EnvDTE / VS COM calls must be performed on the UI thread. Use ThreadHelper
                // to switch to the main thread and perform FindProjectItem and subsequent
                // CodeModel access there to avoid COM/CLR interop corruption.
                ThreadHelper.JoinableTaskFactory.Run(async () =>
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    item = (Dte as DTE2).Solution.FindProjectItem(inputFileName);

                    if (item != null)
                    {   
                        // Keep generation on the UI thread because GenerationService uses EnvDTE CodeModel
                        // and ProjectItem properties which require the UI thread.
                        var dts = GenerationService.GenerateFromProjectItem(item);
                    }
                });

                if (item != null)
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                VSHelpers.WriteOnOutputWindow($"Error during custom tool generation for {inputFileName}: {ex}");
            }

            return new byte[0];
        }
    }
}
