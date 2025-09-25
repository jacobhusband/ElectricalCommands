using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Reflection;

// This attribute tells AutoCAD to run this class when the plugin is loaded.
[assembly: ExtensionApplication(typeof(AutoCADCleanupTool.PluginInitializer))]

namespace AutoCADCleanupTool
{
    public class PluginInitializer : IExtensionApplication
    {
        /// <summary>
        /// This method is called once when the plugin is loaded by AutoCAD.
        /// </summary>
        public void Initialize()
        {
            // Attach our custom assembly resolver to the current application domain.
            // This event will fire whenever the .NET runtime fails to load an assembly.
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolver;
        }

        /// <summary>
        /// This method is called once when the plugin is unloaded by AutoCAD.
        /// </summary>
        public void Terminate()
        {
            // It's good practice to detach the event handler when the application closes.
            AppDomain.CurrentDomain.AssemblyResolve -= AssemblyResolver;
        }

        /// <summary>
        /// Handles the event when an assembly fails to load, and helps the runtime find it.
        /// </summary>
        private Assembly AssemblyResolver(object sender, ResolveEventArgs args)
        {
            // Get the name of the assembly that failed to load (e.g., "Magick.NET.Core").
            var assemblyName = new AssemblyName(args.Name).Name;

            // We are only interested in resolving Magick.NET related assemblies.
            if (!assemblyName.StartsWith("Magick.NET", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            try
            {
                // Get the full path of our main plugin DLL.
                string mainAssemblyPath = Assembly.GetExecutingAssembly().Location;
                string mainAssemblyDir = Path.GetDirectoryName(mainAssemblyPath);

                // Construct the expected path to the required DLL, assuming it's in the same directory.
                string assemblyToLoadPath = Path.Combine(mainAssemblyDir, assemblyName + ".dll");

                // Check if the file exists and, if so, load it.
                if (File.Exists(assemblyToLoadPath))
                {
                    return Assembly.LoadFrom(assemblyToLoadPath);
                }
            }
            catch (System.Exception ex)
            {
                // Log any errors to the AutoCAD command line for easier debugging.
                var editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
                editor?.WriteMessage($"\nAssemblyResolver Error: {ex.Message}");
            }

            // Return null if we couldn't find the assembly.
            return null;
        }
    }
}