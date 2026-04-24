using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace LicenceValidator
{
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "License Validator"),
        ExportMetadata("Description", "Validates Dynamics 365 user license assignments against actual rights, usage and Graph data. Shows optimization opportunities and exports results to Excel."),
        ExportMetadata("Author", "Martin Jäger, Consultics AG"),
        ExportMetadata("Company", "Consultics AG"),
        ExportMetadata("SmallImageBase64", "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"),
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "#0078D4"),
        ExportMetadata("PrimaryFontColor", "White"),
        ExportMetadata("SecondaryFontColor", "#CCE4F7")]
    public class LicenseValidatorPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new LicenseValidatorControl();
        }

        public LicenseValidatorPlugin()
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            if (refAssembly != null)
            {
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}
