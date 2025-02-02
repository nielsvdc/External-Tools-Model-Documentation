using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;

// TODO
// - Import from DMV 1100 (check for missing attributes?)
#pragma warning disable IDE0051 // Remove unused private members
namespace ModelDocumenter
{
    public enum ModelDocumenterErrEnum
    {
        [Description("The operation completed successfully.")]
        ERROR_SUCCESS,
        [Description("One or more arguments are not correct.")]
        ERROR_BAD_ARGUMENTS,
        [Description("This server is not supported.")]
        ERROR_NOT_SUPPORTED,
        [Description("Error exporting vpax file.")]
        ERROR_EXPORT_Vpax,
        [Description("Error installing program.")]
        ERROR_INSTALL
    }

    class Program
    {
        const string SUBFOLDER_VPAXOUTPUT = "VpaxOutput";
        const string SUBFOLDER_TEMPLATE = "Template";
        const string TEMPLATE_FILENAME = "ModelDocumentationTemplate.pbit";
        const string SOURCEPATH_IN_TEMPLATE = @"C:\Power BI Model Documenter";

        static readonly Assembly execAssembly = Assembly.GetExecutingAssembly() ?? Assembly.GetEntryAssembly();
        static readonly AssemblyName assemblyName = execAssembly.GetName() ?? Assembly.GetEntryAssembly().GetName();
        static readonly string assemblyFolderPath = Path.GetDirectoryName(execAssembly.Location);

        static readonly string applicationName = assemblyName.Name;
        static readonly string applicationVersion = $"{assemblyName.Version.Major:0}.{assemblyName.Version.Minor:0}.{assemblyName.Version.MajorRevision:0}";

        internal class ArgumentOptions
        {
            [Option('i', "install", Required = false, HelpText = "This parameter specifies is used for post installation actions.")]
            public bool Install { get; set; }
            [Option('u', "uninstall", Required = false, HelpText = "This parameter specifies is used for uninstallation actions.")]
            public bool Uninstall { get; set; }
            [Option('s', "server", Required = false, HelpText = "This parameter specifies the server name.")]
            public string Server { get; set; }
            [Option('d', "database", Required = false, HelpText = "This parameter specifies the database name.")]
            public string Database { get; set; }
        }

        static void Main(string[] args)
        {
            Console.WriteLine($"---");
            Console.WriteLine($"--- {applicationName} {applicationVersion} for.Net ({execAssembly.ImageRuntimeVersion})");
            Console.WriteLine($"---");

            var plugins = LoadPlugins();

            ModelDocumenterErrEnum exitCode = ModelDocumenterErrEnum.ERROR_SUCCESS;

            var parser = Parser.Default.ParseArguments<ArgumentOptions>(args);
            ArgumentOptions options = null;
            parser.WithParsed(opts =>
                {
                    options = opts;
                })
                .WithNotParsed(errors =>
                {
                    var helpText = HelpText.AutoBuild(parser, h =>
                    {
                        h.AdditionalNewLineAfterOption = false;
                        return HelpText.DefaultParsingErrorsHandler(parser, h);
                    }, e => e);

                    Console.WriteLine(helpText);
                    exitCode = ModelDocumenterErrEnum.ERROR_BAD_ARGUMENTS;
                });

            if (exitCode == ModelDocumenterErrEnum.ERROR_SUCCESS)
            {
                string vpaxFolderPath = Path.Combine(assemblyFolderPath, SUBFOLDER_VPAXOUTPUT);
                string vpaxFilePath = Path.Combine(vpaxFolderPath, $"{applicationName}.vpax");
                string templateFolderPath = Path.Combine(assemblyFolderPath, SUBFOLDER_TEMPLATE);
                string templateFilePath = Path.Combine(templateFolderPath, TEMPLATE_FILENAME);

                if (options.Install)
                {
                    exitCode = RunInstall(vpaxFolderPath, templateFilePath);
                }
                else if (options.Uninstall)
                {
                    RunUninstall(vpaxFolderPath);
                }
                else
                {
                    exitCode = RunProgram(options, vpaxFolderPath, vpaxFilePath, templateFilePath);
                }
            }

            if (exitCode != ModelDocumenterErrEnum.ERROR_SUCCESS)
            {
                Console.WriteLine("Error 0x{0:x4}: {1}", (int)exitCode, exitCode.GetDescription());
                Console.ReadLine();
            }
            Environment.Exit((int)exitCode);
        }

        /// <summary>
        /// Run install procedure
        /// </summary>
        /// <param name="vpaxFolderPath">Path of the VPAX folder</param>
        /// <param name="templateFilePath">File path of the Power BI template</param>
        /// <returns></returns>
        static ModelDocumenterErrEnum RunInstall(string vpaxFolderPath, string templateFilePath)
        {
            try
            {
                Installer.Install(vpaxFolderPath, templateFilePath, SOURCEPATH_IN_TEMPLATE);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e}");
                return ModelDocumenterErrEnum.ERROR_INSTALL;
            }
            return ModelDocumenterErrEnum.ERROR_SUCCESS;
        }

        /// <summary>
        /// Run uninstall procedure
        /// </summary>
        /// <param name="vpaxFolderPath"></param>
        static void RunUninstall(string vpaxFolderPath)
        {
            Installer.Uninstall(vpaxFolderPath);
        }

        /// <summary>
        /// Run program
        /// </summary>
        /// <param name="options">Argument options</param>
        /// <param name="vpaxFolderPath">Path of the VPAX folder</param>
        /// <param name="vpaxFilePath">File path of the VPAX file</param>
        /// <param name="pbiTemplateFilePath">File path of the Power BI template</param>
        /// <returns></returns>
        static ModelDocumenterErrEnum RunProgram(ArgumentOptions options, string vpaxFolderPath, string vpaxFilePath, string pbiTemplateFilePath)
        {
            ModelDocumenterErrEnum exitCode;
            if (options.Server.StartsWith("pbiazure://api.powerbi.com"))
            {
                Console.WriteLine($"Server {options.Server} not supported.");
                exitCode = ModelDocumenterErrEnum.ERROR_NOT_SUPPORTED;
            }
            else
            {
                exitCode = VpaxExport(options.Server, options.Database, vpaxFolderPath, vpaxFilePath);
            }

            if (exitCode == ModelDocumenterErrEnum.ERROR_SUCCESS)
            {
                // Open Power BI template file
                System.Diagnostics.Process.Start(pbiTemplateFilePath);
            }
            return exitCode;
        }

        /// <summary>
        /// Load executable references
        /// </summary>
        /// <returns></returns>
        static IList<Assembly> LoadPlugins()
        {
            List<Assembly> pluginAssemblies = new List<Assembly>();

            foreach (var dll in Directory.EnumerateFiles(AppDomain.CurrentDomain.BaseDirectory, "*.dll"))
            {
                try
                {
                    var pluginAssembly = Assembly.LoadFile(dll);
                    if (pluginAssembly != null)
                    {
                        pluginAssemblies.Add(pluginAssembly);
                        //Console.WriteLine("Successfully loaded plugin " + pluginAssembly.FullName + " from assembly " + Path.GetFileName(dll));
                    }
                }
                catch
                {

                }
            }

            return pluginAssemblies;
        }

        /// <summary>
        /// Export VPAX database with metadata information of the tabular model from Power BI
        /// </summary>
        /// <param name="serverName"></param>
        /// <param name="databaseName"></param>
        /// <param name="vpaxFolderPath"></param>
        /// <param name="vpaxFilePath"></param>
        /// <returns></returns>
        static ModelDocumenterErrEnum VpaxExport(string serverName, string databaseName, string vpaxFolderPath, string vpaxFilePath)
        {
            ModelDocumenterErrEnum exitCode = ModelDocumenterErrEnum.ERROR_SUCCESS;
            bool includeTomModel = true;

            // Create log filename, use filename with .log extension
            if (!Directory.Exists(vpaxFolderPath)) { Directory.CreateDirectory(vpaxFolderPath); }
            string logFile = Path.Combine(vpaxFolderPath, Path.GetFileName(Path.ChangeExtension(vpaxFilePath, ".log")));
            
            try
            {
                Log("Exporting...", logFile);
                Log($"[INFO] Connecting to servername: {serverName}, databaseName: {databaseName}", logFile);
                //
                // Get Dax.Model object from the SSAS engine
                //
                Log("[INFO] Get Dax.Model object from the SSAS engine", logFile);
                Dax.Metadata.Model model = Dax.Model.Extractor.TomExtractor.GetDaxModel(serverName, databaseName, applicationName, applicationVersion);

                //
                // Get TOM model from the SSAS engine
                //
                Log("[INFO] Get TOM model from the SSAS engine", logFile);
                Microsoft.AnalysisServices.Tabular.Database database = includeTomModel ? Dax.Model.Extractor.TomExtractor.GetDatabase(serverName, databaseName) : null;

                // 
                // Create VertiPaq Analyzer views
                //
                Log("[INFO] Create VertiPaq Analyzer views", logFile);
                Dax.ViewVpaExport.Model viewVpa = new Dax.ViewVpaExport.Model(model);

                //
                // Save VPAX file
                // 
                // TODO: export of database should be optional
                Log("[INFO] Exporting VPAX file", logFile);
                Dax.Vpax.Tools.VpaxTools.ExportVpax(vpaxFilePath, model, viewVpa, database);

                Log("Completed", logFile);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Log("[INFO] Exporting VPAX file", logFile);
                exitCode = ModelDocumenterErrEnum.ERROR_EXPORT_Vpax;
            }

            return exitCode;
        }

        public static void Log(string logMessage, string logFile)
        {
            using (StreamWriter w = File.AppendText(logFile))
            {
                w.Write($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
                w.WriteLine($": {logMessage}");
            }
        }
    }
}