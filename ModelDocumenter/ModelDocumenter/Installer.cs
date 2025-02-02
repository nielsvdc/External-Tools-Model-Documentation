using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.Json;
using System.Diagnostics;

namespace ModelDocumenter
{
    static class Installer
    {
        const string POWERBI_EXTERNAL_TOOL_PATH = "Microsoft Shared\\Power BI Desktop\\External Tools";

        static readonly Assembly execAssembly = Assembly.GetCallingAssembly();
        static readonly Version assemblyVersion = execAssembly.GetName().Version;
        static readonly string version = $"{assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.MajorRevision}";
        static readonly string assemblyName = execAssembly.GetName().Name.ToString();
        static readonly string productName = execAssembly.GetCustomAttributes(typeof(AssemblyProductAttribute), false).OfType<AssemblyProductAttribute>().FirstOrDefault().Product;
        static readonly string description = execAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false).OfType<AssemblyDescriptionAttribute>().FirstOrDefault().Description;
        static readonly string company = execAssembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false).OfType<AssemblyCompanyAttribute>().FirstOrDefault().Company;


        static public void Install(string vpaxFolderPath, string templateFilePath, string sourcePathToReplace)
        {
            ConfigureTemplateFile(vpaxFolderPath, templateFilePath, sourcePathToReplace);
            ConfigurePbiToolJson();
        }

        /// <summary>
        /// When uninstalling, remove the VPAX folder by code, as it's not deleted by the installer
        /// </summary>
        /// <param name="vpaxFolderPath"></param>
        static public void Uninstall(string vpaxFolderPath)
        {
            Directory.Delete(vpaxFolderPath, true);
        }

        /// <summary>
        /// Configure the Power BI template file
        /// </summary>
        static void ConfigureTemplateFile(string vpaxFolderPath, string templateFilePath, string sourcePathToReplace)
        {
            Console.WriteLine("Configuring template file...");
            ReplaceSourcePathInPbit(templateFilePath, sourcePathToReplace, vpaxFolderPath);
        }

        /// <summary>
        /// Create the content for the .pbitool.json file and save it
        /// </summary>
        static void ConfigurePbiToolJson()
        {
            Console.WriteLine("Generating Power BI external tool config file...");

            // Get exectuble icon bas64 string
            Icon icon = Icon.ExtractAssociatedIcon(execAssembly.Location);
            string base64String = String.Empty;
            using (Bitmap bitmap = icon.ToBitmap())
            {
                base64String = ConvertBitmapToBase64(bitmap);
            }

            // Config variables
            string arguments = $"--server \"%server%\" --database \"%database%\"";
            string pbiToolJsonFileName = $"{company}_{assemblyName}.pbitool.json";
            string pbiToolJsonPath = Path.Combine(
                Environment.ExpandEnvironmentVariables(@"%CommonProgramFiles(x86)%"),
                POWERBI_EXTERNAL_TOOL_PATH,
                pbiToolJsonFileName
            );

            // Save to .pbitool.json file
            SavePbiJson(version, productName, description, execAssembly.Location, arguments, base64String, pbiToolJsonPath);
        }

        /// <summary>
        /// Convert image to base64 string
        /// </summary>
        /// <param name="bitmap">Bitmap file to convert</param>
        /// <returns>Base64 string</returns>
        static string ConvertBitmapToBase64(Bitmap bitmap)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                byte[] byteArray = memoryStream.ToArray();
                string base64String = Convert.ToBase64String(byteArray);
                string output = $"data:image/png;base64,{base64String}";
                return output;
            }
        }

        /// <summary>
        /// Save .pbitool.json file for this external tool
        /// </summary>
        /// <param name="version">Version of the external tool</param>
        /// <param name="productName">Product name of the external tool</param>
        /// <param name="description">Description of the external tool</param>
        /// <param name="assemblyPath">Path to the executable of the external tool</param>
        /// <param name="arguments">Arguments added by Power BI when starting the tool</param>
        /// <param name="iconData">Base64 of image for the external</param>
        /// <param name="filePath">Path to save the .pbitool.json file</param>
        static void SavePbiJson(
            string version, string productName, string description,
            string assemblyPath, string arguments, string iconData, string filePath
            )
        {
            var pbiTool = new PbiTool
            {
                Version = version,
                ProductName = productName,
                Description = description,
                Path = assemblyPath,
                Arguments = arguments,
                IconData = iconData
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            string jsonString = System.Text.Json.JsonSerializer.Serialize(pbiTool, options);

            try
            {
                File.WriteAllText(filePath, jsonString);
            }
            catch (UnauthorizedAccessException)
            {
                // When in debug mode, you have no permission to write to Microsoft Shared folder, except when running VS in admin mode
                Console.WriteLine($"No permission to write file '{filePath}'. Run this application as administrator.");
            }
        }

        /// <summary>
        /// Class to save the .pbitool.json information as json string
        /// </summary>
        public class PbiTool
        {
            [JsonPropertyName("version")]
            public string Version { get; set; }
            [JsonPropertyName("name")]
            public string ProductName { get; set; }
            [JsonPropertyName("description")]
            public string Description { get; set; }
            [JsonPropertyName("path")]
            public string Path { get; set; }
            [JsonPropertyName("arguments")]
            public string Arguments { get; set; }
            [JsonPropertyName("iconData")]
            public string IconData { get; set; }
        }

        /// <summary>
        /// Replace the source path to the VPAX file in the Power BI template file
        /// </summary>
        /// <param name="templateFilePath">File path to the Power BI template file</param>
        /// <param name="sourcePathToReplace">Path to data VPAX file to replace in the template</param>
        /// <param name="newSourcePath">Path to the local user application path for this project as new source path to the VPAX file</param>
        static void ReplaceSourcePathInPbit(string templateFilePath, string sourcePathToReplace, string newSourcePath)
        {
            try
            {
                string fileToModify = "DataModelSchema";

                using (FileStream pbitStream = new FileStream(templateFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    using (ZipArchive archive = new ZipArchive(pbitStream, ZipArchiveMode.Update, true))
                    {
                        ZipArchiveEntry entry = archive.GetEntry(fileToModify);
                        if (entry != null)
                        {
                            // Use UTF-16 Little Endian encoding for the DataModelschema file
                            UnicodeEncoding encoding = new UnicodeEncoding(false, false);

                            string fileContent;
                            using (StreamReader reader = new StreamReader(entry.Open(), encoding))
                            {
                                fileContent = reader.ReadToEnd();
                            }

                            string modifiedContent = fileContent.Replace(sourcePathToReplace.Replace("\\", "\\\\"), newSourcePath.Replace("\\", "\\\\"));
                            entry.Delete();

                            ZipArchiveEntry newEntry = archive.CreateEntry(fileToModify, CompressionLevel.Optimal);
                            using (StreamWriter writer = new StreamWriter(newEntry.Open(), encoding))
                            {
                                writer.Write(modifiedContent);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
        }
    }
}
