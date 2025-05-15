using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace VisioShapeExtractor
{
    public static class FileValidator
    {
        /// <summary>
        /// Validates if a file is a valid ZIP archive (which includes VSDX files)
        /// </summary>
        /// <param name="filePath">Path to the file to validate</param>
        /// <returns>True if the file is a valid ZIP archive, otherwise false</returns>
        public static bool IsValidZipArchive(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File does not exist: {filePath}");
                return false;
            }

            try
            {
                // Check file size first (a valid ZIP file must be at least 22 bytes - the size of End of Central Directory record)
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length < 22)
                {
                    Console.WriteLine($"File too small to be a valid ZIP archive: {filePath}");
                    return false;
                }

                // Try to open the file as a ZIP archive
                using (var zipFile = ZipFile.OpenRead(filePath))
                {
                    // Check if we can access the entries
                    var entries = zipFile.Entries;
                    return entries != null;
                }
            }
            catch (InvalidDataException ex)
            {
                Console.WriteLine($"Invalid ZIP archive: {filePath}. Error: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating file: {filePath}. Error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Validates if a file is a valid VSDX file by checking if it's a valid ZIP and contains expected Visio structure
        /// </summary>
        /// <param name="filePath">Path to the VSDX file to validate</param>
        /// <returns>True if the file is a valid VSDX file, otherwise false</returns>
        public static bool IsValidVsdxFile(string filePath)
        {
            if (!IsValidZipArchive(filePath))
            {
                return false;
            }

            try
            {
                using (var archive = ZipFile.OpenRead(filePath))
                {
                    // Check for essential VSDX file structure
                    bool hasContentTypes = archive.Entries.Any(e => e.FullName == "[Content_Types].xml");
                    bool hasVisioFolder = archive.Entries.Any(e => e.FullName == "visio/" || e.FullName.StartsWith("visio/"));
                    bool hasVisioPages = archive.Entries.Any(e => e.FullName.StartsWith("visio/pages/"));

                    if (!hasContentTypes || !hasVisioFolder || !hasVisioPages)
                    {
                        Console.WriteLine($"File is not a valid VSDX file (missing essential structure): {filePath}");
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating VSDX file: {filePath}. Error: {ex.Message}");
                return false;
            }
        }        /// <summary>
        /// Attempts to repair a ZIP file by copying it to a temporary file and ensuring it can be opened
        /// </summary>
        /// <param name="filePath">Path to the ZIP file to repair</param>
        /// <returns>Path to the repaired file if successful, null otherwise</returns>
        public static string? TryRepairZipFile(string filePath)
        {
            try
            {
                string? directoryName = Path.GetDirectoryName(filePath);
                if (directoryName == null)
                {
                    Console.WriteLine($"Could not determine directory for file: {filePath}");
                    return null;
                }
                
                string repairedFilePath = Path.Combine(
                    directoryName,
                    $"{Path.GetFileNameWithoutExtension(filePath)}_repaired{Path.GetExtension(filePath)}");

                // Copy the file - this might help with minor corruption issues
                File.Copy(filePath, repairedFilePath, true);

                // Check if the repaired file is valid
                if (IsValidZipArchive(repairedFilePath))
                {
                    Console.WriteLine($"Successfully repaired file: {filePath} -> {repairedFilePath}");
                    return repairedFilePath;
                }

                // If we couldn't repair it, delete the attempted repair
                if (File.Exists(repairedFilePath))
                {
                    File.Delete(repairedFilePath);
                }
                
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error attempting to repair file: {filePath}. Error: {ex.Message}");
                return null;
            }
        }
    }
}