using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using FoundryRulesAndUnits.Extensions;

namespace VisioShapeExtractor
{
    /// <summary>
    /// Exports master stencil information to JSON for reference
    /// </summary>
    public static class MastersExporter
    {
        /// <summary>
        /// Exports the collection of master stencils to a JSON file
        /// </summary>
        /// <param name="masters">Dictionary of master stencils</param>
        /// <param name="vsdxPath">Path of the original VSDX file</param>
        public static void ExportToJson(Dictionary<string, MasterShapeInfo> masters, string vsdxPath)
        {
            if (masters == null || masters.Count == 0)
            {
                return;
            }
            
            try
            {
                // Define output path - same folder as VSDX but with _masters.json suffix
                var outputPath = vsdxPath.Replace(".vsdx", "_masters.json");
                
                // Create a structure to hold all masters by stencil
                var mastersByStencil = GroupMastersByStencil(masters.Values);
                
                // Serialize with pretty printing
                var options = new JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    MaxDepth = 128
                };
                
                string json = JsonSerializer.Serialize(mastersByStencil, options);
                File.WriteAllText(outputPath, json);
                
                $"Exported {masters.Count} master stencils to {outputPath}".WriteSuccess();
            }
            catch (Exception ex)
            {
                $"Error exporting masters: {ex.Message}".WriteError();
            }
        }
        
        /// <summary>
        /// Groups masters by their stencil name for better organization
        /// </summary>
        private static Dictionary<string, List<MasterShapeInfo>> GroupMastersByStencil(IEnumerable<MasterShapeInfo> masters)
        {
            var result = new Dictionary<string, List<MasterShapeInfo>>();
            
            foreach (var master in masters)
            {
                if (!result.ContainsKey(master.StencilName))
                {
                    result[master.StencilName] = new List<MasterShapeInfo>();
                }
                
                result[master.StencilName].Add(master);
            }
            
            return result;
        }
    }
}
