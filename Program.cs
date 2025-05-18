using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using System.Xml;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using DocumentFormat.OpenXml.Drawing.ChartDrawing;
using FoundryRulesAndUnits.Extensions;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor;
public class Program
{
    // Shape information is now stored in BaseShape and its derived classes: Shape1D and Shape2D
    // See BaseShape.cs, Shape1D.cs, Shape2D.cs, and VisioShapes.cs


    static void Main(string[] args)
    {
        var dir = Directory.GetCurrentDirectory();
        dir = dir.Replace(@"bin\Debug\net9.0", "");
            
        try
        {            // Use a default file in the Documents folder if no arguments provided
            string documentsFolder = Path.Combine(dir, "Documents");
            string drawingsFolder = Path.Combine(dir, "DocumentDrawings");
            
            List<string> vsdxFilesList = new List<string>();
            
            // Add files from Documents folder
            if (Directory.Exists(documentsFolder))
            {
                vsdxFilesList.AddRange(Directory.GetFiles(documentsFolder, "*.vsdx"));
            }
            
            // Add files from DocumentDrawings folder
            // if (Directory.Exists(drawingsFolder))
            // {
            //     vsdxFilesList.AddRange(Directory.GetFiles(drawingsFolder, "*.vsdx"));
            // }
            
            string[] vsdxFiles = vsdxFilesList.ToArray();

            if (vsdxFiles == null || vsdxFiles.Length == 0)
            {
                Console.WriteLine("No VSDX files found to process.");
                return;
            }

            foreach (var originalVsdxPath in vsdxFiles)
            {
                try
                {
                    // Create a local variable for the path that might be repaired
                    string vsdxPath = originalVsdxPath;
                    
                    $"Processing Visio file: {vsdxPath}".WriteInfo();
                    
                    var shapeInfos = ExtractShapeInformation(vsdxPath, documentsFolder);
                    
                    if (shapeInfos.Count > 0)
                    {
                        //this function will remap shapeInfos to the into Shape2D and Shape1D items then export as json
                        TransformShapeInformation(shapeInfos, vsdxPath);
                        ExportToCsv(shapeInfos, vsdxPath);
                        $"Processed {shapeInfos.Count} shapes.".WriteSuccess();
                    }
                    else
                    {
                        Console.WriteLine($"No shapes extracted from {vsdxPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file: {originalVsdxPath}");
                    Console.WriteLine($"Error details: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                    // Continue with the next file
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }    private static void TransformShapeInformation(List<ShapeInfo> shapeInfos, string vsdxPath)
    {
        //this function will remap shapeInfos to the into Shape2D and Shape1D items then export as json
        var shape2DList = new Dictionary<string, Shape2D>();
        var shape1DList = new Dictionary<string, Shape1D>();

        // Track statistics for logging purposes
        int shapesWithConnections = 0;
        int shapesWithLayers = 0;
        int shapesWithNoSpecialInfo = 0;
        int totalConnectionPoints = 0;

        // First, count total shapes
        int totalShapes = shapeInfos.Count;
        $"Total shapes being processed: {totalShapes}".WriteInfo();

        // First pass: Create all shapes
        foreach (var shapeInfo in shapeInfos)
        {
            // Track statistics but process ALL shapes
            bool hasConnectionPoints = shapeInfo.ConnectionPointsArray.Count > 0;
            bool hasLayers = shapeInfo.Layers.Count > 0;
            
            if (hasConnectionPoints)
            {
                shapesWithConnections++;
                totalConnectionPoints += shapeInfo.ConnectionPointsArray.Count;
            }
            
            if (hasLayers) shapesWithLayers++;
            if (!hasConnectionPoints && !hasLayers) shapesWithNoSpecialInfo++;

            if (shapeInfo.Is1DShape)
            {
                // Create a new Shape1D object
                var shape1D = new Shape1D
                {
                    Id = shapeInfo.ShapeId,
                    Text = shapeInfo.ShapeText,
                    Master = shapeInfo.MasterName,
                    Type = shapeInfo.ShapeType,
                    ParentId = shapeInfo.ParentShapeId,
                    FromShapeId = shapeInfo.BeginConnectedTo,
                    ToShapeId = shapeInfo.EndConnectedTo,
                    FromCell = shapeInfo.ConnectionPoints,
                    ToCell = shapeInfo.ConnectionPoints,
                    BeginX = shapeInfo.BeginX,
                    BeginY = shapeInfo.BeginY,
                    EndX = shapeInfo.EndX,
                    EndY = shapeInfo.EndY,
                    ConnectionPoints = shapeInfo.ConnectionPointsArray.ToList(),
                    Layers = shapeInfo.Layers.ToList(),
                    LayerMembership = shapeInfo.LayerMembership
                };
                
                // Parse shape data if available
                if (!string.IsNullOrEmpty(shapeInfo.ShapeData))
                {
                    shape1D.ParseShapeData(shapeInfo.ShapeData);
                }
                
                // Add to the dictionary
                if (!shape1DList.ContainsKey(shape1D.Id))
                {
                    shape1DList[shape1D.Id] = shape1D;
                }
            }
            else
            {
                // Create a new Shape2D object
                var shape2D = new Shape2D
                {
                    Id = shapeInfo.ShapeId,
                    Text = shapeInfo.ShapeText,
                    Master = shapeInfo.MasterName,
                    Type = shapeInfo.ShapeType,
                    ParentId = shapeInfo.ParentShapeId,
                    PinX = shapeInfo.PositionX,
                    PinY = shapeInfo.PositionY,
                    Width = shapeInfo.Width,
                    Height = shapeInfo.Height,
                    ConnectionPoints = shapeInfo.ConnectionPointsArray.ToList(),
                    Layers = shapeInfo.Layers.ToList(),
                    LayerMembership = shapeInfo.LayerMembership
                };
                
                // Parse shape data if available
                if (!string.IsNullOrEmpty(shapeInfo.ShapeData))
                {
                    shape2D.ParseShapeData(shapeInfo.ShapeData);
                }

                // Add to the dictionary
                if (!shape2DList.ContainsKey(shape2D.Id))
                {
                    shape2DList[shape2D.Id] = shape2D;
                }
            }
        }

        // Second pass: Process connection points and relationships
        foreach (var connector in shape1DList.Values)
        {
            // For each connector, identify its endpoints and related connection points
            if (!string.IsNullOrEmpty(connector.FromShapeId) && shape2DList.TryGetValue(connector.FromShapeId, out var sourceShape))
            {
                // Try to find the specific connection point that this connector attaches to
                var sourceConnectionPoint = connector.GetSourceConnectionPoint(sourceShape);
                if (sourceConnectionPoint != null)
                {
                    $"Connector {connector.Id} connects from shape {sourceShape.Id} at connection point {sourceConnectionPoint.Id} ({sourceConnectionPoint.X}, {sourceConnectionPoint.Y})".WriteNote();
                }
            }
            
            if (!string.IsNullOrEmpty(connector.ToShapeId) && shape2DList.TryGetValue(connector.ToShapeId, out var targetShape))
            {
                // Try to find the specific connection point that this connector attaches to
                var targetConnectionPoint = connector.GetTargetConnectionPoint(targetShape);
                if (targetConnectionPoint != null)
                {
                    $"Connector {connector.Id} connects to shape {targetShape.Id} at connection point {targetConnectionPoint.Id} ({targetConnectionPoint.X}, {targetConnectionPoint.Y})".WriteNote();
                }
            }
        }

        // use the information in the dictionary to move child shapes to the parent shape
        foreach (var shape2D in shape2DList.Values)
        {
            if (!string.IsNullOrEmpty(shape2D.ParentId) && shape2DList.ContainsKey(shape2D.ParentId))
            {
                var parentShape = shape2DList[shape2D.ParentId];
                parentShape.AddSubShape(shape2D);
            }
        }
        // now do that for 1D shapes    
        foreach (var shape1D in shape1DList.Values)
        {
            if (!string.IsNullOrEmpty(shape1D.ParentId) && shape2DList.ContainsKey(shape1D.ParentId))
            {
                var parentShape = shape2DList[shape1D.ParentId];
                parentShape.AddSubShape(shape1D);
            }
        }

        //now remove shapes from the dictionary if they have a parent
        foreach (var shape2D in shape2DList.Values.ToList())
        {
            if (!string.IsNullOrEmpty(shape2D.ParentId) && shape2DList.ContainsKey(shape2D.ParentId))
            {
                shape2DList.Remove(shape2D.Id);
            }
        }
        
        // Generate connection point usage statistics
        Dictionary<string, int> connectionPointStats = new Dictionary<string, int>();
        foreach (var shape2D in shape2DList.Values)
        {
            var organizedPoints = shape2D.GetOrganizedConnectionPoints();
            foreach (var position in organizedPoints.Keys)
            {
                if (!connectionPointStats.ContainsKey(position))
                {
                    connectionPointStats[position] = 0;
                }
                connectionPointStats[position] += organizedPoints[position].Count;
            }
        }

        VisioShapes visioShapes = new VisioShapes
        {
            filename = vsdxPath,
            Shape2D = shape2DList.Values.ToList(),
            Shape1D = shape1DList.Values.ToList()
        };        // Analyze shape relationships
        AnalyzeShapeRelationships(shape2DList.Values, shape1DList.Values);
        
        // Export the lists to JSON files
        var outputPath = vsdxPath.Replace(".vsdx", ".json");
        var data = DehydrateShapes<VisioShapes>(visioShapes);
        File.WriteAllText(outputPath, data);

        // Output summary statistics
        $"Shape processing summary:".WriteSuccess();
        $"  Total shapes processed: {totalShapes}".WriteInfo();
        $"  Shapes with connection points: {shapesWithConnections}".WriteInfo();
        $"  Total connection points: {totalConnectionPoints}".WriteInfo();
        
        // Display connection point position distribution
        if (connectionPointStats.Count > 0)
        {
            $"  Connection points by position:".WriteInfo();
            foreach (var stat in connectionPointStats)
            {
                $"    {stat.Key}: {stat.Value}".WriteInfo();
            }
        }
        
        $"  Shapes with layer information: {shapesWithLayers}".WriteInfo();
        //$"  Shapes with no connection points or layers: {shapesWithNoSpecialInfo}".WriteInfo();
        $"  Root 2D shapes in output: {shape2DList.Count}".WriteInfo();
        $"  Root 1D shapes in output: {shape1DList.Count}".WriteInfo();
    }    public static string DehydrateShapes<T>(T target) where T : class
    {
        // Use our custom serializer that properly excludes defaults
        return JsonExcludeDefaultSerializer.Serialize(target);
    }

    static void ExportToCsv(List<ShapeInfo> shapeInfos, string outputPath)
    {
        var outputCsvPath = outputPath.Replace(".vsdx", ".csv");

        // Ensure directory exists
        var directory = Path.GetDirectoryName(outputCsvPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Write to CSV - using outputCsvPath instead of outputPath
        using (var writer = new StreamWriter(outputCsvPath))
        using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
        {
            csv.WriteRecords(shapeInfos);
        }
    }

    static List<ShapeInfo> ExtractShapeInformation(string vsdxPath, string xmlOutputDir)
    {
        List<ShapeInfo> shapeInfos = new List<ShapeInfo>();

        try
        {
            // First, verify the file is valid before attempting to open it
            if (!FileValidator.IsValidZipArchive(vsdxPath))
            {
                throw new InvalidDataException($"The file '{vsdxPath}' is not a valid ZIP archive or is corrupt.");
            }
            


            // Use ZipFile to extract the VSDX (which is a ZIP file) and process the XML directly
            using (var archive = ZipFile.OpenRead(vsdxPath))
            {
                // Save all XML files from the archive
                foreach (var entry in archive.Entries.Where(e => e.FullName.EndsWith(".xml")))
                {
                    string outputPath = Path.Combine(xmlOutputDir, entry.FullName.Replace('/', Path.DirectorySeparatorChar));
                    
                    // Ensure directory for this file exists
                    Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? "");
                    
                    // Extract and save the XML file
                    //entry.ExtractToFile(outputPath, overwrite: true);
                }                  
                // 
                // First, get all master information (stencils)
                Dictionary<string, MasterShapeInfo> masters = GetMasterInformation(archive);
                
                // Export master stencil information to JSON if any were found
                if (masters.Count > 0)
                {
                    MastersExporter.ExportToJson(masters, vsdxPath);
                }
                
                // Then, get connection information
                var connections = GetConnectionInformation(archive);

                // Process each page in the document
                foreach (var pageEntry in archive.Entries.Where(e => e.FullName.StartsWith("visio/pages/page") && e.FullName.EndsWith(".xml")))
                {
                    string pageName = Path.GetFileNameWithoutExtension(pageEntry.Name);

                    XDocument pageXml;
                    using (var stream = pageEntry.Open())
                    {
                        pageXml = XDocument.Load(stream);
                    }

                    // Get the page name from the XML if possible
                    var pageElement = pageXml.Descendants().FirstOrDefault(e => e.Name.LocalName == "Page");
                    if (pageElement != null)
                    {
                        var nameAttr = pageElement.Attribute("Name");
                        if (nameAttr != null)
                        {
                            pageName = nameAttr.Value;
                        }
                    }
                      // Extract all shapes from the page
                    var shapes = pageXml.Descendants().Where(e => e.Name.LocalName == "Shape");
                      // Debug: Check if there are any Connection sections in the XML
                    var connectionSections = pageXml.Descendants().Where(e => e.Name.LocalName == "Connections").ToList();
                    if (connectionSections.Any())
                    {
                        $"Found {connectionSections.Count} connection sections in page {pageName}".WriteInfo();
                    }
                    else
                    {
                        $"No connection sections found in page {pageName}".WriteInfo();
                    }
                    
                    // Debug: Check if there are any Layer sections in the XML
                    var layerSections = pageXml.Descendants().Where(e => e.Name.LocalName == "Layers").ToList();
                    var layerElements = pageXml.Descendants().Where(e => e.Name.LocalName == "Layer").ToList();
                    if (layerSections.Any() || layerElements.Any())
                    {
                        $"Found {layerSections.Count} layer sections and {layerElements.Count} layer elements in page {pageName}".WriteInfo();
                    }
                    else
                    {
                        $"No layer information found in page {pageName}".WriteInfo();
                    }

                    // Process all shapes including subshapes
                    ProcessShapes(shapes, pageName, masters, connections, shapeInfos, null, null);
                }
            }
        }
        catch (InvalidDataException ex)
        {
            Console.WriteLine($"Error processing VSDX file: {ex.Message}");
            // Re-throw to let the caller handle repair attempts
            throw;
        }        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error processing VSDX file: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            throw; // Re-throw to let the caller handle repair attempts
        }        // Print summary statistics after processing all shapes
        int totalConnectionPoints = shapeInfos.Sum(s => s.ConnectionPointsArray.Count);
        int totalLayers = shapeInfos.SelectMany(s => s.Layers).Select(l => l.Id).Distinct().Count();
        int shapesWithConnectionPoints = shapeInfos.Count(s => s.ConnectionPointsArray.Count > 0);
        int shapesWithLayers = shapeInfos.Count(s => s.Layers.Count > 0);
        
        "\n===== SHAPE ANALYSIS SUMMARY =====".WriteSuccess();
        $"Total shapes processed: {shapeInfos.Count}".WriteInfo();
        $"Shapes with connection points: {shapesWithConnectionPoints}".WriteInfo();
        $"Total connection points found: {totalConnectionPoints}".WriteInfo();
        $"Shapes assigned to layers: {shapesWithLayers}".WriteInfo();
        $"Unique layers found: {totalLayers}".WriteInfo();
        "==================================\n".WriteSuccess();

        return shapeInfos;
    }

    static void ProcessShapes(IEnumerable<XElement> shapes,
                            string pageName,
                             Dictionary<string, MasterShapeInfo> masters,
                             Dictionary<string, Tuple<string, string, string>> connections,
                             List<ShapeInfo> shapeInfos,
                             string? parentShapeName,
                             string? parentShapeId = null)
    {
        foreach (var shape in shapes)
        {
            string shapeId = shape.Attribute("ID")?.Value ?? "";
            string shapeName = shape.Attribute("Name")?.Value ?? "";

                        // Get shape type
            var oneDAttr = shape.Attribute("Type");
            var shapeType = oneDAttr?.Value ?? "";
            //$"Shape type for {shapeId} is [{shapeType}]".WriteSuccess();

            // Check if it's a 1D shape (connector)

            
            // Look for Cell N="Type" V="1" - the most reliable indicator of 1D behavior
            var typeCell = shape.Descendants()
                .Where(e => e.Name.LocalName == "Cell" && 
                        e.Attribute("N")?.Value == "Type" && 
                        e.Attribute("V")?.Value == "1")
                .FirstOrDefault();

            bool isExplicit1D = typeCell != null;

            ShapeInfo shapeInfo = new ShapeInfo
            {
                ShapeId = shapeId,
                ShapeName = shapeName,
                PageName = pageName,
                ParentShapeId = parentShapeId ?? "",
                ShapeType = shapeType,
                Is1DShape = isExplicit1D,
            };

            //$"{shapeName} ({shapeId}) is Is1DShape {isExplicit1D} shape".WriteInfo();

            // Get shape text
            var textElement = shape.Descendants().FirstOrDefault(e => e.Name.LocalName == "Text");
            shapeInfo.ShapeText = textElement?.Value ?? "";

            // Get position and size information
            var xForm = shape.Elements().FirstOrDefault(e => e.Name.LocalName.Matches("XForm"));
            if (xForm != null)
            {
                shapeInfo.PositionX = GetDoubleValue(xForm.Elements().FirstOrDefault(e => e.Name.LocalName == "PinX"));
                shapeInfo.PositionY = GetDoubleValue(xForm.Elements().FirstOrDefault(e => e.Name.LocalName == "PinY"));
                shapeInfo.Width = GetDoubleValue(xForm.Elements().FirstOrDefault(e => e.Name.LocalName == "Width"));
                shapeInfo.Height = GetDoubleValue(xForm.Elements().FirstOrDefault(e => e.Name.LocalName == "Height"));
            }
            
            // Look for XForm1D section for 1D shapes (connectors)
            var xForm1D = shape.Elements().FirstOrDefault(e => e.Name.LocalName.Matches("XForm1D"));
            if (xForm1D != null)
            {
                shapeInfo.BeginX = GetDoubleValue(xForm1D.Elements().FirstOrDefault(e => e.Name.LocalName == "BeginX"));
                shapeInfo.BeginY = GetDoubleValue(xForm1D.Elements().FirstOrDefault(e => e.Name.LocalName == "BeginY"));
                shapeInfo.EndX = GetDoubleValue(xForm1D.Elements().FirstOrDefault(e => e.Name.LocalName == "EndX"));
                shapeInfo.EndY = GetDoubleValue(xForm1D.Elements().FirstOrDefault(e => e.Name.LocalName == "EndY"));
                $"XForm1D found for shape {shapeId} with BeginX={shapeInfo.BeginX}, BeginY={shapeInfo.BeginY}, EndX={shapeInfo.EndX}, EndY={shapeInfo.EndY}".WriteNote();
            } 
            
            // Also look for Cell elements with N="BeginX", etc. attributes (alternative format)
            var cells = shape.Descendants().Where(e => e.Name.LocalName == "Cell");
            foreach (var cell in cells)
            {
                string cellName = cell.Attribute("N")?.Value ?? "";
                string cellValue = cell.Attribute("V")?.Value ?? "";
                
                double value = 0;
                if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                {
                    switch (cellName)
                    {
                        case "BeginX":
                            shapeInfo.BeginX = value;
                            break;
                        case "BeginY":
                            shapeInfo.BeginY = value;
                            break;
                        case "EndX":
                            shapeInfo.EndX = value;
                            break;
                        case "EndY":
                            shapeInfo.EndY = value;
                            break;
                    }
                }
            }
              // Get master shape information
            var masterIdAttr = shape.Attribute("Master");
            if (masterIdAttr != null && masters.TryGetValue(masterIdAttr.Value, out var masterInfo))
            {
                shapeInfo.MasterName = masterInfo.Name;
                
                // Copy connection points and layer information from master if available
                if (masterInfo.HasConnectionPoints)
                {
                    shapeInfo.ConnectionPointsArray.AddRange(masterInfo.ConnectionPoints);
                }
                
                if (masterInfo.HasLayers)
                {
                    shapeInfo.Layers.AddRange(masterInfo.Layers);
                }                  // Copy shape data properties if any
                if (masterInfo.ShapeData.Count > 0)
                {
                    List<string> masterPropList = new List<string>();
                    foreach (var kvp in masterInfo.ShapeData)
                    {
                        masterPropList.Add($"{kvp.Key}={kvp.Value}");
                    }
                    shapeInfo.ShapeData = string.Join("; ", masterPropList);
                }
            }

            if ( string.IsNullOrEmpty(shapeInfo.MasterName))
            {
                shapeInfo.Is1DShape = false;
            }
            else if (shapeInfo.MasterName.Contains("connector", StringComparison.OrdinalIgnoreCase))
            {
                shapeInfo.Is1DShape = true;
            }
            else
            {
                shapeInfo.Is1DShape = false;
            }
            

            // Get shape data (custom properties)
            var props = shape.Descendants().Where(e => e.Name.LocalName == "Prop");
            var propList = new List<string>();
            foreach (var prop in props)
            {
                var propName = prop.Attribute("Name")?.Value;
                var propValue = prop.Attribute("Value")?.Value;
                if (!string.IsNullOrEmpty(propName) && !string.IsNullOrEmpty(propValue))
                {
                    propList.Add($"{propName}={propValue}");
                }
            }

            // Add the collected properties to ShapeData, preserving the parent shape info if it exists
            if (propList.Count > 0)
            {
                shapeInfo.ShapeData += string.Join("; ", propList);
            }

            // For 1D shapes (connectors), get connection information
            if (shapeInfo.Is1DShape)
            {
                // Look for connection information in the connections dictionary
                if (connections.TryGetValue(shapeInfo.ShapeId, out var connectionInfo))
                {
                    shapeInfo.BeginConnectedTo = connectionInfo.Item1;
                    shapeInfo.EndConnectedTo = connectionInfo.Item2;
                    shapeInfo.ConnectionPoints = connectionInfo.Item3;
                }
            }            // Check for connection or layer sections
            var hasConnectionOrLayerSections = shape.Elements().Any(e => 
                e.Name.LocalName == "Connections" || 
                e.Name.LocalName == "LayerMem" || 
                e.Name.LocalName == "Layers");
                
            // Only log detailed info for shapes that might have connection points or layers
            if (hasConnectionOrLayerSections)
            {
                $"Examining shape {shapeId} ({shapeName}) for connection points and layers:".WriteInfo(1);
                $"  Shape has {shape.Elements().Count()} direct child elements".WriteInfo(2);
                var childElementNames = shape.Elements().Select(e => e.Name.LocalName).Distinct().ToList();
                $"  Child element types: {string.Join(", ", childElementNames)}".WriteInfo(2);
            }
            
            // Extract connection points
            // In Visio, connection points are stored in the Connections section
            var connectionsSection = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "Connections");
            if (connectionsSection != null)
            {
                var rows = connectionsSection.Elements().Where(e => e.Name.LocalName == "Row");
                foreach (var row in rows)
                {
                    var connectionPoint = new ConnectionPoint();
                    connectionPoint.Id = row.Attribute("IX")?.Value ?? "";
                    connectionPoint.Name = row.Attribute("Name")?.Value ?? "";
                    
                    // Extract X, Y positions from cells
                    var cellElements = row.Elements().Where(e => e.Name.LocalName == "Cell");
                    foreach (var cell in cellElements)
                    {
                        string cellName = cell.Attribute("N")?.Value ?? "";
                        string cellValue = cell.Attribute("V")?.Value ?? "";
                        
                        switch (cellName)
                        {
                            case "X":
                                if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
                                    connectionPoint.X = x;
                                break;
                            case "Y":
                                if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                                    connectionPoint.Y = y;
                                break;
                            case "DirX":
                                connectionPoint.DirX = cellValue;
                                break;
                            case "DirY":
                                connectionPoint.DirY = cellValue;
                                break;
                            case "Type":
                                connectionPoint.Type = cellValue;
                                break;
                        }
                    }
                    
                    // Only add non-empty connection points                    if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
                    {                        shapeInfo.ConnectionPointsArray.Add(connectionPoint);
                        $"Found connection point {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for shape {shapeInfo.ShapeId}".WriteNote();
                        $"Shape {shapeInfo.ShapeId} ({shapeInfo.ShapeName}) has connection point: {connectionPoint.Id} at position ({connectionPoint.X}, {connectionPoint.Y})".WriteInfo(2);
                    }
                }
            }
            
            // Look for connection points in Section N="Connection" format (alternate format)
            var sectionConnections = shape.Elements()
                .Where(e => e.Name.LocalName == "Section" && 
                      e.Attribute("N")?.Value == "Connection")
                .ToList();
                
            foreach (var section in sectionConnections)
            {
                var rows = section.Elements().Where(e => e.Name.LocalName == "Row");
                foreach (var row in rows)
                {
                    var connectionPoint = new ConnectionPoint();
                    connectionPoint.Id = row.Attribute("N")?.Value ?? "";
                    
                    // Extract X, Y positions from cells
                    var cellElements = row.Elements().Where(e => e.Name.LocalName == "Cell");
                    foreach (var cell in cellElements)
                    {
                        string cellName = cell.Attribute("N")?.Value ?? "";
                        string cellValue = cell.Attribute("V")?.Value ?? "";
                        
                        switch (cellName)
                        {
                            case "X":
                                // Try to parse as a direct number or a formula
                                if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
                                    connectionPoint.X = x;
                                else
                                    connectionPoint.X = FormulaParser.ParseFormula(cellValue, shapeInfo.Width, shapeInfo.Height);
                                break;
                            case "Y":                                // Try to parse as a direct number or a formula
                                if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                                    connectionPoint.Y = y;
                                else
                                    connectionPoint.Y = FormulaParser.ParseFormula(cellValue, shapeInfo.Width, shapeInfo.Height);
                                break;
                            case "DirX":
                                connectionPoint.DirX = cellValue;
                                break;
                            case "DirY":
                                connectionPoint.DirY = cellValue;
                                break;                            case "Type":
                                connectionPoint.Type = cellValue;
                                break;
                        }
                    }
                    
                    // Only add non-empty connection points
                    if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
                    {
                        shapeInfo.ConnectionPointsArray.Add(connectionPoint);
                        $"Found Section-format connection point {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for shape {shapeInfo.ShapeId}".WriteNote();
                        $"Shape {shapeInfo.ShapeId} ({shapeInfo.ShapeName}) has Section-format connection point: {connectionPoint.Id} at position ({connectionPoint.X}, {connectionPoint.Y})".WriteInfo(2);
                    }
                }
            }
            
            // Look for alternative connection point format - Connection elements
            var connectionElements = shape.Descendants().Where(e => e.Name.LocalName == "Connection");
            foreach (var conn in connectionElements)
            {
                var connectionPoint = new ConnectionPoint();
                connectionPoint.Id = conn.Attribute("ID")?.Value ?? "";
                connectionPoint.Name = conn.Attribute("NameU")?.Value ?? "";
                
                // Extract X, Y from specific child elements if they exist
                var xElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "X");
                var yElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "Y");
                var dirXElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "DirX");
                var dirYElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "DirY");
                var typeElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "Type");
                
                if (xElem != null && double.TryParse(xElem.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double xVal))
                {
                    connectionPoint.X = xVal;
                }
                
                if (yElem != null && double.TryParse(yElem.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double yVal))
                {
                    connectionPoint.Y = yVal;
                }
                
                if (dirXElem != null)
                {
                    connectionPoint.DirX = dirXElem.Value;
                }
                  if (dirYElem != null)
                {
                    connectionPoint.DirY = dirYElem.Value;
                }
                
                if (typeElem != null)
                {
                    connectionPoint.Type = typeElem.Value;
                }
                
                // Only add non-empty connection points
                if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
                {
                    shapeInfo.ConnectionPointsArray.Add(connectionPoint);
                    $"Found direct connection element {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for shape {shapeInfo.ShapeId}".WriteNote();
                }
            }
        }
    }
    
    /// <summary>
    /// Extracts detailed information about master stencils from a Visio document
    /// </summary>
    /// <param name="archive">The ZipArchive containing the VSDX file contents</param>
    /// <returns>Dictionary mapping master IDs to MasterShapeInfo objects</returns>
    static Dictionary<string, MasterShapeInfo> GetMasterInformation(ZipArchive archive)
    {
        // We'll store both simple name dictionary (for backward compatibility) and detailed info
        Dictionary<string, MasterShapeInfo> masters = new Dictionary<string, MasterShapeInfo>();
        Dictionary<string, string> masterNames = new Dictionary<string, string>(); // For backward compatibility        // Debug: Print all entries in the zip file to help diagnose master stencil issues
        $"Archive contains {archive.Entries.Count} entries.".WriteInfo();
        var masterRelatedEntries = archive.Entries.Where(e => e.FullName.Contains("master", StringComparison.OrdinalIgnoreCase) || e.FullName.StartsWith("visio/masters/")).ToList();
        $"Found {masterRelatedEntries.Count} entries related to masters:".WriteInfo();
        foreach (var entry in masterRelatedEntries.Take(10))
        {
            $"  {entry.FullName}".WriteInfo();
        }        // First identify stencil documents to get stencil names
        var stencilDocs = archive.Entries
            .Where(e => e.FullName.StartsWith("visio/masters/") && 
                        e.FullName.EndsWith(".xml") && 
                        !e.FullName.EndsWith("masters.xml") && 
                        !e.FullName.Contains("/_rels/"))
            .ToList();
        
        $"Found {stencilDocs.Count} master stencil documents.".WriteInfo();
        
        foreach (var masterEntry in stencilDocs)
        {
            string stencilName = Path.GetFileNameWithoutExtension(masterEntry.Name);
            stencilName = stencilName.Split('_').FirstOrDefault() ?? "Unknown";
            
            $"Processing master stencil: {stencilName}".WriteInfo();
            
            using (var stream = masterEntry.Open())
            {
                XDocument masterXml = XDocument.Load(stream);
                var masterElements = masterXml.Descendants().Where(e => e.Name.LocalName == "Master");
                int masterCount = 0;

                foreach (var master in masterElements)
                {
                    masterCount++;
                    string masterId = master.Attribute("ID")?.Value ?? "";
                    string masterName = master.Attribute("Name")?.Value ?? "";
                    string baseId = master.Attribute("BaseID")?.Value ?? "";
                    string uniqueId = master.Attribute("UniqueID")?.Value ?? "";

                    // Skip if no valid ID
                    if (string.IsNullOrEmpty(masterId))
                        continue;
                        
                    // Store the simple name mapping for backward compatibility
                    masterNames[masterId] = masterName;
                    
                    // Create a detailed master info object
                    var masterInfo = new MasterShapeInfo
                    {
                        Id = masterId,
                        Name = masterName,
                        StencilName = stencilName,
                        BaseId = baseId,
                        UniqueId = uniqueId,
                        OriginalXml = master
                    };
                    
                    // Check if it's a 1D shape (connector)
                    var typeCell = master.Descendants()
                        .Where(e => e.Name.LocalName == "Cell" && 
                                e.Attribute("N")?.Value == "Type" && 
                                e.Attribute("V")?.Value == "1")
                        .FirstOrDefault();
                    
                    masterInfo.Is1DShape = typeCell != null || masterName.Contains("connector", StringComparison.OrdinalIgnoreCase);
                    
                    // Get the shape type
                    var typeAttr = master.Attribute("Type");
                    masterInfo.Type = typeAttr?.Value ?? "";
                    
                    // Extract size information
                    var xForm = master.Elements().FirstOrDefault(e => e.Name.LocalName.Matches("XForm")) ??
                                master.Descendants().FirstOrDefault(e => e.Name.LocalName.Matches("XForm"));
                    
                    if (xForm != null)
                    {
                        masterInfo.Width = GetDoubleValue(xForm.Elements().FirstOrDefault(e => e.Name.LocalName == "Width"));
                        masterInfo.Height = GetDoubleValue(xForm.Elements().FirstOrDefault(e => e.Name.LocalName == "Height"));
                    }
                    
                    // Extract connection points
                    ExtractConnectionPoints(master, masterInfo);
                    
                    // Extract layer information
                    ExtractLayerInformation(master, masterInfo);
                    
                    // Extract custom properties (shape data)
                    ExtractShapeData(master, masterInfo);
                    
                    // Store the master info
                    masters[masterId] = masterInfo;
                }
                
                $"Processed {masterCount} masters from stencil {stencilName}".WriteInfo();
            }
        }

        // Summary output 
        int totalMasters = masters.Count;
        int connectorsCount = masters.Values.Count(m => m.Is1DShape);
        int withConnectionPoints = masters.Values.Count(m => m.HasConnectionPoints);
        int withLayers = masters.Values.Count(m => m.HasLayers);
        
        $"\n===== MASTER STENCIL ANALYSIS =====".WriteSuccess();
        $"Total master stencils: {totalMasters}".WriteInfo();
        $"Connector masters: {connectorsCount}".WriteInfo();
        $"Masters with connection points: {withConnectionPoints}".WriteInfo();
        $"Masters with layer information: {withLayers}".WriteInfo();
        $"==================================\n".WriteSuccess();
        
        return masters;
    }
    
    /// <summary>
    /// Extracts connection points from a master shape element
    /// </summary>
    static void ExtractConnectionPoints(XElement master, MasterShapeInfo masterInfo)
    {
        // Look for Connection sections
        var connectionSections = master.Descendants().Where(e => e.Name.LocalName == "Connections").ToList();
        
        foreach (var connectionsSection in connectionSections)
        {
            var rows = connectionsSection.Elements().Where(e => e.Name.LocalName == "Row");
            foreach (var row in rows)
            {
                var connectionPoint = new ConnectionPoint();
                connectionPoint.Id = row.Attribute("IX")?.Value ?? "";
                connectionPoint.Name = row.Attribute("Name")?.Value ?? "";
                
                // Extract X, Y positions from cells
                var cellElements = row.Elements().Where(e => e.Name.LocalName == "Cell");
                foreach (var cell in cellElements)
                {
                    string cellName = cell.Attribute("N")?.Value ?? "";
                    string cellValue = cell.Attribute("V")?.Value ?? "";
                    
                    switch (cellName)
                    {
                        case "X":
                            if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
                                connectionPoint.X = x;
                            break;
                        case "Y":
                            if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                                connectionPoint.Y = y;
                            break;
                        case "DirX":
                            connectionPoint.DirX = cellValue;
                            break;
                        case "DirY":
                            connectionPoint.DirY = cellValue;
                            break;
                        case "Type":
                            connectionPoint.Type = cellValue;
                            break;
                    }
                }
                
                // Only add non-empty connection points
                if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
                {
                    masterInfo.ConnectionPoints.Add(connectionPoint);
                    $"Found connection point {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for master {masterInfo.Name}".WriteNote();
                }
            }
        }
        
        // Look for connection points in Section N="Connection" format
        var sectionConnections = master.Descendants()
            .Where(e => e.Name.LocalName == "Section" && 
                  e.Attribute("N")?.Value == "Connection")
            .ToList();
            
        foreach (var section in sectionConnections)
        {
            var rows = section.Elements().Where(e => e.Name.LocalName == "Row");
            foreach (var row in rows)
            {
                var connectionPoint = new ConnectionPoint();
                connectionPoint.Id = row.Attribute("N")?.Value ?? "";
                
                // Extract X, Y positions from cells
                var cellElements = row.Elements().Where(e => e.Name.LocalName == "Cell");
                foreach (var cell in cellElements)
                {
                    string cellName = cell.Attribute("N")?.Value ?? "";
                    string cellValue = cell.Attribute("V")?.Value ?? "";
                    
                    switch (cellName)
                    {
                        case "X":
                            if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
                                connectionPoint.X = x;
                            break;
                        case "Y":
                            if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                                connectionPoint.Y = y;
                            break;
                        case "DirX":
                            connectionPoint.DirX = cellValue;
                            break;
                        case "DirY":
                            connectionPoint.DirY = cellValue;
                            break;
                        case "Type":
                            connectionPoint.Type = cellValue;
                            break;
                    }
                }
                
                // Only add non-empty connection points
                if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
                {
                    masterInfo.ConnectionPoints.Add(connectionPoint);
                    $"Found Section-format connection point {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for master {masterInfo.Name}".WriteNote();
                }
            }
        }
        
        // Also check for alternative connection point format - direct Connection elements
        var connectionElements = master.Descendants().Where(e => e.Name.LocalName == "Connection");
        foreach (var conn in connectionElements)
        {
            var connectionPoint = new ConnectionPoint();
            connectionPoint.Id = conn.Attribute("ID")?.Value ?? "";
            connectionPoint.Name = conn.Attribute("NameU")?.Value ?? "";
            
            // Extract X, Y from specific child elements if they exist
            var xElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "X");
            var yElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "Y");
            var dirXElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "DirX");
            var dirYElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "DirY");
            var typeElem = conn.Elements().FirstOrDefault(e => e.Name.LocalName == "Type");
            
            if (xElem != null && double.TryParse(xElem.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double xVal))
            {
                connectionPoint.X = xVal;
            }
            
            if (yElem != null && double.TryParse(yElem.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double yVal))
            {
                connectionPoint.Y = yVal;
            }
            
            if (dirXElem != null)
            {
                connectionPoint.DirX = dirXElem.Value;
            }
            
            if (dirYElem != null)
            {
                connectionPoint.DirY = dirYElem.Value;
            }
            
            if (typeElem != null)
            {
                connectionPoint.Type = typeElem.Value;
            }
            
            // Only add non-empty connection points
            if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
            {
                masterInfo.ConnectionPoints.Add(connectionPoint);
                $"Found direct connection element {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for master {masterInfo.Name}".WriteNote();
            }
        }
    }
    
    /// <summary>
    /// Extracts layer information from a master shape element
    /// </summary>
    static void ExtractLayerInformation(XElement master, MasterShapeInfo masterInfo)
    {
        // Check for Layer sections
        var layerSections = master.Descendants().Where(e => e.Name.LocalName == "Layers").ToList();
        var layerElements = master.Descendants().Where(e => e.Name.LocalName == "Layer").ToList();
        
        // Process layer elements directly in the master
        foreach (var layerElement in layerElements)
        {
            Layer layer = new Layer();
            layer.Id = layerElement.Attribute("IX")?.Value ?? "";
            
            if (string.IsNullOrEmpty(layer.Id))
                continue;
                
            // Extract layer properties
            var layerCells = layerElement.Elements().Where(e => e.Name.LocalName == "Cell");
            foreach (var cell in layerCells)
            {
                string cellName = cell.Attribute("N")?.Value ?? "";
                string cellValue = cell.Attribute("V")?.Value ?? "";
                
                switch (cellName)
                {
                    case "Name":
                        layer.Name = cellValue;
                        break;
                    case "Status":
                        layer.Status = cellValue;
                        break;
                    case "Visible":
                        layer.Visible = cellValue == "1";
                        break;
                    case "Print":
                        layer.Print = cellValue == "1";
                        break;
                    case "Active":
                        layer.Active = cellValue == "1";
                        break;
                    case "Lock":
                        layer.Lock = cellValue == "1";
                        break;
                    case "Color":
                        layer.Color = cellValue;
                        break;
                }
            }
            
            if (!string.IsNullOrEmpty(layer.Id))
            {
                masterInfo.Layers.Add(layer);
                $"Found layer {layer.Id} ({layer.Name}) for master {masterInfo.Name}".WriteNote();
            }
        }
        
        // Also check for layer membership information
        var layerMem = master.Descendants().Where(e => e.Name.LocalName == "LayerMem").ToList();
        foreach (var layerMemSection in layerMem)
        {
            var layerMembers = layerMemSection.Elements().Where(e => e.Name.LocalName == "LayerMember");
            foreach (var member in layerMembers)
            {
                string layerId = member.Value;
                if (!string.IsNullOrEmpty(layerId) && !masterInfo.Layers.Any(l => l.Id == layerId))
                {
                    // Add as a simple layer reference if we don't have details
                    var layer = new Layer { Id = layerId, Name = $"Layer {layerId}" };
                    masterInfo.Layers.Add(layer);
                    $"Found layer membership {layerId} for master {masterInfo.Name}".WriteNote();
                }
            }
        }
    }
    
    /// <summary>
    /// Extracts shape data (custom properties) from a master shape element
    /// </summary>
    static void ExtractShapeData(XElement master, MasterShapeInfo masterInfo)
    {
        // Get shape data (custom properties)
        var props = master.Descendants().Where(e => e.Name.LocalName == "Prop");
        foreach (var prop in props)
        {
            var propName = prop.Attribute("Name")?.Value;
            var propValue = prop.Attribute("Value")?.Value;
            if (!string.IsNullOrEmpty(propName) && !string.IsNullOrEmpty(propValue))
            {
                masterInfo.AddShapeData(propName, propValue);
            }
        }
        
        // Look for other relevant metadata in cells
        var cells = master.Descendants().Where(e => e.Name.LocalName == "Cell");
        foreach (var cell in cells)
        {
            string cellName = cell.Attribute("N")?.Value ?? "";
            string cellValue = cell.Attribute("V")?.Value ?? "";
            
            // Only store certain important metadata
            switch (cellName)
            {
                case "LineColor":
                case "FillColor":
                case "LinePattern":
                case "FillPattern":
                case "LineWeight":
                case "BeginArrow":
                case "EndArrow":
                case "DisplayName":
                case "Category":
                case "Description":
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        masterInfo.AddShapeData(cellName, cellValue);
                    }
                    break;
            }
        }
    }

    static Dictionary<string, Tuple<string, string, string>> GetConnectionInformation(ZipArchive archive)
    {
        Dictionary<string, Tuple<string, string, string>> connections = new Dictionary<string, Tuple<string, string, string>>();

        foreach (var pageEntry in archive.Entries.Where(e => e.FullName.StartsWith("visio/pages/") && e.FullName.EndsWith(".xml")))
        {
            using (var stream = pageEntry.Open())
            {
                XDocument pageXml = XDocument.Load(stream);
                var connects = pageXml.Descendants().Where(e => e.Name.LocalName == "Connect");

                foreach (var connect in connects)
                {
                    string fromSheet = connect.Attribute("FromSheet")?.Value ?? "";
                    string toSheet = connect.Attribute("ToSheet")?.Value ?? "";
                    string fromPart = connect.Attribute("FromPart")?.Value ?? "";
                    string toPart = connect.Attribute("ToPart")?.Value ?? "";

                    // Create connection information based on FromCell and ToCell values
                    string connectionPoints = $"FromPart={fromPart}, ToPart={toPart}";

                    //$"{fromSheet} -> {toSheet}".WriteInfo();

                    if (!string.IsNullOrEmpty(fromSheet) && !connections.ContainsKey(fromSheet))
                    {
                        connections[fromSheet] = new Tuple<string, string, string>(toSheet, "", connectionPoints);
                    }
                    else if (!string.IsNullOrEmpty(fromSheet) && connections.TryGetValue(fromSheet, out var existing))
                    {
                        connections[fromSheet] = new Tuple<string, string, string>(existing.Item1, toSheet, existing.Item3 + "; " + connectionPoints);
                    }
                }
            }
        }

        return connections;
    }    static double GetDoubleValue(XElement? element)
    {
        if (element == null)
            return 0;

        if (double.TryParse(element.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            return result;

        return 0;
    }
    
    /// <summary>
    /// Analyzes relationships between shapes in the diagram
    /// </summary>
    /// <param name="shapes2D">Collection of 2D shapes</param>
    /// <param name="shapes1D">Collection of 1D shapes (connectors)</param>
    private static void AnalyzeShapeRelationships(IEnumerable<Shape2D> shapes2D, IEnumerable<Shape1D> shapes1D)
    {
        var allShapes = shapes2D.Cast<BaseShape>().Concat(shapes1D.Cast<BaseShape>()).ToList();
        
        // Analyze connections between shapes
        $"\nAnalyzing shape connections:".WriteSuccess();
        var connections = ShapeAnalyzer.AnalyzeConnections(allShapes);
        if (connections.Count > 0)
        {
            $"Found {connections.Count} connections between shapes:".WriteInfo();
            foreach (var conn in connections.Take(5)) // Show only the first 5 to avoid flooding the output
            {
                $"  {conn}".WriteInfo();
            }
            if (connections.Count > 5)
            {
                $"  ... and {connections.Count - 5} more connections".WriteInfo();
            }
        }
        else
        {
            $"No connections found between shapes.".WriteInfo();
        }
        
        // Analyze connection points usage
        $"\nAnalyzing connection points usage:".WriteSuccess();
        var connectionPointsAnalysis = AnalyzeConnectionPointsUsage(shapes2D, shapes1D);
        $"  Total connection points: {connectionPointsAnalysis.totalPoints}".WriteInfo();
        $"  Connection points used as source: {connectionPointsAnalysis.sourcePoints}".WriteInfo();
        $"  Connection points used as target: {connectionPointsAnalysis.targetPoints}".WriteInfo();
        $"  Unused connection points: {connectionPointsAnalysis.unusedPoints}".WriteInfo();
        
        // Group shapes by layer
        $"\nAnalyzing layer membership:".WriteSuccess();
        var layerGroups = ShapeAnalyzer.GroupShapesByLayer(allShapes);
        if (layerGroups.Count > 0)
        {
            $"Found {layerGroups.Count} layers:".WriteInfo();
            foreach (var layer in layerGroups.Take(5))
            {
                $"  Layer {layer.Key}: {layer.Value.Count} shapes".WriteInfo();
            }
            if (layerGroups.Count > 5)
            {
                $"  ... and {layerGroups.Count - 5} more layers".WriteInfo();
            }
        }
        else
        {
            $"No layers found in the diagram.".WriteInfo();
        }
        
        // Find spatial relationships
        $"\nAnalyzing spatial relationships:".WriteSuccess();
        var spatialRelationships = ShapeAnalyzer.FindSpatialRelationships(allShapes);
        if (spatialRelationships.Count > 0)
        {
            $"Found {spatialRelationships.Count} shapes with spatial relationships:".WriteInfo();
            foreach (var relationship in spatialRelationships.Take(5))
            {
                var shape = allShapes.FirstOrDefault(s => s.Id == relationship.Key);
                $"  Shape {shape?.Text ?? relationship.Key} is related to {relationship.Value.Count} other shapes".WriteInfo();
            }
            if (spatialRelationships.Count > 5)
            {
                $"  ... and {spatialRelationships.Count - 5} more relationships".WriteInfo();
            }
        }
        else
        {
            $"No spatial relationships found between shapes.".WriteInfo();
        }
        
        $"\n".WriteInfo();
    }

    /// <summary>
    /// Analyzes how connection points are used in the diagram
    /// </summary>
    /// <param name="shapes2D">Collection of 2D shapes</param>
    /// <param name="shapes1D">Collection of 1D shapes (connectors)</param>
    /// <returns>Statistics about connection point usage</returns>
    private static (int totalPoints, int sourcePoints, int targetPoints, int unusedPoints) AnalyzeConnectionPointsUsage(
        IEnumerable<Shape2D> shapes2D, IEnumerable<Shape1D> shapes1D)
    {
        var allPoints = new Dictionary<string, List<ConnectionPoint>>();
        var usedAsSource = new HashSet<string>();
        var usedAsTarget = new HashSet<string>();
        
        // Collect all connection points
        foreach (var shape in shapes2D)
        {
            if (shape.ConnectionPoints.Count > 0)
            {
                allPoints[shape.Id] = shape.ConnectionPoints;
            }
        }
        
        // Find which connection points are used by connectors
        foreach (var connector in shapes1D)
        {
            if (!string.IsNullOrEmpty(connector.FromShapeId))
            {
                usedAsSource.Add($"{connector.FromShapeId}:{connector.FromCell}");
            }
            
            if (!string.IsNullOrEmpty(connector.ToShapeId))
            {
                usedAsTarget.Add($"{connector.ToShapeId}:{connector.ToCell}");
            }
        }
        
        // Calculate statistics
        int totalPoints = allPoints.Values.Sum(list => list.Count);
        int sourcePoints = usedAsSource.Count;
        int targetPoints = usedAsTarget.Count;
        int unusedPoints = totalPoints - (sourcePoints + targetPoints);
        
        return (totalPoints, sourcePoints, targetPoints, unusedPoints);
    }
}
