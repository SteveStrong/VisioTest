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
    // Define a class to store shape information


    static void Main(string[] args)
    {
        var dir = Directory.GetCurrentDirectory();
        dir = dir.Replace(@"bin\Debug\net9.0", "");
            
        try
        {
            // Use a default file in the Documents folder if no arguments provided
            string documentsFolder = Path.Combine(dir, "Documents");            string[] vsdxFiles = Directory.GetFiles(documentsFolder, "*.vsdx");



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
                    
                    var shapeInfos = ExtractShapeInformation(vsdxPath);
                    
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
    }

    private static void TransformShapeInformation(List<ShapeInfo> shapeInfos, string vsdxPath)
    {
        //this function will remap shapeInfos to the into Shape2D and Shape1D items then export as json
        var shape2DList = new Dictionary<string, Shape2D>();
        var shape1DList = new Dictionary<string, Shape1D>();

        foreach (var shapeInfo in shapeInfos)
        {
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

                // Add to the dictionary
                if (!shape2DList.ContainsKey(shape2D.Id))
                {
                    shape2DList[shape2D.Id] = shape2D;
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

        //now remove shapes from the dictionary if they  have a parent
        foreach (var shape2D in shape2DList.Values.ToList())
        {
            if (!string.IsNullOrEmpty(shape2D.ParentId) && shape2DList.ContainsKey(shape2D.ParentId))
            {
                shape2DList.Remove(shape2D.Id);
            }
        }

        VisioShapes visioShapes = new VisioShapes
        {
            filename = vsdxPath,
            Shape2D = shape2DList.Values.ToList(),
            Shape1D = shape1DList.Values.ToList()
        };

        // Export the lists to JSON files
        var outputPath = vsdxPath.Replace(".vsdx", ".json");
        var data = DehydrateShapes<VisioShapes>(visioShapes);
        File.WriteAllText(outputPath, data);

    }

    public static string DehydrateShapes<T>(T target) where T : class
    {
        var options = new JsonSerializerOptions
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault,
            WriteIndented = true
        };
        return JsonSerializer.Serialize(target, options);
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

    static List<ShapeInfo> ExtractShapeInformation(string vsdxPath)
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
                // foreach (var entry in archive.Entries.Where(e => e.FullName.EndsWith(".xml")))
                // {
                //     string outputPath = Path.Combine(xmlOutputDir, entry.FullName.Replace('/', Path.DirectorySeparatorChar));
                    
                //     // Ensure directory for this file exists
                //     Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? "");
                    
                //     // Extract and save the XML file
                //     entry.ExtractToFile(outputPath, overwrite: true);
                // }
                
                // First, get all master information (stencils)
                Dictionary<string, string> masters = GetMasterInformation(archive);

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
                             Dictionary<string, string> masters,
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
            if (masterIdAttr != null && masters.TryGetValue(masterIdAttr.Value, out var masterName))
            {
                shapeInfo.MasterName = masterName;
            }

            if ( shapeInfo.MasterName.Contains("connector", StringComparison.OrdinalIgnoreCase))
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
            }            // Debug: Print the raw structure of the shape to look for connection points and layers
            $"Examining shape {shapeId} ({shapeName}) for connection points and layers:".WriteInfo(1);
            $"  Shape has {shape.Elements().Count()} direct child elements".WriteInfo(2);
            var childElementNames = shape.Elements().Select(e => e.Name.LocalName).Distinct().ToList();
            $"  Child element types: {string.Join(", ", childElementNames)}".WriteInfo(2);
            
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
                
                if (xElem != null)
                {
                    if (double.TryParse(xElem.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double xVal))
                        connectionPoint.X = xVal;
                }
                
                if (yElem != null)
                {
                    if (double.TryParse(yElem.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double yVal))
                        connectionPoint.Y = yVal;
                }
                  // Only add non-empty connection points
                if (!string.IsNullOrEmpty(connectionPoint.Id) && (connectionPoint.X != 0 || connectionPoint.Y != 0))
                {                    shapeInfo.ConnectionPointsArray.Add(connectionPoint);
                    $"Found alternative connection point {connectionPoint.Id} at ({connectionPoint.X}, {connectionPoint.Y}) for shape {shapeInfo.ShapeId}".WriteNote();
                    $"Shape {shapeInfo.ShapeId} ({shapeInfo.ShapeName}) has alternative connection point: {connectionPoint.Id} at position ({connectionPoint.X}, {connectionPoint.Y})".WriteInfo(2);
                }
            }

            // Extract layer information
            // First, check for Layer membership information
            var layerMember = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "LayerMem");
            if (layerMember != null)
            {
                var layerMembers = layerMember.Elements().Where(e => e.Name.LocalName == "LayerMember");
                List<string> layerIds = new List<string>();
                
                foreach (var member in layerMembers)
                {
                    string layerId = member.Value;
                    if (!string.IsNullOrEmpty(layerId))
                    {
                        layerIds.Add(layerId);
                        
                        // Extract individual layer details if available
                        var layerDetails = shape.Descendants()
                            .Where(e => e.Name.LocalName == "Layer" && e.Attribute("IX")?.Value == layerId)
                            .FirstOrDefault();
                        
                        if (layerDetails != null)
                        {
                            Layer layer = new Layer { Id = layerId };
                            
                            // Extract layer properties from cells
                            var layerCells = layerDetails.Elements().Where(e => e.Name.LocalName == "Cell");
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
                            {                                shapeInfo.Layers.Add(layer);
                                $"Found layer {layer.Id} ({layer.Name}) for shape {shapeInfo.ShapeId}".WriteNote();
                                $"Shape {shapeInfo.ShapeId} ({shapeInfo.ShapeName}) belongs to layer: {layer.Id} - {layer.Name}".WriteInfo(2);
                            }
                        }
                    }
                }
                
                // Store layer membership as a comma-separated list
                if (layerIds.Count > 0)
                {
                    shapeInfo.LayerMembership = string.Join(",", layerIds);
                }
            }
            
            // Also look for layer information from page
            var layerElements = shape.Ancestors()
                .Where(e => e.Name.LocalName == "Page")
                .SelectMany(p => p.Elements().Where(e => e.Name.LocalName == "Layers"))
                .SelectMany(l => l.Elements().Where(e => e.Name.LocalName == "Layer"));
            
            foreach (var layerElement in layerElements)
            {
                Layer layer = new Layer();
                layer.Id = layerElement.Attribute("IX")?.Value ?? "";
                
                if (string.IsNullOrEmpty(layer.Id))
                    continue;
                
                // Check if shape is a member of this layer
                if (shapeInfo.LayerMembership.Contains(layer.Id))
                {
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
                      // Check if we already have this layer
                    if (!shapeInfo.Layers.Any(l => l.Id == layer.Id))
                    {                        shapeInfo.Layers.Add(layer);
                        $"Found page layer {layer.Id} ({layer.Name}) for shape {shapeInfo.ShapeId}".WriteNote();
                        $"Shape {shapeInfo.ShapeId} ({shapeInfo.ShapeName}) belongs to page layer: {layer.Id} - {layer.Name}".WriteInfo(2);
                    }
                }
            }            // Print summary of connection points and layers for this shape
            if (shapeInfo.ConnectionPointsArray.Count > 0 || shapeInfo.Layers.Count > 0)
            {
                $"Shape Summary: {shapeInfo.ShapeId} ({shapeInfo.ShapeName}) has {shapeInfo.ConnectionPointsArray.Count} connection points and belongs to {shapeInfo.Layers.Count} layers".WriteSuccess();
            }

            shapeInfos.Add(shapeInfo);
            // If this is a group shape, process its children with this shape as the parent
            if (shapeInfo.ShapeType == "Group")
            {
                // Process child shapes recursively
                var childShapes = shape.Elements().Where(e => e.Name.LocalName == "Shapes")
                                        .SelectMany(e => e.Elements().Where(c => c.Name.LocalName == "Shape"));

                if (childShapes.Any())
                {
                    ProcessShapes(childShapes, pageName, masters, connections, shapeInfos, shapeName, shapeId);
                }
            }
        }
    }

    static Dictionary<string, string> GetMasterInformation(ZipArchive archive)
    {
        Dictionary<string, string> masters = new Dictionary<string, string>();

        foreach (var masterEntry in archive.Entries.Where(e => e.FullName.StartsWith("visio/masters/") && e.FullName.EndsWith(".xml")))
        {
            using (var stream = masterEntry.Open())
            {
                XDocument masterXml = XDocument.Load(stream);
                var masterElements = masterXml.Descendants().Where(e => e.Name.LocalName == "Master");

                foreach (var master in masterElements)
                {
                    string masterId = master.Attribute("ID")?.Value ?? "";
                    string masterName = master.Attribute("Name")?.Value ?? "";

                    if (!string.IsNullOrEmpty(masterId))
                    {
                        masters[masterId] = masterName;
                    }
                }
            }
        }

        return masters;
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
    }

    static double GetDoubleValue(XElement? element)
    {
        if (element == null)
            return 0;

        if (double.TryParse(element.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            return result;

        return 0;
    }


}
