using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioShapeExtractor
{
    /// <summary>
    /// Provides analysis capabilities for Visio shapes
    /// </summary>
    public static class ShapeAnalyzer
    {
        /// <summary>
        /// Finds all connections between shapes in a diagram
        /// </summary>
        /// <param name="allShapes">Collection of all shapes in the diagram</param>
        /// <returns>A list of connection information</returns>
        public static List<ShapeConnection> AnalyzeConnections(IEnumerable<BaseShape> allShapes)
        {
            var connections = new List<ShapeConnection>();
            var shape1Ds = allShapes.OfType<Shape1D>().ToList();
            var shape2Ds = allShapes.OfType<Shape2D>().ToList();
            
            // Create a dictionary for quick shape lookups
            var shapeDict = shape2Ds.ToDictionary(s => s.Id, s => s);
            
            foreach (var connector in shape1Ds)
            {
                var connection = new ShapeConnection
                {
                    ConnectorId = connector.Id,
                    ConnectorText = connector.Text
                };
                
                // Try to find source shape and connection point
                if (!string.IsNullOrEmpty(connector.FromShapeId) && shapeDict.TryGetValue(connector.FromShapeId, out var sourceShape))
                {
                    connection.FromShapeId = sourceShape.Id;
                    connection.FromShapeText = sourceShape.Text;
                    
                    // Try to identify specific connection point
                    var sourcePoint = connector.GetSourceConnectionPoint(sourceShape);
                    if (sourcePoint != null)
                    {
                        connection.FromConnectionPointId = sourcePoint.Id;
                        connection.FromConnectionPointPosition = GetConnectionPointPosition(sourceShape, sourcePoint);
                    }
                }
                
                // Try to find target shape and connection point
                if (!string.IsNullOrEmpty(connector.ToShapeId) && shapeDict.TryGetValue(connector.ToShapeId, out var targetShape))
                {
                    connection.ToShapeId = targetShape.Id;
                    connection.ToShapeText = targetShape.Text;
                    
                    // Try to identify specific connection point
                    var targetPoint = connector.GetTargetConnectionPoint(targetShape);
                    if (targetPoint != null)
                    {
                        connection.ToConnectionPointId = targetPoint.Id;
                        connection.ToConnectionPointPosition = GetConnectionPointPosition(targetShape, targetPoint);
                    }
                }
                
                connections.Add(connection);
            }
            
            return connections;
        }

        /// <summary>
        /// Determines the position of a connection point relative to its shape
        /// </summary>
        /// <param name="shape">The shape that contains the connection point</param>
        /// <param name="point">The connection point to analyze</param>
        /// <returns>A string describing the position ("top", "right", "bottom", "left", "center", or "unknown")</returns>
        private static string GetConnectionPointPosition(Shape2D shape, ConnectionPoint point)
        {
            double relativeX = (point.X - shape.Left) / shape.Width;
            double relativeY = (point.Y - shape.Bottom) / shape.Height;
            
            const double tolerance = 0.1;
            
            // Check edges first
            if (relativeX < tolerance)
                return "left";
            if (relativeX > 1 - tolerance)
                return "right";
            if (relativeY < tolerance)
                return "bottom";
            if (relativeY > 1 - tolerance)
                return "top";
                
            // If not on an edge, it must be somewhere in the center area
            return "center";
        }

        /// <summary>
        /// Analyzes the connection point usage patterns in the diagram
        /// </summary>
        /// <param name="allShapes">Collection of all shapes in the diagram</param>
        /// <returns>A dictionary with statistics about connection point usage</returns>
        public static Dictionary<string, int> AnalyzeConnectionPointUsage(IEnumerable<BaseShape> allShapes)
        {
            var result = new Dictionary<string, int>
            {
                ["total"] = 0,
                ["used"] = 0,
                ["unused"] = 0,
                ["top"] = 0,
                ["right"] = 0,
                ["bottom"] = 0,
                ["left"] = 0,
                ["center"] = 0
            };
            
            var shape2Ds = allShapes.OfType<Shape2D>().ToList();
            var shape1Ds = allShapes.OfType<Shape1D>().ToList();
            
            // Create sets of used connection points
            var usedPoints = new HashSet<string>(); // Format: "shapeId:connectionPointId"
            
            foreach (var connector in shape1Ds)
            {
                if (!string.IsNullOrEmpty(connector.FromShapeId))
                {
                    string fromConnectionId = Shape1D.ExtractConnectionPointId(connector.FromCell) ?? "";
                    if (!string.IsNullOrEmpty(fromConnectionId))
                    {
                        usedPoints.Add($"{connector.FromShapeId}:{fromConnectionId}");
                    }
                }
                
                if (!string.IsNullOrEmpty(connector.ToShapeId))
                {
                    string toConnectionId = Shape1D.ExtractConnectionPointId(connector.ToCell) ?? "";
                    if (!string.IsNullOrEmpty(toConnectionId))
                    {
                        usedPoints.Add($"{connector.ToShapeId}:{toConnectionId}");
                    }
                }
            }
            
            // Now analyze all shapes and their connection points
            foreach (var shape in shape2Ds)
            {
                var organizedPoints = shape.GetOrganizedConnectionPoints();
                
                // Count total points
                result["total"] += shape.ConnectionPoints.Count;
                
                // Count by position
                foreach (var kvp in organizedPoints)
                {
                    string position = kvp.Key;
                    int count = kvp.Value.Count;
                    
                    if (result.ContainsKey(position))
                    {
                        result[position] += count;
                    }
                }
                
                // Count used vs unused
                int usedCount = 0;
                foreach (var point in shape.ConnectionPoints)
                {
                    if (usedPoints.Contains($"{shape.Id}:{point.Id}"))
                    {
                        usedCount++;
                    }
                }
                
                result["used"] += usedCount;
                result["unused"] += shape.ConnectionPoints.Count - usedCount;
            }
            
            return result;
        }
        
        /// <summary>
        /// Groups shapes by their layer membership
        /// </summary>
        /// <param name="allShapes">Collection of all shapes in the diagram</param>
        /// <returns>Dictionary mapping layer names to lists of shapes in that layer</returns>
        public static Dictionary<string, List<BaseShape>> GroupShapesByLayer(IEnumerable<BaseShape> allShapes)
        {
            var result = new Dictionary<string, List<BaseShape>>();
            
            foreach (var shape in allShapes)
            {
                foreach (var layer in shape.Layers)
                {
                    string layerKey = $"{layer.Id}: {layer.Name}";
                    
                    if (!result.ContainsKey(layerKey))
                    {
                        result[layerKey] = new List<BaseShape>();
                    }
                    
                    result[layerKey].Add(shape);
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// Finds shapes that are potentially spatially related (contained or overlapping)
        /// </summary>
        /// <param name="allShapes">Collection of all shapes in the diagram</param>
        /// <returns>Dictionary mapping shape IDs to lists of related shapes</returns>
        public static Dictionary<string, List<Shape2D>> FindSpatialRelationships(IEnumerable<BaseShape> allShapes)
        {
            var result = new Dictionary<string, List<Shape2D>>();
            var shape2Ds = allShapes.OfType<Shape2D>().ToList();
            
            foreach (var shape in shape2Ds)
            {
                var relatedShapes = new List<Shape2D>();
                
                foreach (var otherShape in shape2Ds)
                {
                    // Skip self-comparison
                    if (shape.Id == otherShape.Id)
                        continue;
                    
                    // Check for containment
                    if (IsContained(shape, otherShape) || IsContained(otherShape, shape))
                    {
                        relatedShapes.Add(otherShape);
                        continue;
                    }
                    
                    // Check for overlap
                    if (DoOverlap(shape, otherShape))
                    {
                        relatedShapes.Add(otherShape);
                    }
                }
                
                if (relatedShapes.Count > 0)
                {
                    result[shape.Id] = relatedShapes;
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// Checks if one shape is completely contained within another
        /// </summary>
        private static bool IsContained(Shape2D outer, Shape2D inner)
        {
            return inner.Left >= outer.Left &&
                   inner.Right <= outer.Right &&
                   inner.Bottom >= outer.Bottom &&
                   inner.Top <= outer.Top;
        }
        
        /// <summary>
        /// Checks if two shapes overlap
        /// </summary>
        private static bool DoOverlap(Shape2D a, Shape2D b)
        {
            // If one rectangle is on the left side of the other
            if (a.Right < b.Left || b.Right < a.Left)
                return false;
            
            // If one rectangle is above the other
            if (a.Bottom > b.Top || b.Bottom > a.Top)
                return false;
            
            return true;
        }
    }
    
    /// <summary>
    /// Represents a connection between two shapes
    /// </summary>
    public class ShapeConnection
    {
        /// <summary>
        /// ID of the connector shape
        /// </summary>
        public string ConnectorId { get; set; } = string.Empty;
        
        /// <summary>
        /// Text label of the connector shape
        /// </summary>
        public string ConnectorText { get; set; } = string.Empty;
        
        /// <summary>
        /// ID of the source shape
        /// </summary>
        public string FromShapeId { get; set; } = string.Empty;
        
        /// <summary>
        /// Text label of the source shape
        /// </summary>
        public string FromShapeText { get; set; } = string.Empty;
        
        /// <summary>
        /// ID of the target shape
        /// </summary>
        public string ToShapeId { get; set; } = string.Empty;
        
        /// <summary>
        /// Text label of the target shape
        /// </summary>
        public string ToShapeText { get; set; } = string.Empty;

        /// <summary>
        /// ID of the connection point on the source shape
        /// </summary>
        public string FromConnectionPointId { get; set; } = string.Empty;
        
        /// <summary>
        /// Position of the connection point on the source shape (top, right, bottom, left, center)
        /// </summary>
        public string FromConnectionPointPosition { get; set; } = string.Empty;
        
        /// <summary>
        /// ID of the connection point on the target shape
        /// </summary>
        public string ToConnectionPointId { get; set; } = string.Empty;
        
        /// <summary>
        /// Position of the connection point on the target shape (top, right, bottom, left, center)
        /// </summary>
        public string ToConnectionPointPosition { get; set; } = string.Empty;
        
        /// <summary>
        /// Returns a string representation of the connection
        /// </summary>
        public override string ToString()
        {
            string from = string.IsNullOrEmpty(FromShapeText) ? FromShapeId : FromShapeText;
            string to = string.IsNullOrEmpty(ToShapeText) ? ToShapeId : ToShapeText;
            string label = string.IsNullOrEmpty(ConnectorText) ? "" : $" ({ConnectorText})";
            
            string fromPoint = string.IsNullOrEmpty(FromConnectionPointPosition) ? 
                "" : $" at {FromConnectionPointPosition}";
                
            string toPoint = string.IsNullOrEmpty(ToConnectionPointPosition) ? 
                "" : $" at {ToConnectionPointPosition}";
                
            return $"{from}{fromPoint} → {to}{toPoint}{label}";
        }
    }
}
