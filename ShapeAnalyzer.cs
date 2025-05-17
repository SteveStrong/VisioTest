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
                
                // Try to find source shape
                if (!string.IsNullOrEmpty(connector.FromShapeId) && shapeDict.TryGetValue(connector.FromShapeId, out var sourceShape))
                {
                    connection.FromShapeId = sourceShape.Id;
                    connection.FromShapeText = sourceShape.Text;
                }
                
                // Try to find target shape
                if (!string.IsNullOrEmpty(connector.ToShapeId) && shapeDict.TryGetValue(connector.ToShapeId, out var targetShape))
                {
                    connection.ToShapeId = targetShape.Id;
                    connection.ToShapeText = targetShape.Text;
                }
                
                connections.Add(connection);
            }
            
            return connections;
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
        public string ConnectorId { get; set; } = string.Empty;
        public string ConnectorText { get; set; } = string.Empty;
        public string FromShapeId { get; set; } = string.Empty;
        public string FromShapeText { get; set; } = string.Empty;
        public string ToShapeId { get; set; } = string.Empty;
        public string ToShapeText { get; set; } = string.Empty;
        
        public override string ToString()
        {
            return $"{FromShapeText} → {ConnectorText} → {ToShapeText}";
        }
    }
}
