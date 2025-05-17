using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor
{
    /// <summary>
    /// Base class for all shapes in a Visio document.
    /// Contains common properties shared between all shape types.
    /// </summary>
    [JsonDerivedType(typeof(Shape2D), typeDiscriminator: "2D")]
    [JsonDerivedType(typeof(Shape1D), typeDiscriminator: "1D")]
    public abstract class BaseShape
    {
        /// <summary>
        /// Unique identifier for the shape
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// Text content of the shape
        /// </summary>
        public string Text { get; set; } = string.Empty;

        /// <summary>
        /// Name of the master shape this shape is based on
        /// </summary>
        public string Master { get; set; } = string.Empty;

        /// <summary>
        /// Type of the shape (e.g., Group, Shape, etc.)
        /// </summary>
        public string Type { get; set; } = string.Empty;

        /// <summary>
        /// ID of the parent shape if this is a child shape
        /// </summary>
        public string ParentId { get; set; } = string.Empty;

        /// <summary>
        /// List of child shapes (for group shapes)
        /// </summary>
        [JsonPropertyName("SubShapes")]
        public List<BaseShape> SubShapes { get; set; } = new List<BaseShape>();

        /// <summary>
        /// Connection points in the shape
        /// </summary>
        public List<ConnectionPoint> ConnectionPoints { get; set; } = new List<ConnectionPoint>();

        /// <summary>
        /// Layers the shape belongs to
        /// </summary>
        public List<Layer> Layers { get; set; } = new List<Layer>();

        /// <summary>
        /// Layer membership string (comma-separated list of layer IDs)
        /// </summary>
        public string LayerMembership { get; set; } = string.Empty;        /// <summary>
        /// Adds a child shape to this shape
        /// </summary>
        /// <param name="shape">The shape to add as a child</param>
        public void AddSubShape(BaseShape shape)
        {
            // SubShapes is already initialized, no need for null check
            if (shape != null && !SubShapes.Contains(shape))
            {
                SubShapes.Add(shape);
            }
        }
        
        /// <summary>
        /// Dictionary of custom shape data properties
        /// </summary>
        [JsonPropertyName("ShapeData")]
        public Dictionary<string, string> ShapeData { get; set; } = new Dictionary<string, string>();
        
        /// <summary>
        /// Adds a custom property to the shape's data
        /// </summary>
        /// <param name="key">Property name</param>
        /// <param name="value">Property value</param>
        public void AddShapeData(string key, string value)
        {
            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
            {
                ShapeData[key] = value;
            }
        }
        
        /// <summary>
        /// Parses a string of key-value pairs separated by semicolons into the ShapeData dictionary
        /// </summary>
        /// <param name="shapeDataString">String in format "key1=value1; key2=value2"</param>
        public void ParseShapeData(string shapeDataString)
        {
            if (string.IsNullOrWhiteSpace(shapeDataString))
                return;
                
            var properties = shapeDataString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var prop in properties)
            {
                var keyValue = prop.Split(new[] { '=' }, 2, StringSplitOptions.RemoveEmptyEntries);
                if (keyValue.Length == 2)
                {
                    var key = keyValue[0].Trim();
                    var value = keyValue[1].Trim();
                    AddShapeData(key, value);
                }
            }
        }
        
        /// <summary>
        /// Generates a summary description of the shape including its key properties
        /// </summary>
        /// <returns>A string describing the shape</returns>
        public virtual string GetShapeSummary()
        {
            var summary = new System.Text.StringBuilder();
            summary.AppendLine($"Shape ID: {Id}");
            summary.AppendLine($"Type: {Type}");
            
            if (!string.IsNullOrEmpty(Text))
                summary.AppendLine($"Text: {Text}");
                
            if (!string.IsNullOrEmpty(Master))
                summary.AppendLine($"Master: {Master}");
                
            if (ConnectionPoints.Count > 0)
                summary.AppendLine($"Connection Points: {ConnectionPoints.Count}");
                
            if (Layers.Count > 0)
                summary.AppendLine($"Layers: {Layers.Count}");
                
            if (SubShapes.Count > 0)
                summary.AppendLine($"Child Shapes: {SubShapes.Count}");
                
            return summary.ToString();
        }
    }
}
