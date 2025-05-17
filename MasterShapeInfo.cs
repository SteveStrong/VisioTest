using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using System.Xml.Linq;

namespace VisioShapeExtractor
{
    /// <summary>
    /// Represents detailed information about a Visio master stencil shape
    /// </summary>
    public class MasterShapeInfo
    {
        /// <summary>
        /// Unique identifier for the master shape
        /// </summary>
        public string Id { get; set; } = string.Empty;
        
        /// <summary>
        /// Name of the master shape
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Name of the stencil this master belongs to
        /// </summary>
        public string StencilName { get; set; } = string.Empty;
        
        /// <summary>
        /// Base ID from the master
        /// </summary>
        public string BaseId { get; set; } = string.Empty;
        
        /// <summary>
        /// Path to the icon image if available
        /// </summary>
        public string IconPath { get; set; } = string.Empty;
        
        /// <summary>
        /// Unique ID in the Visio file
        /// </summary>
        public string UniqueId { get; set; } = string.Empty;

        /// <summary>
        /// Type of the master shape (e.g., "Group", "Shape", etc.)
        /// </summary>
        public string Type { get; set; } = string.Empty;
        
        /// <summary>
        /// Whether this is a 1D shape (connector)
        /// </summary>
        public bool Is1DShape { get; set; }
        
        /// <summary>
        /// Default width of the master shape
        /// </summary>
        public double Width { get; set; }
        
        /// <summary>
        /// Default height of the master shape
        /// </summary>
        public double Height { get; set; }
        
        /// <summary>
        /// Connection points defined on this master
        /// </summary>
        public List<ConnectionPoint> ConnectionPoints { get; set; } = new List<ConnectionPoint>();
        
        /// <summary>
        /// Layers defined on this master
        /// </summary>
        public List<Layer> Layers { get; set; } = new List<Layer>();

        /// <summary>
        /// Additional properties and metadata from the master
        /// </summary>
        [JsonPropertyName("ShapeData")]
        public Dictionary<string, string> ShapeData { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Original XML content of the master shape for reference
        /// </summary>
        [JsonIgnore]
        public XElement OriginalXml { get; set; }
        
        /// <summary>
        /// Check if the master has connection points
        /// </summary>
        [JsonIgnore]
        public bool HasConnectionPoints => ConnectionPoints.Count > 0;
        
        /// <summary>
        /// Check if the master has layers
        /// </summary>
        [JsonIgnore]
        public bool HasLayers => Layers.Count > 0;
        
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
        /// Creates a string representation of the master information
        /// </summary>
        public override string ToString()
        {
            return $"{Name} (ID: {Id}, Type: {Type}, ConnectionPoints: {ConnectionPoints.Count}, Layers: {Layers.Count})";
        }
    }
}
