using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor
{    /// <summary>
    /// Represents a 2D shape in a Visio diagram
    /// </summary>
    public class Shape2D : BaseShape
    {
        /// <summary>
        /// X-coordinate of the shape's pin (center or reference point)
        /// </summary>
        public double PinX { get; set; }

        /// <summary>
        /// Y-coordinate of the shape's pin (center or reference point)
        /// </summary>
        public double PinY { get; set; }

        /// <summary>
        /// Width of the shape
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Height of the shape
        /// </summary>
        public double Height { get; set; }
        
        /// <summary>
        /// Gets the left boundary of the shape
        /// </summary>
        [JsonIgnore]
        public double Left => PinX - (Width / 2);
        
        /// <summary>
        /// Gets the right boundary of the shape
        /// </summary>
        [JsonIgnore]
        public double Right => PinX + (Width / 2);
        
        /// <summary>
        /// Gets the top boundary of the shape
        /// </summary>
        [JsonIgnore]
        public double Top => PinY + (Height / 2);
        
        /// <summary>
        /// Gets the bottom boundary of the shape
        /// </summary>
        [JsonIgnore]
        public double Bottom => PinY - (Height / 2);

        /// <summary>
        /// Gets connection points based on their relative position in the shape
        /// </summary>
        /// <returns>A dictionary of connection points organized by position</returns>
        public Dictionary<string, List<ConnectionPoint>> GetOrganizedConnectionPoints()
        {
            var result = new Dictionary<string, List<ConnectionPoint>>
            {
                ["top"] = new List<ConnectionPoint>(),
                ["right"] = new List<ConnectionPoint>(),
                ["bottom"] = new List<ConnectionPoint>(),
                ["left"] = new List<ConnectionPoint>(),
                ["center"] = new List<ConnectionPoint>()
            };
            
            foreach (var cp in ConnectionPoints)
            {
                // Determine position based on coordinates relative to shape boundaries
                double relativeX = (cp.X - Left) / Width; // 0 = left edge, 1 = right edge
                double relativeY = (cp.Y - Bottom) / Height; // 0 = bottom edge, 1 = top edge
                
                const double tolerance = 0.1;
                
                if (relativeX < tolerance) // Left edge
                {
                    result["left"].Add(cp);
                }
                else if (relativeX > 1 - tolerance) // Right edge
                {
                    result["right"].Add(cp);
                }
                else if (relativeY < tolerance) // Bottom edge
                {
                    result["bottom"].Add(cp);
                }
                else if (relativeY > 1 - tolerance) // Top edge
                {
                    result["top"].Add(cp);
                }
                else // Somewhere in the middle
                {
                    result["center"].Add(cp);
                }
            }
            
            return result;
        }

        /// <summary>
        /// Gets all connectors that connect to this shape
        /// </summary>
        /// <param name="allConnectors">All connectors in the diagram</param>
        /// <returns>A list of connectors that connect to this shape</returns>
        public List<Shape1D> GetConnectors(IEnumerable<Shape1D> allConnectors)
        {
            return allConnectors.Where(c => c.FromShapeId == Id || c.ToShapeId == Id).ToList();
        }

        /// <summary>
        /// Gets all shapes connected to this shape through connectors
        /// </summary>
        /// <param name="allShapes">All shapes in the diagram</param>
        /// <param name="allConnectors">All connectors in the diagram</param>
        /// <returns>A dictionary of connected shapes organized by direction (in/out)</returns>
        public Dictionary<string, List<Shape2D>> GetConnectedShapes(IEnumerable<Shape2D> allShapes, IEnumerable<Shape1D> allConnectors)
        {
            var result = new Dictionary<string, List<Shape2D>>
            {
                ["incoming"] = new List<Shape2D>(),
                ["outgoing"] = new List<Shape2D>()
            };
            
            foreach (var connector in allConnectors)
            {
                if (connector.ToShapeId == Id && !string.IsNullOrEmpty(connector.FromShapeId))
                {
                    var sourceShape = allShapes.FirstOrDefault(s => s.Id == connector.FromShapeId);
                    if (sourceShape != null)
                    {
                        result["incoming"].Add(sourceShape);
                    }
                }
                
                if (connector.FromShapeId == Id && !string.IsNullOrEmpty(connector.ToShapeId))
                {
                    var targetShape = allShapes.FirstOrDefault(s => s.Id == connector.ToShapeId);
                    if (targetShape != null)
                    {
                        result["outgoing"].Add(targetShape);
                    }
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// Checks if this shape contains the specified point
        /// </summary>
        /// <param name="x">X-coordinate</param>
        /// <param name="y">Y-coordinate</param>
        /// <returns>True if the point is inside the shape's boundaries</returns>
        public bool ContainsPoint(double x, double y)
        {
            return x >= Left && x <= Right && y >= Bottom && y <= Top;
        }
        
        /// <summary>
        /// Overrides the base GetShapeSummary method to include 2D-specific information
        /// </summary>
        /// <returns>A string describing the 2D shape</returns>
        public override string GetShapeSummary()
        {
            var summary = new System.Text.StringBuilder(base.GetShapeSummary());
            summary.AppendLine($"Position: ({PinX}, {PinY})");
            summary.AppendLine($"Size: {Width} x {Height}");
            return summary.ToString();
        }
    }
}
