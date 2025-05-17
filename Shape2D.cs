using System;
using System.Collections.Generic;
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
