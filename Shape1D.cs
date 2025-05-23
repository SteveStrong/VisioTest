using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using System.Linq;

namespace VisioShapeExtractor
{    /// <summary>
    /// Represents a 1D shape (connector) in a Visio diagram
    /// </summary>
    public class Shape1D : BaseShape
    {
        /// <summary>
        /// ID of the shape this connector starts from
        /// </summary>
        public string FromShapeId { get; set; } = string.Empty;

        /// <summary>
        /// ID of the shape this connector ends at
        /// </summary>
        public string ToShapeId { get; set; } = string.Empty;

        /// <summary>
        /// Description of connection points at the start of the connector
        /// </summary>
        public string FromCell { get; set; } = string.Empty;

        /// <summary>
        /// Description of connection points at the end of the connector
        /// </summary>
        public string ToCell { get; set; } = string.Empty;

        /// <summary>
        /// X-coordinate of the beginning point
        /// </summary>
        public double BeginX { get; set; }

        /// <summary>
        /// Y-coordinate of the beginning point
        /// </summary>
        public double BeginY { get; set; }

        /// <summary>
        /// X-coordinate of the ending point
        /// </summary>
        public double EndX { get; set; }

        /// <summary>
        /// Y-coordinate of the ending point
        /// </summary>
        public double EndY { get; set; }
        
        /// <summary>
        /// Gets the length of the connector
        /// </summary>
        [JsonIgnore]
        public double Length
        {
            get
            {
                double dx = EndX - BeginX;
                double dy = EndY - BeginY;
                return Math.Sqrt(dx * dx + dy * dy);
            }
        }
        
        /// <summary>
        /// Gets the angle of the connector in degrees
        /// </summary>
        [JsonIgnore]
        public double Angle
        {
            get
            {
                double dx = EndX - BeginX;
                double dy = EndY - BeginY;
                return Math.Atan2(dy, dx) * 180 / Math.PI;
            }
        }
        
        /// <summary>
        /// Gets the midpoint of the connector
        /// </summary>
        /// <returns>A tuple containing the X and Y coordinates of the midpoint</returns>
        public (double X, double Y) GetMidpoint()
        {
            return ((BeginX + EndX) / 2, (BeginY + EndY) / 2);
        }

        /// <summary>
        /// Extracts the connection point ID from FromCell or ToCell
        /// </summary>
        /// <param name="cellInfo">The cell info string (FromCell or ToCell)</param>
        /// <returns>The connection point ID, if it can be parsed</returns>
        public static string? ExtractConnectionPointId(string cellInfo)
        {
            if (string.IsNullOrEmpty(cellInfo))
                return null;
                
            // Try to extract FromPart or ToPart value
            var parts = cellInfo.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                var keyValue = part.Split('=', 2);
                if (keyValue.Length == 2)
                {
                    string key = keyValue[0].Trim();
                    string value = keyValue[1].Trim();
                    
                    if (key.Contains("Part", StringComparison.OrdinalIgnoreCase))
                    {
                        return value;
                    }
                }
            }
            
            return null;
        }
        
        /// <summary>
        /// Gets the connection point used on the source shape
        /// </summary>
        /// <param name="sourceShape">The source shape</param>
        /// <returns>The connection point, if it can be found</returns>
        public ConnectionPoint? GetSourceConnectionPoint(BaseShape? sourceShape)
        {
            if (sourceShape == null || string.IsNullOrEmpty(FromCell))
                return null;
                
            string? connectionPointId = ExtractConnectionPointId(FromCell);
            if (connectionPointId != null)
            {
                return sourceShape.GetConnectionPointById(connectionPointId);
            }
            
            // If we can't find by ID, try to find the nearest connection point
            return sourceShape.GetNearestConnectionPoint(BeginX, BeginY).point;
        }
        
        /// <summary>
        /// Gets the connection point used on the target shape
        /// </summary>
        /// <param name="targetShape">The target shape</param>
        /// <returns>The connection point, if it can be found</returns>
        public ConnectionPoint? GetTargetConnectionPoint(BaseShape? targetShape)
        {
            if (targetShape == null || string.IsNullOrEmpty(ToCell))
                return null;
                
            string? connectionPointId = ExtractConnectionPointId(ToCell);
            if (connectionPointId != null)
            {
                return targetShape.GetConnectionPointById(connectionPointId);
            }
            
            // If we can't find by ID, try to find the nearest connection point
            return targetShape.GetNearestConnectionPoint(EndX, EndY).point;
        }
        
        /// <summary>
        /// Checks if a point is near this connector within the given tolerance
        /// </summary>
        /// <param name="x">X-coordinate of the point</param>
        /// <param name="y">Y-coordinate of the point</param>
        /// <param name="tolerance">Distance tolerance</param>
        /// <returns>True if the point is near the connector</returns>
        public bool IsPointNearLine(double x, double y, double tolerance = 0.1)
        {
            // Calculate the distance from the point to the line
            double dx = EndX - BeginX;
            double dy = EndY - BeginY;
            double len = Math.Sqrt(dx * dx + dy * dy);
            
            // Avoid division by zero
            if (len < 0.0001)
                return Math.Sqrt((x - BeginX) * (x - BeginX) + (y - BeginY) * (y - BeginY)) <= tolerance;
                
            // Calculate the distance from the point to the line
            double distance = Math.Abs((dy * x - dx * y + EndX * BeginY - EndY * BeginX) / len);
            
            // Check if the point is near the line
            if (distance > tolerance)
                return false;
                
            // Check if the point is within the segment bounds
            double dotProduct = ((x - BeginX) * dx + (y - BeginY) * dy) / (dx * dx + dy * dy);
            return dotProduct >= 0 && dotProduct <= 1;
        }
        
        /// <summary>
        /// Overrides the base GetShapeSummary method to include 1D-specific information
        /// </summary>
        /// <returns>A string describing the connector</returns>
        public override string GetShapeSummary()
        {
            var summary = new System.Text.StringBuilder(base.GetShapeSummary());
            summary.AppendLine($"From: ({BeginX}, {BeginY})");
            summary.AppendLine($"To: ({EndX}, {EndY})");
            summary.AppendLine($"Length: {Length:F2}");
            
            if (!string.IsNullOrEmpty(FromShapeId))
                summary.AppendLine($"Connected From: {FromShapeId}");
            
            if (!string.IsNullOrEmpty(ToShapeId))
                summary.AppendLine($"Connected To: {ToShapeId}");
            
            return summary.ToString();
        }
    }
}
