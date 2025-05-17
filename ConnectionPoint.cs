using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

public class ConnectionPoint
{
    public string Id { get; set; } = string.Empty;
    
    public string Name { get; set; } = string.Empty;
    
    public string Type { get; set; } = string.Empty;
    
    public double X { get; set; }
    
    public double Y { get; set; }
    
    public string DirX { get; set; } = string.Empty;
    
    public string DirY { get; set; } = string.Empty;
}
