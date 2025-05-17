using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

public class Layer
{
    public string Id { get; set; } = string.Empty;
    
    public string Name { get; set; } = string.Empty;
    
    public int Index { get; set; }
    
    public bool Visible { get; set; }
    
    public bool Print { get; set; }
    
    public bool Active { get; set; }
    
    public bool Lock { get; set; }
    
    public string Color { get; set; } = string.Empty;
    
    public string Status { get; set; } = string.Empty;
}
