using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

public class ShapeInfo
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ShapeId { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ShapeName { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ParentShapeId { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ShapeText { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ShapeType { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string MasterName { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Is1DShape { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string BeginConnectedTo { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string EndConnectedTo { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ConnectionPoints { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string PageName { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double PositionX { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double PositionY { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double Width { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double Height { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double BeginX { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double BeginY { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double EndX { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double EndY { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ShapeData { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public List<ConnectionPoint> ConnectionPointsArray { get; set; } = new List<ConnectionPoint>();
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public List<Layer> Layers { get; set; } = new List<Layer>();
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string LayerMembership { get; set; } = string.Empty;
}

public class VisioShapes
{
    public string filename { get; set; } =  string.Empty;
    public List<Shape2D> Shape2D { get; set; } = new();
    public List<Shape1D> Shape1D { get; set; } = new();
}

[JsonDerivedType(typeof(Shape2D), typeDiscriminator: "2D")]
[JsonDerivedType(typeof(Shape1D), typeDiscriminator: "1D")]
public class ShapeSheet
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string Id { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string Text { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string Master { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string Type { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ParentId { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public List<ShapeSheet> SubShapes { get; set; } = null!;
    
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public List<ConnectionPoint> ConnectionPoints { get; set; } = new List<ConnectionPoint>();
    
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public List<Layer> Layers { get; set; } = new List<Layer>();
    
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string LayerMembership { get; set; } = string.Empty;

    public void AddSubShape(ShapeSheet shape)
    {
        SubShapes ??= new List<ShapeSheet>();
        if (SubShapes.Contains(shape))
            return;
        SubShapes.Add(shape);
    }
}

public class Shape2D : ShapeSheet
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double PinX { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double PinY { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double Width { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double Height { get; set; }
}

public class Shape1D : ShapeSheet
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string FromShapeId { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ToShapeId { get; set; } = string.Empty;
    
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string FromCell { get; set; } = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string ToCell { get; set; } = string.Empty;
    
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double BeginX { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double BeginY { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double EndX { get; set; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double EndY { get; set; }
}
