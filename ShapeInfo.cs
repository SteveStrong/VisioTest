using System.Collections.Generic;

namespace VisioShapeExtractor;

public class ShapeInfo
{
    public string ShapeId { get; set; } = string.Empty;
    public string ShapeName { get; set; } = string.Empty;
    public string ParentShapeId { get; set; } = string.Empty;
    public string ShapeText { get; set; } = string.Empty;
    public string ShapeType { get; set; } = string.Empty;
    public string MasterName { get; set; } = string.Empty;
    public bool Is1DShape { get; set; }
    public string BeginConnectedTo { get; set; } = string.Empty;
    public string EndConnectedTo { get; set; } = string.Empty;
    public string ConnectionPoints { get; set; } = string.Empty;
    public string PageName { get; set; } = string.Empty;
    public double PositionX { get; set; }
    public double PositionY { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public string ShapeData { get; set; } = string.Empty;
}

public class ShapeSheet
{
    public string Id { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public string Master { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public string ParentId { get; set; } = string.Empty;
    public List<ShapeSheet> SubShapes { get; set; } = new List<ShapeSheet>();
}

public class Shape2D : ShapeSheet
{

    public double PinX { get; set; }
    public double PinY { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }

}

public class Shape1D : ShapeSheet
{
    public string FromShapeId { get; set; } = string.Empty;
    public string ToShapeId { get; set; } = string.Empty;
    
    public string FromCell { get; set; } = string.Empty;
    public string ToCell { get; set; } = string.Empty;
}
