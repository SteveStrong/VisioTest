using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

// Removed EmptyCollectionConverter - replaced with ListConverter

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
    public double BeginX { get; set; }
    public double BeginY { get; set; }
    public double EndX { get; set; }
    public double EndY { get; set; }
    public string ShapeData { get; set; } = string.Empty;
    public List<ConnectionPoint> ConnectionPointsArray { get; set; } = new List<ConnectionPoint>();
    public List<Layer> Layers { get; set; } = new List<Layer>();
    public string LayerMembership { get; set; } = string.Empty;
}
