using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

public class EmptyCollectionConverter<T> : JsonConverter<List<T>>
{
    public override List<T> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        // Default deserialization behavior
        return JsonSerializer.Deserialize<List<T>>(ref reader, options) ?? new List<T>();
    }

    public override void Write(Utf8JsonWriter writer, List<T> value, JsonSerializerOptions options)
    {
        // Don't write anything if the list is empty
        if (value == null || value.Count == 0)
        {
            writer.WriteNullValue();
        }
        else
        {
            JsonSerializer.Serialize(writer, value, options);
        }
    }
}

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
    [JsonConverter(typeof(EmptyCollectionConverter<ConnectionPoint>))]
    public List<ConnectionPoint> ConnectionPointsArray { get; set; } = new List<ConnectionPoint>();
    [JsonConverter(typeof(EmptyCollectionConverter<Layer>))]
    public List<Layer> Layers { get; set; } = new List<Layer>();
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
    public string Id { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public string Master { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public string ParentId { get; set; } = string.Empty;
    public List<ShapeSheet> SubShapes { get; set; } = null!;
    
    [JsonConverter(typeof(EmptyCollectionConverter<ConnectionPoint>))]
    public List<ConnectionPoint> ConnectionPoints { get; set; } = new List<ConnectionPoint>();
    
    [JsonConverter(typeof(EmptyCollectionConverter<Layer>))]
    public List<Layer> Layers { get; set; } = new List<Layer>();
    
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
    
    public double BeginX { get; set; }
    public double BeginY { get; set; }
    public double EndX { get; set; }
    public double EndY { get; set; }
}
