using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

/// <summary>
/// A specialized converter that filters out properties with default values
/// </summary>
public class PropertyFilterConverter : JsonConverterFactory
{
    public override bool CanConvert(Type typeToConvert)
    {
        // This converter handles complex objects (classes)
        return typeToConvert.IsClass && typeToConvert != typeof(string);
    }

    public override JsonConverter CreateConverter(Type typeToConvert, JsonSerializerOptions options)
    {
        var converterType = typeof(PropertyFilterConverter<>).MakeGenericType(typeToConvert);
        return (JsonConverter)Activator.CreateInstance(converterType)!;
    }
}

public class PropertyFilterConverter<T> : JsonConverter<T> where T : class
{
    public override T? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        // For reading, delegate to the default behavior
        var newOptions = new JsonSerializerOptions(options);
        
        // Remove this converter to avoid recursion
        foreach (var converter in newOptions.Converters.ToList())
        {
            if (converter is PropertyFilterConverter)
            {
                newOptions.Converters.Remove(converter);
                break;
            }
        }
        
        return JsonSerializer.Deserialize<T>(ref reader, newOptions);
    }

    public override void Write(Utf8JsonWriter writer, T value, JsonSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteNullValue();
            return;
        }

        // Create a new options instance without this converter to avoid recursion
        var newOptions = new JsonSerializerOptions(options);
        foreach (var converter in newOptions.Converters.ToList())
        {
            if (converter is PropertyFilterConverter)
            {
                newOptions.Converters.Remove(converter);
                break;
            }
        }

        writer.WriteStartObject();

        // Get all properties from the type
        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanRead);

        foreach (var property in properties)
        {
            // Skip properties with [JsonIgnore] attribute
            if (property.GetCustomAttributes(typeof(JsonIgnoreAttribute), true).Any())
                continue;

            object? propertyValue = property.GetValue(value);
            
            // Skip if property has default value
            if (propertyValue == null || 
                IsDefaultValue(propertyValue, property.PropertyType) ||
                (propertyValue is string str && string.IsNullOrEmpty(str)) ||
                (propertyValue is System.Collections.ICollection collection && collection.Count == 0))
            {
                continue;
            }

            // Write property name
            writer.WritePropertyName(GetPropertyName(property, options));
            
            // Write property value using the options without this converter
            JsonSerializer.Serialize(writer, propertyValue, property.PropertyType, newOptions);
        }

        writer.WriteEndObject();
    }

    private string GetPropertyName(PropertyInfo property, JsonSerializerOptions options)
    {
        // Respect the naming policy for property names
        var attribute = property.GetCustomAttribute<JsonPropertyNameAttribute>();
        if (attribute != null)
            return attribute.Name;

        string propertyName = property.Name;
        if (options.PropertyNamingPolicy != null)
            return options.PropertyNamingPolicy.ConvertName(propertyName);
            
        return propertyName;
    }    private bool IsDefaultValue(object value, Type type)
    {
        // For value types, compare with default value
        if (type.IsValueType)
        {
            // Always treat numeric zero values as default
            if (value is int intVal && intVal == 0) return true;
            if (value is double doubleVal && doubleVal == 0.0) return true;
            if (value is float floatVal && floatVal == 0.0f) return true;
            if (value is decimal decimalVal && decimalVal == 0.0m) return true;
            if (value is bool boolVal && boolVal == false) return true;
            
            // For other value types, compare with default
            var defaultValue = Activator.CreateInstance(type);
            return value.Equals(defaultValue);
        }
        
        // String handling - treat empty or whitespace as default
        if (value is string stringVal)
        {
            return string.IsNullOrWhiteSpace(stringVal);
        }
        
        // For reference types, null is the default
        return value == null;
    }
}
