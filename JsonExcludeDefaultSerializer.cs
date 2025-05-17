using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisioShapeExtractor;

/// <summary>
/// A serializer that's specifically designed to exclude default values from serialization output
/// </summary>
public static class JsonExcludeDefaultSerializer
{
    public static string Serialize<T>(T value, bool writeIndented = true)
    {
        var options = new JsonSerializerOptions
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = writeIndented,
        };
        
        return JsonSerializer.Serialize(value, options);
    }
    
    private class JsonExcludeDefaultValueConverter : JsonConverterFactory
    {
        public override bool CanConvert(Type typeToConvert)
        {
            // This converter handles all types except strings (handled by EmptyStringConverter)
            return typeToConvert != typeof(string);
        }

        public override JsonConverter CreateConverter(Type typeToConvert, JsonSerializerOptions options)
        {
            // Create a JsonExcludeDefaultValueConverter<T> instance for the specific type
            var converterType = typeof(JsonExcludeDefaultValueConverter<>).MakeGenericType(typeToConvert);
            return (JsonConverter)Activator.CreateInstance(converterType)!;
        }
    }

    private class JsonExcludeDefaultValueConverter<T> : JsonConverter<T>
    {
        public override T Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            // For reading, use the default behavior
            var newOptions = new JsonSerializerOptions(options);
            // Remove this converter to avoid recursion
            foreach (var converter in newOptions.Converters.ToList())
            {
                if (converter is JsonExcludeDefaultValueConverter)
                {
                    newOptions.Converters.Remove(converter);
                    break;
                }
            }
            
            return JsonSerializer.Deserialize<T>(ref reader, newOptions)!;
        }

        public override void Write(Utf8JsonWriter writer, T value, JsonSerializerOptions options)
        {
            // Skip writing if the value is the default for its type
            if (EqualityComparer<T>.Default.Equals(value, default(T)!))
            {
                // Don't write anything for default values
                return;
            }

            // For collections, check if empty
            if (value is System.Collections.ICollection collection && collection.Count == 0)
            {
                // Don't serialize empty collections
                return;
            }

            // Special handling for numeric types with value 0
            if ((typeof(T) == typeof(int) && (int)(object)value! == 0) ||
                (typeof(T) == typeof(double) && (double)(object)value! == 0) ||
                (typeof(T) == typeof(float) && (float)(object)value! == 0) ||
                (typeof(T) == typeof(decimal) && (decimal)(object)value! == 0))
            {
                // Skip zeros
                return;
            }

            // For writing, use the default behavior but without this converter
            var newOptions = new JsonSerializerOptions(options);
            // Remove this converter to avoid recursion
            foreach (var converter in newOptions.Converters.ToList())
            {
                if (converter is JsonExcludeDefaultValueConverter)
                {
                    newOptions.Converters.Remove(converter);
                    break;
                }
            }

            JsonSerializer.Serialize(writer, value, newOptions);
        }
    }
}
