using System.Text.Json;

static class JsonExport
{
    public static void Write(string path, int rowCount, int columnCount, Func<int, int, string> cellText)
    {
        var headers = new string[columnCount];
        for (int c = 1; c <= columnCount; c++)
        {
            headers[c - 1] = cellText(1, c);
        }

        using FileStream stream = File.Create(path);
        using var writer = new Utf8JsonWriter(stream);
        writer.WriteStartArray();
        for (int r = 2; r <= rowCount; r++)
        {
            writer.WriteStartObject();
            for (int c = 1; c <= columnCount; c++)
            {
                writer.WriteString(headers[c - 1], cellText(r, c));
            }
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
    }
}
