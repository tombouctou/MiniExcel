using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace MiniExcelTests.Helpers;

public static class ExcelContainsChecker
{
    public static bool Contains(byte[] data, IEnumerable<string> values)
    {
        var existing = GetExistingValues(data);

        return values.All(x => existing.ContainsKey(x));
    }

    public static Dictionary<string, string> GetExistingValues(byte[] data)
    {
        var stream = new MemoryStream();
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        Dictionary<string, string> existing = new();

        try
        {
            using var reader = ExcelReaderFactory.CreateReader(
                stream,
                new()
                {
                    FallbackEncoding = Encoding.ASCII,
                });

            do
            {
                var sheetName = reader.Name;
                var colCount = reader.FieldCount;

                while (reader.Read())
                {
                    for (var i = 0; i < colCount; ++i)
                    {
                        var value = "";

                        try
                        {
                            value = reader.GetString(i);
                        }
                        catch (Exception)
                        {
                            // no handler needed
                        }

                        if (value == "")
                        {
                            try
                            {
                                value = reader.GetDouble(i).ToString(CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                                // no handler needed
                            }
                        }

                        if (value != null && value.Length > 0)
                        {
                            existing[value] = value;
                        }
                    }
                }
            } while (reader.NextResult());
        }
        catch (Exception e)
        {
            throw new InvalidOperationException(e.Message, e);
        }

        return existing;
    }
}