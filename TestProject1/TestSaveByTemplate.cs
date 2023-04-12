using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MiniExcelTests.Helpers;

namespace MiniExcelTests;

public class TestSaveByTemplate
{
    private const string ValidFileTwoSheets =
        "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";

    private const string SomeValue = "123";

    public TestSaveByTemplate()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    public static byte[] DecodeBase64(string value)
    {
        return string.IsNullOrEmpty(value) ? Array.Empty<byte>() : Convert.FromBase64String(value);
    }

    [Fact]
    public async Task Test_Applies_Ok_SingleSheet()
    {
        var stream = new MemoryStream();
        var dict = new Dictionary<string, object>
        {
            ["CID"] = SomeValue
        };
        const string sheet = "sheet1.xml";
        await stream.SaveAsByTemplateAsync(DecodeBase64(ValidFileTwoSheets), dict, new OpenXmlConfiguration { Sheet = sheet });
        stream.Position = 0;
        Assert.True(stream.Length > 0);
        var wasReplaced = ExcelContainsChecker.Contains(stream.ToArray(), new List<string> { SomeValue });
        Assert.True(wasReplaced);
    }

    [Fact]
    public async Task Test_NotApplies_Ok_OtherSheet()
    {
        var stream = new MemoryStream();
        var dict = new Dictionary<string, object>
        {
            ["CID"] = SomeValue
        };
        const string sheet = "s2";
        await stream.SaveAsByTemplateAsync(DecodeBase64(ValidFileTwoSheets), dict, new OpenXmlConfiguration { Sheet = sheet });
        stream.Position = 0;
        Assert.True(stream.Length > 0);
        var wasReplaced = ExcelContainsChecker.Contains(stream.ToArray(), new List<string> { SomeValue });
        Assert.False(wasReplaced);
    }

    [Fact]
    public async Task Test_Applies_Ok_NoSheet()
    {
        var stream = new MemoryStream();
        var dict = new Dictionary<string, object>
        {
            ["CID"] = SomeValue
        };
        await stream.SaveAsByTemplateAsync(DecodeBase64(ValidFileTwoSheets), dict);
        stream.Position = 0;
        Assert.True(stream.Length > 0);
        var wasReplaced = ExcelContainsChecker.Contains(stream.ToArray(), new List<string> { SomeValue });
        Assert.True(wasReplaced);
    }
}