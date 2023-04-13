
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlTemplate : IExcelTemplateAsync
    {
        private static readonly XmlNamespaceManager _ns;
        private static readonly Regex _isExpressionRegex;
        static ExcelOpenXmlTemplate()
        {
            _isExpressionRegex = new Regex("(?<={{).*?(?=}})");
            _ns = new XmlNamespaceManager(new NameTable());
            _ns.AddNamespace("x", Config.SpreadsheetmlXmlns);
        }

        private readonly Stream _stream;
        private readonly OpenXmlConfiguration _configuration;

        public ExcelOpenXmlTemplate(Stream stream, IConfiguration configuration)
        {
            _stream = stream;
            _configuration = (OpenXmlConfiguration)configuration?? OpenXmlConfiguration.DefaultConfig;
        }

        public void SaveAsByTemplate(string templatePath, object value)
        {
            using var stream = FileHelper.OpenSharedRead(templatePath);
            SaveAsByTemplateImpl(stream, value);
        }
        public void SaveAsByTemplate(byte[] templateBytes, object value)
        {
            using Stream stream = new MemoryStream(templateBytes);
            SaveAsByTemplateImpl(stream, value);
        }

        public void SaveAsByTemplateImpl(Stream templateStream, object value)
        {
            //only support xlsx         
            Dictionary<string, object> values;
            if (value is Dictionary<string, object> objects)
            {
                values = objects;
                foreach (var key in values.Keys)
                {
                    var v = values[key];
                    if (v is IDataReader reader)
                    {
                        values[key] = TypeHelper.ConvertToEnumerableDictionary(reader).ToList();
                    }
                }
            }
            else
            {
                var type = value.GetType();
                var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                values = props.ToDictionary(p => p.Name, p => p.GetValue(value));
            }

            {
                templateStream.CopyTo(_stream);

                var reader = new ExcelOpenXmlSheetReader(_stream,null);
                var archive = new ExcelOpenXmlZip(_stream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);
                {
                    //read sharedString
                    var sharedStrings = reader._sharedStrings;

                    //read all xlsx sheets
                    var sheets = archive.zipFile.Entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                        || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                    foreach (var sheet in sheets)
                    {
                        var sheetStream = sheet.Open();
                        var fullName = sheet.FullName;

                        XRowInfos = new List<XRowInfo>(); //every time need to use new XRowInfos or it'll cause duplicate problem: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
                        XMergeCellInfos = new Dictionary<string, XMergeCell>();
                        NewXMergeCellInfos = new List<XMergeCell>();

                        if (!string.IsNullOrEmpty(_configuration.Sheet) && sheet.Name != _configuration.Sheet)
                        {
                            var entryCopy = archive.zipFile.CreateEntry(fullName);
                            using var zipStreamCopy = entryCopy.Open();

                            CopySheetXmlImpl(sheet, zipStreamCopy, sheetStream);
                            continue;
                        }

                        var entry = archive.zipFile.CreateEntry(fullName);
                        using var zipStream = entry.Open();
                        GenerateSheetXmlImpl(sheet, zipStream, sheetStream, values, sharedStrings);
                    }
                }

                archive.zipFile.Dispose();
            }
        }

        public Task SaveAsByTemplateAsync(string templatePath, object value,CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => SaveAsByTemplate(templatePath, value),cancellationToken);
        }

        public Task SaveAsByTemplateAsync(byte[] templateBtyes, object value,CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => SaveAsByTemplate(templateBtyes, value),cancellationToken);
        }
    }
}
