using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

[assembly: AssemblyTitle("noWPS")]
[assembly: AssemblyDescription("Converts WPS DISPIMG pictures in XLSX files to standard Excel pictures.")]
[assembly: AssemblyCompany("noWPS")]
[assembly: AssemblyProduct("noWPS")]
[assembly: AssemblyCopyright("")]
[assembly: ComVisible(false)]
[assembly: Guid("b41fe1fd-f233-4327-a40b-e986df5f9ed4")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace WpsDispImgToExcel
{
    internal static class Program
    {
        private static readonly XNamespace NsSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace NsRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace NsOfficeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace NsXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private static readonly XNamespace NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace NsEtc = "http://www.wps.cn/officeDocument/2017/etCustomData";

        private const string DrawingRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        private const string ImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
        private const long EmusPerPixel = 9525;
        private const long EmusPerPoint = 12700;
        private const int ThumbnailPaddingPixels = 2;

        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            var files = new List<string>();
            foreach (string arg in args)
            {
                string trimmed = arg.Trim('"');
                if (File.Exists(trimmed))
                {
                    files.Add(trimmed);
                }
            }

            if (files.Count == 0)
            {
                using (var dialog = new OpenFileDialog())
                {
                    dialog.Title = "Выберите WPS Excel-файлы с фотографиями";
                    dialog.Filter = "Excel-файлы (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы (*.*)|*.*";
                    dialog.Multiselect = true;
                    dialog.InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                    if (dialog.ShowDialog() != DialogResult.OK)
                    {
                        return 1;
                    }
                    files.AddRange(dialog.FileNames);
                }
            }

            var results = new List<string>();
            bool hasError = false;
            foreach (string file in files)
            {
                try
                {
                    ConversionResult result = ConvertWorkbook(file);
                    results.Add(string.Format("{0}: готово, картинок: {1}. Новый файл: {2}", Path.GetFileName(file), result.ImageCount, result.OutputPath));
                    Console.WriteLine(results[results.Count - 1]);
                }
                catch (Exception ex)
                {
                    hasError = true;
                    string message = string.Format("{0}: ошибка: {1}", Path.GetFileName(file), ex.Message);
                    results.Add(message);
                    Console.Error.WriteLine(message);
                }
            }

            ShowTopMostMessage(string.Join(Environment.NewLine, results.ToArray()), "WPS фото в Excel",
                hasError ? MessageBoxIcon.Warning : MessageBoxIcon.Information);

            return hasError ? 2 : 0;
        }

        private static void ShowTopMostMessage(string text, string caption, MessageBoxIcon icon)
        {
            using (var owner = new Form())
            {
                owner.TopMost = true;
                owner.ShowInTaskbar = false;
                owner.StartPosition = FormStartPosition.Manual;
                owner.Location = new System.Drawing.Point(-32000, -32000);
                owner.Size = new System.Drawing.Size(1, 1);
                owner.Show();
                owner.Activate();

                MessageBox.Show(owner, text, caption, MessageBoxButtons.OK, icon);
            }
        }

        private static ConversionResult ConvertWorkbook(string inputPath)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("Input file not found.", inputPath);
            }

            string extension = Path.GetExtension(inputPath);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Поддерживаются только .xlsx и .xlsm.");
            }

            string outputPath = BuildOutputPath(inputPath);
            File.Copy(inputPath, outputPath, true);

            int convertedImages;
            using (ZipArchive archive = ZipFile.Open(outputPath, ZipArchiveMode.Update))
            {
                Dictionary<string, string> imageById = LoadCellImageMap(archive);
                if (imageById.Count == 0)
                {
                    throw new InvalidOperationException("Внутри файла не найден WPS-каталог картинок cellimages.xml.");
                }

                convertedImages = ConvertSheets(archive, imageById);
                if (convertedImages == 0)
                {
                    throw new InvalidOperationException("В листах не найдены ячейки DISPIMG.");
                }
            }

            return new ConversionResult(outputPath, convertedImages);
        }

        private static Dictionary<string, string> LoadCellImageMap(ZipArchive archive)
        {
            ZipArchiveEntry cellImagesEntry = archive.GetEntry("xl/cellimages.xml");
            ZipArchiveEntry relsEntry = archive.GetEntry("xl/_rels/cellimages.xml.rels");
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (cellImagesEntry == null || relsEntry == null)
            {
                return result;
            }

            XDocument relsDoc = LoadXml(relsEntry);
            var targetByRelId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement rel in relsDoc.Root.Elements(NsRelationships + "Relationship"))
            {
                string id = (string)rel.Attribute("Id");
                string type = (string)rel.Attribute("Type");
                string target = (string)rel.Attribute("Target");
                if (!string.IsNullOrEmpty(id) && type == ImageRelType && !string.IsNullOrEmpty(target))
                {
                    targetByRelId[id] = NormalizeZipPath(CombineZipPath("xl", target));
                }
            }

            XDocument cellImagesDoc = LoadXml(cellImagesEntry);
            foreach (XElement cellImage in cellImagesDoc.Descendants(NsEtc + "cellImage"))
            {
                XElement pic = cellImage.Element(NsXdr + "pic");
                if (pic == null)
                {
                    continue;
                }

                XElement cNvPr = pic.Descendants(NsXdr + "cNvPr").FirstOrDefault();
                XElement blip = pic.Descendants(NsA + "blip").FirstOrDefault();
                if (cNvPr == null || blip == null)
                {
                    continue;
                }

                string imageId = (string)cNvPr.Attribute("name");
                string relId = (string)blip.Attribute(NsOfficeRel + "embed");
                string target;
                if (!string.IsNullOrEmpty(imageId) &&
                    !string.IsNullOrEmpty(relId) &&
                    targetByRelId.TryGetValue(relId, out target) &&
                    archive.GetEntry(target) != null)
                {
                    result[imageId] = target;
                }
            }

            return result;
        }

        private static int ConvertSheets(ZipArchive archive, Dictionary<string, string> imageById)
        {
            var worksheetEntries = archive.Entries
                .Where(e => Regex.IsMatch(e.FullName, @"^xl/worksheets/[^/]+\.xml$", RegexOptions.IgnoreCase))
                .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
                .Select(e => e.FullName)
                .ToList();

            int total = 0;
            int nextDrawingIndex = GetNextDrawingIndex(archive);
            var imageDimensionCache = new Dictionary<string, ImageDimensions>(StringComparer.OrdinalIgnoreCase);

            foreach (string sheetPath in worksheetEntries)
            {
                ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                XDocument sheetDoc = LoadXml(sheetEntry);
                SheetMetrics sheetMetrics = SheetMetrics.FromSheet(sheetDoc);
                List<CellImagePlacement> placements = FindDispImgCells(sheetDoc, imageById);
                if (placements.Count == 0)
                {
                    continue;
                }

                SheetDrawingInfo drawingInfo = EnsureSheetDrawing(archive, sheetDoc, sheetPath, ref nextDrawingIndex);
                XDocument drawingDoc = drawingInfo.DrawingDocument;
                XDocument drawingRelsDoc = drawingInfo.DrawingRelationshipsDocument;

                int nextPicId = GetNextPictureId(drawingDoc);
                foreach (CellImagePlacement placement in placements)
                {
                    ImageDimensions imageDimensions = GetImageDimensions(archive, placement.ImagePath, imageDimensionCache);
                    string imageRelId = NextRelationshipId(drawingRelsDoc);
                    string relTarget = MakeRelativePath("xl/drawings", placement.ImagePath);
                    drawingRelsDoc.Root.Add(new XElement(NsRelationships + "Relationship",
                        new XAttribute("Id", imageRelId),
                        new XAttribute("Type", ImageRelType),
                        new XAttribute("Target", relTarget)));

                    drawingDoc.Root.Add(CreatePictureAnchor(placement, imageRelId, nextPicId++, sheetMetrics, imageDimensions));
                    ClearCellValue(placement.Cell);
                    total++;
                }

                SaveXml(archive, drawingInfo.DrawingPath, drawingDoc);
                SaveXml(archive, drawingInfo.DrawingRelationshipsPath, drawingRelsDoc);
                SaveXml(archive, sheetPath, sheetDoc);
                EnsureDrawingContentType(archive, drawingInfo.DrawingPath);
                EnsureImageContentDefaults(archive, placements.Select(p => p.ImagePath));
            }

            return total;
        }

        private static List<CellImagePlacement> FindDispImgCells(XDocument sheetDoc, Dictionary<string, string> imageById)
        {
            var result = new List<CellImagePlacement>();
            Regex idRegex = new Regex(@"ID_[A-F0-9]{16,}", RegexOptions.IgnoreCase);

            foreach (XElement cell in sheetDoc.Descendants(NsSpreadsheet + "c"))
            {
                string cellText = string.Concat(cell.DescendantNodes().OfType<XText>().Select(t => t.Value));
                if (cellText.IndexOf("DISPIMG", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                Match match = idRegex.Match(cellText);
                if (!match.Success)
                {
                    continue;
                }

                string imageId = match.Value;
                string imagePath;
                if (!imageById.TryGetValue(imageId, out imagePath))
                {
                    continue;
                }

                string address = (string)cell.Attribute("r");
                int rowIndex;
                int columnIndex;
                if (!TryParseCellAddress(address, out rowIndex, out columnIndex))
                {
                    continue;
                }

                result.Add(new CellImagePlacement(cell, address, rowIndex, columnIndex, imageId, imagePath));
            }

            return result;
        }

        private static ImageDimensions GetImageDimensions(ZipArchive archive, string imagePath, Dictionary<string, ImageDimensions> cache)
        {
            ImageDimensions cached;
            if (cache.TryGetValue(imagePath, out cached))
            {
                return cached;
            }

            ImageDimensions result = new ImageDimensions(1, 1);
            ZipArchiveEntry entry = archive.GetEntry(imagePath);
            if (entry != null)
            {
                try
                {
                    using (Stream stream = entry.Open())
                    using (System.Drawing.Image image = System.Drawing.Image.FromStream(stream, false, false))
                    {
                        result = new ImageDimensions(Math.Max(1, image.Width), Math.Max(1, image.Height));
                    }
                }
                catch
                {
                    result = new ImageDimensions(1, 1);
                }
            }

            cache[imagePath] = result;
            return result;
        }

        private static SheetDrawingInfo EnsureSheetDrawing(ZipArchive archive, XDocument sheetDoc, string sheetPath, ref int nextDrawingIndex)
        {
            string sheetDir = PathGetDirectoryNameZip(sheetPath);
            string sheetFileName = PathGetFileNameZip(sheetPath);
            string sheetRelsPath = CombineZipPath(sheetDir, "_rels/" + sheetFileName + ".rels");
            XDocument sheetRelsDoc = LoadOrCreateRelationships(archive, sheetRelsPath);

            XElement drawingElement = sheetDoc.Root.Elements(NsSpreadsheet + "drawing").FirstOrDefault();
            string drawingPath = null;

            if (drawingElement != null)
            {
                string drawingRelId = (string)drawingElement.Attribute(NsOfficeRel + "id");
                XElement drawingRel = sheetRelsDoc.Root.Elements(NsRelationships + "Relationship")
                    .FirstOrDefault(r => (string)r.Attribute("Id") == drawingRelId);
                if (drawingRel != null)
                {
                    drawingPath = NormalizeZipPath(CombineZipPath(sheetDir, (string)drawingRel.Attribute("Target")));
                }
            }

            if (string.IsNullOrEmpty(drawingPath) || archive.GetEntry(drawingPath) == null)
            {
                drawingPath = NextAvailableDrawingPath(archive, ref nextDrawingIndex);
                string drawingRelId = NextRelationshipId(sheetRelsDoc);
                sheetRelsDoc.Root.Add(new XElement(NsRelationships + "Relationship",
                    new XAttribute("Id", drawingRelId),
                    new XAttribute("Type", DrawingRelType),
                    new XAttribute("Target", MakeRelativePath(sheetDir, drawingPath))));

                if (drawingElement == null)
                {
                    drawingElement = new XElement(NsSpreadsheet + "drawing", new XAttribute(NsOfficeRel + "id", drawingRelId));
                    InsertWorksheetDrawingElement(sheetDoc, drawingElement);
                }
                else
                {
                    drawingElement.SetAttributeValue(NsOfficeRel + "id", drawingRelId);
                }

                SaveXml(archive, sheetRelsPath, sheetRelsDoc);
            }

            XDocument drawingDoc = LoadOrCreateDrawing(archive, drawingPath);
            string drawingRelsPath = "xl/drawings/_rels/" + PathGetFileNameZip(drawingPath) + ".rels";
            XDocument drawingRelsDoc = LoadOrCreateRelationships(archive, drawingRelsPath);

            return new SheetDrawingInfo(drawingPath, drawingRelsPath, drawingDoc, drawingRelsDoc);
        }

        private static XElement CreatePictureAnchor(CellImagePlacement placement, string imageRelId, int pictureId, SheetMetrics sheetMetrics, ImageDimensions imageDimensions)
        {
            int zeroColumn = placement.ColumnIndex - 1;
            int zeroRow = placement.RowIndex - 1;
            string pictureName = "WPS " + placement.ImageId;
            long cellWidth = Math.Max(EmusPerPixel, sheetMetrics.GetColumnWidthEmu(placement.ColumnIndex));
            long cellHeight = Math.Max(EmusPerPixel, sheetMetrics.GetRowHeightEmu(placement.RowIndex));
            long padding = Math.Min(ThumbnailPaddingPixels * EmusPerPixel, Math.Min(cellWidth, cellHeight) / 8);
            long availableWidth = Math.Max(EmusPerPixel, cellWidth - padding * 2);
            long availableHeight = Math.Max(EmusPerPixel, cellHeight - padding * 2);
            double imageAspect = imageDimensions.Width / (double)imageDimensions.Height;
            double boxAspect = availableWidth / (double)availableHeight;
            long pictureWidth;
            long pictureHeight;

            if (imageAspect >= boxAspect)
            {
                pictureWidth = availableWidth;
                pictureHeight = (long)Math.Round(availableWidth / imageAspect);
            }
            else
            {
                pictureHeight = availableHeight;
                pictureWidth = (long)Math.Round(availableHeight * imageAspect);
            }

            pictureWidth = Math.Max(EmusPerPixel, Math.Min(pictureWidth, availableWidth));
            pictureHeight = Math.Max(EmusPerPixel, Math.Min(pictureHeight, availableHeight));

            long columnOffset = padding + (availableWidth - pictureWidth) / 2;
            long rowOffset = padding + (availableHeight - pictureHeight) / 2;

            return new XElement(NsXdr + "oneCellAnchor",
                new XElement(NsXdr + "from",
                    new XElement(NsXdr + "col", zeroColumn),
                    new XElement(NsXdr + "colOff", columnOffset),
                    new XElement(NsXdr + "row", zeroRow),
                    new XElement(NsXdr + "rowOff", rowOffset)),
                new XElement(NsXdr + "ext",
                    new XAttribute("cx", pictureWidth),
                    new XAttribute("cy", pictureHeight)),
                new XElement(NsXdr + "pic",
                    new XElement(NsXdr + "nvPicPr",
                        new XElement(NsXdr + "cNvPr",
                            new XAttribute("id", pictureId),
                            new XAttribute("name", pictureName),
                            new XAttribute("descr", placement.ImageId)),
                        new XElement(NsXdr + "cNvPicPr",
                            new XElement(NsA + "picLocks", new XAttribute("noChangeAspect", "1")))),
                    new XElement(NsXdr + "blipFill",
                        new XElement(NsA + "blip", new XAttribute(NsOfficeRel + "embed", imageRelId)),
                        new XElement(NsA + "stretch", new XElement(NsA + "fillRect"))),
                    new XElement(NsXdr + "spPr",
                        new XElement(NsA + "prstGeom",
                            new XAttribute("prst", "rect"),
                            new XElement(NsA + "avLst")))),
                new XElement(NsXdr + "clientData"));
        }

        private static void ClearCellValue(XElement cell)
        {
            cell.Elements(NsSpreadsheet + "f").Remove();
            cell.Elements(NsSpreadsheet + "v").Remove();
            cell.Elements(NsSpreadsheet + "is").Remove();
            XAttribute typeAttribute = cell.Attribute("t");
            if (typeAttribute != null)
            {
                typeAttribute.Remove();
            }
        }

        private static void InsertWorksheetDrawingElement(XDocument sheetDoc, XElement drawingElement)
        {
            string[] afterDrawingElementNames = new[]
            {
                "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
                "controls", "webPublishItems", "tableParts", "extLst"
            };

            XElement before = sheetDoc.Root.Elements()
                .FirstOrDefault(e => afterDrawingElementNames.Contains(e.Name.LocalName));
            if (before != null)
            {
                before.AddBeforeSelf(drawingElement);
            }
            else
            {
                sheetDoc.Root.Add(drawingElement);
            }
        }

        private static XDocument LoadOrCreateDrawing(ZipArchive archive, string path)
        {
            ZipArchiveEntry entry = archive.GetEntry(path);
            if (entry != null)
            {
                return LoadXml(entry);
            }

            return new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(NsXdr + "wsDr",
                    new XAttribute(XNamespace.Xmlns + "xdr", NsXdr.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "a", NsA.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", NsOfficeRel.NamespaceName)));
        }

        private static XDocument LoadOrCreateRelationships(ZipArchive archive, string path)
        {
            ZipArchiveEntry entry = archive.GetEntry(path);
            if (entry != null)
            {
                return LoadXml(entry);
            }

            return new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(NsRelationships + "Relationships"));
        }

        private static void EnsureDrawingContentType(ZipArchive archive, string drawingPath)
        {
            XDocument doc = LoadXml(archive.GetEntry("[Content_Types].xml"));
            string partName = "/" + drawingPath;
            bool exists = doc.Root.Elements(NsContentTypes + "Override")
                .Any(e => string.Equals((string)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
            if (!exists)
            {
                doc.Root.Add(new XElement(NsContentTypes + "Override",
                    new XAttribute("PartName", partName),
                    new XAttribute("ContentType", DrawingContentType)));
                SaveXml(archive, "[Content_Types].xml", doc);
            }
        }

        private static void EnsureImageContentDefaults(ZipArchive archive, IEnumerable<string> imagePaths)
        {
            XDocument doc = LoadXml(archive.GetEntry("[Content_Types].xml"));
            bool changed = false;
            foreach (string imagePath in imagePaths)
            {
                string ext = Path.GetExtension(imagePath).TrimStart('.').ToLowerInvariant();
                if (string.IsNullOrEmpty(ext))
                {
                    continue;
                }

                string contentType = null;
                if (ext == "jpg" || ext == "jpeg")
                {
                    contentType = "image/jpeg";
                }
                else if (ext == "png")
                {
                    contentType = "image/png";
                }
                else if (ext == "gif")
                {
                    contentType = "image/gif";
                }
                else if (ext == "bmp")
                {
                    contentType = "image/bmp";
                }

                if (contentType == null)
                {
                    continue;
                }

                bool exists = doc.Root.Elements(NsContentTypes + "Default")
                    .Any(e => string.Equals((string)e.Attribute("Extension"), ext, StringComparison.OrdinalIgnoreCase));
                if (!exists)
                {
                    doc.Root.Add(new XElement(NsContentTypes + "Default",
                        new XAttribute("Extension", ext),
                        new XAttribute("ContentType", contentType)));
                    changed = true;
                }
            }

            if (changed)
            {
                SaveXml(archive, "[Content_Types].xml", doc);
            }
        }

        private static int GetNextPictureId(XDocument drawingDoc)
        {
            int max = 1;
            foreach (XElement cNvPr in drawingDoc.Descendants(NsXdr + "cNvPr"))
            {
                int id;
                if (int.TryParse((string)cNvPr.Attribute("id"), out id) && id >= max)
                {
                    max = id + 1;
                }
            }
            return max;
        }

        private static int GetNextDrawingIndex(ZipArchive archive)
        {
            int max = 0;
            Regex regex = new Regex(@"^xl/drawings/drawing(\d+)\.xml$", RegexOptions.IgnoreCase);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                Match match = regex.Match(entry.FullName);
                if (match.Success)
                {
                    int index;
                    if (int.TryParse(match.Groups[1].Value, out index) && index > max)
                    {
                        max = index;
                    }
                }
            }
            return max + 1;
        }

        private static string NextAvailableDrawingPath(ZipArchive archive, ref int nextDrawingIndex)
        {
            while (true)
            {
                string path = "xl/drawings/drawing" + nextDrawingIndex + ".xml";
                nextDrawingIndex++;
                if (archive.GetEntry(path) == null)
                {
                    return path;
                }
            }
        }

        private static string NextRelationshipId(XDocument relsDoc)
        {
            int max = 0;
            Regex regex = new Regex(@"^rId(\d+)$", RegexOptions.IgnoreCase);
            foreach (XElement rel in relsDoc.Root.Elements(NsRelationships + "Relationship"))
            {
                string id = (string)rel.Attribute("Id");
                Match match = regex.Match(id ?? string.Empty);
                if (match.Success)
                {
                    int value;
                    if (int.TryParse(match.Groups[1].Value, out value) && value > max)
                    {
                        max = value;
                    }
                }
            }
            return "rId" + (max + 1);
        }

        private static bool TryParseCellAddress(string address, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrEmpty(address))
            {
                return false;
            }

            Match match = Regex.Match(address, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return false;
            }

            string letters = match.Groups[1].Value.ToUpperInvariant();
            for (int i = 0; i < letters.Length; i++)
            {
                column = column * 26 + (letters[i] - 'A' + 1);
            }

            return int.TryParse(match.Groups[2].Value, out row) && row > 0 && column > 0;
        }

        private static string BuildOutputPath(string inputPath)
        {
            string dir = Path.GetDirectoryName(inputPath);
            string name = Path.GetFileNameWithoutExtension(inputPath);
            string ext = Path.GetExtension(inputPath);
            string candidate = Path.Combine(dir, name + "_excel" + ext);
            int index = 2;
            while (File.Exists(candidate))
            {
                candidate = Path.Combine(dir, name + "_excel_" + index + ext);
                index++;
            }
            return candidate;
        }

        private static XDocument LoadXml(ZipArchiveEntry entry)
        {
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
            {
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            }
        }

        private static void SaveXml(ZipArchive archive, string path, XDocument doc)
        {
            ZipArchiveEntry oldEntry = archive.GetEntry(path);
            if (oldEntry != null)
            {
                oldEntry.Delete();
            }

            ZipArchiveEntry newEntry = archive.CreateEntry(path, CompressionLevel.Optimal);
            using (Stream stream = newEntry.Open())
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                OmitXmlDeclaration = false
            }))
            {
                doc.Save(writer);
            }
        }

        private static string NormalizeZipPath(string path)
        {
            var parts = new List<string>();
            foreach (string rawPart in path.Replace('\\', '/').Split('/'))
            {
                if (rawPart.Length == 0 || rawPart == ".")
                {
                    continue;
                }
                if (rawPart == "..")
                {
                    if (parts.Count > 0)
                    {
                        parts.RemoveAt(parts.Count - 1);
                    }
                    continue;
                }
                parts.Add(rawPart);
            }
            return string.Join("/", parts.ToArray());
        }

        private static string CombineZipPath(string basePath, string relativePath)
        {
            if (string.IsNullOrEmpty(basePath))
            {
                return NormalizeZipPath(relativePath);
            }
            return NormalizeZipPath(basePath.TrimEnd('/') + "/" + relativePath);
        }

        private static string MakeRelativePath(string fromDirectory, string toPath)
        {
            string[] fromParts = NormalizeZipPath(fromDirectory).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string[] toParts = NormalizeZipPath(toPath).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            int common = 0;
            while (common < fromParts.Length &&
                   common < toParts.Length &&
                   string.Equals(fromParts[common], toParts[common], StringComparison.OrdinalIgnoreCase))
            {
                common++;
            }

            var rel = new List<string>();
            for (int i = common; i < fromParts.Length; i++)
            {
                rel.Add("..");
            }
            for (int i = common; i < toParts.Length; i++)
            {
                rel.Add(toParts[i]);
            }
            return rel.Count == 0 ? "." : string.Join("/", rel.ToArray());
        }

        private static string PathGetDirectoryNameZip(string path)
        {
            int slash = path.LastIndexOf('/');
            return slash < 0 ? string.Empty : path.Substring(0, slash);
        }

        private static string PathGetFileNameZip(string path)
        {
            int slash = path.LastIndexOf('/');
            return slash < 0 ? path : path.Substring(slash + 1);
        }

        private sealed class SheetMetrics
        {
            private readonly double defaultColumnWidth;
            private readonly double defaultRowHeightPoints;
            private readonly List<ColumnRule> columnRules;
            private readonly Dictionary<int, RowRule> rowRules;

            private SheetMetrics(double defaultColumnWidth, double defaultRowHeightPoints, List<ColumnRule> columnRules, Dictionary<int, RowRule> rowRules)
            {
                this.defaultColumnWidth = defaultColumnWidth;
                this.defaultRowHeightPoints = defaultRowHeightPoints;
                this.columnRules = columnRules;
                this.rowRules = rowRules;
            }

            public static SheetMetrics FromSheet(XDocument sheetDoc)
            {
                XElement sheetFormat = sheetDoc.Root.Element(NsSpreadsheet + "sheetFormatPr");
                double defaultColumnWidth = ParseDouble(sheetFormat == null ? null : sheetFormat.Attribute("defaultColWidth"), 8.43);
                double defaultRowHeightPoints = ParseDouble(sheetFormat == null ? null : sheetFormat.Attribute("defaultRowHeight"), 15.0);
                var columnRules = new List<ColumnRule>();
                XElement cols = sheetDoc.Root.Element(NsSpreadsheet + "cols");
                if (cols != null)
                {
                    foreach (XElement col in cols.Elements(NsSpreadsheet + "col"))
                    {
                        int min = ParseInt(col.Attribute("min"), 1);
                        int max = ParseInt(col.Attribute("max"), min);
                        double width = ParseDouble(col.Attribute("width"), defaultColumnWidth);
                        bool hidden = IsTrue(col.Attribute("hidden"));
                        columnRules.Add(new ColumnRule(min, max, width, hidden));
                    }
                }

                var rowRules = new Dictionary<int, RowRule>();
                foreach (XElement row in sheetDoc.Descendants(NsSpreadsheet + "row"))
                {
                    int rowIndex = ParseInt(row.Attribute("r"), 0);
                    if (rowIndex <= 0)
                    {
                        continue;
                    }

                    double height = ParseDouble(row.Attribute("ht"), defaultRowHeightPoints);
                    bool hidden = IsTrue(row.Attribute("hidden"));
                    rowRules[rowIndex] = new RowRule(height, hidden);
                }

                return new SheetMetrics(defaultColumnWidth, defaultRowHeightPoints, columnRules, rowRules);
            }

            public long GetColumnWidthEmu(int columnIndex)
            {
                double width = defaultColumnWidth;
                bool hidden = false;
                foreach (ColumnRule rule in columnRules)
                {
                    if (columnIndex >= rule.Min && columnIndex <= rule.Max)
                    {
                        width = rule.Width;
                        hidden = rule.Hidden;
                        break;
                    }
                }

                if (hidden)
                {
                    return 0;
                }

                return Math.Max(1, ExcelColumnWidthToPixels(width)) * EmusPerPixel;
            }

            public long GetRowHeightEmu(int rowIndex)
            {
                double heightPoints = defaultRowHeightPoints;
                bool hidden = false;
                RowRule rowRule;
                if (rowRules.TryGetValue(rowIndex, out rowRule))
                {
                    heightPoints = rowRule.HeightPoints;
                    hidden = rowRule.Hidden;
                }

                if (hidden)
                {
                    return 0;
                }

                return Math.Max(1, (long)Math.Round(heightPoints * EmusPerPoint));
            }

            private static int ExcelColumnWidthToPixels(double width)
            {
                if (width <= 0)
                {
                    return 0;
                }

                if (width < 1)
                {
                    return (int)Math.Floor(width * 12.0 + 0.5);
                }

                return (int)Math.Floor(width * 7.0 + 5.0);
            }

            private static int ParseInt(XAttribute attribute, int fallback)
            {
                int value;
                if (attribute != null && int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }

                return fallback;
            }

            private static double ParseDouble(XAttribute attribute, double fallback)
            {
                double value;
                if (attribute != null && double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }

                return fallback;
            }

            private static bool IsTrue(XAttribute attribute)
            {
                if (attribute == null)
                {
                    return false;
                }

                return attribute.Value == "1" || attribute.Value.Equals("true", StringComparison.OrdinalIgnoreCase);
            }

            private sealed class ColumnRule
            {
                public ColumnRule(int min, int max, double width, bool hidden)
                {
                    Min = min;
                    Max = max;
                    Width = width;
                    Hidden = hidden;
                }

                public int Min { get; private set; }
                public int Max { get; private set; }
                public double Width { get; private set; }
                public bool Hidden { get; private set; }
            }

            private sealed class RowRule
            {
                public RowRule(double heightPoints, bool hidden)
                {
                    HeightPoints = heightPoints;
                    Hidden = hidden;
                }

                public double HeightPoints { get; private set; }
                public bool Hidden { get; private set; }
            }
        }

        private sealed class ImageDimensions
        {
            public ImageDimensions(int width, int height)
            {
                Width = width;
                Height = height;
            }

            public int Width { get; private set; }
            public int Height { get; private set; }
        }

        private sealed class ConversionResult
        {
            public ConversionResult(string outputPath, int imageCount)
            {
                OutputPath = outputPath;
                ImageCount = imageCount;
            }

            public string OutputPath { get; private set; }
            public int ImageCount { get; private set; }
        }

        private sealed class CellImagePlacement
        {
            public CellImagePlacement(XElement cell, string address, int rowIndex, int columnIndex, string imageId, string imagePath)
            {
                Cell = cell;
                Address = address;
                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
                ImageId = imageId;
                ImagePath = imagePath;
            }

            public XElement Cell { get; private set; }
            public string Address { get; private set; }
            public int RowIndex { get; private set; }
            public int ColumnIndex { get; private set; }
            public string ImageId { get; private set; }
            public string ImagePath { get; private set; }
        }

        private sealed class SheetDrawingInfo
        {
            public SheetDrawingInfo(string drawingPath, string drawingRelationshipsPath, XDocument drawingDocument, XDocument drawingRelationshipsDocument)
            {
                DrawingPath = drawingPath;
                DrawingRelationshipsPath = drawingRelationshipsPath;
                DrawingDocument = drawingDocument;
                DrawingRelationshipsDocument = drawingRelationshipsDocument;
            }

            public string DrawingPath { get; private set; }
            public string DrawingRelationshipsPath { get; private set; }
            public XDocument DrawingDocument { get; private set; }
            public XDocument DrawingRelationshipsDocument { get; private set; }
        }
    }
}
