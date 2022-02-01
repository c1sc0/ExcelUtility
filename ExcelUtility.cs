using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelUtil
{
    public class ExcelUtility :IDisposable
    {
        private const int BlackCircleCoordinateX = 7;
        private readonly string _directory;
        private readonly XDocument _drawingVml;
        private readonly XDocument _drawingVmlRelations;
        private readonly XDocument _drawingXml;

        public ExcelUtility(string path)
        {
            var tempPath = Path.GetTempPath();
            _directory = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(path));
            if (Directory.Exists(_directory))
                Directory.Delete(_directory, true);
            ZipFile.ExtractToDirectory(path, _directory);
            _drawingXml = GetDrawingXDocument($"{_directory}/xl/drawings/drawing1.xml").Result;
            _drawingVml = GetDrawingXDocument($"{_directory}/xl/drawings/vmlDrawing1.vml").Result;
            _drawingVmlRelations = GetDrawingXDocument($@"{_directory}/xl/drawings/_rels/vmlDrawing1.vml.rels").Result;
        }

        public Task<bool> GetCheckboxStateByNameAsync(string name) => Task.FromResult(IsCheckBoxChecked(name));

        public async Task<bool> GetRadioButtonStateByNameAsync(string name)
        {
            var radioButtonRelationId = GetRadioButtonRelationId(name);
            var imageName = GetEmfImageName(radioButtonRelationId);
            return await IsRadioButtonChecked(imageName);
        }

        private bool IsCheckBoxChecked(string name)
        {
            var checkBoxIdentifier = CheckBoxIdentifier(name);
            return _drawingVml.Descendants("{urn:schemas-microsoft-com:vml}shape")
                .FirstOrDefault(_ => _.Attribute("id").Value == checkBoxIdentifier)
                ?.Descendants("{urn:schemas-microsoft-com:office:excel}Checked").FirstOrDefault()?.Value != null;
        }

        private static bool GetRadioButtonState(Bitmap image) => image.GetPixel(BlackCircleCoordinateX, image.Height / 2) == Color.FromArgb(255, Color.Black);

        private string CheckBoxIdentifier(string name)
        {
            try
            {
                return _drawingXml.Descendants()
                    .FirstOrDefault(_ => _.Name.LocalName.Contains("cNvPr") && _.Attribute("name").Value.Contains(name))
                    .Descendants("{http://schemas.microsoft.com/office/drawing/2010/main}compatExt").FirstOrDefault()
                    .Attribute("spid")?.Value;
            }
            catch (NullReferenceException)
            {
                throw new ControlNotFoundException(name);
            }
        }

        private string GetEmfImageName(string radioButtonRelationId)
        {
            return Path.GetFileName(_drawingVmlRelations.Root.Descendants()
                .FirstOrDefault(_ => _.Attribute("Id")?.Value == radioButtonRelationId).Attribute("Target").Value);
        }

        private string GetRadioButtonRelationId(string name)
        {
            return _drawingVml.Descendants("{urn:schemas-microsoft-com:vml}shape")
                .FirstOrDefault(_ => _.Attribute("id").Value == name)
                ?.Descendants().FirstOrDefault().Attributes().FirstOrDefault().Value;
        }

        private async Task<Bitmap> LoadImage(string imageName)
        {
            await using var fs = new FileStream(Path.Combine($"{_directory}/xl/media", imageName), FileMode.Open);
            return new Bitmap(fs);
        }

        private async Task<bool> IsRadioButtonChecked(string imageName)
        {
            var emfPicture = await LoadImage(imageName);
            return GetRadioButtonState(emfPicture);
        }

        private static async Task<XDocument> GetDrawingXDocument(string path)
        {
            await using var fileStream = new FileStream(path, FileMode.Open);
            var xDoc = await XDocument.LoadAsync(
                fileStream, LoadOptions.None,
                new CancellationToken());
            return xDoc;
        }

        public void Dispose()
        {
            Directory.Delete(_directory, true);
        }
    }
}