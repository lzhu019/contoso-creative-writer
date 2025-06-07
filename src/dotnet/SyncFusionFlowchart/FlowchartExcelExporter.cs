using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Syncfusion.XlsIO;

namespace SyncFusionFlowchart
{
    public class FlowchartExcelExporter
    {
        public void ConvertJsonToExcelFlowchart(string json, string outputPath)
        {
            var data = JsonConvert.DeserializeObject<DiagramData>(json);
            if (data == null)
            {
                throw new ArgumentException("Invalid JSON input.", nameof(json));
            }

            using var excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            var nodeShapes = new Dictionary<string, IShape>();
            foreach (var node in data.Nodes)
            {
                var shapeType = MapShape(node.Shape);
                var shape = sheet.Shapes.AddShape(shapeType, node.X, node.Y, node.Width, node.Height);
                shape.Text = node.Text;
                nodeShapes[node.Id] = shape;
            }

            foreach (var connector in data.Connectors)
            {
                if (nodeShapes.TryGetValue(connector.SourceId, out var source) &&
                    nodeShapes.TryGetValue(connector.TargetId, out var target))
                {
                    var line = sheet.Shapes.AddConnector(ExcelConnectorType.Straight, 0, 0, 10, 10);
                    line.BeginConnect(source, 1);
                    line.EndConnect(target, 1);
                }
            }

            workbook.SaveAs(outputPath);
        }

        private ExcelShapeType MapShape(string shape)
        {
            return shape?.ToLowerInvariant() switch
            {
                "terminator" => ExcelShapeType.FlowChartTerminator,
                "process" => ExcelShapeType.FlowChartProcess,
                "decision" => ExcelShapeType.FlowChartDecision,
                _ => ExcelShapeType.Rectangle
            };
        }
    }

    public class DiagramData
    {
        [JsonProperty("nodes")]
        public List<Node> Nodes { get; set; } = new();

        [JsonProperty("connectors")]
        public List<Connector> Connectors { get; set; } = new();
    }

    public class Node
    {
        [JsonProperty("id")]
        public string Id { get; set; } = string.Empty;

        [JsonProperty("text")]
        public string Text { get; set; } = string.Empty;

        [JsonProperty("shape")]
        public string Shape { get; set; } = string.Empty;

        [JsonProperty("x")]
        public double X { get; set; }

        [JsonProperty("y")]
        public double Y { get; set; }

        [JsonProperty("width")]
        public double Width { get; set; } = 60;

        [JsonProperty("height")]
        public double Height { get; set; } = 40;
    }

    public class Connector
    {
        [JsonProperty("sourceId")]
        public string SourceId { get; set; } = string.Empty;

        [JsonProperty("targetId")]
        public string TargetId { get; set; } = string.Empty;
    }
}
