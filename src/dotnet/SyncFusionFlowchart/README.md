# SyncFusion Flowchart Excel Exporter

This sample project demonstrates how to parse a SyncFusion diagram JSON payload and generate a flowchart in an Excel file using the Syncfusion.XlsIO library.

## Usage

1. Add the `SyncFusionFlowchart` project to your solution or reference the compiled assembly.
2. Call `FlowchartExcelExporter.ConvertJsonToExcelFlowchart` passing the JSON string and desired output path for the Excel file.

The JSON payload is expected to contain a list of `nodes` and `connectors` similar to the diagram serialization format used by SyncFusion diagrams.
