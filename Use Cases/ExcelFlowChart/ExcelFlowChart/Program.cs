using Syncfusion.XlsIO;
using Syncfusion.Drawing;

/// <summary>
/// Builds an Order Fulfillment flowchart in an Excel worksheet using Syncfusion XlsIO.
/// </summary>
internal class Program
{
    /// <summary>
    /// Entry point. Creates a workbook, draws the flowchart, applies colors, and saves the file.
    /// </summary>
    static void Main()
    {
        // Initialize ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            // Use XLSX format by default
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Create a new workbook with one worksheet
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Name & present the canvas
            worksheet.Name = "Order Fulfillment Workflow";
            worksheet.IsGridLinesVisible = false;

            // ----- Shapes (row/col are 1-based anchors; height/width are in points) -----
            // Center column = 9, Right column = 14, as per your layout.
            IShape start = AddFlowChartShape(worksheet, 2, 9, 50, 170, "Start", AutoShapeType.FlowChartTerminator);
            IShape receiveOrder = AddFlowChartShape(worksheet, 6, 9, 50, 170, "Receive Order", AutoShapeType.FlowChartProcess);
            IShape checkInv = AddFlowChartShape(worksheet, 10, 9, 50, 170, "Check Inventory", AutoShapeType.FlowChartProcess);
            IShape invAvailable = AddFlowChartShape(worksheet, 14, 9, 50, 170, "Inventory Available?", AutoShapeType.FlowChartDecision);

            // No branch (left/vertical)
            IShape noNotify = AddFlowChartShape(worksheet, 18, 9, 50, 170, "Notify Customer", AutoShapeType.FlowChartProcess);
            IShape backOrCan = AddFlowChartShape(worksheet, 24, 9, 50, 170, "Backorder or Cancel", AutoShapeType.FlowChartProcess);
            IShape leftEnd = AddFlowChartShape(worksheet, 30, 9, 50, 170, "End", AutoShapeType.FlowChartTerminator);

            // Yes branch (right/vertical)
            IShape payment = AddFlowChartShape(worksheet, 14, 14, 50, 170, "Process Payment", AutoShapeType.FlowChartProcess);
            IShape packed = AddFlowChartShape(worksheet, 18, 14, 50, 170, "Pack Order", AutoShapeType.FlowChartProcess);
            IShape shipped = AddFlowChartShape(worksheet, 24, 14, 50, 170, "Ship Order", AutoShapeType.FlowChartProcess);
            IShape yesNotify = AddFlowChartShape(worksheet, 30, 14, 50, 170, "Notify Customer", AutoShapeType.FlowChartProcess);

            // ----- Connectors -----
            // Decision → branches: fromSite/toSite mapping (0:Top, 1:Right, 2:Bottom, 3:Left)
            Connect(worksheet, start, 2, receiveOrder, 0, true);  
            Connect(worksheet, receiveOrder, 2, checkInv, 0, true);
            Connect(worksheet, checkInv, 2, invAvailable, 0, true);
            Connect(worksheet, invAvailable, 1, payment, 3, true); 
            Connect(worksheet, invAvailable, 2, noNotify, 0, true);

            // Left chain (No branch)
            Connect(worksheet, noNotify, 2, backOrCan, 0, true);
            Connect(worksheet, backOrCan, 2, leftEnd, 0, true);

            // Right chain (Yes branch)
            Connect(worksheet, payment, 2, packed, 0, true);
            Connect(worksheet, packed, 2, shipped, 0, true);
            Connect(worksheet, shipped, 2, yesNotify, 0, true);
            Connect(worksheet, yesNotify, 3, leftEnd, 1, false); // Left to End; arrow at beginning to indicate flow into End

            // ----- Decision branch labels -----
            // "Yes" near the right-going branch from the decision
            AddLabel(worksheet, 14, 12, "Yes");

            // "No" near the downward branch from the decision
            AddLabel(worksheet, 17, 9, "No");

            // Colors
            Color colorTerminator = ColorTranslator.FromHtml("#10B981"); // Emerald 500
            Color colorProcess = ColorTranslator.FromHtml("#3B82F6"); // Blue 500
            Color colorDecision = ColorTranslator.FromHtml("#F59E0B"); // Amber 500
            Color colorShip = ColorTranslator.FromHtml("#8B5CF6"); // Violet 500
            Color colorNotify = ColorTranslator.FromHtml("#14B8A6"); // Teal 500
            Color colorLine = ColorTranslator.FromHtml("#1F2937"); // Slate 800 (borders/connectors)
            Color colorAccent = ColorTranslator.FromHtml("#2563EB"); // Blue 600 (reserved if needed)

            // Apply fills (lines are set in AddFlowChartShape)
            start.Fill.ForeColor = colorTerminator;
            receiveOrder.Fill.ForeColor = colorProcess;
            checkInv.Fill.ForeColor = colorProcess;
            invAvailable.Fill.ForeColor = colorDecision;

            noNotify.Fill.ForeColor = colorNotify;
            backOrCan.Fill.ForeColor = colorNotify;
            leftEnd.Fill.ForeColor = colorTerminator;

            payment.Fill.ForeColor = colorProcess;
            packed.Fill.ForeColor = colorProcess;
            shipped.Fill.ForeColor = colorShip;
            yesNotify.Fill.ForeColor = colorNotify;

            // Save the workbook
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OrderFulfillmentFlowchart.xlsx");
            workbook.SaveAs(outputPath);
        }
    }

    /// <summary>
    /// Adds a flowchart shape to <paramref name="worksheet"/> at the given anchor.
    /// </summary>
    /// <param name="worksheet">Target worksheet.</param>
    /// <param name="row">Top anchor row (1-based).</param>
    /// <param name="col">Left anchor column (1-based).</param>
    /// <param name="height">Shape height in points.</param>
    /// <param name="width">Shape width in points.</param>
    /// <param name="text">Caption rendered inside the shape.</param>
    /// <param name="flowChartShapeType">Flowchart/auto shape type.</param>
    /// <returns>The created <see cref="IShape"/>.</returns>
    public static IShape AddFlowChartShape(IWorksheet worksheet, int row, int col, int height, int width, string text, AutoShapeType flowChartShapeType)
    {
        // Create the shape anchored at (row, col)
        IShape flowChartShape = worksheet.Shapes.AddAutoShapes(flowChartShapeType, row, col, height, width);

        // Basic style
        flowChartShape.Line.Weight = 1.25;

        // Content & alignment
        flowChartShape.TextFrame.TextRange.Text = text;
        flowChartShape.TextFrame.HorizontalAlignment = ExcelHorizontalAlignment.CenterMiddle;
        flowChartShape.TextFrame.VerticalAlignment = ExcelVerticalAlignment.Middle;

        return flowChartShape;
    }


    /// <summary>
    /// Adds a simple text label (transparent rectangle) for annotations such as "Yes"/"No".
    /// </summary>
    /// <param name="worksheet">Target worksheet.</param>
    /// <param name="row">Top anchor row (1-based).</param>
    /// <param name="col">Left anchor column (1-based).</param>
    /// <param name="text">Label text.</param>
    /// <returns>The created label <see cref="IShape"/>.</returns>
    public static IShape AddLabel(IWorksheet worksheet, int row, int col, string text)
    {
        IShape rectangle = worksheet.Shapes.AddAutoShapes(AutoShapeType.Rectangle, row, col, 14, 40);

        // Make it text-only
        rectangle.Fill.Transparency = 1f;
        rectangle.Line.Visible = false;

        // Content & alignment
        rectangle.TextFrame.TextRange.Text = text;
        rectangle.TextFrame.HorizontalAlignment = ExcelHorizontalAlignment.CenterMiddle;
        rectangle.TextFrame.VerticalAlignment = ExcelVerticalAlignment.Middle;

        // Nudge to the right to avoid overlapping connectors
        rectangle.Left += 30;

        return rectangle;
    }

    /// <summary>
    /// Connects two shapes with a straight connector.
    /// </summary>
    /// <param name="worksheet">Worksheet on which to draw the connector.</param>
    /// <param name="from">Start shape.</param>
    /// <param name="fromSite">Connection site on start shape (0:Top, 1:Right, 2:Bottom, 3:Left).</param>
    /// <param name="to">End shape.</param>
    /// <param name="toSite">Connection site on end shape (0:Top, 1:Right, 2:Bottom, 3:Left).</param>
    /// <param name="isEnd">If <c>true</c>, adds the arrow head at the <b>end</b>; otherwise at the <b>beginning</b>.</param>
    /// <returns>The created connector <see cref="IShape"/>.</returns>
    public static IShape Connect(IWorksheet worksheet, IShape from, int fromSite, IShape to, int toSite, bool isEnd)
    {
        // Absolute positions (in points) on the worksheet
        PointF startPoint = GetConnectionPoint(from, fromSite);
        PointF endPoint = GetConnectionPoint(to, toSite);

        // Bounding box for the straight connector
        float left = Math.Min(startPoint.X, endPoint.X);
        float top = Math.Min(startPoint.Y, endPoint.Y);
        double width = Math.Abs(endPoint.X - startPoint.X);
        double height = Math.Abs(endPoint.Y - startPoint.Y);

        // Draw a straight connector and place it
        IShape connector = worksheet.Shapes.AddAutoShapes(AutoShapeType.StraightConnector, 1, 1, (int)height, (int)width);
        connector.Left = (int)left;
        connector.Top = (int)top;

        // Arrow style
        if (isEnd)
            connector.Line.EndArrowHeadStyle = ExcelShapeArrowStyle.LineArrow;
        else
            connector.Line.BeginArrowHeadStyle = ExcelShapeArrowStyle.LineArrow;

        connector.Line.Weight = 1.25;
        return connector;
    }

    /// <summary>
    /// Computes a connection point (in points) on the boundary of <paramref name="shape"/> for site indices:
    /// 0 = Top, 1 = Right, 2 = Bottom, 3 = Left; any other value returns the center.
    /// </summary>
    /// <param name="shape">The source shape.</param>
    /// <param name="site">Connection site index.</param>
    /// <returns>Absolute point coordinates (Left/Top in points) on the sheet where the connector should attach.</returns>
    private static PointF GetConnectionPoint(IShape shape, int site)
    {
        float x = shape.Left;
        float y = shape.Top;

        switch (site)
        {
            case 0: // Top
                x += shape.Width / 2;
                break;
            case 1: // Right
                x += shape.Width;
                y += shape.Height / 2;
                break;
            case 2: // Bottom
                x += shape.Width / 2;
                y += shape.Height;
                break;
            case 3: // Left
                y += shape.Height / 2;
                break;
            default: // Center
                x += shape.Width / 2;
                y += shape.Height / 2;
                break;
        }

        return new PointF(x, y);
    }
}
