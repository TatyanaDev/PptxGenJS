import pptxgen from "pptxgenjs";

export async function createPptx(): Promise<Buffer> {
  const pptx = new pptxgen();

  // Define a custom slide layout for the presentation
  pptx.defineLayout({ name: "CUSTOM", width: 13.33, height: 7.5 });
  pptx.layout = "CUSTOM";

  const slide = pptx.addSlide();

  // Define color palette variables for consistent styling
  const lightGreen = "E8F1EB"; // Light green for specific elements
  const darkGreen = "58A65C"; // Dark green for specific elements
  const blueText = "1A294B"; // Dark blue for main text/headers
  const bgColor = "F8F9FB"; // Light grey background color
  const white = "FFFFFF"; // White color
  const black = "000000"; // Black color
  const green = "53A457"; // Green color

  // Set the background color of the slide
  slide.background = { fill: bgColor };

  // Define general layout parameters
  const margin = 0.4; // Margin from the slide edges (in inches)
  const contentWidth = 13.33 - 2 * margin; // Calculated width for content area
  const contentHeight = 7.5 - 2 * margin; // Calculated height for content area

  // Add the main title of the slide
  slide.addText("The Dependencies Dilemma", {
    x: margin,
    y: margin,
    w: contentWidth,
    h: 0.6, // Height of the title text box
    fontSize: 30,
    bold: true,
    color: blueText,
  });

  // Add a rounded rectangle shape for the quote block
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: margin + 0.7, // Positioned below the main title
    w: contentWidth,
    h: 0.7, // Height of the quote block
    fill: { color: white }, // White background for the block
    rectRadius: 0.1, // Radius for rounded corners
    shadow: {
      type: "outer",
      angle: 180, // Shadow angle (180 degrees points downwards)
      blur: 10, // Blur radius of the shadow
      offset: 0.05, // Offset distance of the shadow
      opacity: 0.15, // Opacity of the shadow
      color: black, // Color of the shadow (black)
      // Note: pptxgenjs's shadow property tends to apply shadow to all sides
      // even with specific angles. For a bottom-only shadow, a custom approach
      // with multiple shapes might be needed, but this setting was preferred stylistically.
    },
  });

  // Add the quote text inside the quote block
  slide.addText(
    [
      {
        text: "The value of an initiative isn't just its immediate impact,",
        options: { fontSize: 16, italic: true, color: black },
      },
      {
        text: " but what it unlocks ",
        options: { fontSize: 16, italic: true, color: green }, // Green text for "unlocks"
      },
      {
        text: "🔓", // Unicode lock emoji
        options: { fontSize: 16 },
      },
      {
        text: ".",
        options: { fontSize: 16, italic: true, color: black },
      },
    ],
    {
      x: margin + 0.3, // X-position of text within the block (offset from block's x)
      y: margin + 0.7, // Y-position of text (same as block's y)
      w: contentWidth - 0.6, // Width of text area (block width minus double x-offset)
      h: 0.7, // Height of text area (same as block height)
      valign: "middle", // Vertical alignment of text
      shape: pptx.ShapeType.rect, // Text container shape
    }
  );

  // Calculate dimensions and positions for the two-column layout below the quote
  const columnGap = 0.4; // Gap between the left and right columns
  const rightW = (contentWidth - columnGap) * 0.5; // Width of the right column (half of available width minus gap)
  const topY = margin + 1.6; // Starting Y-coordinate for the columns content
  const gridH = contentHeight - 1.6; // Height available for the column content
  const leftW = (contentWidth - columnGap) * 0.5; // Width of the left column

  // Define gaps and heights for the three blocks in the left column
  const blockGap = 0.2; // Vertical gap between blocks
  const blockH1 = (gridH - 2 * blockGap) * 0.4; // Height of the first block (40% of available height)
  const blockH2 = (gridH - 2 * blockGap) * 0.4; // Height of the second block (40% of available height)
  const blockH3 = (gridH - 2 * blockGap) * 0.2; // Height of the third block (20% of available height)

  // Calculate Y-coordinates for each block in the left column
  const blockY1 = topY;
  const blockY2 = blockY1 + blockH1 + blockGap;
  const blockY3 = blockY2 + blockH2 + blockGap;

  // Left Column: "Real-World Example" block (top block)
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY1,
    w: leftW,
    h: blockH1,
    fill: { color: white }, // White background
    line: { color: "DDDDDD" }, // Light grey border
    rectRadius: 0.1,
  });

  // Text content for "Real-World Example" block
  slide.addText(
    [
      { text: "Real-World Example\n", options: { fontSize: 14, bold: true, color: blueText } },
      {
        text: "A fintech startup invested in comprehensive KYC infrastructure that enabled:\n",
        options: { fontSize: 12, color: black },
      },
      { text: "• Launch in 4 new countries within 12 months\n", options: { fontSize: 12, color: black } },
      { text: "• Add 3 regulated financial products\n", options: { fontSize: 12, color: black } },
      { text: "• Partner with 2 major banks\n", options: { fontSize: 12, color: black } },
      { text: "• Achieve compliance in weeks instead of months", options: { fontSize: 12, color: black } },
    ],
    {
      x: margin + 0.2, // X-offset for text inside the block
      y: blockY1 + 0.2, // Y-offset for text inside the block
      w: leftW - 0.4, // Text width (block width minus double x-offset)
      h: blockH1 - 0.4, // Text height (block height minus double y-offset)
      valign: "top", // Vertical alignment of text
    }
  );

  // Left Column: "Dependency Mapping" block (middle block)
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY2,
    w: leftW,
    h: blockH2,
    fill: { color: white },
    rectRadius: 0.1,
    shadow: {
      type: "outer",
      angle: 180,
      blur: 10,
      offset: 0.05,
      opacity: 0.15,
      color: black,
    },
  });

  // Title for "Dependency Mapping" block
  slide.addText("Dependency Mapping", {
    x: margin + 0.2,
    y: blockY2 + 0.2,
    w: leftW - 0.4,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: blueText,
  });

  // Parameters for the bulleted list within "Dependency Mapping"
  const bulletColor = "4285F4"; // Blue color for bullets
  const bulletSize = 0.15; // Diameter of the bullet circle
  const itemFontSize = 12; // Font size for list items
  const textStartX = margin + 0.2 + bulletSize + 0.15; // X-start for text (bullet x + bullet size + small gap)
  const listYStart = blockY2 + 0.6; // Starting Y-coordinate for the list
  const lineHeight = 0.4; // Vertical space allocated for each list item

  const items = ["Foundation capabilities vs. surface features", "Regulatory infrastructure unlocks market expansion", "Compliance systems enable product diversification"];

  // Add the bulleted list items
  items.forEach((text, idx) => {
    const y = listYStart + idx * lineHeight;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: margin + 0.2,
      y: y + 0.1, // Adjust Y to vertically center bullet with text
      w: bulletSize,
      h: bulletSize,
      fill: { color: bulletColor },
      line: { color: bulletColor },
    });
    slide.addText(text, {
      x: textStartX,
      y,
      w: leftW - (textStartX - margin), // Width of text area
      h: lineHeight, // Height of text area
      fontSize: itemFontSize,
      color: black,
      valign: "middle",
    });
  });

  // Left Column: "Key Insight" block (bottom block)
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY3,
    w: leftW,
    h: blockH3,
    fill: { color: white },
  });
  // These two shapes create a layered effect for the "Key Insight" block,
  // giving it a distinct visual style with a green border/background.
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY3,
    w: leftW,
    h: blockH3,
    fill: { color: darkGreen }, // Dark green background
    rectRadius: 0.1,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin + 0.1, // Slightly offset to create a border effect
    y: blockY3,
    w: leftW - 0.1, // Reduced width
    h: blockH3,
    fill: { color: lightGreen }, // Light green inner fill
    rectRadius: 0.1,
  });

  // Text content for "Key Insight" block
  slide.addText(
    [
      { text: "Key Insight: ", options: { bold: true, color: darkGreen, fontSize: 12 } },
      {
        text: "Foundation investments create exponential value through what they unlock, not just their direct impact.",
        options: { color: black, fontSize: 12 },
      },
    ],
    {
      x: margin + 0.3, // X-offset for text
      y: blockY3 + 0.2, // Y-offset for text
      w: leftW - 0.6, // Text width
      h: blockH3 - 0.4, // Text height
      valign: "top",
      align: "left",
    }
  );

  // Right Column: "Feature Enablement Tree" block
  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin + leftW + columnGap, // Positioned to the right of the left column
    y: topY + 0.4, // Y-position, slightly offset from topY
    w: rightW,
    h: gridH - 1, // Height of the tree container
    fill: { color: white },
    rectRadius: 0.1,
    shadow: {
      type: "outer",
      angle: 180,
      blur: 10,
      offset: 0.05,
      opacity: 0.15,
      color: black,
    },
  });

  // Parameters for the "Feature Enablement Tree" diagram
  const treeX = margin + leftW + columnGap; // X-start for the tree diagram
  const treeY = topY + 0.5; // Y-start for the tree diagram
  const treeW = rightW; // Width of the tree diagram area
  const colCount = 4; // Number of columns for the top two levels
  const boxW = 1.2; // Width of each individual box in the tree
  const boxH = 0.4; // Height of each individual box in the tree
  const spacingX = (treeW - colCount * boxW) / (colCount + 1); // Horizontal spacing between boxes
  const spacingY = 0.35; // Vertical spacing between levels

  // Calculate Y-coordinates for each level of the tree
  const levelsY = Array.from({ length: 4 }, (_, i) => treeY + 0.5 + i * (boxH + spacingY));
  // Calculate X-coordinates for columns (used by top two levels)
  const colsX = Array.from({ length: colCount }, (_, i) => treeX + spacingX + i * (boxW + spacingX));

  // Define colors for different levels/categories in the tree
  const colors = {
    revenue: "A52A2A", // Brownish-red for Revenue
    product: "F4B400", // Yellow/Orange for Product
    compliance: "0F9D58", // Green for Compliance
    foundation: "4285F4", // Blue for Foundation
  };

  // Data for each level of the tree, with line breaks for multi-line text
  const revenue = ["Banking-as-a-\nService", "White-Label\nSolutions", "Cross-Border\nPayments", "Institutional\nTrading"];
  const products = ["International\nMarkets", "Business\nBanking", "Investment\nPlatform", "Lending\nProducts"];
  const compliance = ["AML\nMonitoring", "Regulatory\nReporting", "Risk\nAssessment"];
  const foundation = ["KYC/Identity\nVerification"];

  // Helper function to draw a single box with text
  function drawBox(text: string, x: number, y: number, color: string, w = boxW) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h: boxH,
      fill: { color },
    });

    slide.addText(text, {
      x,
      y,
      w,
      h: boxH,
      align: "center",
      valign: "middle",
      fontSize: 8,
      bold: true,
      color: white,
    });
  }

  // Add title for the "Feature Enablement Tree"
  slide.addText("Feature Enablement Tree", {
    x: treeX,
    y: treeY,
    w: treeW,
    h: 0.4,
    align: "center",
    valign: "middle",
    color: blueText,
    fontSize: 14,
  });

  // Draw Revenue boxes (top level)
  revenue.forEach((text, i) => drawBox(text, colsX[i], levelsY[0], colors.revenue));

  // Draw Product boxes (second level)
  products.forEach((text, i) => drawBox(text, colsX[i], levelsY[1], colors.product));

  // Draw lines connecting Revenue to Product boxes (vertical lines)
  for (let i = 0; i < colCount; i++) {
    const centerX = colsX[i] + boxW / 2;
    const fromY = levelsY[0] + boxH;
    const toY = levelsY[1];
    slide.addShape(pptx.ShapeType.line, {
      x: centerX,
      y: fromY,
      w: 0, // Vertical line, so width is 0
      h: toY - fromY, // Height is the difference in Y-coordinates
      line: { color: colors.product, width: 1.5 },
    });
  }

  // Calculate positions for Compliance boxes (third level)
  const productsLeftX = colsX[0];
  const productsRightX = colsX[3] + boxW;
  const complianceAreaCenter = (productsLeftX + productsRightX) / 2; // Center of the product area
  const complianceTotalWidth = 3 * boxW + 2 * spacingX; // Total width for 3 compliance boxes with gaps
  const complianceStartX = complianceAreaCenter - complianceTotalWidth / 2; // Starting X for compliance boxes
  const complianceXs = [0, 1, 2].map((i) => complianceStartX + i * (boxW + spacingX)); // X-coordinates for compliance boxes

  // Draw Compliance boxes
  compliance.forEach((text, i) => drawBox(text, complianceXs[i], levelsY[2], colors.compliance));

  // Draw lines connecting Product to Compliance boxes (diagonal lines)
  // This uses the Math.min/Math.abs and flipH/flipV approach for better PowerPoint compatibility.
  [
    { from: 0, to: 0 }, // Product 0 (International Markets) connects to Compliance 0 (AML Monitoring)
    { from: 1, to: 0 }, // Product 1 (Business Banking) connects to Compliance 0 (AML Monitoring)
    { from: 2, to: 1 }, // Product 2 (Investment Platform) connects to Compliance 1 (Regulatory Reporting)
    { from: 3, to: 2 }, // Product 3 (Lending Products) connects to Compliance 2 (Risk Assessment)
  ].forEach(({ from, to }) => {
    const x1 = colsX[from] + boxW / 2; // Center X of source Product box
    const y1 = levelsY[1] + boxH; // Bottom Y of source Product box
    const x2 = complianceXs[to] + boxW / 2; // Center X of target Compliance box
    const y2 = levelsY[2]; // Top Y of target Compliance box

    slide.addShape(pptx.ShapeType.line, {
      x: Math.min(x1, x2), // Start X for the line (minimum of source/target X)
      y: Math.min(y1, y2), // Start Y for the line (minimum of source/target Y)
      w: Math.abs(x2 - x1), // Width of the bounding box for the line
      h: Math.abs(y2 - y1), // Height of the bounding box for the line
      flipH: x2 < x1, // Flip horizontally if target X is less than source X
      flipV: y2 < y1, // Flip vertically if target Y is less than source Y
      line: {
        color: colors.compliance,
        width: 1.5,
      },
    });
  });

  // Calculate position for Foundation box (bottom level)
  const foundationW = boxW * 2 + spacingX; // Width of the Foundation box (2 standard boxes + 1 spacing)
  const foundationCenterX = complianceXs[1] + boxW / 2; // Center X of the middle Compliance box
  const foundationX = foundationCenterX - foundationW / 2; // Starting X for the Foundation box

  // Draw Foundation box
  drawBox(foundation[0], foundationX, levelsY[3], colors.foundation, foundationW);

  // Draw lines connecting Compliance to Foundation box (diagonal lines)
  // This also uses the Math.min/Math.abs and flipH/flipV approach for better PowerPoint compatibility.
  complianceXs.forEach((x) => {
    const fromX = x + boxW / 2; // Center X of source Compliance box
    const fromY = levelsY[2] + boxH; // Bottom Y of source Compliance box
    const toY = levelsY[3]; // Top Y of target Foundation box
    const delta = fromX - foundationCenterX; // Horizontal distance from compliance box center to foundation box center
    const toX = foundationCenterX + delta * 0.8; // Adjusted target X for a more visually appealing connection

    slide.addShape(pptx.ShapeType.line, {
      x: Math.min(fromX, toX),
      y: Math.min(fromY, toY),
      w: Math.abs(toX - fromX),
      h: Math.abs(toY - fromY),
      flipH: toX < fromX,
      flipV: toY < fromY,
      line: {
        color: colors.foundation,
        width: 1.5,
      },
    });
  });

  // Define legend items
  const legendItems = [
    { label: "Foundation", color: colors.foundation },
    { label: "Compliance", color: colors.compliance },
    { label: "Products", color: colors.product },
    { label: "Revenue", color: colors.revenue },
  ];

  // Parameters for the legend
  const squareSize = 0.2; // Size of the color square in the legend
  const labelW = 1.2; // Width allocated for the text label
  const itemW = squareSize + labelW; // Total width for one legend item (square + label)
  const totalLegendW = legendItems.length * itemW; // Total width of the entire legend
  const whiteBoxX = margin + leftW + columnGap; // X-coordinate of the right column's container
  const whiteBoxW = rightW; // Width of the right column's container
  const legendX = whiteBoxX + (whiteBoxW - totalLegendW) / 2; // X-position to center the legend horizontally
  const legendY = levelsY[3] + boxH + 0.3; // Y-position for the legend (below the foundation box)

  // Add legend items (color squares and text labels)
  legendItems.forEach((item, i) => {
    const x = legendX + i * itemW; // X-position for the current legend item

    // Add the color square
    slide.addShape(pptx.ShapeType.rect, {
      x: x + squareSize, // Offset to the right to align with text
      y: legendY,
      w: squareSize,
      h: squareSize,
      fill: { color: item.color },
    });

    // Add the text label
    slide.addText(item.label, {
      x: x + squareSize + squareSize, // Position text after the square with a small gap
      y: legendY,
      w: labelW,
      h: squareSize, // Set text height to match square height for visual balance
      fontSize: 12, // Increased font size as requested
      valign: "middle", // Vertically center text
      color: black,
    });
  });

  // Generate the PowerPoint presentation as a Node.js Buffer
  const result = await pptx.write({ outputType: "nodebuffer" });

  return Buffer.from(result as ArrayBuffer);
}
