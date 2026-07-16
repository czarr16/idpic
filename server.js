const express = require("express");
const multer  = require("multer");
const cors    = require("cors");
const path    = require("path");
const fs      = require("fs");

const { Document, Packer, Paragraph, ImageRun, TextWrappingType,
        HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom } = require("docx");

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ─────────────────────────────────────────
// Shared constants (match all reference .docx templates)
// ─────────────────────────────────────────
const EMU_PER_INCH  = 914400;
const TWIP_PER_INCH = 1440;
const PAGE_W_TWIP   = 11906;
const PAGE_H_TWIP   = 16838;
const MARGIN_TWIP   = 173;       // 0.12"
const MARGIN_IN     = MARGIN_TWIP / TWIP_PER_INCH;
const BORDER_COLOR  = "D9D9D9"; // White, Background 1, Darker 15%
const BORDER_W_EMU  = 3175;     // 0.25pt

function inchToEmu(inches) {
  return Math.round(inches * EMU_PER_INCH);
}

function floatingImg(buffer, wIn, hIn, xIn, yIn) {
  return new ImageRun({
    data: buffer,
    transformation: {
      width:  inchToEmu(wIn) / 9525,
      height: inchToEmu(hIn) / 9525,
    },
    floating: {
      horizontalPosition: {
        relative: HorizontalPositionRelativeFrom.PAGE,
        offset: inchToEmu(xIn),
      },
      verticalPosition: {
        relative: VerticalPositionRelativeFrom.PAGE,
        offset: inchToEmu(yIn),
      },
      wrap: { type: TextWrappingType.NONE },
      behindDocument: false,
      zIndex: 1,
    },
    outline: {
      type: "solidFill",
      solidFillType: "rgb",
      value: BORDER_COLOR,
      width: BORDER_W_EMU,
    },
  });
}

// ─────────────────────────────────────────
// Package layout builders
// ─────────────────────────────────────────
function buildPackage1(buf) {
  // 8× 1×1, single row
  const images = [];
  for (let i = 0; i < 8; i++)
    images.push(floatingImg(buf, 1, 1, MARGIN_IN + i * 1, MARGIN_IN));
  return images;
}

function buildPackage2(buf) {
  // 8× 2×2, 4 cols × 2 rows
  const images = [];
  for (let r = 0; r < 2; r++)
    for (let c = 0; c < 4; c++)
      images.push(floatingImg(buf, 2, 2, MARGIN_IN + c * 2, MARGIN_IN + r * 2));
  return images;
}

function buildPackage3(buf) {
  // 4× 2×2 top row + 8× 1×1 bottom row
  const images = [];
  for (let c = 0; c < 4; c++)
    images.push(floatingImg(buf, 2, 2, MARGIN_IN + c * 2, MARGIN_IN));
  for (let c = 0; c < 8; c++)
    images.push(floatingImg(buf, 1, 1, MARGIN_IN + c * 1, MARGIN_IN + 2));
  return images;
}

function buildPackage4(buf) {
  // 10× passport (1.3694" × 1.7694"), 5 cols × 2 rows
  const pw     = 1252220 / EMU_PER_INCH;
  const ph     = 1617980 / EMU_PER_INCH;
  const startX = 524510  / EMU_PER_INCH;
  const stepX  = 1257935 / EMU_PER_INCH;
  const rowY   = [17780 / EMU_PER_INCH, 1641475 / EMU_PER_INCH];
  const images = [];
  for (let r = 0; r < 2; r++)
    for (let c = 0; c < 5; c++)
      images.push(floatingImg(buf, pw, ph, startX + c * stepX, rowY[r]));
  return images;
}

const BUILDERS = { 1: buildPackage1, 2: buildPackage2, 3: buildPackage3, 4: buildPackage4 };

// ─────────────────────────────────────────
// POST /generate
// Body: multipart/form-data
//   photo    — image file (PNG/JPG)
//   package  — 1 | 2 | 3 | 4
// Returns: .docx file download
// ─────────────────────────────────────────
app.post("/generate", upload.single("photo"), async (req, res) => {
  try {
    const packageId = parseInt(req.body.package, 10);
    const builder   = BUILDERS[packageId];

    if (!builder) {
      return res.status(400).json({ error: `Unknown package: ${packageId}. Valid: 1–4` });
    }
    if (!req.file) {
      return res.status(400).json({ error: "No photo uploaded." });
    }

    const imageBuffer = req.file.buffer;
    const images      = builder(imageBuffer);

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size:   { width: PAGE_W_TWIP, height: PAGE_H_TWIP },
            margin: { top: MARGIN_TWIP, right: MARGIN_TWIP, bottom: MARGIN_TWIP, left: MARGIN_TWIP },
          },
        },
        children: [new Paragraph({ children: images })],
      }],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader("Content-Type",        "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="package_${packageId}_layout.docx"`);
    res.send(Buffer.from(buffer));

  } catch (err) {
    console.error("Generation error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────
// Start server
// ─────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  ID Photo Studio running at http://localhost:${PORT}\n`);
});
