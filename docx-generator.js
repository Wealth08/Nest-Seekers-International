/**
 * docx-generator.js
 * Nest Seekers International — Rental Application
 * Generates and downloads a formatted .docx Word document
 * from the submitted form data using docx.js (CDN build).
 *
 * Usage: call generateAndDownloadDocx(data) where data is
 * a plain object of { fieldLabel: value } pairs.
 */

async function generateAndDownloadDocx(data) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
    HeadingLevel, PageOrientation
  } = docx;

  // ── Colour palette ───────────────────────────────────────────────────
  const GOLD  = "C9A84C";
  const DARK  = "0A0A0A";
  const WHITE = "FFFFFF";
  const LIGHT = "FAF8F4";
  const MUTED = "6B6B6B";

  // ── Helpers ──────────────────────────────────────────────────────────
  const thin = { style: BorderStyle.SINGLE, size: 1, color: "DDD5C8" };
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const cellBorders = { top: thin, bottom: thin, left: thin, right: thin };
  const noBorders   = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  function labelCell(text) {
    return new TableCell({
      width: { size: 3000, type: WidthType.DXA },
      borders: cellBorders,
      shading: { fill: "F0EBE0", type: ShadingType.CLEAR },
      margins: { top: 90, bottom: 90, left: 130, right: 130 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        children: [new TextRun({ text, bold: true, size: 18, color: MUTED, font: "Arial" })]
      })]
    });
  }

  function valueCell(text) {
    return new TableCell({
      width: { size: 6360, type: WidthType.DXA },
      borders: cellBorders,
      shading: { fill: LIGHT, type: ShadingType.CLEAR },
      margins: { top: 90, bottom: 90, left: 130, right: 130 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        children: [new TextRun({ text: String(text || "—"), size: 20, color: DARK, font: "Arial" })]
      })]
    });
  }

  function dataRow(label, value) {
    return new TableRow({ children: [labelCell(label), valueCell(value)] });
  }

  function sectionHeaderRow(title) {
    return new TableRow({
      children: [new TableCell({
        columnSpan: 2,
        borders: noBorders,
        shading: { fill: DARK, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 130, right: 130 },
        children: [new Paragraph({
          children: [new TextRun({ text: title.toUpperCase(), bold: true, size: 20, color: GOLD, font: "Arial" })]
        })]
      })]
    });
  }

  function spacerRow() {
    return new TableRow({
      children: [new TableCell({
        columnSpan: 2,
        borders: noBorders,
        shading: { fill: WHITE, type: ShadingType.CLEAR },
        margins: { top: 40, bottom: 40, left: 0, right: 0 },
        children: [new Paragraph({ children: [] })]
      })]
    });
  }

  // ── Build main data table ────────────────────────────────────────────
  const tableRows = [
    // Section 1
    sectionHeaderRow("1. Personal Information"),
    dataRow("Full Name",           `${data["First Name"] || ""} ${data["Last Name"] || ""}`.trim()),
    dataRow("Email Address",       data["Email Address"]),
    dataRow("Phone Number",        data["Phone Number"]),
    dataRow("Date of Birth",       data["Date of Birth"]),
    dataRow("Nationality",         data["Nationality"]),
    dataRow("ID / SSN",            data["ID or SSN"]),
    spacerRow(),

    // Section 2
    sectionHeaderRow("2. Current Accommodation"),
    dataRow("Current Address",             data["Current Address"]),
    dataRow("Reason for Moving",           data["Reason for Moving"]),
    dataRow("Previously Rented from NSI",  data["Previously Rented from NSI"]),
    dataRow("Pets",                        data["Pets"]),
    spacerRow(),

    // Section 3
    sectionHeaderRow("3. Preferred Property"),
    dataRow("Preferred Location",  data["Preferred Location"]),
    dataRow("Move-in Date",        data["Move-in Date"]),
    dataRow("Monthly Budget",      data["Monthly Budget"]),
    dataRow("Bedrooms Required",   data["Bedrooms Required"]),
    spacerRow(),

    // Section 4
    sectionHeaderRow("4. Employment & Income"),
    dataRow("Employment Status",    data["Employment Status"]),
    dataRow("Employer / School",    data["Employer or School"]),
    dataRow("Annual Income",        data["Annual Income"] ? "$" + data["Annual Income"] : "—"),
    dataRow("Length of Employment", data["Length of Employment"]),
    spacerRow(),

    // Section 5
    sectionHeaderRow("5. Tenancy History"),
    dataRow("Ever Been Evicted",        data["Ever Been Evicted"]),
    dataRow("In Debt to Landlord",      data["In Debt to Landlord"]),
    dataRow("Previous Landlord / Ref",  data["Previous Landlord Reference"]),
    dataRow("Additional Notes",         data["Additional Notes"]),
    spacerRow(),

    // Section 6
    sectionHeaderRow("6. Application Fee & Payment"),
    dataRow("Fee Amount",     "$100 (Refundable)"),
    dataRow("Payment Method", data["Payment Method"]),
    spacerRow(),

    // Declaration
    sectionHeaderRow("Declaration"),
    new TableRow({
      children: [new TableCell({
        columnSpan: 2,
        borders: cellBorders,
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        margins: { top: 110, bottom: 110, left: 130, right: 130 },
        children: [new Paragraph({
          children: [new TextRun({
            text: data["Declaration Agreed"] === "Yes - I agree"
              ? "✓  I confirm the above declaration is true and I agree to the terms of this application."
              : "Declaration not confirmed.",
            size: 18,
            color: data["Declaration Agreed"] === "Yes - I agree" ? "2E7D32" : "C0392B",
            font: "Arial",
            bold: true,
          })]
        })]
      })]
    }),
  ];

  const submittedAt = new Date().toLocaleString();
  const refNumber   = "REF-" + Date.now();

  // ── Build document ───────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Arial", size: 22, color: DARK } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [

        // ── Title block ──
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: "NEST SEEKERS INTERNATIONAL", bold: true, size: 36, color: DARK, font: "Arial" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: "RENTAL APPLICATION — OFFICIAL COPY", size: 22, color: GOLD, font: "Arial" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 60 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 1 } },
          children: [
            new TextRun({ text: `Submitted: ${submittedAt}     `, size: 17, color: MUTED, font: "Arial" }),
            new TextRun({ text: refNumber,                         size: 17, color: MUTED, font: "Arial" }),
          ]
        }),

        // ── Spacer ──
        new Paragraph({ children: [], spacing: { before: 200, after: 0 } }),

        // ── Main data table ──
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3000, 6360],
          rows: tableRows,
        }),

        // ── Spacer ──
        new Paragraph({ children: [], spacing: { before: 300, after: 0 } }),

        // ── Footer note ──
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: `© ${new Date().getFullYear()} Nest Seekers International · nestseekersinternationals@gmail.com`,
            size: 16, color: MUTED, font: "Arial",
          })]
        }),
      ]
    }]
  });

  // ── Pack and trigger download ────────────────────────────────────────
  const buffer   = await Packer.toBlob(doc);
  const lastName = (data["Last Name"] || "Applicant").replace(/\s+/g, "_");
  const fileName = `NSI_Application_${lastName}_${Date.now()}.docx`;

  const url  = URL.createObjectURL(buffer);
  const link = document.createElement("a");
  link.href     = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
