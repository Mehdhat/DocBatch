import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  HeadingLevel,
} from "docx";

const PORT = 3000;

// Lazy initialization of Gemini client to prevent startup crashes when API key is missing
let aiClient: GoogleGenAI | null = null;

function getGeminiClient(): GoogleGenAI {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      throw new Error(
        "GEMINI_API_KEY environment variable is required. Please set it in the Secrets panel in Google AI Studio."
      );
    }
    aiClient = new GoogleGenAI({
      apiKey: key,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }
  return aiClient;
}

// Convert extracted JSON structure into a styled Word Document using docx
function buildWordDocument(data: any): Document {
  const children: any[] = [];

  // 1. Document Title
  if (data.title) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.TITLE,
        children: [
          new TextRun({
            text: data.title,
            bold: true,
            size: 36, // 18pt
            font: "Arial",
            color: "0f172a", // Slate 900
          }),
        ],
        spacing: {
          before: 0,
          after: 300,
        },
      })
    );
  }

  // 2. Document Elements
  if (Array.isArray(data.elements)) {
    for (const el of data.elements) {
      if (!el.type) continue;

      switch (el.type) {
        case "title":
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: el.text || "",
                  bold: true,
                  size: 32, // 16pt
                  font: "Arial",
                  color: "1e3a8a", // Dark Blue
                }),
              ],
              spacing: { before: 200, after: 200 },
            })
          );
          break;

        case "heading1":
          children.push(
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [
                new TextRun({
                  text: el.text || "",
                  bold: true,
                  size: 28, // 14pt
                  font: "Arial",
                  color: "1e3a8a", // Dark Blue
                }),
              ],
              spacing: { before: 300, after: 120 },
            })
          );
          break;

        case "heading2":
          children.push(
            new Paragraph({
              heading: HeadingLevel.HEADING_2,
              children: [
                new TextRun({
                  text: el.text || "",
                  bold: true,
                  size: 24, // 12pt
                  font: "Arial",
                  color: "334155", // Slate 700
                }),
              ],
              spacing: { before: 240, after: 100 },
            })
          );
          break;

        case "paragraph":
          if (el.text) {
            children.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: el.text,
                    size: 22, // 11pt
                    font: "Calibri",
                    color: "334155",
                  }),
                ],
                spacing: { after: 160, line: 276 }, // 1.15 line spacing
              })
            );
          }
          break;

        case "bullet_list":
          if (Array.isArray(el.listItems)) {
            for (const item of el.listItems) {
              if (!item) continue;
              children.push(
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "•  ",
                      bold: true,
                      size: 22,
                      font: "Calibri",
                      color: "2563eb", // Vibrant Blue bullet
                    }),
                    new TextRun({
                      text: item,
                      size: 22,
                      font: "Calibri",
                      color: "334155",
                    }),
                  ],
                  spacing: { after: 80 },
                  indent: { left: 360 },
                })
              );
            }
          }
          break;

        case "numbered_list":
          if (Array.isArray(el.listItems)) {
            el.listItems.forEach((item, index) => {
              if (!item) return;
              children.push(
                new Paragraph({
                  children: [
                    new TextRun({
                      text: `${index + 1}.  `,
                      bold: true,
                      size: 22,
                      font: "Calibri",
                      color: "2563eb",
                    }),
                    new TextRun({
                      text: item,
                      size: 22,
                      font: "Calibri",
                      color: "334155",
                    }),
                  ],
                  spacing: { after: 80 },
                  indent: { left: 360 },
                })
              );
            });
          }
          break;

        case "callout":
          if (el.text) {
            children.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: el.text,
                    italics: true,
                    size: 22,
                    font: "Calibri",
                    color: "1e293b",
                  }),
                ],
                spacing: { before: 120, after: 120 },
                indent: { left: 540, right: 540 },
              })
            );
          }
          break;

        case "table":
          if (el.tableData && Array.isArray(el.tableData.rows)) {
            const hasHeaders = Array.isArray(el.tableData.headers) && el.tableData.headers.length > 0;
            const rows: TableRow[] = [];

            const borderStyle = {
              style: BorderStyle.SINGLE,
              size: 4,
              color: "cbd5e1", // Slate 300
            };

            const cellBorders = {
              top: borderStyle,
              bottom: borderStyle,
              left: borderStyle,
              right: borderStyle,
            };

            // Header row
            if (hasHeaders) {
              const headerCells = el.tableData.headers.map((hdr: string) => {
                return new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: hdr || "",
                          bold: true,
                          size: 20, // 10pt
                          font: "Arial",
                          color: "ffffff",
                        }),
                      ],
                      alignment: AlignmentType.LEFT,
                    }),
                  ],
                  shading: {
                    fill: "1e3a8a", // Navy Blue
                  },
                  margins: {
                    top: 120,
                    bottom: 120,
                    left: 150,
                    right: 150,
                  },
                  borders: cellBorders,
                });
              });
              rows.push(new TableRow({ children: headerCells }));
            }

            // Data rows
            el.tableData.rows.forEach((row: string[], rowIndex: number) => {
              if (!Array.isArray(row)) return;
              const cells = row.map((cellText: string) => {
                return new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: cellText || "",
                          size: 20,
                          font: "Calibri",
                          color: "334155",
                        }),
                      ],
                    }),
                  ],
                  shading: {
                    fill: rowIndex % 2 === 0 ? "f8fafc" : "ffffff", // Clean zebra striping
                  },
                  margins: {
                    top: 100,
                    bottom: 100,
                    left: 150,
                    right: 150,
                  },
                  borders: cellBorders,
                });
              });
              rows.push(new TableRow({ children: cells }));
            });

            const table = new Table({
              rows: rows,
              width: {
                size: 100,
                type: WidthType.PERCENTAGE,
              },
            });
            children.push(table);
          }
          break;

        default:
          break;
      }
    }
  }

  // Fallback if no content was extracted
  if (children.length === 0) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "Document empty. No text content could be extracted from the source file.",
            italics: true,
            color: "64748b",
          }),
        ],
      })
    );
  }

  return new Document({
    sections: [
      {
        properties: {},
        children: children,
      },
    ],
  });
}

async function startServer() {
  const app = express();

  // Increase payload limit to handle large images and PDF base64 bulk uploads
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ limit: "50mb", extended: true }));

  // API route for status/health check
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // API route to perform conversion using Gemini and docx
  app.post("/api/convert", async (req, res) => {
    try {
      const { fileBase64, fileName, mimeType } = req.body;

      if (!fileBase64 || !fileName || !mimeType) {
        return res.status(400).json({
          error: "Missing required parameters: fileBase64, fileName, and mimeType are required.",
        });
      }

      // Check API Key
      const ai = getGeminiClient();

      // Clean the base64 prefix if present
      const cleanBase64 = fileBase64.includes(",")
        ? fileBase64.split(",")[1]
        : fileBase64;

      // Call Gemini 3.5 Flash for Multimodal OCR and Structured Layout Extraction
      const filePart = {
        inlineData: {
          data: cleanBase64,
          mimeType: mimeType,
        },
      };

      const promptPart = {
        text: `Perform a high-fidelity visual and content extraction of the provided document.
Your goal is to recreate this document as an editable Microsoft Word document.
Extract all text, headings, paragraphs, lists, and tables exactly as they appear, preserving their logical reading order.
Analyze the layout and format the response according to the JSON schema.
Do not omit any text. For scanned documents, perform complete OCR.
For tabular grids or lists, capture them as structured 'table' or 'bullet_list'/'numbered_list' objects.
If a block is emphasized or represents notes, make it a 'callout'.`,
      };

      // Define standard JSON schema to get structured data back from Gemini
      const responseSchema = {
        type: Type.OBJECT,
        properties: {
          title: {
            type: Type.STRING,
            description: "A suitable title for the document. If a clear title is present, use it. Otherwise, generate a concise descriptive title.",
          },
          elements: {
            type: Type.ARRAY,
            description: "The blocks of text in logical reading order.",
            items: {
              type: Type.OBJECT,
              properties: {
                type: {
                  type: Type.STRING,
                  enum: [
                    "title",
                    "heading1",
                    "heading2",
                    "paragraph",
                    "bullet_list",
                    "numbered_list",
                    "table",
                    "callout",
                  ],
                  description: "The type of content block.",
                },
                text: {
                  type: Type.STRING,
                  description: "The text content for 'title', 'heading1', 'heading2', 'paragraph', or 'callout'. Leave empty for tables or lists.",
                },
                listItems: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING },
                  description: "The string items for lists. Leave empty for non-list types.",
                },
                tableData: {
                  type: Type.OBJECT,
                  properties: {
                    headers: {
                      type: Type.ARRAY,
                      items: { type: Type.STRING },
                      description: "The table column headers.",
                    },
                    rows: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.ARRAY,
                        items: { type: Type.STRING },
                      },
                      description: "2D array of rows containing cells.",
                    },
                  },
                  description: "The table data structure. Leave empty for non-table types.",
                },
              },
              required: ["type"],
            },
          },
        },
        required: ["title", "elements"],
      };

      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: { parts: [filePart, promptPart] },
        config: {
          responseMimeType: "application/json",
          responseSchema: responseSchema,
          systemInstruction: "You are an expert document parser. Convert images and PDFs into a high-fidelity logical structure that can be reconstructed into an elegant Word document.",
        },
      });

      const responseText = response.text;
      if (!responseText) {
        throw new Error("Empty response from Gemini model.");
      }

      // Parse JSON structure
      const docStructure = JSON.parse(responseText);

      // Build docx Document
      const doc = buildWordDocument(docStructure);

      // Compile docx document into base64
      const buffer = await Packer.toBuffer(doc);
      const outputBase64 = buffer.toString("base64");

      // Respond with base64 data and suggested filename
      const originalBaseName = fileName.substring(0, fileName.lastIndexOf(".")) || fileName;
      const targetFileName = `${originalBaseName}.docx`;

      res.json({
        success: true,
        base64: outputBase64,
        fileName: targetFileName,
        docStructure: docStructure, // Also return the extracted structure for UI previewing!
      });
    } catch (error: any) {
      console.error("Conversion Error:", error);
      res.status(500).json({
        error: error.message || "An unexpected error occurred during document conversion.",
      });
    }
  });

  // Vite development integration or static files serving
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error("Failed to start server:", err);
});
