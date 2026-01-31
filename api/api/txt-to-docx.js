import formidable from "formidable";
import fs from "fs";
import { Document, Packer, Paragraph } from "docx";

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  const form = formidable();

  form.parse(req, async (err, fields, files) => {
    if (err) {
      res.status(500).json({ error: "File parse failed" });
      return;
    }

    const file = files.file;
    if (!file) {
      res.status(400).json({ error: "No file uploaded" });
      return;
    }

    const content = fs.readFileSync(file.filepath, "utf8");

    const paragraphs = content
      .split(/\r?\n/)
      .map(line => new Paragraph(line));

    const doc = new Document({
      sections: [
        {
          children: paragraphs,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="converted.docx"'
    );

    res.send(buffer);
  });
}
