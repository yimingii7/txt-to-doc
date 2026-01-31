import { Document, Packer, Paragraph, TextRun } from "docx";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).send("Method Not Allowed");
    return;
  }

  try {
    let body = "";
    for await (const chunk of req) {
      body += chunk;
    }

    const lines = body.split(/\r?\n/);

    const doc = new Document({
      sections: [
        {
          children: lines.map(
            line =>
              new Paragraph({
                children: [new TextRun(line)],
              })
          ),
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
  } catch (err) {
    res.status(500).send("Conversion failed");
  }
}
