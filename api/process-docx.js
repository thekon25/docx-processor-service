// api/process-docx.js
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const formidable = require('formidable');
const fs = require('fs');

module.exports = async (req, res) => {
  // CORS заголовки для n8n
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const form = formidable({ multiples: false });
  
  form.parse(req, async (err, fields, files) => {
    if (err) {
      return res.status(500).json({ error: 'Failed to parse form data' });
    }

    try {
      const templateFile = files.template;
      const templateBuffer = fs.readFileSync(templateFile.filepath);
      
      const data = JSON.parse(fields.data);
      
      const zip = new PizZip(templateBuffer);
      
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        nullGetter: () => ''
      });
      
      doc.setData(data);
      doc.render();
      
      const processedBuffer = doc.getZip().generate({
        type: 'nodebuffer',
        compression: 'DEFLATE'
      });
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', 'attachment; filename="processed.docx"');
      res.send(processedBuffer);
      
    } catch (error) {
      console.error('Processing error:', error);
      res.status(500).json({ error: 'Failed to process document', details: error.message });
    }
  });
};
