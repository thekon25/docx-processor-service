// api/process-docx.js
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const formidable = require('formidable');
const fs = require('fs');
const path = require('path');

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

  try {
    const form = new formidable.IncomingForm();
    form.parse(req, async (err, fields, files) => {
      if (err) {
        console.error('Form parse error:', err);
        return res.status(500).json({ error: 'Failed to parse form data', message: err.message });
      }

      try {
        // Проверим наличие файла и данных
        if (!files.template) {
          return res.status(400).json({ error: 'Missing template file' });
        }

        if (!fields.data) {
          return res.status(400).json({ error: 'Missing data field' });
        }

        const templateFile = files.template[0] || files.template;
        const templatePath = templateFile.filepath;
        
        // Читаем файл
        const templateBuffer = fs.readFileSync(templatePath);
        
        // Парсим JSON данные
        let data;
        const dataField = Array.isArray(fields.data) ? fields.data[0] : fields.data;
        try {
          data = typeof dataField === 'string' ? JSON.parse(dataField) : dataField;
        } catch (parseErr) {
          console.error('JSON parse error:', parseErr);
          return res.status(400).json({ error: 'Invalid JSON in data field', message: parseErr.message });
        }
        
        console.log('Template buffer size:', templateBuffer.length);
        console.log('Data keys:', Object.keys(data));
        
        // Обработка документа
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
        
        console.log('Processed buffer size:', processedBuffer.length);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="processed.docx"');
        res.send(processedBuffer);
        
      } catch (error) {
        console.error('Processing error:', error);
        console.error('Error stack:', error.stack);
        res.status(500).json({ 
          error: 'Failed to process document', 
          message: error.message,
          stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
      }
    });
  } catch (error) {
    console.error('Main error:', error);
    res.status(500).json({ error: error.message });
  }
};
