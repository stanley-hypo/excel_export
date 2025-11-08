import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { mergeExcelTemplate } from './lib/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.get('/', (_req, res) => {
	res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/export', upload.single('template'), async (req, res) => {
	try {
		const jsonText = req.body?.json ?? '';
		let data;
		try {
			data = jsonText ? JSON.parse(jsonText) : {};
		} catch (err) {
			res.status(400).send('Invalid JSON');
			return;
		}
		if (!req.file) {
			res.status(400).send('Missing template file');
			return;
		}
		const output = await mergeExcelTemplate(req.file.buffer, data);
		res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		res.setHeader('Content-Disposition', 'attachment; filename="export.xlsx"');
		res.send(output);
	} catch (err) {
		// eslint-disable-next-line no-console
		console.error(err);
		res.status(500).send('Failed to export Excel');
	}
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
	// eslint-disable-next-line no-console
	console.log(`Server listening on http://localhost:${PORT}`);
});

