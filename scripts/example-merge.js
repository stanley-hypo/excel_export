import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { mergeExcelTemplate } from '../lib/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function main() {
	const templatePath = process.argv[2] || path.join(__dirname, 'sample-template.xlsx');
	const outputPath = process.argv[3] || path.join(__dirname, 'output.xlsx');
	const dataJson = process.argv[4] || '{"name":"Stanley","order":{"total":199.99}}';
	let data;
	try {
		data = JSON.parse(dataJson);
	} catch {
		console.error('Invalid JSON passed as third argument.');
		process.exit(1);
	}
	const template = await fs.readFile(templatePath);
	const result = await mergeExcelTemplate(template, data);
	await fs.writeFile(outputPath, result);
	console.log(`Wrote ${outputPath}`);
}

main().catch((err) => {
	console.error(err);
	process.exit(1);
});

