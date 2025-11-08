export const runtime = 'nodejs';

import { docxToPdf } from '../../../lib/wordToPdf.js';

export async function POST(request) {
	try {
		const { searchParams } = new URL(request.url);
		const engine = searchParams.get('engine') || undefined; // 'libreoffice' | 'mammoth'
		const form = await request.formData();
		const file = form.get('docx');
		if (!file || typeof file.arrayBuffer !== 'function') {
			return new Response('Missing DOCX file', { status: 400 });
		}
		const buf = Buffer.from(await file.arrayBuffer());
		const pdf = await docxToPdf(buf, { engine });
		return new Response(pdf, {
			headers: {
				'Content-Type': 'application/pdf',
				'Content-Disposition': 'attachment; filename="export.pdf"'
			}
		});
	} catch (err) {
		console.error(err);
		return new Response('Failed to convert Word to PDF', { status: 500 });
	}
}


