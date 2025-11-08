export const runtime = 'nodejs';

import { mergeExcelTemplate } from '../../../lib/index.js';

export async function POST(request) {
	try {
		const form = await request.formData();
		const jsonText = form.get('json') || '';
		let data = {};
		if (jsonText) {
			try {
				data = JSON.parse(jsonText);
			} catch {
				return new Response('Invalid JSON', { status: 400 });
			}
		}
		const file = form.get('template');
		if (!file || typeof file.arrayBuffer !== 'function') {
			return new Response('Missing template file', { status: 400 });
		}
		const buf = Buffer.from(await file.arrayBuffer());
		const output = await mergeExcelTemplate(buf, data);
		return new Response(output, {
			headers: {
				'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				'Content-Disposition': 'attachment; filename="export.xlsx"'
			}
		});
	} catch (err) {
		console.error(err);
		return new Response('Failed to export Excel', { status: 500 });
	}
}

