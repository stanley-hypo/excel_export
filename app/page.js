'use client';

import { useCallback } from 'react';

export default function Page() {
	const onSubmit = useCallback((e) => {
		const form = e.currentTarget;
		const text = form.json?.value?.trim();
		if (text) {
			try { JSON.parse(text); } catch {
				e.preventDefault();
				alert('JSON is invalid');
			}
		}
	}, []);

	return (
		<main style={{ margin: 32, fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif' }}>
			<h1 style={{ margin: '0 0 16px', fontSize: 22 }}>Excel Merge ({'{{placeholder}}'})</h1>
			<form method="post" action="/api/export" encType="multipart/form-data" onSubmit={onSubmit} style={{ display: 'grid', gap: 12, maxWidth: 720 }}>
				<div style={{ display: 'grid', gap: 8 }}>
					<label htmlFor="json" style={{ fontWeight: 600 }}>JSON Data</label>
					<textarea id="json" name="json" placeholder='{"name":"Stanley","total":123}' style={{ width: '100%', minHeight: 200, fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Consolas, \"Liberation Mono\", monospace' }}>
{`{
  "name": "Stanley",
  "email": "stanley@example.com",
  "order": {
    "id": 1001,
    "total": 199.99
  }
}`}
					</textarea>
					<div style={{ color: '#666', fontSize: 12, opacity: .9 }}>
						Placeholders support dot/bracket paths, e.g. <code>{'{{name}}'}</code>, <code>{'{{order.id}}'}</code>
					</div>
				</div>
				<div style={{ display: 'grid', gap: 8 }}>
					<label htmlFor="template" style={{ fontWeight: 600 }}>Excel Template (.xlsx)</label>
					<input id="template" name="template" type="file" accept=".xlsx" required />
					<div style={{ color: '#666' }}>Put placeholders directly inside cells, like <code>{'{{name}}'}</code> or <code>{'{{order.total}}'}</code></div>
				</div>
				<div>
					<button type="submit" style={{ padding: '10px 14px', fontWeight: 600, borderRadius: 8, border: '1px solid #999', cursor: 'pointer' }}>
						Export
					</button>
				</div>
			</form>
		</main>
	);
}

