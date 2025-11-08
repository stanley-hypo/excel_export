import ExcelJS from 'exceljs';
import fs from 'fs/promises';
import path from 'path';

/**
 * Resolve a value from an object by a dotted/bracket path, e.g. "user.name", "items[0].price"
 */
function resolveByPath(source, rawPath) {
	if (source == null || !rawPath) return undefined;
	const tokens = [];
	let buf = '';
	let i = 0;
	while (i < rawPath.length) {
		const ch = rawPath[i];
		if (ch === '.') {
			if (buf) tokens.push(buf), buf = '';
			i += 1;
		} else if (ch === '[') {
			if (buf) tokens.push(buf), buf = '';
			i += 1;
			let bracket = '';
			while (i < rawPath.length && rawPath[i] !== ']') {
				bracket += rawPath[i];
				i += 1;
			}
			// skip ']'
			i += 1;
			const idx = bracket.replace(/^['"]|['"]$/g, '');
			tokens.push(idx);
		} else {
			buf += ch;
			i += 1;
		}
	}
	if (buf) tokens.push(buf);
	let cur = source;
	for (const key of tokens) {
		if (cur == null) return undefined;
		cur = cur[key];
	}
	return cur;
}

/**
 * Replace {{placeholders}} in a string using data.
 * - Supports dotted/bracket paths like {{user.name}} or {{items[0].price}}
 * - Trims spaces inside the braces
 */
function replacePlaceholders(text, data, { onMissing = 'empty', toString = defaultToString } = {}) {
	if (typeof text !== 'string' || text.indexOf('{{') === -1) return text;
	const pattern = /{{\s*([^}]+?)\s*}}/g;
	return text.replace(pattern, (_m, expr) => {
		const val = resolveByPath(data, String(expr).trim());
		if (val === undefined || val === null) {
			if (onMissing === 'keep') return _m;
			if (onMissing === 'empty') return '';
			if (typeof onMissing === 'function') return onMissing(expr);
			return '';
		}
		return toString(val);
	});
}

function defaultToString(value) {
	if (value instanceof Date) return value.toISOString();
	if (typeof value === 'object') return JSON.stringify(value);
	return String(value);
}

/**
 * Load a workbook from Buffer or file path.
 * @param {Buffer|string} template
 * @returns {Promise<ExcelJS.Workbook>}
 */
async function loadWorkbook(template) {
	const workbook = new ExcelJS.Workbook();
	if (Buffer.isBuffer(template)) {
		await workbook.xlsx.load(template);
		return workbook;
	}
	if (typeof template === 'string') {
		const abs = path.resolve(template);
		const buf = await fs.readFile(abs);
		await workbook.xlsx.load(buf);
		return workbook;
	}
	throw new Error('Unsupported template input. Provide a Buffer or file path.');
}

/**
 * Merge placeholders in an Excel template with JSON data.
 * - Replaces {{...}} in all string cells across all worksheets.
 * - Leaves non-string cells untouched (except when text was derived).
 *
 * @param {Buffer|string} template Buffer or path to .xlsx template
 * @param {object} data JSON object used for replacement
 * @param {object} [options]
 * @param {'empty'|'keep'|function} [options.onMissing='empty'] Behavior when a placeholder path is missing
 * @param {function} [options.valueToString] Custom stringify for inserted values
 * @returns {Promise<Buffer>} Resulting .xlsx as a Buffer
 */
export async function mergeExcelTemplate(template, data, options = {}) {
	const workbook = await loadWorkbook(template);
	const replaceOpts = {
		onMissing: options.onMissing ?? 'empty',
		toString: options.valueToString ?? defaultToString
	};

	for (const sheet of workbook.worksheets) {
		sheet.eachRow({ includeEmpty: false }, (row) => {
			row.eachCell({ includeEmpty: false }, (cell) => {
				// Prefer the displayed text for replacement to avoid missing rich text cases
				const currentText = typeof cell.text === 'string' ? cell.text : (typeof cell.value === 'string' ? cell.value : null);
				if (typeof currentText === 'string' && currentText.includes('{{')) {
					const newText = replacePlaceholders(currentText, data, replaceOpts);
					cell.value = newText;
				}
			});
		});
	}
	const out = await workbook.xlsx.writeBuffer();
	return Buffer.from(out);
}

export default {
	mergeExcelTemplate
};

