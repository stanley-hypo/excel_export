import mammoth from 'mammoth';
import puppeteer from 'puppeteer';
import os from 'os';
import fs from 'fs/promises';
import path from 'path';
import { spawn } from 'child_process';

/**
 * Convert a .docx Buffer to a PDF Buffer.
 * The conversion is performed by:
 * 1) Converting DOCX -> HTML using mammoth (best-effort fidelity)
 * 2) Rendering HTML -> PDF using Puppeteer (Chromium)
 *
 * @param {Buffer} docxBuffer
 * @param {object} [options]
 * @param {object} [options.pdf] Options passed to page.pdf (e.g. { format: 'A4' })
 * @returns {Promise<Buffer>}
 */
export async function docxToPdf(docxBuffer, options = {}) {
	if (!Buffer.isBuffer(docxBuffer)) {
		throw new Error('docxToPdf: expected a Buffer input for the .docx file');
	}

	// Prefer higher-fidelity conversion using LibreOffice if available or explicitly requested
	const preferEngine = options.engine || process.env.WORD_TO_PDF_ENGINE || 'auto';
	if (preferEngine === 'libreoffice' || preferEngine === 'auto') {
		const sofficeCmd = await getSofficeCommand();
		const canUseLibre = Boolean(sofficeCmd);
		if (canUseLibre) {
			return await convertWithLibreOffice(docxBuffer, sofficeCmd);
		}
		if (preferEngine === 'libreoffice') {
			throw new Error('LibreOffice (soffice) not found on PATH. Install it or set engine to "mammoth".');
		}
	}

	// Fallback: DOCX -> HTML (Mammoth) -> PDF (Puppeteer)
	// Step 1: DOCX -> HTML
	const { value: html } = await mammoth.convertToHtml({ buffer: docxBuffer });

	// Basic wrapper HTML to ensure sane defaults when printing to PDF
	const pageHtml =
		'<!doctype html>' +
		'<html>' +
		'<head>' +
		'<meta charset="utf-8">' +
		'<meta name="viewport" content="width=device-width, initial-scale=1">' +
		'<style>' +
		'body { font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 32px; color: #111; }' +
		'table { border-collapse: collapse; }' +
		'td, th { border: 1px solid #ddd; padding: 6px 8px; }' +
		'img { max-width: 100%; }' +
		'</style>' +
		'</head>' +
		'<body>' +
		html +
		'</body>' +
		'</html>';

	// Step 2: HTML -> PDF using headless Chromium
	const browser = await puppeteer.launch({
		args: ['--no-sandbox', '--disable-setuid-sandbox']
	});
	try {
		const page = await browser.newPage();
		await page.setContent(pageHtml, { waitUntil: 'networkidle0' });
		const pdfBuffer = await page.pdf({
			format: 'A4',
			printBackground: true,
			...options.pdf
		});
		return pdfBuffer;
	} finally {
		await browser.close();
	}
}

async function getSofficeCommand() {
	// 1) Try 'soffice' on PATH
	const tryCmd = async (cmd) => {
		try {
			await new Promise((resolve, reject) => {
				const child = spawn(cmd, ['--version'], { stdio: 'ignore' });
				child.on('error', reject);
				child.on('exit', () => resolve());
			});
			return cmd;
		} catch {
			return null;
		}
	};
	const onPath = await tryCmd('soffice');
	if (onPath) return onPath;

	// 2) Try common macOS application path
	const macPath = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
	try {
		await fs.access(macPath);
		const ok = await tryCmd(macPath);
		if (ok) return ok;
	} catch {}

	return null;
}

async function convertWithLibreOffice(docxBuffer, sofficeCmd) {
	const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docx2pdf-'));
	const inputPath = path.join(tmpDir, 'input.docx');
	const outputPath = path.join(tmpDir, 'input.pdf');
	try {
		await fs.writeFile(inputPath, docxBuffer);
		// Convert using soffice
		await new Promise((resolve, reject) => {
			const args = [
				'--headless',
				'--nologo',
				'--nolockcheck',
				'--nodefault',
				'--nofirststartwizard',
				'--convert-to',
				'pdf:writer_pdf_Export',
				'--outdir',
				tmpDir,
				inputPath
			];
			const child = spawn(sofficeCmd, args, { stdio: 'ignore' });
			child.on('error', reject);
			child.on('exit', (code) => {
				if (code === 0) resolve();
				else reject(new Error(`soffice exited with code ${code}`));
			});
		});
		const pdf = await fs.readFile(outputPath);
		return pdf;
	} finally {
		// Best-effort cleanup
		try { await fs.rm(tmpDir, { recursive: true, force: true }); } catch {}
	}
}

export default {
	docxToPdf
};


