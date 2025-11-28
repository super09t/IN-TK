const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { spawn, spawnSync } = require('child_process');
const Tesseract = require('tesseract.js');

// Trang đích (ASP.NET WebForms)
const TARGET_URL = 'https://pus1.customs.gov.vn/BarcodeContainer/BarcodeContainer.aspx';

// Đầu vào mẫu (có thể override bằng tham số dòng lệnh)
const DEFAULT_INPUTS = {
	MaDoanhNghiep: '0314404243001',
	SoToKhai: '307766164220',
	MaHQ: '01dd',
	txtNgayToKhai: '15/09/2025',
	txtCaptcha: ''
};

let ocrOptions = {
	engine: 'tesseract',
	python: 'python',
	script: null,
	fallback: true
};
let pythonWorker = null;
let pythonWorkerReady = false;
let pythonWorkerQueue = [];
let pythonReadyPromise = null;
let pythonReader = null;

function now() { return new Date().toISOString(); }
function ms(s, e) { return `${(e - s)}ms`; }

function parseArgs() {
	const args = process.argv.slice(2);
	const kv = {};
	for (let i = 0; i < args.length; i++) {
		const a = args[i];
		if (a.startsWith('--')) {
			const key = a.replace(/^--/, '');
			const next = args[i + 1];
			if (next && !next.startsWith('--')) {
				kv[key] = next;
				i++;
			} else {
				kv[key] = 'true';
			}
		}
	}
	return kv;
}

function initLogger(args){
	try{
		const logPath = args && args.log ? path.resolve(args.log) : path.resolve(process.cwd(),'log.txt');
		// reset file
		try{ fs.writeFileSync(logPath, `# Log started at ${now()}\n`, {encoding:'utf-8'});}catch{}
		const origLog = console.log;
		const origErr = console.error;
		console.log = (...a)=>{
			try{ fs.appendFileSync(logPath, a.map(x=> typeof x==='string'?x:JSON.stringify(x)).join(' ')+'\n', {encoding:'utf-8'});}catch{}
			origLog(...a);
		};
		console.error = (...a)=>{
			try{ fs.appendFileSync(logPath, a.map(x=> typeof x==='string'?x:JSON.stringify(x)).join(' ')+'\n', {encoding:'utf-8'});}catch{}
			origErr(...a);
		};
		console.log(`[${now()}] Ghi log vào: ${logPath}`);
	}catch{}
}

// Danh sách file tạm để dọn sau khi chạy xong
const tempFiles = [];

function sanitizeCaptchaText(value) {
	return (value || '').replace(/\s+/g, '').replace(/[^A-Za-z0-9]/g, '');
}

function getOcrLabel(engine = ocrOptions.engine) {
	const value = (engine || '').toLowerCase();
	if (value === 'paddle') return 'PaddleOCR';
	if (value === 'easy') return 'EasyOCR';
	return 'Tesseract';
}

function setOcrOptions(opts = {}) {
	const next = { ...ocrOptions };
	if (opts.engine) {
		const norm = String(opts.engine).toLowerCase();
		if (['paddle', 'easy', 'tesseract'].includes(norm)) {
			next.engine = norm;
		}
	}
	if (opts.python) {
		next.python = opts.python;
	}
	if (opts.script) {
		next.script = opts.script;
	}
	if (typeof opts.fallback === 'boolean') {
		next.fallback = opts.fallback;
	}
	if (next.engine === 'tesseract') {
		next.script = null;
	}
	const engineChanged = next.engine !== ocrOptions.engine || next.script !== ocrOptions.script;
	ocrOptions = next;
	if (engineChanged) {
		closePythonWorker();
	}
}

function isFallbackEnabled() {
	return ocrOptions.fallback !== false;
}

function closePythonWorker() {
	if (pythonReader) {
		try { pythonReader.close(); } catch (_) {}
		pythonReader = null;
	}
	if (pythonWorker) {
		try {
			if (pythonWorker.stdin && !pythonWorker.killed) {
				pythonWorker.stdin.write('__EXIT__\n');
				pythonWorker.stdin.end();
			}
		} catch (_) {}
		try { pythonWorker.kill('SIGTERM'); } catch (_) {}
		pythonWorker = null;
	}
	pythonWorkerReady = false;
	pythonWorkerQueue.forEach((job) => {
		try {
			if (job && job.resolve) job.resolve('');
		} catch (_) {}
	});
	pythonWorkerQueue = [];
	pythonReadyPromise = null;
}

function ensurePythonWorker() {
	if (pythonWorker && pythonWorkerReady) {
		return Promise.resolve(true);
	}
	if (pythonReadyPromise) {
		return pythonReadyPromise;
	}

	pythonReadyPromise = new Promise((resolve) => {
		let settled = false;
		const finish = (value) => {
			if (!settled) {
				settled = true;
				resolve(value);
			}
		};

		const scriptPath = ocrOptions.script;
		if (!scriptPath || !fs.existsSync(scriptPath)) {
			console.error(`[${now()}] Script OCR không tồn tại: ${scriptPath}`);
			closePythonWorker();
			finish(false);
			return;
		}

		try {
			const spawnArgs = [scriptPath, '--server'];
			pythonWorker = spawn(ocrOptions.python, spawnArgs, { stdio: ['pipe', 'pipe', 'pipe'] });
		} catch (err) {
			console.error(`[${now()}] Không khởi động được ${getOcrLabel()}:`, err);
			closePythonWorker();
			finish(false);
			return;
		}

		if (!pythonWorker || !pythonWorker.stdin) {
			console.error(`[${now()}] ${getOcrLabel()} worker không hợp lệ.`);
			closePythonWorker();
			finish(false);
			return;
		}

		if (pythonWorker.stderr) {
			pythonWorker.stderr.on('data', (data) => {
				const txt = data.toString().trim();
				if (txt.length) {
					console.log(`[Python OCR stderr] ${txt}`);
				}
			});
		}

		pythonWorker.on('error', (err) => {
			console.error(`[${now()}] ${getOcrLabel()} worker error:`, err);
			closePythonWorker();
			finish(false);
		});

		pythonWorker.on('exit', (code) => {
			console.log(`[${now()}] ${getOcrLabel()} worker thoát với mã ${code}`);
			closePythonWorker();
			finish(false);
		});

		const rl = readline.createInterface({ input: pythonWorker.stdout });
		pythonReader = rl;

		rl.on('line', (line) => {
			const text = (line || '').trim();
			if (!pythonWorkerReady) {
				if (text === '__READY__') {
					pythonWorkerReady = true;
					console.log(`[${now()}] ${getOcrLabel()} worker sẵn sàng.`);
					finish(true);
					return;
				}
				if (text === '__ERROR__') {
					console.error(`[${now()}] ${getOcrLabel()} worker báo lỗi khi khởi tạo.`);
					closePythonWorker();
					finish(false);
					return;
				}
				pythonWorkerReady = true;
				finish(true);
				if (!text) {
					return;
				}
			}
			const job = pythonWorkerQueue.shift();
			if (job && job.resolve) {
				job.resolve(text);
			} else {
				console.log(`[${now()}] ${getOcrLabel()} trả về nhưng không có yêu cầu tương ứng: '${text}'`);
			}
		});
	}).finally(() => {
		pythonReadyPromise = null;
	});

	return pythonReadyPromise;
}

async function runPythonOcr(filePath) {
	const ok = await ensurePythonWorker().catch(() => false);
	if (!ok || !pythonWorker || !pythonWorker.stdin) {
		if (ocrOptions.engine !== 'tesseract' && isFallbackEnabled()) {
			console.log(`[${now()}] ${getOcrLabel()} không khả dụng, chuyển sang Tesseract.`);
			setOcrOptions({ engine: 'tesseract' });
		} else {
			console.log(`[${now()}] ${getOcrLabel()} không khả dụng và fallback bị tắt.`);
		}
		return '';
	}
	return await new Promise((resolve) => {
		const job = {
			resolve: (text) => resolve(sanitizeCaptchaText(text)),
			reject: () => resolve('')
		};
		pythonWorkerQueue.push(job);
		try {
			pythonWorker.stdin.write(`${filePath}\n`);
		} catch (err) {
			console.error(`[${now()}] Không ghi được vào ${getOcrLabel()} stdin:`, err);
			closePythonWorker();
			if (isFallbackEnabled()) {
				resolve('');
			} else {
				console.log(`[${now()}] Không thể ghi vào ${getOcrLabel()} stdin và fallback bị tắt.`);
				resolve('');
			}
		}
	});
}

async function runTesseract(filePath) {
	const start = Date.now();
	try {
		const { data } = await Tesseract.recognize(filePath, 'eng', {
			logger: () => {},
			tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789',
			tessedit_pageseg_mode: '8',
			tessedit_ocr_engine_mode: '3',
			tessedit_char_blacklist: '.,;:!?@#$%^&*()_+-=[]{}|\\:";\'<>?/~`',
			tessedit_resize_factor: '2'
		});
		const text = sanitizeCaptchaText(data && data.text ? data.text : '');
		console.log(`[${now()}] Tesseract OCR xong (${ms(start, Date.now())}): '${text}'`);
		return text;
	} catch (err) {
		console.error(`[${now()}] Lỗi Tesseract OCR:`, err);
		return '';
	}
}

function parseCsv(filePath) {
	const raw = fs.readFileSync(filePath, 'utf8');
	const lines = raw.split(/\r?\n/).filter((l) => l.trim().length > 0);
	if (lines.length === 0) return [];
	const headers = lines[0].split(',').map((h) => h.trim());
	const items = [];
	for (let i = 1; i < lines.length; i++) {
		const cols = lines[i].split(',');
		const obj = {};
		headers.forEach((h, idx) => { obj[h] = (cols[idx] || '').trim(); });
		items.push(obj);
	}
	return items;
}

function resolveOutName(item, pattern, fallback) {
	if (!pattern) return fallback;
	return pattern
		.replace('{MaDoanhNghiep}', item.MaDoanhNghiep || '')
		.replace('{SoToKhai}', item.SoToKhai || '')
		.replace('{MaHQ}', item.MaHQ || '')
		.replace('{NgayToKhai}', item.txtNgayToKhai || item.NgayToKhai || '');
}

async function fillAndSubmit(page, inputs) {
	console.log(`[${now()}] Điền form & submit: DN=${inputs.MaDoanhNghiep}, TK=${inputs.SoToKhai}, HQ=${inputs.MaHQ}, Ngay=${inputs.txtNgayToKhai}`);
	const s = Date.now();
	await page.waitForSelector('#MaDoanhNghiep', { state: 'visible', timeout: 10000 });
	await page.fill('#MaDoanhNghiep', inputs.MaDoanhNghiep);
	await page.fill('#SoToKhai', inputs.SoToKhai);
	await page.fill('#MaHQ', inputs.MaHQ);
	await page.fill('#txtNgayToKhai', inputs.txtNgayToKhai);
	if (inputs.txtCaptcha) {
		await page.fill('#txtCaptcha', inputs.txtCaptcha);
	}

	// Nhấn submit và hiển thị tiến trình
	const submit = page.locator('#Button1');
	await submit.scrollIntoViewIfNeeded();
	await Promise.all([
		page.waitForLoadState('networkidle', { timeout: 500 }).catch(() => {}),
		submit.click()
	]);
	console.log(`[${now()}] Submit xong (${ms(s, Date.now())})`);

	// Chờ ảnh loading ẩn đi nếu có
	const loading = page.locator('#div_Img_LoadingInformation');
	if (await loading.count()) {
		await loading.waitFor({ state: 'hidden', timeout: 1000 }).catch(() => {});
	}

	// Chờ vùng kết quả có dữ liệu (ví dụ `#lbl_TieuDeBangKe` hoặc bảng GridView)
	await page.waitForSelector('#div3, #GridView_IsNotContainer', { timeout: 1000 }).catch(() => {});
}

async function promptCaptcha(page, outDir = process.cwd()) {
	const img = page.locator('#Img1');
	await img.waitFor({ state: 'visible', timeout: 1000 });
	const filePath = path.join(outDir, `captcha_${Date.now()}.png`);
	await img.screenshot({ path: filePath });
	tempFiles.push(filePath);
	console.log(`[${now()}] Đã lưu ảnh captcha: ${filePath}`);

	const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
	const answer = await new Promise((resolve) => rl.question('Nhập mã captcha: ', (ans) => resolve(ans)));
	rl.close();
	return (answer || '').trim();
}

async function refreshCaptcha(page) {
	console.log(`[${now()}] Làm mới captcha`);
	await page.evaluate(() => {
		const img = document.getElementById('Img1');
		if (img) {
			const base = 'GenerateCaptcha.aspx?random=';
			img.src = base + Date.now();
		}
	});
	await page.waitForTimeout(500);
}

async function hasCaptchaError(page) {
	// Kiểm tra validator yêu cầu captcha hoặc vẫn còn ảnh captcha hiện (không có kết quả)
	try {
		const v5 = page.locator('#RequiredFieldValidator5');
		if (await v5.count()) {
			const vis = await v5.evaluate((el) => getComputedStyle(el).visibility !== 'hidden');
			if (vis) return true;
		}
		// Kiểm tra thông báo lỗi hiển thị trên trang (tiếng Việt)
		const err = await page.evaluate(() => {
			const txt = document.body.innerText || '';
			return /Chưa\s+nhập\s+đúng\s+mã\s+như\s+hình\s+trên/i.test(txt);
		});
		if (err) return true;
	} catch {}
	// Có thể sai captcha không hiện validator; ta dựa vào việc không có kết quả và vẫn còn form
	try {
		const hasResult = await page.locator('#div3').count();
		return !hasResult;
	} catch { return true; }
}

async function ocrCaptcha(page, outDir = process.cwd()) {
	const img = page.locator('#Img1');
	await img.waitFor({ state: 'visible', timeout: 1000 });
	const filePath = path.join(outDir, `captcha_ocr_${Date.now()}.png`);
	
	// Chụp ảnh captcha
	await img.screenshot({ path: filePath });
	tempFiles.push(filePath);
	console.log(`[${now()}] Đã chụp ảnh captcha: ${filePath}`);
	
	if (ocrOptions.engine !== 'tesseract') {
		const engineLabel = getOcrLabel();
		const pythonStart = Date.now();
		const pythonText = await runPythonOcr(filePath);
		console.log(`[${now()}] ${engineLabel} captcha xong (${ms(pythonStart, Date.now())}): '${pythonText}'`);
		if (pythonText) {
			return pythonText;
		}
		if (!isFallbackEnabled()) {
			console.log(`[${now()}] ${engineLabel} không nhận được kết quả hợp lệ và fallback bị tắt.`);
			return '';
		}
		console.log(`[${now()}] ${engineLabel} không nhận được kết quả hợp lệ, fallback sang Tesseract.`);
	}

	return await runTesseract(filePath);
}

function getCaptchaSrcKeySync() {
	try {
		const el = document.getElementById('Img1');
		return el ? el.getAttribute('src') || '' : '';
	} catch { return ''; }
}
async function getCaptchaSrcKey(page) {
	try {
		return await page.evaluate(() => {
			const el = document.getElementById('Img1');
			return el ? el.getAttribute('src') || '' : '';
		});
	} catch { return ''; }
}
async function waitForCaptchaChanged(page, prevKey, timeoutMs = 1000) {
	const start = Date.now();
	while (Date.now() - start < timeoutMs) {
		const key = await getCaptchaSrcKey(page);
		if (key && key !== prevKey) return true;
		await page.waitForTimeout(500);
	}
	return false;
}

async function tryAutoCaptcha(page, inputs, maxTries = 2) {
	for (let i = 1; i <= maxTries; i++) {
		console.log(`Thử OCR captcha lần ${i}/${maxTries}...`);
		
		// Thử OCR captcha hiện tại
		let code = await ocrCaptcha(page).catch(() => '');
		
		// Nếu OCR không đọc được, refresh và thử lại
		if (!code) {
			console.log('OCR không đọc được captcha, refresh ảnh captcha...');
			await refreshCaptcha(page);
			code = await ocrCaptcha(page).catch(() => '');
			if (!code) {
				console.log('Vẫn không đọc được captcha sau refresh, chờ captcha mới...');
				await page.waitForTimeout(500);
				continue;
			}
		}
		
		// Kiểm tra độ dài captcha hợp lý (thường 4-8 ký tự)
		if (code.length < 3 || code.length > 10) {
			console.log(`Captcha có độ dài không hợp lý (${code.length} ký tự): '${code}', thử lại...`);
			await page.waitForTimeout(500);
			continue;
		}
		
		console.log(`OCR đọc được: ${code}`);
		inputs.txtCaptcha = code;
		
		await fillAndSubmit(page, inputs);
		await forceShowResult(page);
		
		// Xác nhận cả 2 điều kiện: không có thông báo sai captcha và có nội dung kết quả
		const bad = await hasCaptchaError(page).catch(() => true);
		let hasContent = false;
		if (!bad) {
			hasContent = await waitForNonEmptyResult(page, 1000);
		}
		if (!bad && hasContent) {
			console.log('Captcha đúng!');
			return true;
		}
		console.log('Captcha sai hoặc chưa có dữ liệu, chờ captcha mới...');
		await page.waitForTimeout(500);
	}
	console.log('Hết lần thử OCR, auto-only sẽ bỏ qua nếu được bật');
	return false;
}

async function forceShowResult(page) {
	// Buộc hiển thị các khối kết quả nếu bị ẩn do CSS
	await page.evaluate(() => {
		['KetQuaTraCuu', 'div3', 'div4', 'Panel_IsNotContainer'].forEach((id) => {
			const el = document.getElementById(id);
			if (el && el.style && el.style.display === 'none') el.style.display = '';
		});
	});
}

async function saveResultPdf(page, filePath) {
	// Kết quả nằm trong `#div3` + `#div4` + ghi chú (divGhiChu1/2 nếu có)
	await page.waitForSelector('#div3', { timeout: 3000 }).catch(() => {});
	await page.evaluate(() => {
		const blocks = [];
		const add = (id) => { const el = document.getElementById(id); if (el) blocks.push(el.innerHTML); };
		add('div3');
		add('div4');
		add('divGhiChu1');
		add('divGhiChu2');
		const html = `<link rel=\"stylesheet\" type=\"text/css\" href=\"css/Layout.css\" />` + blocks.join('');
		document.body.innerHTML = html;
	});
	await page.emulateMedia({ media: 'print' });
	const s = Date.now();
	await page.pdf({ path: filePath, printBackground: true, preferCSSPageSize: true });
	console.log(`[${now()}] Đã lưu PDF: ${filePath} (${ms(s, Date.now())})`);
}

async function getResultInnerText(page) {
	try {
		return await page.evaluate(() => {
			const el = document.querySelector('#div3');
			return el ? el.innerText.trim() : '';
		});
	} catch { return ''; }
}

async function waitForNonEmptyResult(page, timeoutMs = 5000) {
	const start = Date.now();
	while (Date.now() - start < timeoutMs) {
		const txt = await getResultInnerText(page);
		if (txt && txt.length > 50) return true;
		await page.waitForTimeout(500);
	}
	return false;
}

async function saveDebugArtifacts(page, outDir, soTK) {
	try {
		const png = path.join(outDir, `MV_DEBUG_${soTK}.png`);
		await page.screenshot({ path: png, fullPage: true });
		const html = path.join(outDir, `MV_DEBUG_${soTK}.html`);
		const content = await page.content();
		fs.writeFileSync(html, content, { encoding: 'utf-8' });
		console.log(`[${now()}] Đã lưu debug: ${png}, ${html}`);
	} catch (e) { console.error('Lỗi lưu debug:', e); }
}

async function run() {
	const args = parseArgs();
	initLogger(args);
	const engineArg = args['ocr-engine'];
	const disableFallback = (args['no-fallback'] || '').toString().toLowerCase() === 'true';
	setOcrOptions({
		engine: engineArg ? engineArg.toLowerCase() : ocrOptions.engine,
		python: args['ocr-python'] || ocrOptions.python,
		script: args['ocr-script'] ? path.resolve(args['ocr-script']) : ocrOptions.script,
		fallback: !disableFallback
	});
	console.log(`[${now()}] Sử dụng OCR engine: ${getOcrLabel()} (fallback=${isFallbackEnabled()})`);
	const inputs = { ...DEFAULT_INPUTS };
	if (args.MaDoanhNghiep) inputs.MaDoanhNghiep = args.MaDoanhNghiep;
	if (args.SoToKhai) inputs.SoToKhai = args.SoToKhai;
	if (args.MaHQ) inputs.MaHQ = args.MaHQ;
	if (args.NgayToKhai) inputs.txtNgayToKhai = args.NgayToKhai; // alias
	if (args.txtNgayToKhai) inputs.txtNgayToKhai = args.txtNgayToKhai;
	if (args.captcha) inputs.txtCaptcha = args.captcha;
	const autoOnly = (args['auto-only'] || '').toString().toLowerCase() === 'true';
	const ocrTries = autoOnly ? 4 : (args['ocr-tries'] ? parseInt(args['ocr-tries'],10) || 2 : 2);

	const headlessArg = (args['headless'] || '').toString().toLowerCase();
	const headless = headlessArg === 'false' ? false : true;
	console.log(`[${now()}] Khởi chạy Chromium headless=${headless}`);
	const browser = await chromium.launch({ headless }); // Cho phép tắt headless để debug
	// Hỗ trợ tải trạng thái (cookies/localStorage) để tránh nhập captcha lại
	const contextOptions = {};
	if (args['load-state']) {
		contextOptions.storageState = path.resolve(args['load-state']);
		console.log(`[${now()}] Đã tải trạng thái phiên: ${contextOptions.storageState}`);
	}
	const context = await browser.newContext(contextOptions);
	const page = await context.newPage();

	try {
		// Batch mode trong 1 phiên
		if (args.batch) {
			const list = parseCsv(path.resolve(args.batch));
			if (!list.length) throw new Error('Batch rỗng hoặc CSV không hợp lệ');
			await page.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
			for (let idx = 0; idx < list.length; idx++) {
				const row = list[idx];
				const inItem = {
					MaDoanhNghiep: row.MaDoanhNghiep || inputs.MaDoanhNghiep,
					SoToKhai: row.SoToKhai || inputs.SoToKhai,
					MaHQ: row.MaHQ || inputs.MaHQ,
					txtNgayToKhai: row.NgayToKhai || row.txtNgayToKhai || inputs.txtNgayToKhai,
					txtCaptcha: ''
				};

				console.log(`\n[${now()}] ==== Bắt đầu hàng ${idx + 1}/${list.length} - TK=${inItem.SoToKhai} ==== `);
				const t0 = Date.now();

				// Không reload trước; đảm bảo dùng captcha mới nhất
				await page.waitForTimeout(500);

				let solved = false;
				const disableOcr = args['no-ocr'] === 'true';
				if (!disableOcr) {
					solved = await tryAutoCaptcha(page, inItem, ocrTries);
				}
				if (!solved) {
					if (autoOnly || args['no-prompt'] === 'true') {
						console.log(`Bỏ qua hàng ${idx + 1} do không giải được captcha (auto-only/no-prompt).`);
						continue;
					}
					for (let attempt = 1; attempt <= 3; attempt++) {
						console.log(`[${now()}] Hỏi người dùng nhập captcha (attempt ${attempt}/3)`);
						const cap = await promptCaptcha(page);
						inItem.txtCaptcha = cap;
						await fillAndSubmit(page, inItem);
						await forceShowResult(page);
						const bad = await hasCaptchaError(page).catch(() => false);
						if (!bad) { solved = true; break; }
						await page.waitForTimeout(2000);
					}
				}
				if (!solved) {
					console.log(`Không thể xử lý hàng ${idx + 1}. Thời gian: ${ms(t0, Date.now())}`);
					continue;
				}

				// Đợi nội dung thực sự có dữ liệu, nếu rỗng thì thử submit lại 1 lần
				let hasContent = await waitForNonEmptyResult(page, 5000);
				if (!hasContent) {
					console.log(`[${now()}] Kết quả rỗng, thử submit lại một lần...`);
					await fillAndSubmit(page, inItem);
					await forceShowResult(page);
					hasContent = await waitForNonEmptyResult(page, 5000);
				}

				const out = resolveOutName(
					inItem,
					args['out-pattern'],
					args.out || `BarcodeContainer_${inItem.SoToKhai || idx + 1}.pdf`
				);
				if (!hasContent) {
					console.log(`[${now()}] Vẫn không có dữ liệu sau khi retry. Lưu debug & bỏ qua hàng này.`);
					await saveDebugArtifacts(page, path.dirname(path.resolve(out)), inItem.SoToKhai || `${idx+1}`);
					// Không reload; chỉ refresh captcha để đồng bộ tình trạng mới
					await page.waitForTimeout(2000);
					continue;
				}
				await saveResultPdf(page, out);
				console.log(`(${idx + 1}/${list.length}) Đã xuất PDF: ${out}`);
				// Chờ thêm trước khi chuyển sang dòng tiếp theo
				await page.waitForTimeout(100);
				console.log(`Hoàn tất hàng ${idx + 1} trong ${ms(t0, Date.now())}`);
				// Không reload; refresh captcha để đảm bảo ảnh mới cho lần sau
				await page.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
				await page.waitForTimeout(600);
			}

			// Xoá file captcha tạm nếu có
			for (const f of tempFiles) {
				try { fs.existsSync(f) && fs.unlinkSync(f); } catch (_) {}
			}
			console.log('Đã dọn dẹp file captcha tạm');
			// Lưu trạng thái phiên nếu được yêu cầu
			if (args['save-state']) {
				const savePath = path.resolve(args['save-state']);
				await context.storageState({ path: savePath });
				console.log(`Đã lưu trạng thái phiên: ${savePath}`);
			}
			return;
		}

		await page.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });

		// Nếu chưa có captcha được truyền, thử OCR trước, sau đó mới hỏi người dùng
		if (!inputs.txtCaptcha) {
			const disableOcr = args['no-ocr'] === 'true';
			let solved = false;
			if (!disableOcr) {
				console.log('Đang thử nhận dạng captcha tự động...');
				solved = await tryAutoCaptcha(page, inputs, ocrTries);
			}
			if (!solved) {
				if (autoOnly || args['no-prompt'] === 'true') {
					console.log('Auto-only/no-prompt: bỏ qua nhập captcha thủ công (single).');
				} else {
					for (let attempt = 1; attempt <= 3; attempt++) {
						console.log(`[${now()}] Hỏi người dùng nhập captcha (attempt ${attempt}/3)`);
						const cap = await promptCaptcha(page);
						inputs.txtCaptcha = cap;
						await fillAndSubmit(page, inputs);
						await forceShowResult(page);
						const captchaBad = await hasCaptchaError(page).catch(() => false);
						if (!captchaBad) break;
						console.log(`Captcha có thể sai, thử lại (${attempt}/3)...`);
						inputs.txtCaptcha = '';
						await page.waitForTimeout(300);
					}
				}
			}
		} else {
			await fillAndSubmit(page, inputs);
		}
		await forceShowResult(page);

		// Đợi nội dung có dữ liệu, nếu rỗng thì retry 1 lần trước khi lưu
		let ok = await waitForNonEmptyResult(page, 5000);
		if (!ok) {
			console.log(`[${now()}] Kết quả rỗng (single), thử submit lại một lần...`);
			await fillAndSubmit(page, inputs);
			await forceShowResult(page);
			ok = await waitForNonEmptyResult(page, 5000);
		}

		const out = args.out || `BarcodeContainer_${inputs.SoToKhai}.pdf`;
		if (!ok) {
			console.log(`[${now()}] Vẫn rỗng sau retry (single). Lưu debug.`);
			await saveDebugArtifacts(page, path.dirname(path.resolve(out)), inputs.SoToKhai || 'single');
		} else {
			await saveResultPdf(page, out);
			console.log(`Đã xuất PDF: ${out}`);
		}

		// Xoá file captcha tạm nếu có
		for (const f of tempFiles) {
			try { fs.existsSync(f) && fs.unlinkSync(f); } catch (_) {}
		}
		console.log('Đã dọn dẹp file captcha tạm (single mode)');

		// Lưu trạng thái phiên nếu được yêu cầu
		if (args['save-state']) {
			const savePath = path.resolve(args['save-state']);
			await context.storageState({ path: savePath });
			console.log(`Đã lưu trạng thái phiên: ${savePath}`);
		}
	} catch (err) {
		console.error('Lỗi:', err);
		process.exitCode = 1;
	} finally {
		// Trường hợp lỗi, vẫn cố dọn file tạm
		for (const f of tempFiles) {
			try { fs.existsSync(f) && fs.unlinkSync(f); } catch (_) {}
		}
		console.log('Đã dọn dẹp file captcha tạm (trường hợp lỗi)');
		closePythonWorker();
		await browser.close();
	}
}

if (require.main === module) {
	run();
}
