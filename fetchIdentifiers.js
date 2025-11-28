const { chromium } = require('playwright');
const readline = require('readline');
const fs = require('fs');
const path = require('path');
const Tesseract = require('tesseract.js');

const LOGIN_URL = 'https://pus.customs.gov.vn/faces/Login';
const CAPTCHA_DIR = path.join(process.cwd(), 'output', 'captcha_ids');
const SELECTORS = {
  username: '#pt1\\:it1\\:\\:content',
  password: '#pt1\\:it2\\:\\:content',
  loginType: '#pt1\\:rsoLoginType\\:_1',
  captcha: '#pt1\\:it42\\:\\:content',
  captchaImg: '#pt1\\:i5',
  captchaRefresh: '#pt1\\:cmlRefresh',
  loginBtn: '#pt1\\:cbLogin',
  moduleLink: '#pt1\\:dc7\\:dinhDanhhangHoa > div > table > tbody > tr > td.x18i > a',
  newBtn: '#pt1\\:b1',
  flowSelect: '#pt1\\:soc2\\:\\:content',
  dnSelect: '#pt1\\:soc1\\:\\:content',
  mstInput: '#pt1\\:it5\\:\\:content',
  confirmBtn: '#pt1\\:cb3',
  codeField: '#pt1\\:it11\\:\\:content',
  closeBtn: '#pt1\\:b4',
};

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stderr,
  terminal: false,
  crlfDelay: Infinity,
});

function parseArgs() {
  const args = process.argv.slice(2);
  const map = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg.startsWith('--')) {
      const key = arg.slice(2);
      const next = args[i + 1];
      if (next && !next.startsWith('--')) {
        map[key] = next;
        i++;
      } else {
        map[key] = 'true';
      }
    }
  }
  return map;
}

function log(...messages) {
  console.error('[fetchIdentifiers]', ...messages);
}

function sendResult(items) {
  process.stdout.write('RESULT ' + JSON.stringify({ success: true, items }) + '\n');
}

function sendError(message) {
  process.stdout.write('ERROR ' + JSON.stringify({ message }) + '\n');
}

function sanitizeCaptchaText(value) {
  return (value || '').replace(/\s+/g, '').replace(/[^A-Za-z0-9]/g, '');
}

async function ensureDir(dirPath) {
  try {
    await fs.promises.mkdir(dirPath, { recursive: true });
  } catch (_) {}
}

async function captureCaptchaImage(page) {
  if (!page || !SELECTORS.captchaImg) {
    return '';
  }
  try {
    const img = page.locator(SELECTORS.captchaImg);
    await img.waitFor({ state: 'visible', timeout: 5000 });
    await ensureDir(CAPTCHA_DIR);
    const filePath = path.join(
      CAPTCHA_DIR,
      `captcha_${Date.now()}_${Math.random().toString(36).slice(2)}.png`
    );
    await img.screenshot({ path: filePath });
    log(`Đã lưu ảnh captcha: ${filePath}`);
    return filePath;
  } catch (err) {
    log(`Không thể chụp ảnh captcha: ${err.message || err}`);
    return '';
  }
}

function requestCaptcha(meta) {
  return new Promise((resolve) => {
    rl.once('line', (line) => resolve((line || '').trim()));
    if (meta && typeof meta === 'object') {
      process.stdout.write(`CAPTCHA ${JSON.stringify(meta)}\n`);
    } else {
      process.stdout.write('CAPTCHA\n');
    }
  });
}

async function recognizeCaptchaWithTesseract(imagePath) {
  if (!imagePath) {
    return '';
  }
  try {
    const start = Date.now();
    const result = await Tesseract.recognize(imagePath, 'eng', {
      tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789',
    });
    const text = sanitizeCaptchaText(result?.data?.text || '');
    log(`[Captcha OCR] Tesseract đọc "${text}" trong ${Date.now() - start}ms (${imagePath})`);
    return text;
  } catch (err) {
    log(`[Captcha OCR] Lỗi Tesseract: ${err.message || err}`);
    return '';
  }
}

async function autoSolveCaptcha(page, maxAttempts = 3) {
  let lastImage = '';
  for (let i = 0; i < maxAttempts; i++) {
    const imagePath = await captureCaptchaImage(page);
    lastImage = imagePath || lastImage;
    if (imagePath) {
      const guess = await recognizeCaptchaWithTesseract(imagePath);
      if (guess) {
        return { value: guess, image: imagePath };
      }
    }
    if (i < maxAttempts - 1) {
      try {
        await page.click(SELECTORS.captchaRefresh);
      } catch (err) {
        log(`Không thể refresh captcha tự động: ${err.message || err}`);
      }
      await page.waitForTimeout(1200);
    }
  }
  return { value: '', image: lastImage };
}

async function waitForSelectorSafe(page, selector, timeout = 15000) {
  try {
    await page.waitForSelector(selector, { timeout });
    return true;
  } catch (_) {
    return false;
  }
}

async function selectOptionSafe(page, selector, preferredLabels = []) {
  try {
    await page.waitForSelector(selector, { timeout: 10000 });
  } catch (_) {
    return;
  }
  for (const label of preferredLabels) {
    try {
      await page.selectOption(selector, { label });
      return;
    } catch (_) {
      // continue
    }
  }
  const fallbackValue = await page.$eval(
    selector,
    (el) => {
      const options = Array.from(el.options || []);
      if (options.length > 1) {
        return options[1].value;
      }
      return options.length ? options[0].value : '';
    }
  ).catch(() => '');
  if (fallbackValue) {
    try {
      await page.selectOption(selector, fallbackValue);
    } catch (_) {
      // ignore
    }
  }
}

async function performLogin(page, username, password) {
  page.waitForTimeout(800);
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' });
  const loginReady = await waitForSelectorSafe(page, SELECTORS.username, 5000);
  if (!loginReady) {
    log('Trang Login chưa hiển thị ô nhập, reload lại trang login.');
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(1500);
    await page.waitForSelector(SELECTORS.username, { timeout: 20000 });
  }
  await page.fill(SELECTORS.username, username);
  await page.fill(SELECTORS.password, password);
  try {
    await page.click(SELECTORS.loginType);
  } catch (_) {}

  const maxAttempts = 6;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    let captchaValue = '';
    let captchaImage = '';
    const autoResult = await autoSolveCaptcha(page, 2);
    captchaValue = autoResult.value;
    captchaImage = autoResult.image;
    if (!captchaValue) {
      log('OCR tự động không đọc được captcha, chuyển sang hỏi Python/GUI.');
      const fallback = await requestCaptcha(
        captchaImage ? { image: captchaImage } : null
      );
      captchaValue = fallback;
    }
    if (!captchaValue) {
      throw new Error('Captcha trống.');
    }
    await page.fill(SELECTORS.captcha, captchaValue);
    log(`Thử captcha lần ${attempt + 1} với mã "${captchaValue}"`);
    await Promise.all([
      page.click(SELECTORS.loginBtn),
      page.waitForTimeout(800),
    ]);
    const success = await waitForSelectorSafe(page, SELECTORS.moduleLink, 15000);
    if (success) {
      log('Đăng nhập thành công.');
      return;
    }
    const stillLogin = await page.$(SELECTORS.captcha);
    if (!stillLogin) {
      throw new Error('Không xác định được trạng thái sau đăng nhập.');
    }
    log(`Captcha sai (lần ${attempt + 1}). Yêu cầu nhập lại.`);
    try {
      await page.click(SELECTORS.captchaRefresh);
    } catch (_) {}
    await page.waitForTimeout(1200);
  }
  throw new Error('Đăng nhập thất bại do captcha không hợp lệ.');
}

async function openModule(page) {
  await page.waitForSelector(SELECTORS.moduleLink, { timeout: 30000 });
  await page.click(SELECTORS.moduleLink);
  await page.waitForTimeout(1500);
  await page.waitForSelector(SELECTORS.newBtn, { timeout: 30000 });
}

async function openIssueForm(page) {
  const targetUrl = 'https://pus.customs.gov.vn/faces/SoDinhDanh';
  const maxTries = 3;
  for (let attempt = 0; attempt < maxTries; attempt++) {
    try {
      await page.click(SELECTORS.newBtn);
    } catch (err) {
      log(`Không thể bấm nút cấp mới: ${err.message || err}`);
    }
    const ready = await waitForSelectorSafe(page, SELECTORS.flowSelect, 8000);
    if (ready) {
      return true;
    }
    log(`Không thấy flowSelect sau khi bấm cấp mới (lần ${attempt + 1}/${maxTries}), reload lại trang SoDinhDanh.`);
    try {
      await page.goto(targetUrl, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(2000);
      await page.waitForSelector(SELECTORS.newBtn, { timeout: 20000 });
    } catch (err) {
      log(`Lỗi khi reload trang SoDinhDanh: ${err.message || err}`);
    }
  }
  throw new Error('Không mở được form cấp mới sau nhiều lần thử.');
}

async function issueSingle(page, username) {
  await openIssueForm(page);
  await selectOptionSafe(page, SELECTORS.flowSelect, ['Xuất khẩu', 'Xu?t kh?u']);
  await selectOptionSafe(page, SELECTORS.dnSelect, ['DN XNK']);
  await page.fill(SELECTORS.mstInput, username);
  await page.click(SELECTORS.confirmBtn);
  await page.waitForSelector(SELECTORS.codeField, { timeout: 30000 });
  const code = await page.$eval(SELECTORS.codeField, (el) => (el.value || el.textContent || '').trim());
  if (!code) {
    throw new Error('Không đọc được số định danh.');
  }
  const timestamp = new Date().toISOString();
  await page.click(SELECTORS.closeBtn);
  await page.waitForTimeout(500);
  return { code, time: timestamp };
}

async function withRetry(fn, attempts = 3) {
  let lastError;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn(i);
    } catch (err) {
      lastError = err;
      log(`Thử lần ${i + 1}/${attempts} thất bại: ${err.message || err}`);
    }
  }
  throw lastError;
}

async function main() {
  const args = parseArgs();
  const username = (args.username || '').trim();
  const password = (args.password || username || '').trim();
  const count = Math.max(1, parseInt(args.count || '1', 10));
  const headlessArg = (args.headless || 'false').toString().toLowerCase();
  const headless = headlessArg === 'true';

  if (!username) {
    throw new Error('Thiếu tham số --username.');
  }

  let browser;
  try {
    browser = await chromium.launch({ headless });
    const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const page = await context.newPage();
    await performLogin(page, username, password);
    await openModule(page);

    const items = [];
    for (let i = 0; i < count; i++) {
      const record = await withRetry(() => issueSingle(page, username), 2);
      items.push(record);
    }
    sendResult(items);
  } finally {
    rl.close();
    if (browser) {
      try {
        await browser.close();
      } catch (_) {}
    }
  }
}

main().catch((err) => {
  rl.close();
  sendError(err && err.message ? err.message : String(err));
  process.exit(1);
});
