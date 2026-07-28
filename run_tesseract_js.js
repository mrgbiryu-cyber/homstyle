const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const RUN_DIR = path.join(ROOT, 'poc_full_run');
const MANIFEST_PATH = path.join(RUN_DIR, 'image_manifest.json');
const IMAGE_DIR = path.join(RUN_DIR, 'images_flat');
const OUTPUT_PATH = path.join(RUN_DIR, 'ocr_tesseract_js.json');
const PACKAGE_ROOT = path.join(ROOT, '.tools', 'tesseract-js');
const LANG_PATH = path.join(PACKAGE_ROOT, 'lang-data');
const CACHE_PATH = path.join(PACKAGE_ROOT, 'cache');
const { createWorker, OEM, PSM } = require(path.join(PACKAGE_ROOT, 'node_modules', 'tesseract.js'));


function parseArgs() {
  const args = process.argv.slice(2);
  const limitIndex = args.indexOf('--limit');
  return {
    resume: args.includes('--resume'),
    limit: limitIndex >= 0 ? Number(args[limitIndex + 1] || 0) : 0,
  };
}


function saveOutput(metadata, items) {
  const tempPath = `${OUTPUT_PATH}.tmp`;
  fs.writeFileSync(tempPath, JSON.stringify({ metadata, items }, null, 2), 'utf8');
  fs.renameSync(tempPath, OUTPUT_PATH);
}


async function main() {
  const args = parseArgs();
  let manifest = JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf8'))
    .filter((row) => row.download_status === 'SUCCESS');
  if (args.limit > 0) manifest = manifest.slice(0, args.limit);

  const existing = new Map();
  if (args.resume && fs.existsSync(OUTPUT_PATH)) {
    const old = JSON.parse(fs.readFileSync(OUTPUT_PATH, 'utf8'));
    for (const row of old.items || []) {
      if (row.status === 'SUCCESS') existing.set(row.file, row);
    }
  }

  fs.mkdirSync(CACHE_PATH, { recursive: true });
  const initStarted = performance.now();
  const worker = await createWorker(['kor', 'eng'], OEM.LSTM_ONLY, {
    langPath: LANG_PATH,
    cachePath: CACHE_PATH,
    gzip: true,
  });
  await worker.setParameters({
    tessedit_pageseg_mode: PSM.AUTO,
    preserve_interword_spaces: '1',
  });
  const initMs = Math.round(performance.now() - initStarted);
  const pkg = JSON.parse(
    fs.readFileSync(path.join(PACKAGE_ROOT, 'node_modules', 'tesseract.js', 'package.json'), 'utf8'),
  );

  const metadata = {
    engine: 'Tesseract.js',
    engine_note: 'WebAssembly port of Tesseract; not the native Windows CLI build.',
    tesseract_js_version: pkg.version,
    language: 'kor+eng',
    oem: 'LSTM_ONLY',
    psm: 'AUTO',
    traineddata: 'tessdata 4.0.0 compressed packages',
    init_ms: initMs,
    started_at: new Date().toISOString(),
    input_count: manifest.length,
  };
  const items = [];
  const runStarted = performance.now();

  try {
    for (let i = 0; i < manifest.length; i += 1) {
      const entry = manifest[i];
      const filename = entry.file;
      if (existing.has(filename)) {
        items.push(existing.get(filename));
        process.stdout.write(`[${String(i + 1).padStart(2, '0')}/${manifest.length}] ${filename}: REUSED\n`);
        continue;
      }

      const started = performance.now();
      let row;
      try {
        const result = await worker.recognize(path.join(IMAGE_DIR, filename));
        const text = String(result.data.text || '').trim();
        const lines = text ? text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean) : [];
        row = {
          product_id: entry.product_id || '',
          file: filename,
          role: entry.role || '',
          language: 'kor+eng',
          status: 'SUCCESS',
          line_count: lines.length,
          character_count: text.length,
          mean_confidence: Number.isFinite(result.data.confidence)
            ? Number(result.data.confidence.toFixed(6))
            : null,
          text,
          lines,
          error: '',
        };
      } catch (error) {
        row = {
          product_id: entry.product_id || '',
          file: filename,
          role: entry.role || '',
          language: 'kor+eng',
          status: 'ERROR',
          line_count: 0,
          character_count: 0,
          mean_confidence: null,
          text: '',
          lines: [],
          error: `${error.name || 'Error'}: ${error.message || String(error)}`,
        };
      }

      row.elapsed_ms = Math.round(performance.now() - started);
      items.push(row);
      metadata.processed_count = items.length;
      metadata.elapsed_ms = Math.round(performance.now() - runStarted);
      saveOutput(metadata, items);
      process.stdout.write(
        `[${String(i + 1).padStart(2, '0')}/${manifest.length}] ${filename}: ${row.status} `
          + `lines=${row.line_count} chars=${row.character_count} elapsed=${row.elapsed_ms}ms\n`,
      );
    }
  } finally {
    await worker.terminate();
  }

  Object.assign(metadata, {
    completed_at: new Date().toISOString(),
    processed_count: items.length,
    success_count: items.filter((row) => row.status === 'SUCCESS').length,
    error_count: items.filter((row) => row.status === 'ERROR').length,
    nonempty_count: items.filter((row) => Boolean(row.text)).length,
    elapsed_ms: Math.round(performance.now() - runStarted),
  });
  saveOutput(metadata, items);
  process.stdout.write(`${JSON.stringify(metadata, null, 2)}\n`);
}


main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
