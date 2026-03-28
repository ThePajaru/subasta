const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const { execFile } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3000;

// Uploads directory — Vercel has read-only filesystem, only /tmp is writable
const UPLOADS_DIR = process.env.VERCEL ? '/tmp/uploads' : path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Multer config for file uploads
const storage = multer.diskStorage({
    destination: UPLOADS_DIR,
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname);
    }
});
const upload = multer({ storage });

// ============================================================
// EXCEL PARSER
// ============================================================
function parseExcelFile(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref']);

    // Find header row (row 6, 0-indexed)
    const headerRowIdx = 6;
    const headers = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: headerRowIdx, c })];
        if (cell && cell.v) {
            headers[String(cell.v).trim()] = c;
        }
    }

    console.log('Detected headers:', headers);

    // Column mapping
    const colMap = {
        matricula: headers['Matrícula'] ?? 2,
        marca: headers['Marca'] ?? 3,
        modelo: headers['Modelo'] ?? 4,
        motor: headers['Motor'] ?? 5,
        derivado: headers['Derivado'] ?? 6,
        acabado: headers['Nivel de acabado'] ?? 7,
        tipoIVA: headers['Tipo de IVA'] ?? 8,
        carroceria: headers['Carrocería'] ?? 9,
        combustible: headers['Combustible'] ?? 10,
        color: headers['Color'] ?? 11,
        cv: headers['CV'] ?? 13,
        kilometraje: headers['Kilometraje'] ?? 14,
        fechaMatriculacion: headers['Fecha Matriculación'] ?? 15,
        uso: headers['Uso'] ?? 21,
        equipamiento: headers['Equipamiento'] ?? 22,
        procedencia: headers['Procedencia'] ?? 24,
        clasificacionDanos: headers['Clasificación de daños'] ?? 26,
        bcaLink: headers['Ver vínculo del lote'] ?? 27,
        precioGuia: headers['Precio guía'] ?? 28,
        precioVentaPublico: headers['Precio de venta al público'] ?? 29,
    };

    const cars = [];
    for (let r = headerRowIdx + 1; r <= range.e.r; r++) {
        const getVal = (col) => {
            const cell = sheet[XLSX.utils.encode_cell({ r, c: col })];
            return cell ? (cell.v !== undefined ? cell.v : '') : '';
        };

        const marca = String(getVal(colMap.marca)).trim();
        if (!marca) continue;

        const km = getVal(colMap.kilometraje);
        const cv = getVal(colMap.cv);
        const modelo = String(getVal(colMap.modelo)).trim();
        const derivado = String(getVal(colMap.derivado)).trim();
        const combustible = String(getVal(colMap.combustible)).trim();
        const fechaStr = String(getVal(colMap.fechaMatriculacion)).trim();
        const clasificacionDanos = String(getVal(colMap.clasificacionDanos)).trim();

        // Parse year from date
        let year = 0;
        if (fechaStr) {
            const parts = fechaStr.split('/');
            if (parts.length >= 3) {
                year = parseInt(parts[2]);
            } else if (parts.length === 2) {
                year = parseInt(parts[1]);
            }
            if (year < 100) year += 2000;
        }

        // Extract damage level number
        let damageLevel = 0;
        const dmgMatch = clasificacionDanos.match(/(\d+)/);
        if (dmgMatch) damageLevel = parseInt(dmgMatch[1]);

        cars.push({
            id: r,
            matricula: String(getVal(colMap.matricula)).trim(),
            marca,
            modelo,
            motor: String(getVal(colMap.motor)).trim(),
            derivado,
            carroceria: String(getVal(colMap.carroceria)).trim(),
            combustible,
            color: String(getVal(colMap.color)).trim(),
            cv: typeof cv === 'number' ? cv : parseInt(cv) || 0,
            kilometraje: typeof km === 'number' ? km : parseInt(String(km).replace(/\./g, '')) || 0,
            year,
            fechaMatriculacion: fechaStr,
            uso: String(getVal(colMap.uso)).trim(),
            procedencia: String(getVal(colMap.procedencia)).trim(),
            clasificacionDanos,
            damageLevel,
            bcaLink: String(getVal(colMap.bcaLink)).trim(),
            precioGuia: getVal(colMap.precioGuia) || null,
            precioVentaPublico: getVal(colMap.precioVentaPublico) || null,
        });
    }

    return cars;
}

// ============================================================
// COCHES.NET SEARCH URL BUILDER (for user reference links)
// ============================================================

function normalizeFuelType(fuel) {
    const f = fuel.toLowerCase();
    if (f.includes('diesel')) return 'diesel';
    if (f.includes('gasolina') || f.includes('petrol') || f.includes('gasoline')) return 'gasolina';
    if (f.includes('eléctric') || f.includes('electric') || f.includes('bev')) return 'electrico';
    if (f.includes('híbrido') || f.includes('hybrid') || f.includes('hev') || f.includes('phev')) return 'hibrido';
    if (f.includes('gnc') || f.includes('glp') || f.includes('gas')) return 'gas';
    return '';
}

// Build a human-clickable coches.net search URL (for reference only)
function buildCochesNetSearchURL(car) {
    const base = 'https://www.coches.net/segunda-mano/';
    let marca = car.marca.toLowerCase()
        .replace(/é/g, 'e').replace(/á/g, 'a').replace(/í/g, 'i').replace(/ó/g, 'o').replace(/ú/g, 'u')
        .replace(/ñ/g, 'n').replace(/\s+/g, '-');

    // Extract just the primary model name
    let modeloWords = car.modelo.toLowerCase()
        .replace(/ diesel/gi, '').replace(/ gasolina/gi, '').replace(/ híbrido/gi, '')
        .replace(/é/g, 'e').replace(/á/g, 'a').replace(/í/g, 'i').replace(/ó/g, 'o').replace(/ú/g, 'u')
        .replace(/ñ/g, 'n').replace(/[^a-z0-9\s]/g, '')
        .trim().split(/\s+/);
    
    let realModel = modeloWords[0];
    if (modeloWords.length > 1) {
        const first = modeloWords[0];
        if (['serie', 'clase', 'grand', 'santa', 'range', 'land'].includes(first)) {
            realModel = first + '-' + modeloWords[1];
        }
    }
    let modelo = realModel.replace(/-+/g, '-').replace(/-$/, '');

    let url = `${base}${marca}-${modelo}/`;
    const params = new URLSearchParams();

    if (car.year > 2000) {
        params.append('MinYear', Math.max(car.year - 1, 1990));
        params.append('MaxYear', car.year + 1);
    }
    if (car.kilometraje > 0) {
        params.append('MinKms', Math.max(0, car.kilometraje - 20000));
        params.append('MaxKms', car.kilometraje + 20000);
    }
    const fuelType = normalizeFuelType(car.combustible);
    if (fuelType === 'diesel') params.append('FuelTypeIds', '2');
    else if (fuelType === 'gasolina') params.append('FuelTypeIds', '1');
    else if (fuelType === 'electrico') params.append('FuelTypeIds', '4');
    else if (fuelType === 'hibrido') params.append('FuelTypeIds', '5');
    if (car.cv > 0) {
        params.append('MinPower', car.cv);
        params.append('MaxPower', car.cv);
    }
    const isAuto = /AUTO|AUT|DSG|EDC|EAT|CVT|TRONIC|GEARTRONIC/i.test(car.derivado || '') || /AUTO|AUT|DSG|EDC|EAT|CVT|TRONIC|GEARTRONIC/i.test(car.modelo || '');
    params.append('TransmissionTypeId%5B0%5D', isAuto ? '2' : '1');
    params.append('fi', 'Price');
    params.append('or', '1');
    const paramStr = params.toString();
    return paramStr ? `${url}?${paramStr}` : url;
}

// ============================================================
// APIFY SCRAPER - Coches.net via API (reliable, no bot blocks)
// ============================================================

// Helper to normalize strings (remove accents, lowercase, etc.)
function normalizeStr(str) {
    if (!str) return '';
    return str.toLowerCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove accents
        .replace(/[^a-z0-9\s]/g, '') // Remove special chars
        .trim();
}

// Build the search URL for a car with robust parameters
function buildSearchURL(car) {
    // 1. Try to load from coches_makes_models.json first!
    try {
        const dictPath = path.join(__dirname, 'coches_makes_models.json');
        if (fs.existsSync(dictPath)) {
            const allModels = JSON.parse(fs.readFileSync(dictPath, 'utf8'));
            const targetMakeNorm = normalizeStr(car.marca);
            const makeModels = allModels.filter(m => normalizeStr(m.make) === targetMakeNorm || targetMakeNorm.includes(normalizeStr(m.make)));
            
            if (makeModels.length > 0) {
                const targetModelNorm = normalizeStr(car.modelo || '');
                let bestMatch = null;
                makeModels.sort((a, b) => b.model.length - a.model.length);
                
                for (const m of makeModels) {
                    const mNorm = normalizeStr(m.model);
                    if (targetModelNorm === mNorm || targetModelNorm.startsWith(mNorm + ' ') || targetModelNorm.includes(' ' + mNorm + ' ')) {
                        bestMatch = m;
                        break;
                    }
                }
                if (!bestMatch) {
                    for (const m of makeModels) {
                        const mNorm = normalizeStr(m.model);
                        if (targetModelNorm.includes(mNorm)) {
                            bestMatch = m;
                            break;
                        }
                    }
                }

                if (bestMatch) {
                    console.log(`[ID MATCH] ${car.marca} ${car.modelo} -> MakeId: ${bestMatch.makeId}, ModelId: ${bestMatch.modelId}`);
                    
                    let url = `https://www.coches.net/segunda-mano/?MakeIds%5B0%5D=${bestMatch.makeId}&ModelIds%5B0%5D=${bestMatch.modelId}`;
                    
                    if (car.year > 2000) url += `&MinYear=${Math.max(car.year - 1, 1990)}&MaxYear=${car.year + 1}`;
                    if (car.kilometraje > 0) {
                        url += `&MinKms=${Math.max(0, car.kilometraje - 30000)}&MaxKms=${car.kilometraje + 30000}`;
                    }
                    const f = normalizeFuelType(car.combustible);
                    if (f === 'diesel') url += `&FuelTypeIds%5B0%5D=2`;
                    else if (f === 'gasolina') url += `&FuelTypeIds%5B0%5D=1`;
                    else if (f === 'electrico') url += `&FuelTypeIds%5B0%5D=3`;
                    else if (f === 'hibrido') url += `&FuelTypeIds%5B0%5D=6`;
                    else if (f === 'gas') url += `&FuelTypeIds%5B0%5D=4`;
                    
                    url += '&fi=Price&or=1';
                    return url;
                }
            }
        }
    } catch (e) { console.error('Error reading ID dictionary:', e.message); }

    // 2. Fallback to slug method
    const base = 'https://www.coches.net/segunda-mano/';
    let marca = car.marca.toLowerCase().replace(/é/g, 'e').replace(/á/g, 'a').replace(/í/g, 'i').replace(/ó/g, 'o').replace(/ú/g, 'u').replace(/ñ/g, 'n').replace(/\s+/g, '-');
    let modeloWords = car.modelo.toLowerCase().replace(/ diesel/gi, '').replace(/ gasolina/gi, '').replace(/ híbrido/gi, '').replace(/é/g, 'e').replace(/á/g, 'a').replace(/í/g, 'i').replace(/ó/g, 'o').replace(/ú/g, 'u').replace(/ñ/g, 'n').replace(/[^a-z0-9\s]/g, '').trim().split(/\s+/);
    
    let realModel = modeloWords[0];
    if (modeloWords.length > 1 && ['serie', 'clase', 'grand', 'santa', 'range', 'land'].includes(modeloWords[0])) {
        realModel = modeloWords[0] + '-' + modeloWords[1];
    }
    let modelo = realModel.replace(/-+/g, '-').replace(/-$/, '');

    let url = `${base}${marca}-${modelo}/`;
    const params = new URLSearchParams();
    if (car.year > 2000) {
        params.append('MinYear', Math.max(car.year - 1, 1990));
        params.append('MaxYear', car.year + 1);
    }
    const fuelType = normalizeFuelType(car.combustible);
    if (fuelType === 'diesel') params.append('FuelTypeIds', '2');
    else if (fuelType === 'gasolina') params.append('FuelTypeIds', '1');
    params.append('fi', 'Price');
    params.append('or', '1');
    
    const paramStr = params.toString();
    return paramStr ? `${url}?${paramStr}` : url;
}


// ============================================================
// CURL-BASED FETCHER — más fiable que Puppeteer para coches.net
// ============================================================
// Proxy config — DataImpulse residential proxies
// Set PROXY_USER and PROXY_PASS as environment variables in Vercel dashboard
const PROXY_USER = process.env.PROXY_USER || '';
const PROXY_PASS = process.env.PROXY_PASS || '';
const PROXY_HOST = process.env.PROXY_HOST || 'gw.dataimpulse.com';
const PROXY_PORT = process.env.PROXY_PORT || '823';

function fetchWithCurl(url) {
    return new Promise((resolve, reject) => {
        const args = [
            '-s',
            '--compressed',
            '-L',
            '--max-time', '30',
            '-H', 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
            '-H', 'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            '-H', 'Accept-Language: es-ES,es;q=0.9,en;q=0.8',
            '-H', 'Sec-Ch-Ua: "Google Chrome";v="131"',
            '-H', 'Sec-Ch-Ua-Mobile: ?0',
            '-H', 'Sec-Ch-Ua-Platform: "Windows"',
            '-H', 'Sec-Fetch-Dest: document',
            '-H', 'Sec-Fetch-Mode: navigate',
            '-H', 'Sec-Fetch-Site: none',
            '-H', 'Sec-Fetch-User: ?1',
            '-H', 'Upgrade-Insecure-Requests: 1',
        ];

        // Use proxy if credentials are configured — Spanish IPs, rotates on every request
        if (PROXY_USER && PROXY_PASS) {
            args.push('--proxy', `http://${PROXY_USER}__cr.es:${PROXY_PASS}@${PROXY_HOST}:${PROXY_PORT}`);
        }

        args.push(url);

        execFile('curl', args, { maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
            if (error) return reject(new Error(`curl error: ${error.message} | stderr: ${stderr}`));
            resolve(stdout);
        });
    });
}

// Safely decodes the JSON string embedded inside JSON.parse("...") in HTML
function decodeInitialProps(encoded) {
    // Method 1: let JSON.parse handle all escape sequences natively
    try {
        const jsonStr = JSON.parse('"' + encoded + '"');
        return JSON.parse(jsonStr);
    } catch (_) {}
    // Method 2: manual unescape — order matters: \\ must go first
    try {
        const jsonStr = encoded
            .replace(/\\\\/g, '\x00')  // protect escaped backslashes
            .replace(/\\"/g, '"')
            .replace(/\\n/g, '\n')
            .replace(/\\r/g, '\r')
            .replace(/\\t/g, '\t')
            .replace(/\\u([0-9a-fA-F]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
            .replace(/\x00/g, '\\');   // restore single backslashes
        return JSON.parse(jsonStr);
    } catch (_) {}
    return null;
}

function extractPricesFromHtml(html) {
    let data = null;

    const m1 = html.match(/window\.__INITIAL_PROPS__\s*=\s*JSON\.parse\("([\s\S]+?)"\)\s*;/);
    if (m1) {
        data = decodeInitialProps(m1[1]);
        if (data) console.log('[SCRAPER] Matched pattern 1');
    }

    if (!data) {
        const m2 = html.match(/window\.__INITIAL_PROPS__\s*=\s*JSON\.parse\('([\s\S]+?)'\)\s*;/);
        if (m2) {
            try { data = JSON.parse(m2[1].replace(/\\'/g, "'")); } catch (_) {}
            if (data) console.log('[SCRAPER] Matched pattern 2');
        }
    }

    if (!data) {
        const m3 = html.match(/window\.__INITIAL_PROPS__\s*=\s*(\{[\s\S]+?\})\s*;[\s\S]*?<\/script>/);
        if (m3) {
            try { data = JSON.parse(m3[1]); } catch (_) {}
            if (data) console.log('[SCRAPER] Matched pattern 3');
        }
    }

    if (!data) return null;

    const items =
        data?.initialResults?.items ||
        data?.props?.initialResults?.items ||
        data?.pageProps?.initialResults?.items ||
        data?.listings?.items ||
        data?.results?.items ||
        [];

    const prices = items
        .map(p => p.price ?? p.Price ?? p.pvp ?? p.salePrice ?? p.priceInfo?.price ?? null)
        .filter(p => typeof p === 'number' && p > 0);

    return { items, prices, totalResults: data?.initialResults?.totalResults ?? items.length };
}

// Rate limiter
// - Sin proxy: batches de 4 con pausa de 60s y 12s entre requests
// - Con proxy: solo 1s entre requests (cada uno va con IP diferente)
let lastRequestTime = 0;
let requestCountInBatch = 0;
const BATCH_SIZE = 4;
const BATCH_PAUSE_MS = 60000;
const MIN_DELAY_MS = PROXY_USER ? 1000 : 12000;

async function waitForRateLimit() {
    if (!PROXY_USER) {
        // Sin proxy: pausas largas entre batches
        if (requestCountInBatch > 0 && requestCountInBatch % BATCH_SIZE === 0) {
            console.log(`[RATE] Pausa de lote. Esperando ${BATCH_PAUSE_MS / 1000}s...`);
            await new Promise(r => setTimeout(r, BATCH_PAUSE_MS));
            requestCountInBatch = 0;
            lastRequestTime = 0;
        }
    }

    const now = Date.now();
    const wait = Math.max(0, lastRequestTime + MIN_DELAY_MS - now);
    if (wait > 0) await new Promise(r => setTimeout(r, wait));
    lastRequestTime = Date.now();
    requestCountInBatch++;
}

async function scrapeMarketPrices(car) {
    const searchURL = buildSearchURL(car);
    console.log(`[SCRAPER] ${car.marca} ${car.modelo} → ${searchURL}`);

    await waitForRateLimit();

    try {
        const html = await fetchWithCurl(searchURL);
        console.log(`[SCRAPER] ${html.length} bytes`);

        // Guide says blocked response is ~8,715 bytes (thin client-side shell without __INITIAL_PROPS__)
        if (html.length < 15000) {
            console.warn(`[SCRAPER] Bloqueado (${html.length} bytes) — respuesta sin __INITIAL_PROPS__`);
            return { url: searchURL, prices: [], success: false };
        }

        const result = extractPricesFromHtml(html);

        if (!result || result.items.length === 0) {
            console.warn(`[SCRAPER] Sin resultados para ${car.marca} ${car.modelo}`);
            return { url: searchURL, prices: [], success: false };
        }

        // Post-filter by exact CV (coches.net ignores power params in URL)
        let filteredItems = result.items;
        if (car.cv > 0) {
            const cvFiltered = filteredItems.filter(i => i.hp === car.cv);
            console.log(`[SCRAPER] Filtrado por CV exacto ${car.cv}: ${cvFiltered.length}/${result.items.length} anuncios`);
            filteredItems = cvFiltered;
        }

        const prices = filteredItems
            .map(i => i.price ?? i.Price ?? null)
            .filter(p => typeof p === 'number' && p > 0);

        if (prices.length === 0) {
            console.warn(`[SCRAPER] Sin precios para ${car.marca} ${car.modelo}`);
            return { url: searchURL, prices: [], success: false };
        }

        const sorted = [...new Set(prices)].sort((a, b) => a - b);
        console.log(`[SCRAPER] ${sorted.length} precios: ${sorted.slice(0, 8).join(', ')}`);

        return {
            url: searchURL,
            prices: sorted,
            totalResults: result.totalResults,
            success: true,
            rawItems: result.items.length,
            method: 'curl',
        };

    } catch (error) {
        console.error(`[SCRAPER ERROR] ${car.marca} ${car.modelo}: ${error.message}`);
        return { url: searchURL, prices: [], success: false };
    }
}

// ============================================================
// BCA AUCTION FEES (Tarifas de Servicios v. 1 junio 2025)
// ============================================================
const BCA_FEE_TABLE = [
    { from: 0,     to: 249.99,   fee: 102 },
    { from: 250,   to: 499.99,   fee: 135 },
    { from: 500,   to: 749.99,   fee: 181 },
    { from: 750,   to: 999.99,   fee: 198 },
    { from: 1000,  to: 1249.99,  fee: 215 },
    { from: 1250,  to: 1499.99,  fee: 237 },
    { from: 1500,  to: 1749.99,  fee: 257 },
    { from: 1750,  to: 1999.99,  fee: 279 },
    { from: 2000,  to: 2499.99,  fee: 296 },
    { from: 2500,  to: 2999.99,  fee: 314 },
    { from: 3000,  to: 3499.99,  fee: 337 },
    { from: 3500,  to: 3999.99,  fee: 340 },
    { from: 4000,  to: 4499.99,  fee: 343 },
    { from: 4500,  to: 4999.99,  fee: 347 },
    { from: 5000,  to: 5499.99,  fee: 350 },
    { from: 5500,  to: 5999.99,  fee: 353 },
    { from: 6000,  to: 6499.99,  fee: 357 },
    { from: 6500,  to: 6999.99,  fee: 361 },
    { from: 7000,  to: 7499.99,  fee: 364 },
    { from: 7500,  to: 7999.99,  fee: 367 },
    { from: 8000,  to: 8499.99,  fee: 370 },
    { from: 8500,  to: 8999.99,  fee: 375 },
    { from: 9000,  to: 9499.99,  fee: 378 },
    { from: 9500,  to: 9999.99,  fee: 381 },
    { from: 10000, to: 10499.99, fee: 384 },
    { from: 10500, to: 10999.99, fee: 388 },
    { from: 11000, to: 11499.99, fee: 392 },
    { from: 11500, to: 11999.99, fee: 395 },
    { from: 12000, to: 12999.99, fee: 398 },
    { from: 13000, to: 13999.99, fee: 401 },
    { from: 14000, to: 14999.99, fee: 405 },
    { from: 15000, to: 15999.99, fee: 409 },
    { from: 16000, to: 16999.99, fee: 434 },
    { from: 17000, to: 17999.99, fee: 459 },
    { from: 18000, to: 18999.99, fee: 484 },
    { from: 19000, to: 19999.99, fee: 509 },
];

// Tasa de adquisición + IVA 21%
// Para >= 20.000€: 2,5% del precio de adjudicación + IVA
function calculateBCAAcquisitionFee(bidPrice) {
    if (bidPrice <= 0) return 0;
    if (bidPrice >= 20000) return Math.round(bidPrice * 0.025 * 1.21);
    const entry = BCA_FEE_TABLE.find(e => bidPrice >= e.from && bidPrice <= e.to);
    const feeExcl = entry ? entry.fee : BCA_FEE_TABLE[BCA_FEE_TABLE.length - 1].fee;
    return Math.round(feeExcl * 1.21);
}

// Gestión y Transferencia BCA:
//   - Tasas de tráfico (sin IVA): 55,70€
//   - Honorarios de gestión (+IVA 21%): 69,77 * 1.21 = 84,42€
const BCA_GESTION_FEE = Math.round(55.70 + 69.77 * 1.21); // ~140€

// ============================================================
// PRICE CALCULATOR
// ============================================================
function calculateMarketPrice(prices) {
    if (!prices || prices.length === 0) {
        return { marketPrice: 0, confidence: 0, method: 'none', stats: {} };
    }

    const sorted = [...prices].sort((a, b) => a - b);
    const n = sorted.length;

    const sum = sorted.reduce((a, b) => a + b, 0);
    const mean = sum / n;
    const median = n % 2 === 0 ? (sorted[n / 2 - 1] + sorted[n / 2]) / 2 : sorted[Math.floor(n / 2)];

    // Take the 5 cheapest
    const topCheapest = sorted.slice(0, 5);
    const sumTop = topCheapest.reduce((a, b) => a + b, 0);
    const avgTop = sumTop / topCheapest.length;

    const q1 = sorted[Math.floor(n * 0.25)];
    const q3 = sorted[Math.floor(n * 0.75)];
    const iqr = q3 - q1;

    let marketPrice = Math.round(avgTop);
    let method = `Media de los ${topCheapest.length} más baratos`;
    let confidence = Math.min(95, 40 + topCheapest.length * 10);

    if (n > 5 && marketPrice > sorted[0] * 2) {
        marketPrice = sorted[0];
        method = 'El más barato (gran disparidad)';
    }

    return {
        marketPrice,
        confidence,
        method,
        cheapestPrices: topCheapest,
        stats: {
            count: n,
            filteredCount: topCheapest.length,
            mean: Math.round(mean),
            median: Math.round(median),
            min: sorted[0],
            max: sorted[n - 1],
            q1: Math.round(q1),
            q3: Math.round(q3),
            iqr: Math.round(iqr),
        }
    };
}

function calculateBuyPrice(marketPrice, damageLevel, targetMargin = 0.35) {
    if (!marketPrice) return { buyPrice: 0, estimatedProfit: 0 };

    const damageCostEstimates = { 0: 0, 1: 200, 2: 500, 3: 1200, 4: 2500, 5: 4500 };
    const repairCost = damageCostEstimates[damageLevel] ?? damageCostEstimates[3];
    const preparationCost = 300;

    // Iterative solve: acquisition fee depends on bid price, so estimate once then refine
    const fixedWithoutAcq = preparationCost + repairCost + BCA_GESTION_FEE;
    const bidEst = Math.max(0, Math.round(marketPrice / (1 + targetMargin) - fixedWithoutAcq));
    const acqFeeEst = calculateBCAAcquisitionFee(bidEst);

    const totalFixed = fixedWithoutAcq + acqFeeEst;
    const buyPrice = Math.max(0, Math.round(marketPrice / (1 + targetMargin) - totalFixed));

    // Recalculate fee for the refined bid price (one more iteration for accuracy)
    const bcaAcquisitionFee = calculateBCAAcquisitionFee(buyPrice);
    const totalFixedCosts = preparationCost + repairCost + BCA_GESTION_FEE + bcaAcquisitionFee;
    const buyPriceFinal = Math.max(0, Math.round(marketPrice / (1 + targetMargin) - totalFixedCosts));

    const totalOutOfPocket = buyPriceFinal + bcaAcquisitionFee + BCA_GESTION_FEE;
    const estimatedProfit = marketPrice - buyPriceFinal - totalFixedCosts;
    const actualMargin = buyPriceFinal > 0 ? (estimatedProfit / buyPriceFinal) * 100 : 0;

    return {
        buyPrice: buyPriceFinal,
        sellingPrice: marketPrice,
        repairCost,
        preparationCost,
        bcaAcquisitionFee,
        bcaGestionFee: BCA_GESTION_FEE,
        totalOutOfPocket,
        totalFixedCosts,
        estimatedProfit: Math.round(estimatedProfit),
        actualMargin: Math.round(actualMargin * 10) / 10,
        targetMargin: targetMargin * 100,
    };
}

// ============================================================
// API ROUTES
// ============================================================

// Upload Excel file
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

        const cars = parseExcelFile(req.file.path);
        console.log(`Parsed ${cars.length} cars from Excel`);

        res.json({
            success: true,
            count: cars.length,
            cars,
            auctionName: req.file.originalname,
        });
    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Analyze a single car
app.post('/api/analyze', async (req, res) => {
    try {
        const { car, targetMargin } = req.body;
        if (!car) return res.status(400).json({ error: 'No car data provided' });

        const margin = targetMargin || 0.35;
        const scrapeResult = await scrapeMarketPrices(car);
        const priceAnalysis = calculateMarketPrice(scrapeResult.prices);
        const buyCalc = calculateBuyPrice(priceAnalysis.marketPrice, car.damageLevel, margin);

        res.json({
            success: true,
            car,
            scrapeResult: {
                url: scrapeResult.url,
                pricesFound: scrapeResult.prices.length,
                success: scrapeResult.success,
            },
            priceAnalysis,
            buyCalculation: buyCalc,
        });
    } catch (error) {
        console.error('Analysis error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Batch analyze
app.post('/api/analyze-batch', async (req, res) => {
    try {
        const { cars, targetMargin } = req.body;
        if (!cars || !cars.length) return res.status(400).json({ error: 'No cars provided' });

        res.writeHead(200, {
            'Content-Type': 'text/event-stream',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
        });

        const margin = targetMargin || 0.35;
        const results = [];

        for (let i = 0; i < cars.length; i++) {
            const car = cars[i];
            res.write(`data: ${JSON.stringify({ type: 'progress', current: i + 1, total: cars.length, car: `${car.marca} ${car.modelo}` })}\n\n`);

            try {
                const scrapeResult = await scrapeMarketPrices(car);
                const priceAnalysis = calculateMarketPrice(scrapeResult.prices);
                const buyCalc = calculateBuyPrice(priceAnalysis.marketPrice, car.damageLevel, margin);

                const result = {
                    car,
                    scrapeResult: {
                        url: scrapeResult.url,
                        pricesFound: scrapeResult.prices.length,
                        success: scrapeResult.success,
                    },
                    priceAnalysis,
                    buyCalculation: buyCalc,
                };

                results.push(result);
                res.write(`data: ${JSON.stringify({ type: 'result', index: i, result })}\n\n`);

            } catch (err) {
                console.error(`Error analyzing car ${i}:`, err.message);
                results.push({
                    car,
                    error: err.message,
                    scrapeResult: { success: false },
                    priceAnalysis: { marketPrice: 0, confidence: 0 },
                    buyCalculation: { buyPrice: 0 },
                });
            }
        }

        res.write(`data: ${JSON.stringify({ type: 'complete', results })}\n\n`);
        res.end();

    } catch (error) {
        console.error('Batch analysis error:', error);
        res.write(`data: ${JSON.stringify({ type: 'error', message: error.message })}\n\n`);
        res.end();
    }
});


// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Debug endpoint: tests scraping for a specific car URL and returns raw diagnostic info
app.post('/api/debug-scrape', async (req, res) => {
    try {
        const { url } = req.body;
        if (!url) return res.status(400).json({ error: 'url required' });

        const html = await fetchWithCurl(url);

        const hasInitialProps = html.includes('__INITIAL_PROPS__');
        const m1 = html.match(/window\.__INITIAL_PROPS__\s*=\s*JSON\.parse\("([\s\S]+?)"\)\s*;/);
        const m3 = html.match(/window\.__INITIAL_PROPS__\s*=\s*(\{[\s\S]+?\})\s*;[\s\S]*?<\/script>/);

        let parsedOk = false;
        let itemCount = 0;
        let pricesFound = [];
        let parseError = null;

        if (m1) {
            try {
                const jsonStr = JSON.parse('"' + m1[1] + '"');
                const data = JSON.parse(jsonStr);
                const items = data?.initialResults?.items || data?.props?.initialResults?.items || data?.pageProps?.initialResults?.items || [];
                parsedOk = true;
                itemCount = items.length;
                pricesFound = items.slice(0, 10).map(i => i.price ?? i.pvp ?? i.salePrice ?? '?');
            } catch (e) { parseError = e.message; }
        }

        res.json({
            htmlLength: html.length,
            hasInitialProps,
            patternMatch: { m1: !!m1, m3: !!m3 },
            parsedOk,
            itemCount,
            pricesFound,
            parseError,
            first500: html.slice(0, 500),
            snippetAroundProps: hasInitialProps
                ? html.substring(Math.max(0, html.indexOf('__INITIAL_PROPS__') - 20), html.indexOf('__INITIAL_PROPS__') + 120)
                : null,
        });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ============================================================
// START SERVER
// ============================================================
app.listen(PORT, () => {
    console.log(`\n🚗 BCA Price Analyzer running at http://localhost:${PORT}\n`);
    console.log(`Features:`);
    console.log(`  ✅ Upload BCA Excel files`);
    console.log(`  ✅ Direct scraping from coches.net (no external APIs)`);
    console.log(`  ✅ Price calculation based on 5 cheapest listings`);
    console.log(`  ✅ Recommended buy price for 30-40% profit margin`);
    console.log(`  ✅ Damage level cost adjustment\n`);
});
