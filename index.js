const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');

// =================================================================
// 1. HELPER FUNCTIONS (UTILITIES)
// =================================================================

// รอโหลดไฟล์และเปลี่ยนชื่อ
async function waitForDownloadAndRename(downloadPath, newFileName, maxWaitMs = 300000) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;
    const checkInterval = 10000; 
    let waittime = 0;

    while (waittime < maxWaitMs) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => 
            (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
            !f.endsWith('.crdownload') && 
            !f.startsWith('DTC_Completed_') &&
            !f.startsWith('Converted_')
        );
        
        if (downloadedFile) {
            // รออีกนิดเพื่อให้แน่ใจว่าเขียนไฟล์เสร็จ
            await new Promise(r => setTimeout(r, 10000));
            const stats = fs.statSync(path.join(downloadPath, downloadedFile));
            if (stats.size > 0) {
                console.log(`   ✅ File detected: ${downloadedFile}`);
                break; 
            }
        }
        
        await new Promise(resolve => setTimeout(resolve, checkInterval));
        waittime += checkInterval;
    }

    if (!downloadedFile) throw new Error(`Download timeout for ${newFileName}`);

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    // ลบไฟล์เก่าถ้ามี
    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    
    // แปลงเป็น CSV ทันที
    const csvFileName = `Converted_${newFileName.replace('.xls', '.csv')}`;
    const csvPath = path.join(downloadPath, csvFileName);
    await convertToCsv(newPath, csvPath);
    
    return csvPath;
}

// แปลงไฟล์ Excel/HTML เป็น CSV
async function convertToCsv(sourcePath, destPath) {
    try {
        console.log(`   🔄 Converting to CSV: ${path.basename(destPath)}`);
        const buffer = fs.readFileSync(sourcePath);
        let rows = [];

        // เช็คว่าเป็น XLSX (Binary) หรือไม่
        const isXLSX = buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4B;

        if (isXLSX) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1);
            worksheet.eachRow((row) => {
                const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
                rows.push(rowValues.map(v => (v && typeof v === 'object') ? (v.text || v.result || '') : String(v || '').trim()));
            });
        } else {
            // HTML Format (JSDOM)
            const content = buffer.toString('utf8');
            const dom = new JSDOM(content);
            const table = dom.window.document.querySelector('table');
            if (table) {
                const trs = Array.from(table.querySelectorAll('tr'));
                rows = trs.map(tr => Array.from(tr.querySelectorAll('td, th')).map(td => td.textContent.replace(/\s+/g, ' ').trim()));
            }
        }

        if (rows.length > 0) {
            let csvContent = '\uFEFF'; // BOM
            rows.forEach(row => {
                const escapedRow = row.map(cell => cell.includes(',') || cell.includes('"') ? `"${cell.replace(/"/g, '""')}"` : cell);
                csvContent += escapedRow.join(',') + '\n';
            });
            fs.writeFileSync(destPath, csvContent, 'utf8');
        }
    } catch (e) {
        console.warn(`   ⚠️ CSV Conversion warning: ${e.message}`);
    }
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

function zipFiles(sourceDir, outPath, filesToZip) {
    return new Promise((resolve, reject) => {
        const output = fs.createWriteStream(outPath);
        const archive = archiver('zip', { zlib: { level: 9 } });
        output.on('close', () => resolve(outPath));
        archive.on('error', (err) => reject(err));
        archive.pipe(output);
        filesToZip.forEach(file => archive.file(path.join(sourceDir, file), { name: file }));
        archive.finalize();
    });
}

// =================================================================
// 2. DATA PROCESSING & REPORT LOGIC (UPDATED)
// =================================================================

// แปลงเวลา (รองรับภาษาไทย และ HH:MM:SS)
function parseDuration(str) {
    if (!str) return 0;
    str = str.toString().trim();
    if (str.includes('นาที') || str.includes('ชม') || str.includes('วินาที')) {
        const h = (str.match(/(\d+)\s*ชม/) || [0,0])[1];
        const m = (str.match(/(\d+)\s*นาที/) || [0,0])[1];
        const s = (str.match(/(\d+)\s*วินาที/) || [0,0])[1];
        return (parseInt(h) * 3600) + (parseInt(m) * 60) + parseInt(s);
    }
    if (str.includes(':')) {
        const parts = str.split(':').map(Number);
        return parts.length === 3 ? (parts[0]*3600 + parts[1]*60 + parts[2]) : (parts[0]*60 + parts[1]);
    }
    return 0;
}

// แปลงวินาทีกลับเป็น HH:MM:SS
function formatTime(sec) {
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = sec % 60;
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

// อ่าน CSV แบบดิบ (ใช้สำหรับ Process Logic ใหม่)
function readCsvRaw(filePath, skipLines) {
    if (!fs.existsSync(filePath)) return [];
    const content = fs.readFileSync(filePath, 'utf8');
    const lines = content.split('\n').slice(skipLines).join('\n');
    return parse(lines, { columns: false, skip_empty_lines: true, relax_column_count: true, bom: true });
}

// ฟังก์ชัน Pivot (รวมเวลา) -> Return Top 10
function aggregateData(filePath, skipLines, colIdx) {
    if (!filePath || !fs.existsSync(filePath)) return { list: [], totalEvents: 0 };
    
    const records = readCsvRaw(filePath, skipLines);
    const pivot = {};
    let totalCount = 0;

    records.forEach(row => {
        // CSV Parse จะ return array 0-based.
        // ปรับปรุง: ใช้ index ตามที่ config มาเลย ไม่ต้องลบ 1 ซ้ำซ้อน ถ้า config เป็น 0-based แล้ว
        // หรือถ้า config เป็น 1-based (แบบ Excel) ก็ลบ 1 ถูกแล้ว แต่ต้องเช็คให้แน่ใจ
        // สมมติว่า config ที่ให้มาคือ 1-based (Col A=1)
        const license = row[colIdx.license - 1];
        const duration = row[colIdx.duration - 1];

        // เพิ่ม logging เพื่อ debug ถ้าไม่เจอข้อมูล
        // if (!license) console.log('Row missing license:', row);

        if (!license || !duration || license.includes('รวม') || license.includes('รถทั้งหมด')) return;

        const key = license.trim();
        const sec = parseDuration(duration);

        if (!pivot[key]) {
            pivot[key] = { 
                license: key, 
                station: colIdx.station ? (row[colIdx.station - 1] || '-') : '-',
                count: 0, 
                time: 0 
            };
        }
        pivot[key].count++;
        pivot[key].time += sec;
        totalCount++;
    });

    // เรียงตามเวลา (มาก->น้อย) และตัดมา 10 อันดับ
    const sortedList = Object.values(pivot)
        .sort((a, b) => b.time - a.time)
        .slice(0, 10);

    return { list: sortedList, totalEvents: totalCount };
}

// ฟังก์ชันดึงรายการเหตุการณ์ (Critical) -> Return Top 10
function getCriticalEvents(filePath, skipLines, colIdx) {
    if (!filePath || !fs.existsSync(filePath)) return [];
    
    const records = readCsvRaw(filePath, skipLines);
    const events = [];

    records.forEach(row => {
        const license = row[colIdx.license - 1];
        if (!license || license.includes('รวม')) return;

        events.push({
            license: license.trim(),
            desc: `${row[colIdx.v1 - 1]} -> ${row[colIdx.v2 - 1]} km/h`
        });
    });

    return events.slice(0, 10);
}

// =================================================================
// 3. MAIN SCRIPT
// =================================================================

(async () => {
    // ใส่ Secret ตรงนี้ หรือใช้ process.env ถ้า Run ผ่าน Github Action
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    
    // Validation
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing Secrets (DTC_USERNAME, DTC_PASSWORD).');
        process.exit(1);
    }

    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation...');
    
    const browser = await puppeteer.launch({
        headless: true, // ตั้งเป็น false ถ้าอยากเห็นจอ
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    await page.setViewport({ width: 1920, height: 1080 });

    try {
        // --- STEP 1: Login ---
        console.log('🔑 Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'networkidle2' });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        await Promise.all([
            page.click('#btnLogin'),
            page.waitForNavigation({ waitUntil: 'networkidle2' })
        ]);
        console.log('   Login Success');

        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;

        // --- STEP 2-6: Download Reports ---
        // (ส่วนนี้ใช้ Code เดิมของคุณที่ทำงานได้ดีแล้ว ผมย่อให้กระชับ)
        
        // Report 1: Over Speed
        console.log('📊 R1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max');
        await new Promise(r => setTimeout(r, 3000));
        await page.evaluate((s, e) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = s;
            document.getElementById('date10').value = e;
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            var sel = document.getElementById('ddl_truck');
            for(var o of sel.options) if(o.text.includes('ทั้งหมด')) sel.value = o.value;
            sel.dispatchEvent(new Event('change'));
        }, startDateTime, endDateTime);
        await page.evaluate(() => typeof sertch_data === 'function' ? sertch_data() : document.querySelector("span[onclick='sertch_data();']").click());
        // Wait for generation
        console.log('   Waiting for report generation (2 mins)...');
        await new Promise(r => setTimeout(r, 1100000)); 
        await page.evaluate(() => document.getElementById('btnexport').click());
        const file1 = await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // Report 2: Idling
        console.log('📊 R2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9');
        await new Promise(r => setTimeout(r, 3000));
        await page.evaluate((s, e) => {
            document.getElementById('date9').value = s;
            document.getElementById('date10').value = e;
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '10';
            var sel = document.getElementById('ddl_truck');
            for(var o of sel.options) if(o.text.includes('ทั้งหมด')) sel.value = o.value;
            sel.dispatchEvent(new Event('change'));
        }, startDateTime, endDateTime);
        await page.click('td:nth-of-type(6) > span');
        console.log('   Waiting for report generation (2 mins)...');
        await new Promise(r => setTimeout(r, 1100000));
        await page.evaluate(() => document.getElementById('btnexport').click());
        const file2 = await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

        // Report 3: Sudden Brake
        console.log('📊 R3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9');
        await new Promise(r => setTimeout(r, 3000));
        await page.evaluate((s, e) => {
            document.getElementById('date9').value = s;
            document.getElementById('date10').value = e;
            var sel = document.getElementById('ddl_truck');
            for(var o of sel.options) if(o.text.includes('ทั้งหมด')) sel.value = o.value;
            sel.dispatchEvent(new Event('change'));
        }, startDateTime, endDateTime);
        await page.click('td:nth-of-type(6) > span');
        console.log('   Waiting for report generation (2 mins)...');
        await new Promise(r => setTimeout(r, 1100000));
        await page.evaluate(() => {
            const b = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('Excel'));
            if(b) b.click();
        });
        const file3 = await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // Report 4: Harsh Start (Optional handling)
        console.log('📊 R4: Harsh Start...');
        let file4 = null; // ย้ายมาประกาศข้างนอก try-catch
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#date9');
            await new Promise(r => setTimeout(r, 3000));
            await page.evaluate((s, e) => {
                document.getElementById('date9').value = s;
                document.getElementById('date10').value = e;
                var sel = document.getElementById('ddl_truck');
                for(var o of sel.options) if(o.text.includes('ทั้งหมด')) sel.value = o.value;
                sel.dispatchEvent(new Event('change'));
            }, startDateTime, endDateTime);
            await page.evaluate(() => typeof sertch_data === 'function' ? sertch_data() : document.querySelector('td:nth-of-type(6) > span').click());
            console.log('   Waiting for report generation (2 mins)...');
            await new Promise(r => setTimeout(r, 1100000));
            await page.evaluate(() => {
                const b = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('Excel'));
                if(b) b.click();
            });
            file4 = await waitForDownloadAndRename(downloadPath, 'Report4_HarshStart.xls');
        } catch (e) {
            console.log('   ⚠️ Report 4 skipped or empty:', e.message);
        }

        // Report 5: Forbidden
        console.log('📊 R5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9');
        await new Promise(r => setTimeout(r, 3000));
        await page.evaluate((s, e) => {
            document.getElementById('date9').value = s;
            document.getElementById('date10').value = e;
            var sel = document.getElementById('ddl_truck');
            for(var o of sel.options) if(o.text.includes('ทั้งหมด')) sel.value = o.value;
            sel.dispatchEvent(new Event('change'));
            // Select Forbidden Zone
            var allS = document.getElementsByTagName('select');
            for(var sl of allS) for(var op of sl.options) if(op.text.includes('พิ้น')) { sl.value = op.value; sl.dispatchEvent(new Event('change')); break; }
        }, startDateTime, endDateTime);
        await new Promise(r => setTimeout(r, 3000)); // Wait for zones load
        await page.evaluate(() => {
            var allS = document.getElementsByTagName('select');
            for(var sl of allS) for(var op of sl.options) if(op.text.includes('สถานีทั้งหมด')) { sl.value = op.value; sl.dispatchEvent(new Event('change')); break; }
        });
        await page.click('td:nth-of-type(7) > span');
        console.log('   Waiting for report generation (2 mins)...');
        await new Promise(r => setTimeout(r, 1100000));
        await page.evaluate(() => document.getElementById('btnexport').click());
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');


        // =================================================================
        // STEP 7: Generate PDF Summary (IMPROVED LOGIC)
        // =================================================================
        console.log('📑 Step 7: Generating Final PDF Report...');

        // Config Data Structure (1-based index to match user habits, handled in aggregateData)
        // ตรวจสอบ Index จากไฟล์ CSV ตัวอย่าง:
        // Report 1: Col 2 (Name) -> Index 2, Col 5 (Duration) -> Index 5 ??
        // ลองเช็คจาก snippet: "ลำดับ,ชื่อรถ,วันที่บันทึก..." -> ชื่อรถคือ Col 2 (Index 1 in 0-based)
        // "รวมเวลา" คือ Col 5 (Index 4 in 0-based)
        // ดังนั้น Config: license: 2, duration: 5 (ถ้า function ลบ 1 ให้)
        
        // แก้ไข DATA_CONFIG ให้ตรงกับ CSV output จริงๆ (ดูจาก snippet)
        // Report 1 (OverSpeed): Col 2=ชื่อรถ, Col 5=รวมเวลา -> 1-based: 2, 5
        // Report 2 (Idling): Col 2=ชื่อรถ, Col 5=รวมเวลา -> 1-based: 2, 5
        // Report 3 (Sudden): Col 2=ชื่อรถ, Col 3=ทะเบียน, Col 5=v1, Col 6=v2 -> 1-based: 3 (ทะเบียน), 5, 6
        // Report 4 (Harsh): Col 2=ชื่อรถ, Col 3=ทะเบียน, Col 5=v1, Col 6=v2 -> 1-based: 3 (ทะเบียน), 5, 6
        // Report 5 (Forbidden): Col 2=ทะเบียน, Col 5=สถานี, Col 10=รวมเวลา -> 1-based: 2, 5, 10
        
        const DATA_CONFIG = {
            OVERSPEED: { skip: 6, license: 2, duration: 5 }, 
            IDLING: { skip: 7, license: 2, duration: 5 },    
            SUDDEN: { skip: 5, license: 3, v1: 5, v2: 6 },   
            HARSH: { skip: 5, license: 3, v1: 5, v2: 6 },    
            FORBIDDEN: { skip: 1, license: 2, station: 5, duration: 10 } 
        };

        // 1. Process Data
        const r1 = aggregateData(file1, DATA_CONFIG.OVERSPEED.skip, DATA_CONFIG.OVERSPEED);
        const r2 = aggregateData(file2, DATA_CONFIG.IDLING.skip, DATA_CONFIG.IDLING);
        const r5 = aggregateData(file5, DATA_CONFIG.FORBIDDEN.skip, DATA_CONFIG.FORBIDDEN);
        
        const r3_list = getCriticalEvents(file3, DATA_CONFIG.SUDDEN.skip, DATA_CONFIG.SUDDEN);
        const r4_list = (file4) ? getCriticalEvents(file4, DATA_CONFIG.HARSH.skip, DATA_CONFIG.HARSH) : [];

        // Count Critical
        const totalCritical = (readCsvRaw(file3, DATA_CONFIG.SUDDEN.skip).length) + 
                              (file4 ? readCsvRaw(file4, DATA_CONFIG.HARSH.skip).length : 0);

        const maxIdleCar = r2.list.length > 0 ? r2.list[0] : { time: 0, license: '-' };

        // 2. Generate HTML
        const html = `
        <!DOCTYPE html>
        <html lang="th">
        <head>
            <meta charset="UTF-8">
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;600;700&display=swap" rel="stylesheet">
            <style>
                @page { size: A4; margin: 0; }
                body { font-family: 'Noto Sans Thai', sans-serif; margin: 0; background: #fff; color: #333; -webkit-print-color-adjust: exact; }
                .page { width: 210mm; height: 296mm; position: relative; page-break-after: always; padding: 40px; box-sizing: border-box; }
                h1 { color: #1E40AF; font-size: 28px; margin: 0 0 10px 0; text-align: center; }
                .sub-head { text-align: center; color: #6B7280; margin-bottom: 40px; }
                .header-banner { background: #1E40AF; color: white; padding: 12px 20px; font-size: 20px; font-weight: bold; margin-bottom: 20px; border-radius: 4px; }
                .bg-orange { background: #F59E0B; }
                .bg-red { background: #DC2626; }
                .bg-purple { background: #9333EA; }
                .grid-2x2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 40px; }
                .card { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 25px; text-align: center; }
                .card-val { font-size: 42px; font-weight: bold; margin: 10px 0; }
                .text-blue { color: #1E40AF; } .text-orange { color: #F59E0B; } .text-red { color: #DC2626; } .text-purple { color: #9333EA; }
                .chart-box { margin: 20px 0; }
                .bar-row { display: flex; align-items: center; margin-bottom: 8px; font-size: 12px; }
                .bar-lbl { width: 180px; text-align: right; padding-right: 15px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 600; }
                .bar-trk { flex-grow: 1; background: #F1F5F9; height: 20px; border-radius: 4px; overflow: hidden; }
                .bar-fill { height: 100%; display: flex; align-items: center; justify-content: flex-end; padding-right: 8px; color: white; font-weight: bold; font-size: 11px; }
                table { width: 100%; border-collapse: collapse; font-size: 12px; }
                th { color: white; padding: 8px; text-align: left; }
                td { padding: 8px; border-bottom: 1px solid #E2E8F0; }
                tr:nth-child(even) { background: #F8FAFC; }
                .num { text-align: center; }
                .time { text-align: right; font-family: monospace; font-weight: 600; }
                .badge { padding: 2px 8px; border-radius: 10px; color: white; font-size: 10px; }
            </style>
        </head>
        <body>
            <!-- PAGE 1: Summary -->
            <div class="page">
                <h1>รายงานสรุปพฤติกรรมการขับขี่</h1>
                <div class="sub-head">Fleet Safety Report | ประจำวันที่: ${todayStr}</div>
                <div class="header-banner" style="text-align:center;">บทสรุปผู้บริหาร (Executive Summary)</div>
                <div class="grid-2x2">
                    <div class="card"><div style="font-weight:bold;">Over Speed (ครั้ง)</div><div class="card-val text-blue">${r1.totalEvents}</div><div style="font-size:12px; color:#666;">เหตุการณ์ทั้งหมด</div></div>
                    <div class="card"><div style="font-weight:bold;">Max Idling (นาที)</div><div class="card-val text-orange">${Math.round(maxIdleCar.time / 60)}m</div><div style="font-size:12px; color:#666;">${maxIdleCar.license}</div></div>
                    <div class="card"><div style="font-weight:bold;">Critical Events</div><div class="card-val text-red">${totalCritical}</div><div style="font-size:12px; color:#666;">เบรก/ออกตัว กระชาก</div></div>
                    <div class="card"><div style="font-weight:bold;">พื้นที่ห้ามจอด</div><div class="card-val text-purple">${r5.totalEvents}</div><div style="font-size:12px; color:#666;">จำนวนครั้ง</div></div>
                </div>
            </div>

            <!-- PAGE 2: Over Speed -->
            <div class="page">
                <div class="header-banner">1. การใช้ความเร็วเกินกำหนด (Top 10 by Duration)</div>
                <div class="chart-box">
                    ${r1.list.map(x => `<div class="bar-row"><div class="bar-lbl">${x.license}</div><div class="bar-trk"><div class="bar-fill" style="width:${(x.time/r1.list[0].time)*100}%; background:#1E40AF;">${formatTime(x.time)}</div></div></div>`).join('')}
                </div>
                <table><thead><tr><th style="background:#1E40AF; width:40px;">No.</th><th style="background:#1E40AF;">ทะเบียนรถ</th><th style="background:#1E40AF;" class="num">ครั้ง</th><th style="background:#1E40AF;" class="time">รวมเวลา</th></tr></thead><tbody>
                    ${r1.list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td class="num">${x.count}</td><td class="time">${formatTime(x.time)}</td></tr>`).join('')}
                </tbody></table>
            </div>

            <!-- PAGE 3: Idling -->
            <div class="page">
                <div class="header-banner bg-orange">2. การจอดไม่ดับเครื่อง (Top 10 by Duration)</div>
                <div class="chart-box">
                    ${r2.list.map(x => `<div class="bar-row"><div class="bar-lbl">${x.license}</div><div class="bar-trk"><div class="bar-fill" style="width:${(x.time/r2.list[0].time)*100}%; background:#F59E0B;">${formatTime(x.time)}</div></div></div>`).join('')}
                </div>
                <table><thead><tr><th style="background:#F59E0B; width:40px;">No.</th><th style="background:#F59E0B;">ทะเบียนรถ</th><th style="background:#F59E0B;" class="num">ครั้ง</th><th style="background:#F59E0B;" class="time">รวมเวลา</th></tr></thead><tbody>
                    ${r2.list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td class="num">${x.count}</td><td class="time">${formatTime(x.time)}</td></tr>`).join('')}
                </tbody></table>
            </div>

            <!-- PAGE 4: Critical -->
            <div class="page">
                <div class="header-banner bg-red">3. เหตุการณ์วิกฤต (Critical Safety Events)</div>
                <h3 style="color:#DC2626; border-left:4px solid #DC2626; padding-left:10px;">3.1 เบรกกะทันหัน (Sudden Brake)</h3>
                <table><thead><tr><th style="background:#DC2626; width:40px;">No.</th><th style="background:#DC2626;">ทะเบียนรถ</th><th style="background:#DC2626;">รายละเอียด</th><th style="background:#DC2626;" class="num">ระดับ</th></tr></thead><tbody>
                    ${r3_list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td>${x.desc}</td><td class="num"><span class="badge bg-red">High</span></td></tr>`).join('')}
                </tbody></table>
                <h3 style="color:#F59E0B; border-left:4px solid #F59E0B; padding-left:10px; margin-top:30px;">3.2 ออกตัวกระชาก (Harsh Start)</h3>
                <table><thead><tr><th style="background:#F59E0B; width:40px;">No.</th><th style="background:#F59E0B;">ทะเบียนรถ</th><th style="background:#F59E0B;">รายละเอียด</th><th style="background:#F59E0B;" class="num">ระดับ</th></tr></thead><tbody>
                    ${r4_list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td>${x.desc}</td><td class="num"><span class="badge bg-orange">Medium</span></td></tr>`).join('')}
                </tbody></table>
            </div>

            <!-- PAGE 5: Forbidden -->
            <div class="page">
                <div class="header-banner bg-purple">4. พื้นที่ห้ามจอด (Top 10 by Duration)</div>
                <div class="chart-box">
                    ${r5.list.map(x => `<div class="bar-row"><div class="bar-lbl">${x.license}</div><div class="bar-trk"><div class="bar-fill" style="width:${(x.time/r5.list[0].time)*100}%; background:#9333EA;">${formatTime(x.time)}</div></div></div>`).join('')}
                </div>
                <table><thead><tr><th style="background:#9333EA; width:40px;">No.</th><th style="background:#9333EA;">ทะเบียนรถ</th><th style="background:#9333EA;">สถานีล่าสุด</th><th style="background:#9333EA;" class="num">ครั้ง</th><th style="background:#9333EA;" class="time">รวมเวลา</th></tr></thead><tbody>
                    ${r5.list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td>${x.station}</td><td class="num">${x.count}</td><td class="time">${formatTime(x.time)}</td></tr>`).join('')}
                </tbody></table>
            </div>
        </body>
        </html>`;

        await page.setContent(html, { waitUntil: 'networkidle0' });
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true,
            margin: { top: 0, bottom: 0, left: 0, right: 0 }
        });
        console.log(`✅ PDF Generated: ${pdfPath}`);

        // --- STEP 8: Email ---
        console.log('📧 Sending Email...');
        const allFiles = fs.readdirSync(downloadPath);
        const csvsToZip = allFiles.filter(f => f.startsWith('Converted_') && f.endsWith('.csv'));
        const zipName = `DTC_Report_${todayStr.replace(/ /g, '_')}.zip`;
        const zipPath = path.join(downloadPath, zipName);
        
        if (csvsToZip.length > 0) await zipFiles(downloadPath, zipPath, csvsToZip);

        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: { user: EMAIL_USER, pass: EMAIL_PASS }
        });

        await transporter.sendMail({
            from: `"DTC Reporter" <${EMAIL_USER}>`,
            to: EMAIL_TO,
            subject: `รายงานสรุปพฤติกรรมการขับขี่ (Fleet Safety Report) - ${todayStr}`,
            text: `เรียน ผู้เกี่ยวข้อง\n\nนำส่งรายงานประจำวัน (PDF + CSV)\n\nขอบคุณครับ`,
            attachments: [
                { filename: zipName, path: zipPath },
                { filename: 'Fleet_Safety_Report.pdf', path: pdfPath }
            ]
        });
        console.log(`✅ Email Sent!`);

    } catch (err) {
        console.error('❌ Fatal Error:', err);
    } finally {
        await browser.close();
    }
})();
