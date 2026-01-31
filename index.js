/**
 * DTC Automation Script
 * Version: 4.3.0 (CSV Fix & New PDF Logic Integration)
 * Last Updated: 30/01/2026
 * Features: 
 * - Strict Hard Wait
 * - Robust XLSX -> CSV Conversion
 * - PDF Generation using user-provided logic
 */

const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');

// --- Helper Functions ---

// 1. ฟังก์ชันรอโหลดไฟล์ และแปลงเป็น CSV
async function waitForDownloadAndRename(downloadPath, newFileName, maxWaitMs = 300000) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;
    const checkInterval = 10000; 
    let waittime = 0;

    // วนลูปรอไฟล์
    while (waittime < maxWaitMs) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => 
            (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
            !f.endsWith('.crdownload') && 
            !f.startsWith('DTC_Completed_') &&
            !f.startsWith('Converted_')
        );
        
        if (downloadedFile) {
            console.log(`   ✅ File detected: ${downloadedFile} (${waittime/1000}s)`);
            break; 
        }
        
        await new Promise(resolve => setTimeout(resolve, checkInterval));
        waittime += checkInterval;
    }

    if (!downloadedFile) throw new Error(`Download timeout for ${newFileName}`);

    await new Promise(resolve => setTimeout(resolve, 10000)); // รอเขียนไฟล์

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    const stats = fs.statSync(oldPath);
    if (stats.size === 0) throw new Error(`Downloaded file is empty!`);

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    
    // แปลงเป็น CSV (UTF-8)
    const csvFileName = `Converted_${newFileName.replace('.xls', '.csv')}`;
    const csvPath = path.join(downloadPath, csvFileName);
    await convertToCsv(newPath, csvPath);
    
    return csvPath;
}

// 2. ฟังก์ชันแปลงไฟล์ (รองรับทั้ง HTML Table และ XLSX Binary) -> CSV
async function convertToCsv(sourcePath, destPath) {
    try {
        console.log(`   🔄 Converting to CSV...`);
        const buffer = fs.readFileSync(sourcePath);
        let rows = [];

        // ตรวจสอบว่าเป็น XLSX (Zip based) หรือไม่ (Signature: PK)
        const isXLSX = buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4B;

        if (isXLSX) {
            console.log('      - Type: Binary XLSX (Using ExcelJS)');
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1); // อ่าน Sheet แรก
            
            worksheet.eachRow((row) => {
                // ExcelJS เริ่ม index 1
                const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
                rows.push(rowValues.map(v => {
                    if (v === null || v === undefined) return '';
                    if (typeof v === 'object') return v.text || v.result || ''; // Handle Rich Text/Formula
                    return String(v).trim();
                }));
            });
        } else {
            console.log('      - Type: HTML Table (Using JSDOM)');
            const content = buffer.toString('utf8');
            const dom = new JSDOM(content);
            const table = dom.window.document.querySelector('table');
            if (table) {
                const trs = Array.from(table.querySelectorAll('tr'));
                rows = trs.map(tr => 
                    Array.from(tr.querySelectorAll('td, th')).map(td => td.textContent.replace(/\s+/g, ' ').trim())
                );
            } else {
                console.warn('      ⚠️ No table found in HTML/Text file.');
            }
        }

        if (rows.length > 0) {
            // เขียน CSV พร้อม BOM
            let csvContent = '\uFEFF'; 
            rows.forEach(row => {
                const escapedRow = row.map(cell => {
                    if (cell.includes(',') || cell.includes('"') || cell.includes('\n')) {
                        return `"${cell.replace(/"/g, '""')}"`;
                    }
                    return cell;
                });
                csvContent += escapedRow.join(',') + '\n';
            });
            fs.writeFileSync(destPath, csvContent, 'utf8');
            console.log(`   ✅ CSV Created: ${path.basename(destPath)}`);
        } else {
            console.warn('   ⚠️ No data extracted for CSV conversion.');
        }

    } catch (e) {
        console.warn(`   ⚠️ CSV Conversion error: ${e.message}`);
    }
}

// 3. ฟังก์ชันรอตารางข้อมูล
async function waitForTableData(page, minRows = 2, timeout = 300000) {
    console.log(`   Waiting for table data (Max ${timeout/1000}s)...`);
    try {
        await page.waitForFunction((min) => {
            const rows = document.querySelectorAll('table tr');
            const bodyText = document.body.innerText;
            if (bodyText.includes('ไม่พบข้อมูล') || bodyText.includes('No data found')) return true; 
            return rows.length >= min; 
        }, { timeout: timeout }, minRows);
        console.log('   ✅ Table data populated.');
    } catch (e) {
        console.warn('   ⚠️ Wait for table data timed out.');
    }
}

// --- HELPER: แปลงเวลาไทย/HH:MM:SS เป็นวินาที (Code ของคุณ) ---
function parseDurationToSeconds(timeStr) {
    if (!timeStr) return 0;
    
    // กรณี 1: "0 ชม. 1 นาที 31 วินาที"
    const thaiMatch = timeStr.match(/(?:(\d+)\s*ชม\.)?\s*(?:(\d+)\s*นาที)?\s*(?:(\d+)\s*วินาที)?/);
    if (thaiMatch && timeStr.includes('นาที')) {
        const h = parseInt(thaiMatch[1] || 0);
        const m = parseInt(thaiMatch[2] || 0);
        const s = parseInt(thaiMatch[3] || 0);
        return (h * 3600) + (m * 60) + s;
    }

    // กรณี 2: "00:11:19" (HH:MM:SS)
    if (timeStr.includes(':')) {
        const parts = timeStr.split(':').map(Number);
        if (parts.length === 3) return (parts[0] * 3600) + (parts[1] * 60) + parts[2];
        if (parts.length === 2) return (parts[0] * 60) + parts[1]; // MM:SS
    }

    return 0;
}

// --- HELPER: แปลงวินาทีกลับเป็น HH:MM:SS ---
function formatSeconds(totalSeconds) {
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

// --- FUNCTION: อ่านและประมวลผล CSV (Code ของคุณ) ---
function processCSV(filePath, skipLines, colMap) {
    try {
        if (!fs.existsSync(filePath)) {
            console.warn(`File not found: ${filePath}`);
            return [];
        }
        
        const fileContent = fs.readFileSync(filePath, 'utf8');
        // ข้ามบรรทัด Header ด้านบน
        const lines = fileContent.split('\n').slice(skipLines).join('\n');
        
        const records = parse(lines, {
            columns: false,
            skip_empty_lines: true,
            relax_column_count: true,
            bom: true
        });

        return records.map(row => {
            const data = {};
            for (const [key, index] of Object.entries(colMap)) {
                // index ที่ส่งมาเป็น 1-based (Excel Style) ต้องลบ 1
                const idx = parseInt(index) - 1; 
                data[key] = row[idx] ? row[idx].trim() : '';
            }
            return data;
        }).filter(r => r.license); // กรองแถวว่างที่ไม่มีทะเบียนรถ
    } catch (err) {
        console.error(`Error reading ${filePath}:`, err.message);
        return [];
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

// --- Main Script ---

(async () => {
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing Secrets.');
        process.exit(1);
    }

    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (Revise PDF + Strict Wait)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(3600000); 
    page.setDefaultTimeout(3600000);
    
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    
    await page.setViewport({ width: 1920, height: 1080 });
    await page.emulateTimezone('Asia/Bangkok');

    try {
        // Step 1: Login
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`🕒 Global Time Settings: ${startDateTime} to ${endDateTime}`);

        // --- Step 2 to 6: DOWNLOAD REPORTS ---
        
        // REPORT 1: Over Speed
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        
        // Hard Wait 10s before fill
        await new Promise(r => setTimeout(r, 10000));

        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) {
                document.getElementById('ddlMinute').value = '1';
                document.getElementById('ddlMinute').dispatchEvent(new Event('change'));
            }
            var selectElement = document.getElementById('ddl_truck'); 
            var options = selectElement.options; 
            for (var i = 0; i < options.length; i++) { 
                if (options[i].text.includes('ทั้งหมด')) { selectElement.value = options[i].value; break; } 
            } 
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
        }, startDateTime, endDateTime);

        console.log('   Searching Report 1...');
        await page.evaluate(() => {
            if(typeof sertch_data === 'function') sertch_data();
            else document.querySelector("span[onclick='sertch_data();']").click();
        });

        // Hard Wait 5 Mins after search
        console.log('   ⏳ Waiting 5 mins...');
        await new Promise(resolve => setTimeout(resolve, 300000));
        
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        console.log('   Exporting Report 1...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        // Convert to CSV
        const file1 = await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // REPORT 2: Idling
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 10000));

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '10';
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);
        
        await page.click('td:nth-of-type(6) > span');
        
        // Hard Wait 3 mins
        console.log('   ⏳ Waiting 3 mins (Strict)...');
        await new Promise(r => setTimeout(r, 180000));

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file2 = await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

        // REPORT 3: Sudden Brake
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 10000));

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);
        
        await page.click('td:nth-of-type(6) > span');
        
        // Hard Wait 3 mins
        console.log('   ⏳ Waiting 3 mins (Strict)...'); 
        await new Promise(r => setTimeout(r, 180000)); 

        await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const b = btns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
            if (b) b.click(); else document.querySelector('#table button:nth-of-type(3)')?.click();
        });
        const file3 = await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // REPORT 4: Harsh Start
        console.log('📊 Processing Report 4: Harsh Start...');
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#date9', { visible: true, timeout: 60000 });
            await new Promise(r => setTimeout(r, 10000));
            
            console.log('   Setting Report 4 Conditions (Programmatic)...');
            await page.evaluate((start, end) => {
                document.getElementById('date9').value = start;
                document.getElementById('date10').value = end;
                document.getElementById('date9').dispatchEvent(new Event('change'));
                document.getElementById('date10').dispatchEvent(new Event('change'));
                const select = document.getElementById('ddl_truck');
                if (select) {
                    let found = false;
                    for (let i = 0; i < select.options.length; i++) {
                        if (select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                            select.selectedIndex = i; found = true; break;
                        }
                    }
                    if (!found && select.options.length > 0) select.selectedIndex = 0;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    if (typeof $ !== 'undefined' && $(select).data('select2')) { $(select).trigger('change'); }
                }
            }, startDateTime, endDateTime);
            
            await page.evaluate(() => {
                if (typeof sertch_data === 'function') { sertch_data(); } else { document.querySelector('td:nth-of-type(6) > span').click(); }
            });
            
            // Hard Wait 3 Mins
            console.log('   ⏳ Waiting 3 mins (Strict)...');
            await new Promise(r => setTimeout(r, 180000));
            
            console.log('   Clicking Export Report 4...');
            await page.evaluate(() => {
                const xpathResult = document.evaluate('//*[@id="table"]/div[1]/button[3]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                const btn = xpathResult.singleNodeValue;
                if (btn) btn.click();
                else {
                    const allBtns = Array.from(document.querySelectorAll('button'));
                    const excelBtn = allBtns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
                    if (excelBtn) excelBtn.click(); else throw new Error("Cannot find Export button for Report 4");
                }
            });
            const file4 = await waitForDownloadAndRename(downloadPath, 'Report4_HarshStart.xls');
        } catch (error) {
            console.error('❌ Report 4 Failed:', error.message);
        }

        // REPORT 5: Forbidden
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 10000));
        
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            
            // 1. รถทั้งหมด
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
            
            // 2. พื้นที่ห้ามเข้า (Updated: Fix typo "พิ้น")
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) { 
                for(var i=0; i<s.options.length; i++) { 
                    const txt = s.options[i].text;
                    if(txt.includes('พิ้น')) { 
                        s.value = s.options[i].value; 
                        s.dispatchEvent(new Event('change', { bubbles: true })); 
                        break; 
                    } 
                } 
            }
        }, startDateTime, endDateTime);
        
        await new Promise(r => setTimeout(r, 10000));
        await page.evaluate(() => {
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) { for(var i=0; i<s.options.length; i++) { if(s.options[i].text.includes('สถานีทั้งหมด')) { s.value = s.options[i].value; s.dispatchEvent(new Event('change', { bubbles: true })); break; } } }
        });
        
        await page.click('td:nth-of-type(7) > span');
        
        // Hard Wait 3 mins
        console.log('   ⏳ Waiting 3 mins (Strict)...');
        await new Promise(r => setTimeout(r, 180000));
        
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        // Convert to CSV
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 7: Generate PDF Summary (UPDATED WITH YOUR LOGIC)
        // =================================================================
        console.log('📑 Step 7: Generating PDF Summary...');

        const FILES_CSV = {
            OVERSPEED: file1,
            IDLING: file2,
            SUDDEN_BRAKE: file3,
            HARSH_START: typeof file4 !== 'undefined' ? file4 : '',
            PROHIBITED: file5
        };

        // 1. Process Report 1: Over Speed
        // (CSV Index: 0=No, 1=License, ..., Last=Duration)
        const rawSpeed = processCSV(FILES_CSV.OVERSPEED, 5, { license: 1, duration: 10 }); 
        const speedStats = {};
        rawSpeed.forEach(r => {
            if (!speedStats[r.license]) speedStats[r.license] = { count: 0, time: 0, license: r.license };
            speedStats[r.license].count++;
            speedStats[r.license].time += parseDurationToSeconds(r.duration);
        });
        const topSpeed = Object.values(speedStats).sort((a, b) => b.count - a.count).slice(0, 5);
        const totalOverSpeed = rawSpeed.length;

        // 2. Process Report 2: Idling
        const rawIdling = processCSV(FILES_CSV.IDLING, 6, { license: 1, duration: 4 });
        const idleStats = {};
        rawIdling.forEach(r => {
            if (!idleStats[r.license]) idleStats[r.license] = { count: 0, time: 0, license: r.license };
            idleStats[r.license].count++;
            idleStats[r.license].time += parseDurationToSeconds(r.duration);
        });
        const topIdle = Object.values(idleStats).sort((a, b) => b.time - a.time).slice(0, 5);
        const maxIdleCar = topIdle.length > 0 ? topIdle[0] : { time: 0, license: '-' };

        // 3. Process Report 3 & 4
        const rawBrake = fs.existsSync(FILES_CSV.SUDDEN_BRAKE) ? processCSV(FILES_CSV.SUDDEN_BRAKE, 4, { license: 2, v_start: 4, v_end: 5 }) : [];
        const rawStart = (FILES_CSV.HARSH_START && fs.existsSync(FILES_CSV.HARSH_START)) ? processCSV(FILES_CSV.HARSH_START, 4, { license: 2, v_start: 4, v_end: 5 }) : [];
        
        const criticalEvents = [
            ...rawBrake.map(r => ({ ...r, type: 'Sudden Brake', level: 'High' })),
            ...rawStart.map(r => ({ ...r, type: 'Harsh Start', level: 'Medium' }))
        ];

        // 4. Process Report 5
        const rawForbidden = processCSV(FILES_CSV.PROHIBITED, 5, { license: 1, station: 4, duration: 9 });
        const forbiddenList = rawForbidden.map(r => ({
            license: r.license,
            station: r.station,
            timeSec: parseDurationToSeconds(r.duration),
            timeStr: r.duration
        })).sort((a, b) => b.timeSec - a.timeSec).slice(0, 8);
        
        const forbiddenChartStats = {};
        rawForbidden.forEach(r => {
            if(!forbiddenChartStats[r.license]) forbiddenChartStats[r.license] = 0;
            forbiddenChartStats[r.license] += parseDurationToSeconds(r.duration);
        });
        const topForbiddenChart = Object.entries(forbiddenChartStats)
            .map(([license, time]) => ({ license, time }))
            .sort((a, b) => b.time - a.time).slice(0, 5);

        // --- HTML GENERATION ---
        const today = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
        
        const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;600;700&display=swap" rel="stylesheet">
            <style>
            @page { size: A4; margin: 0; }
            body { font-family: 'Noto Sans Thai', sans-serif; margin: 0; padding: 0; background: #fff; color: #333; }
            .page { width: 210mm; height: 296mm; position: relative; page-break-after: always; overflow: hidden; }
            .content { padding: 40px; }
            
            .header-banner { background: #1E40AF; color: white; padding: 15px 40px; font-size: 24px; font-weight: bold; margin-bottom: 30px; }
            h1 { font-size: 32px; color: #1E40AF; margin-bottom: 10px; }
            
            .grid-2x2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 50px; }
            .card { background: #F8FAFC; border-radius: 12px; padding: 30px; text-align: center; border: 1px solid #E2E8F0; }
            .card-title { font-size: 18px; font-weight: bold; margin-bottom: 10px; }
            .card-value { font-size: 48px; font-weight: bold; margin: 10px 0; }
            .card-sub { font-size: 14px; color: #64748B; }
            
            .c-blue { color: #1E40AF; }
            .c-orange { color: #F59E0B; }
            .c-red { color: #DC2626; }
            .c-purple { color: #9333EA; }
            
            .chart-container { margin: 40px 0; }
            .bar-row { display: flex; align-items: center; margin-bottom: 15px; }
            .bar-label { width: 180px; text-align: right; padding-right: 15px; font-weight: 600; font-size: 14px; }
            .bar-track { flex-grow: 1; background: #F1F5F9; height: 30px; border-radius: 4px; overflow: hidden; }
            .bar-fill { height: 100%; display: flex; align-items: center; justify-content: flex-end; padding-right: 10px; color: white; font-size: 12px; font-weight: bold; }
            
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th { background: #1E40AF; color: white; padding: 12px; text-align: left; }
            td { padding: 10px; border-bottom: 1px solid #E2E8F0; }
            tr:nth-child(even) { background: #F8FAFC; }
            .risk-High { color: #DC2626; font-weight: bold; }
            .risk-Medium { color: #F59E0B; font-weight: bold; }
            </style>
        </head>
        <body>

            <!-- Page 1: Executive Summary -->
            <div class="page">
            <div style="text-align: center; padding-top: 60px;">
                <h1 style="font-size: 48px;">รายงานสรุปพฤติกรรมการขับขี่</h1>
                <div style="font-size: 24px; color: #64748B;">Fleet Safety & Telematics Analysis Report</div>
                <div style="margin-top: 20px; font-size: 18px;">ประจำวันที่: ${today}</div>
            </div>

            <div class="content">
                <div class="header-banner" style="margin-top: 40px; text-align: center;">บทสรุปผู้บริหาร (Executive Summary)</div>
                
                <div class="grid-2x2">
                <div class="card">
                    <div class="card-title c-blue">Over Speed (ครั้ง)</div>
                    <div class="card-value c-blue">${totalOverSpeed}</div>
                    <div class="card-sub">เหตุการณ์ทั้งหมด</div>
                </div>
                <div class="card">
                    <div class="card-title c-orange">Max Idling (สูงสุด)</div>
                    <div class="card-value c-orange">${Math.round(maxIdleCar.time / 60)}m</div>
                    <div class="card-sub">${maxIdleCar.license}</div>
                </div>
                <div class="card">
                    <div class="card-title c-red">Critical Events</div>
                    <div class="card-value c-red">${criticalEvents.length}</div>
                    <div class="card-sub">เบรก/ออกตัว กระชาก</div>
                </div>
                <div class="card">
                    <div class="card-title c-purple">พื้นที่ห้ามจอด</div>
                    <div class="card-value c-purple">${rawForbidden.length}</div>
                    <div class="card-sub">จำนวนครั้งทั้งหมด</div>
                </div>
                </div>
            </div>
            </div>

            <!-- Page 2: Over Speed -->
            <div class="page">
            <div class="header-banner">1. การใช้ความเร็วเกินกำหนด (Over Speed Analysis)</div>
            <div class="content">
                <h3>Top 5 Over Speed Frequency</h3>
                <div class="chart-container">
                ${topSpeed.map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.count / (topSpeed[0]?.count || 1)) * 100}%; background: #1E40AF;">${item.count}</div>
                    </div>
                    </div>
                `).join('')}
                </div>

                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้ง</th><th>รวมเวลา</th></tr>
                </thead>
                <tbody>
                    ${topSpeed.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>${item.count}</td>
                        <td>${formatSeconds(item.time)}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

            <!-- Page 3: Idling -->
            <div class="page">
            <div class="header-banner">2. การจอดไม่ดับเครื่อง (Idling Analysis)</div>
            <div class="content">
                <h3>Top 5 Idling Duration (Minutes)</h3>
                <div class="chart-container">
                ${topIdle.map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.time / (topIdle[0]?.time || 1)) * 100}%; background: #F59E0B;">${Math.round(item.time / 60)}m</div>
                    </div>
                    </div>
                `).join('')}
                </div>

                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้ง</th><th>รวมเวลา</th></tr>
                </thead>
                <tbody>
                    ${topIdle.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>${item.count}</td>
                        <td>${formatSeconds(item.time)}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

            <!-- Page 4: Critical Events -->
            <div class="page">
            <div class="header-banner">3. เหตุการณ์วิกฤต (Critical Safety Events)</div>
            <div class="content">
                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>รายละเอียด</th><th>ประเภท</th></tr>
                </thead>
                <tbody>
                    ${criticalEvents.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>Speed: ${item.v_start} &#8594; ${item.v_end} km/h</td>
                        <td class="risk-${item.level}">${item.type}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

            <!-- Page 5: Prohibited Parking -->
            <div class="page">
            <div class="header-banner">4. รายงานพื้นที่ห้ามจอด (Prohibited Parking Area Report)</div>
            <div class="content">
                <h3>Top 5 Prohibited Area Duration</h3>
                <div class="chart-container">
                ${topForbiddenChart.map(item => `
                    <div class="bar-row">
                    <div class="bar-label">${item.license}</div>
                    <div class="bar-track">
                        <div class="bar-fill" style="width: ${(item.time / (topForbiddenChart[0]?.time || 1)) * 100}%; background: #9333EA;">${Math.round(item.time / 60)}m</div>
                    </div>
                    </div>
                `).join('')}
                </div>

                <table>
                <thead>
                    <tr><th>No.</th><th>ทะเบียนรถ</th><th>ชื่อสถานี</th><th>รวมเวลา</th></tr>
                </thead>
                <tbody>
                    ${forbiddenList.map((item, idx) => `
                    <tr>
                        <td>${idx + 1}</td>
                        <td>${item.license}</td>
                        <td>${item.station}</td>
                        <td>${item.timeStr}</td>
                    </tr>
                    `).join('')}
                </tbody>
                </table>
            </div>
            </div>

        </body>
        </html>
        `;

        await page.setContent(html, { waitUntil: 'networkidle0' });
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true
        });
        console.log(`   ✅ PDF Generated: ${pdfPath}`);

        // =================================================================
        // STEP 8: Zip & Email
        // =================================================================
        console.log('📧 Step 8: Zipping CSVs & Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        // เลือกเฉพาะ CSV ที่แปลงแล้ว (Converted_...csv)
        const csvsToZip = allFiles.filter(f => f.startsWith('Converted_') && f.endsWith('.csv'));

        if (csvsToZip.length > 0 || fs.existsSync(pdfPath)) {
            const zipName = `DTC_Report_Data_${today.replace(/ /g, '_')}.zip`;
            const zipPath = path.join(downloadPath, zipName);
            
            if(csvsToZip.length > 0) {
                await zipFiles(downloadPath, zipPath, csvsToZip);
            }

            const attachments = [];
            if (fs.existsSync(zipPath)) attachments.push({ filename: zipName, path: zipPath });
            if (fs.existsSync(pdfPath)) attachments.push({ filename: 'Fleet_Safety_Analysis_Report.pdf', path: pdfPath });

            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงานสรุปพฤติกรรมการขับขี่ (Fleet Safety Report) - ${today}`,
                text: `เรียน ผู้เกี่ยวข้อง\n\nระบบส่งรายงานประจำวัน (06:00 - 18:00) ดังแนบ:\n1. ไฟล์ข้อมูลดิบ CSV (อยู่ใน Zip)\n2. ไฟล์ PDF สรุปภาพรวม\n\nขอบคุณครับ\nDTC Automation Bot`,
                attachments: attachments
            });
            console.log(`   ✅ Email Sent Successfully!`);
        } else {
            console.warn('⚠️ No files to send!');
        }

        console.log('🧹 Cleanup...');
        // fs.rmSync(downloadPath, { recursive: true, force: true });
        console.log('   ✅ Cleanup Complete.');

    } catch (err) {
        console.error('❌ Fatal Error:', err);
        await page.screenshot({ path: path.join(downloadPath, 'fatal_error.png') });
        process.exit(1);
    } finally {
        await browser.close();
        console.log('🏁 Browser Closed.');
    }
})();
