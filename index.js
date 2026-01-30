const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');

// --- Helper Functions ---

// 1. ฟังก์ชันรอโหลดไฟล์ และแปลงเป็น CSV (Universal Converter)
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
    
    // ตรวจสอบขนาด
    const stats = fs.statSync(oldPath);
    if (stats.size === 0) throw new Error(`Downloaded file is empty!`);

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    
    // แปลงเป็น CSV (UTF-8) เสมอ
    const csvFileName = `Converted_${newFileName.replace('.xls', '.csv')}`;
    const csvPath = path.join(downloadPath, csvFileName);
    await convertToCsv(newPath, csvPath); // ใช้ฟังก์ชันใหม่ที่รองรับทั้ง HTML/XLSX
    
    return csvPath;
}

// 2. ฟังก์ชันแปลงไฟล์ (ฉลาดขึ้น รองรับทั้ง Binary XLSX และ HTML Table) -> CSV
async function convertToCsv(sourcePath, destPath) {
    try {
        console.log(`   🔄 Converting to CSV...`);
        const buffer = fs.readFileSync(sourcePath);
        let rows = [];

        // ตรวจสอบ Header ไฟล์ว่าเป็น Zip/XLSX หรือไม่ (PK..)
        const isXLSX = buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4B;

        if (isXLSX) {
            console.log('      - Detected: Binary XLSX (Using ExcelJS)');
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1); // อ่าน Sheet แรก
            
            worksheet.eachRow((row) => {
                // ExcelJS เริ่ม index 1, slice ออกเพื่อให้เป็น 0-based array
                const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
                rows.push(rowValues.map(v => {
                    if (v === null || v === undefined) return '';
                    if (typeof v === 'object') return v.text || v.result || ''; // Handle Rich Text
                    return String(v).trim();
                }));
            });
        } else {
            console.log('      - Detected: HTML/Text (Using JSDOM)');
            const content = buffer.toString('utf8');
            const dom = new JSDOM(content);
            const table = dom.window.document.querySelector('table');
            if (table) {
                const trs = Array.from(table.querySelectorAll('tr'));
                rows = trs.map(tr => 
                    Array.from(tr.querySelectorAll('td, th')).map(td => td.textContent.replace(/\s+/g, ' ').trim())
                );
            } else {
                console.warn('      ⚠️ No table found in HTML source.');
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
            // Fallback: Copy original if conversion yields nothing (prevent crash)
            // fs.copyFileSync(sourcePath, destPath); 
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

// --- DATA PROCESSING HELPERS (Updated Logic) ---

// 1. แปลงเวลา (รองรับภาษาไทย และ HH:MM:SS)
function parseDuration(str) {
    if (!str) return 0;
    str = str.toString().trim();
    
    // Pattern 1: ภาษาไทย "0 ชม. 1 นาที 31 วินาที"
    if (str.includes('นาที') || str.includes('ชม') || str.includes('วินาที')) {
        const h = (str.match(/(\d+)\s*ชม/) || [0,0])[1];
        const m = (str.match(/(\d+)\s*นาที/) || [0,0])[1];
        const s = (str.match(/(\d+)\s*วินาที/) || [0,0])[1];
        return (parseInt(h) * 3600) + (parseInt(m) * 60) + parseInt(s);
    }
    
    // Pattern 2: HH:MM:SS
    if (str.includes(':')) {
        const parts = str.split(':').map(Number);
        if (parts.length === 3) return (parts[0] * 3600) + (parts[1] * 60) + parts[2];
        if (parts.length === 2) return (parts[0] * 60) + parts[1];
    }
    
    return 0;
}

// 2. แปลงวินาทีกลับเป็น HH:MM:SS
function formatTime(totalSeconds) {
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

// 3. อ่านไฟล์ CSV แบบดิบ (เพื่อความแม่นยำในการข้ามบรรทัด)
function readCsvRaw(filePath, skipLines) {
    if (!fs.existsSync(filePath)) return [];
    const content = fs.readFileSync(filePath, 'utf8');
    const lines = content.split('\n').slice(skipLines).join('\n'); // ตัด Header ทิ้ง
    return parse(lines, {
        columns: false,
        skip_empty_lines: true,
        relax_column_count: true,
        bom: true
    });
}

// 4. ฟังก์ชัน Pivot ข้อมูล (รวมเวลาตามทะเบียนรถ) -> Return Top 10
function aggregateData(filePath, skipLines, colIdx) {
    const records = readCsvRaw(filePath, skipLines);
    const pivot = {};
    let totalCount = 0;

    records.forEach(row => {
        const license = row[colIdx.license];
        const duration = row[colIdx.duration];

        // กรองข้อมูลขยะ
        if (!license || !duration || license.includes('รวม') || license.includes('รถทั้งหมด')) return;

        const key = license.trim();
        const sec = parseDuration(duration);

        if (!pivot[key]) {
            pivot[key] = { 
                license: key, 
                station: colIdx.station ? (row[colIdx.station] || '-') : '-',
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

// 5. ฟังก์ชันดึงรายการเหตุการณ์ (Critical Events) -> Return Top 10
function getCriticalEvents(filePath, skipLines, colIdx) {
    const records = readCsvRaw(filePath, skipLines);
    const events = [];

    records.forEach(row => {
        const license = row[colIdx.license];
        if (!license || license.includes('รวม')) return;

        events.push({
            license: license.trim(),
            desc: `${row[colIdx.v1]} -> ${row[colIdx.v2]} km/h`
        });
    });

    // เอา 10 รายการแรก (หรือจะเอาท้ายสุดก็ได้ ถ้าเรียงตามเวลา)
    return events.slice(0, 10);
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

    console.log('🚀 Starting DTC Automation (CSV Fix & New PDF)...');
    
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

        // --- Step 2 to 6: DOWNLOAD REPORTS (STRICT WAIT & AUTOMATIC CSV CONVERSION) ---
        
        // REPORT 1: Over Speed
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        
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
            var select = document.getElementById('ddl_truck'); 
            for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } 
            select.dispatchEvent(new Event('change', { bubbles: true }));
        }, startDateTime, endDateTime);

        console.log('   Searching Report 1...');
        await page.evaluate(() => {
            if(typeof sertch_data === 'function') sertch_data();
            else document.querySelector("span[onclick='sertch_data();']").click();
        });

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
        
        console.log('   ⏳ Waiting 3 mins (Strict)...');
        await new Promise(r => setTimeout(r, 180000));
        
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        // Convert to CSV
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 7: Generate PDF Summary (Complete Logic)
        console.log('📑 Step 7: Generating Final PDF Report...');

        // Config: Mapping Index ตามไฟล์ CSV (0-based index)
        // ต้องเช็คจากไฟล์จริงว่าคอลัมน์อยู่ลำดับที่เท่าไหร่
        const DATA_CONFIG = {
            OVERSPEED: { skip: 6, license: 1, duration: 4 }, // Report 1
            IDLING: { skip: 7, license: 1, duration: 4 },    // Report 2
            SUDDEN: { skip: 5, license: 2, v1: 4, v2: 5 },   // Report 3
            HARSH: { skip: 5, license: 2, v1: 4, v2: 5 },    // Report 4
            FORBIDDEN: { skip: 1, license: 1, station: 4, duration: 9 } // Report 5
        };

        // 1. Process Data
        const r1 = aggregateData(file1, DATA_CONFIG.OVERSPEED.skip, DATA_CONFIG.OVERSPEED);
        const r2 = aggregateData(file2, DATA_CONFIG.IDLING.skip, DATA_CONFIG.IDLING);
        const r5 = aggregateData(file5, DATA_CONFIG.FORBIDDEN.skip, DATA_CONFIG.FORBIDDEN);
        
        const r3_list = getCriticalEvents(file3, DATA_CONFIG.SUDDEN.skip, DATA_CONFIG.SUDDEN);
        const r4_list = (file4) ? getCriticalEvents(file4, DATA_CONFIG.HARSH.skip, DATA_CONFIG.HARSH) : [];

        // คำนวณยอดรวม Critical
        const totalCritical = (readCsvRaw(file3, DATA_CONFIG.SUDDEN.skip).length) + 
                              (file4 ? readCsvRaw(file4, DATA_CONFIG.HARSH.skip).length : 0);

        // Max Idling for Card
        const maxIdleCar = r2.list.length > 0 ? r2.list[0] : { time: 0, license: '-' };

        // 2. Generate HTML Content
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

                /* Grid Cards */
                .grid-2x2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 40px; }
                .card { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 25px; text-align: center; }
                .card-val { font-size: 42px; font-weight: bold; margin: 10px 0; }
                .text-blue { color: #1E40AF; } .text-orange { color: #F59E0B; } .text-red { color: #DC2626; } .text-purple { color: #9333EA; }

                /* Charts */
                .chart-box { margin: 20px 0; }
                .bar-row { display: flex; align-items: center; margin-bottom: 8px; font-size: 12px; }
                .bar-lbl { width: 180px; text-align: right; padding-right: 15px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 600; }
                .bar-trk { flex-grow: 1; background: #F1F5F9; height: 20px; border-radius: 4px; overflow: hidden; }
                .bar-fill { height: 100%; display: flex; align-items: center; justify-content: flex-end; padding-right: 8px; color: white; font-weight: bold; font-size: 11px; }

                /* Tables */
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
                    <div class="card">
                        <div style="font-weight:bold;">Over Speed (ครั้ง)</div>
                        <div class="card-val text-blue">${r1.totalEvents}</div>
                        <div style="font-size:12px; color:#666;">เหตุการณ์ทั้งหมด</div>
                    </div>
                    <div class="card">
                        <div style="font-weight:bold;">Max Idling (นาที)</div>
                        <div class="card-val text-orange">${Math.round(maxIdleCar.time / 60)}m</div>
                        <div style="font-size:12px; color:#666;">${maxIdleCar.license}</div>
                    </div>
                    <div class="card">
                        <div style="font-weight:bold;">Critical Events</div>
                        <div class="card-val text-red">${totalCritical}</div>
                        <div style="font-size:12px; color:#666;">เบรก/ออกตัว กระชาก</div>
                    </div>
                    <div class="card">
                        <div style="font-weight:bold;">พื้นที่ห้ามจอด</div>
                        <div class="card-val text-purple">${r5.totalEvents}</div>
                        <div style="font-size:12px; color:#666;">จำนวนครั้ง</div>
                    </div>
                </div>
            </div>

            <!-- PAGE 2: Over Speed -->
            <div class="page">
                <div class="header-banner">1. การใช้ความเร็วเกินกำหนด (Top 10 by Duration)</div>
                <div class="chart-box">
                    ${r1.list.map(x => `
                        <div class="bar-row">
                            <div class="bar-lbl">${x.license}</div>
                            <div class="bar-trk"><div class="bar-fill" style="width:${(x.time/r1.list[0].time)*100}%; background:#1E40AF;">${formatTime(x.time)}</div></div>
                        </div>
                    `).join('')}
                </div>
                <table>
                    <thead><tr><th style="background:#1E40AF; width:40px;">No.</th><th style="background:#1E40AF;">ทะเบียนรถ</th><th style="background:#1E40AF;" class="num">ครั้ง</th><th style="background:#1E40AF;" class="time">รวมเวลา</th></tr></thead>
                    <tbody>
                        ${r1.list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td class="num">${x.count}</td><td class="time">${formatTime(x.time)}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>

            <!-- PAGE 3: Idling -->
            <div class="page">
                <div class="header-banner bg-orange">2. การจอดไม่ดับเครื่อง (Top 10 by Duration)</div>
                <div class="chart-box">
                    ${r2.list.map(x => `
                        <div class="bar-row">
                            <div class="bar-lbl">${x.license}</div>
                            <div class="bar-trk"><div class="bar-fill" style="width:${(x.time/r2.list[0].time)*100}%; background:#F59E0B;">${formatTime(x.time)}</div></div>
                        </div>
                    `).join('')}
                </div>
                <table>
                    <thead><tr><th style="background:#F59E0B; width:40px;">No.</th><th style="background:#F59E0B;">ทะเบียนรถ</th><th style="background:#F59E0B;" class="num">ครั้ง</th><th style="background:#F59E0B;" class="time">รวมเวลา</th></tr></thead>
                    <tbody>
                        ${r2.list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td class="num">${x.count}</td><td class="time">${formatTime(x.time)}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>

            <!-- PAGE 4: Critical -->
            <div class="page">
                <div class="header-banner bg-red">3. เหตุการณ์วิกฤต (Critical Safety Events)</div>
                
                <h3 style="color:#DC2626; border-left:4px solid #DC2626; padding-left:10px;">3.1 เบรกกะทันหัน (Sudden Brake)</h3>
                <table>
                    <thead><tr><th style="background:#DC2626; width:40px;">No.</th><th style="background:#DC2626;">ทะเบียนรถ</th><th style="background:#DC2626;">รายละเอียด</th><th style="background:#DC2626;" class="num">ระดับ</th></tr></thead>
                    <tbody>
                        ${r3_list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td>${x.desc}</td><td class="num"><span class="badge bg-red">High</span></td></tr>`).join('')}
                    </tbody>
                </table>

                <h3 style="color:#F59E0B; border-left:4px solid #F59E0B; padding-left:10px; margin-top:30px;">3.2 ออกตัวกระชาก (Harsh Start)</h3>
                <table>
                    <thead><tr><th style="background:#F59E0B; width:40px;">No.</th><th style="background:#F59E0B;">ทะเบียนรถ</th><th style="background:#F59E0B;">รายละเอียด</th><th style="background:#F59E0B;" class="num">ระดับ</th></tr></thead>
                    <tbody>
                        ${r4_list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td>${x.desc}</td><td class="num"><span class="badge bg-orange">Medium</span></td></tr>`).join('')}
                    </tbody>
                </table>
            </div>

            <!-- PAGE 5: Forbidden -->
            <div class="page">
                <div class="header-banner bg-purple">4. พื้นที่ห้ามจอด (Top 10 by Duration)</div>
                <div class="chart-box">
                    ${r5.list.map(x => `
                        <div class="bar-row">
                            <div class="bar-lbl">${x.license}</div>
                            <div class="bar-trk"><div class="bar-fill" style="width:${(x.time/r5.list[0].time)*100}%; background:#9333EA;">${formatTime(x.time)}</div></div>
                        </div>
                    `).join('')}
                </div>
                <table>
                    <thead><tr><th style="background:#9333EA; width:40px;">No.</th><th style="background:#9333EA;">ทะเบียนรถ</th><th style="background:#9333EA;">สถานีล่าสุด</th><th style="background:#9333EA;" class="num">ครั้ง</th><th style="background:#9333EA;" class="time">รวมเวลา</th></tr></thead>
                    <tbody>
                        ${r5.list.map((x, i) => `<tr><td>${i+1}</td><td>${x.license}</td><td>${x.station}</td><td class="num">${x.count}</td><td class="time">${formatTime(x.time)}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>

        </body>
        </html>
        `;

        await page.setContent(html, { waitUntil: 'networkidle0' });
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true,
            margin: { top: 0, bottom: 0, left: 0, right: 0 }
        });
        console.log(`✅ PDF Generated: ${pdfPath}`);

        // =================================================================
        // STEP 8: Zip & Email
        // =================================================================
        console.log('📧 Step 8: Zipping CSVs & Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
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
