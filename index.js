/**
 * DTC Automation Script
 * Version: 4.3.1 (Refactored & Cleaned)
 * Features: 
 * - Strict Hard Wait
 * - Robust XLSX -> CSV Conversion
 * - PDF Generation using user-provided logic
 * - Auto Zip & Email
 */

const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');

// =============================================================================
// SECTION 1: HELPER FUNCTIONS (ฟังก์ชันช่วยทำงานต่างๆ)
// =============================================================================

// 1.1 ฟังก์ชันสำหรับดึงวันที่ปัจจุบัน (DD/MM/YYYY)
function getTodayFormatted() {
    const date = new Date();
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

// 1.2 ฟังก์ชันรอโหลดไฟล์ และเปลี่ยนชื่อไฟล์
async function waitForDownloadAndRename(downloadPath, newFileName, maxWaitMs = 300000) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;
    const checkInterval = 10000; 
    let waittime = 0;

    // วนลูปรอจนกว่าจะเจอไฟล์
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

    await new Promise(resolve => setTimeout(resolve, 10000)); // รอให้เขียนไฟล์เสร็จสมบูรณ์

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

// 1.3 ฟังก์ชันแปลงไฟล์ Excel/HTML Table เป็น CSV
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
                const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
                rows.push(rowValues.map(v => {
                    if (v === null || v === undefined) return '';
                    if (typeof v === 'object') return v.text || v.result || ''; 
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

// 1.4 ฟังก์ชันรอตารางข้อมูลบนหน้าเว็บ
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

// 1.5 ฟังก์ชัน Zip ไฟล์
function zipFiles(sourceDir, outPath, filesToZip) {
    return new Promise((resolve, reject) => {
        const output = fs.createWriteStream(outPath);
        const archive = archiver('zip', { zlib: { level: 9 } });

        output.on('close', () => {
            console.log(`   ✅ Zip created: ${path.basename(outPath)} (${archive.pointer()} bytes)`);
            resolve();
        });

        archive.on('error', (err) => reject(err));

        archive.pipe(output);
        filesToZip.forEach(file => {
            archive.file(path.join(sourceDir, file), { name: file });
        });
        archive.finalize();
    });
}

// 1.6 ฟังก์ชันแปลงเวลา (Thai Text -> Seconds)
function parseThaiDurationToSeconds(str) {
    if (!str || typeof str !== 'string') return 0;
    let seconds = 0;
    const hourMatch = str.match(/(\d+)\s*ชม\./);
    const minMatch = str.match(/(\d+)\s*นาที/);
    const secMatch = str.match(/(\d+)\s*วินาที/);

    if (hourMatch) seconds += parseInt(hourMatch[1]) * 3600;
    if (minMatch) seconds += parseInt(minMatch[1]) * 60;
    if (secMatch) seconds += parseInt(secMatch[1]);
    return seconds;
}

// 1.7 ฟังก์ชันแปลงเวลา (HH:mm:ss -> Seconds)
function parseColonDurationToSeconds(str) {
    if (!str || typeof str !== 'string') return 0;
    const parts = str.split(':').map(Number);
    if (parts.length !== 3) return 0;
    return (parts[0] * 3600) + (parts[1] * 60) + parts[2];
}

// 1.8 ฟังก์ชันแปลงเวลา Forbidden Parking (Day:Hour:Min -> Seconds)
function parseForbiddenDurationToSeconds(str) {
    if (!str || typeof str !== 'string') return 0;
    const parts = str.split(':').map(Number);
    if (parts.length !== 3) return 0;
    return (parts[0] * 86400) + (parts[1] * 3600) + (parts[2] * 60);
}

// 1.9 ฟังก์ชันแปลงวินาทีกลับเป็นข้อความสวยๆ
function formatSecondsToText(totalSeconds) {
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    
    if (h > 0) return `${h} ชม. ${m} น.`;
    if (m > 0) return `${m} น. ${s} วิ.`;
    return `${s} วิ.`;
}

// 1.10 ฟังก์ชันอ่าน CSV แบบข้าม Header ขยะ (หาคำว่า "ลำดับ")
function readCleanCSV(filePath) {
    if (!fs.existsSync(filePath)) return [];
    
    const fileContent = fs.readFileSync(filePath, 'utf8');
    const lines = fileContent.split('\n');
    
    let headerIndex = -1;
    for (let i = 0; i < Math.min(lines.length, 20); i++) {
        if (lines[i].includes('ลำดับ') && (lines[i].includes('ชื่อรถ') || lines[i].includes('ทะเบียนรถ'))) {
            headerIndex = i;
            break;
        }
    }

    if (headerIndex === -1) {
        console.warn(`⚠️ Warning: Could not find valid header in ${path.basename(filePath)}`);
        return [];
    }

    const cleanCSVContent = lines.slice(headerIndex).join('\n');
    
    try {
        return parse(cleanCSVContent, {
            columns: true,
            skip_empty_lines: true,
            relax_quotes: true
        });
    } catch (e) {
        console.error(`❌ Error parsing CSV ${path.basename(filePath)}:`, e.message);
        return [];
    }
}

// =============================================================================
// SECTION 2: MAIN SCRIPT EXECUTION
// =============================================================================

(async () => {
    // 2.1 Setup & Config
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing Secrets (DTC_USERNAME/PASSWORD).');
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

    const todayStr = getTodayFormatted(); // Format: DD/MM/YYYY
    const startDateTime = `${todayStr} 06:00`;
    const endDateTime = `${todayStr} 18:00`;
    
    try {
        // 2.2 Login
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
        console.log(`🕒 Global Time Settings: ${startDateTime} to ${endDateTime}`);

        // --- 2.3 DOWNLOAD REPORTS LOOP ---
        
        // REPORT 1: Over Speed
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
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

        console.log('   ⏳ Waiting 5 mins...');
        await new Promise(resolve => setTimeout(resolve, 300000));
        
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        console.log('   Exporting Report 1...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

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
        await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

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
        await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // REPORT 4: Harsh Start
        console.log('📊 Processing Report 4: Harsh Start...');
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#date9', { visible: true, timeout: 60000 });
            await new Promise(r => setTimeout(r, 10000));
            
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
            await waitForDownloadAndRename(downloadPath, 'Report4_HarshStart.xls');
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
            
            // 2. พื้นที่ห้ามเข้า
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
            for(var s of allSelects) { 
                for(var i=0; i<s.options.length; i++) { 
                    if(s.options[i].text.includes('สถานีทั้งหมด')) { 
                        s.value = s.options[i].value; 
                        s.dispatchEvent(new Event('change', { bubbles: true })); 
                        break; 
                    } 
                } 
            }
        });
        
        await page.click('td:nth-of-type(7) > span');
        console.log('   ⏳ Waiting 3 mins (Strict)...');
        await new Promise(r => setTimeout(r, 180000));
        
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 2.4: Generate PDF Summary
        // =================================================================
        console.log('7. Processing Data & Generating PDF Report...');

        // 1. อ่านข้อมูลจาก CSV
        const rawOverSpeed = readCleanCSV(path.join(downloadPath, 'Converted_Report1_OverSpeed.csv'));
        const rawIdling = readCleanCSV(path.join(downloadPath, 'Converted_Report2_Idling.csv'));
        const rawSudden = readCleanCSV(path.join(downloadPath, 'Converted_Report3_SuddenBrake.csv'));
        const rawHarsh = readCleanCSV(path.join(downloadPath, 'Converted_Report4_HarshStart.csv'));
        const rawForbidden = readCleanCSV(path.join(downloadPath, 'Converted_Report5_ForbiddenParking.csv'));

        // 2. ประมวลผล - A. Over Speed Analysis
        const overSpeedMap = new Map();
        rawOverSpeed.forEach(row => {
            if (!row['ชื่อรถ'] || row['ชื่อรถ'].trim() === 'รวม' || !row['ลำดับ']) return;
            const carId = row['ชื่อรถ'] || row['ทะเบียนรถ'];
            const duration = parseThaiDurationToSeconds(row['รวมเวลา']);
            
            if (!overSpeedMap.has(carId)) overSpeedMap.set(carId, { count: 0, duration: 0 });
            const data = overSpeedMap.get(carId);
            data.count += 1;
            data.duration += duration;
        });
        const topOverSpeed = Array.from(overSpeedMap.entries())
            .map(([car, data]) => ({ car, ...data }))
            .sort((a, b) => b.duration - a.duration)
            .slice(0, 10);

        // 3. ประมวลผล - B. Idling Analysis
        const idlingMap = new Map();
        rawIdling.forEach(row => {
            if (!row['ชื่อรถ'] || row['ชื่อรถ'].trim() === 'รวม' || !row['ลำดับ']) return;
            const carId = row['ชื่อรถ'];
            const duration = parseColonDurationToSeconds(row['รวมเวลา']);
            
            if (!idlingMap.has(carId)) idlingMap.set(carId, { count: 0, duration: 0 });
            const data = idlingMap.get(carId);
            data.count += 1;
            data.duration += duration;
        });
        const topIdling = Array.from(idlingMap.entries())
            .map(([car, data]) => ({ car, ...data }))
            .sort((a, b) => b.duration - a.duration)
            .slice(0, 10);

        // 4. ประมวลผล - C. Forbidden Parking Analysis
        const forbiddenMap = new Map();
        rawForbidden.forEach(row => {
            if (!row['ทะเบียนรถ'] || row['ทะเบียนรถ'].trim() === 'รวม' || !row['ลำดับ']) return;
            const carId = row['ทะเบียนรถ'];
            const location = row['ชื่อสถานี'] || '-';
            const rawTime = row['รวมเวลาในสถานี(วัน:ชั่วโมง:นาที)'];
            const duration = parseForbiddenDurationToSeconds(rawTime);

            if (!forbiddenMap.has(carId)) forbiddenMap.set(carId, { count: 0, duration: 0, location: location });
            const data = forbiddenMap.get(carId);
            data.count += 1;
            data.duration += duration;
        });
        const topForbidden = Array.from(forbiddenMap.entries())
            .map(([car, data]) => ({ car, ...data }))
            .sort((a, b) => b.duration - a.duration)
            .slice(0, 10);

        // 5. ประมวลผล - D. Critical Events
        const listSudden = rawSudden.filter(row => row['ลำดับ'] && row['ทะเบียนรถ'] && row['ทะเบียนรถ'] !== 'รวม');
        const listHarsh = rawHarsh.filter(row => row['ลำดับ'] && row['ทะเบียนรถ'] && row['ทะเบียนรถ'] !== 'รวม');

        // 6. Stats Summary
        const totalOverSpeedEvents = rawOverSpeed.filter(r => r['ลำดับ']).length;
        const totalIdlingEvents = rawIdling.filter(r => r['ลำดับ']).length;
        const totalForbiddenEvents = rawForbidden.filter(r => r['ลำดับ']).length;
        const totalCriticalEvents = listSudden.length + listHarsh.length;

        // 7. สร้าง HTML Content
        const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body { font-family: 'Sarabun', sans-serif; padding: 20px; color: #333; }
                h1, h2 { color: #004085; border-bottom: 2px solid #004085; padding-bottom: 5px; }
                h3 { color: #555; margin-top: 20px; }
                .summary-box { display: flex; justify-content: space-between; margin-bottom: 30px; }
                .card { background: #f8f9fa; padding: 15px; border-radius: 8px; width: 22%; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
                .card h4 { margin: 0; color: #666; font-size: 14px; }
                .card .val { font-size: 24px; font-weight: bold; color: #0056b3; margin-top: 5px; }
                table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #004085; color: white; text-align: center; }
                td { text-align: center; }
                .text-left { text-align: left; }
                .warning { color: #d9534f; font-weight: bold; }
                .page-break { page-break-before: always; }
            </style>
        </head>
        <body>
            <div style="text-align: center; margin-bottom: 30px;">
                <h1>รายงานสรุปพฤติกรรมการขับขี่ (Fleet Safety Report)</h1>
                <p>ประจำวันที่: ${todayStr}</p>
            </div>

            <!-- Executive Summary -->
            <h2>บทสรุปผู้บริหาร (Executive Summary)</h2>
            <div class="summary-box">
                <div class="card">
                    <h4>Over Speed (ครั้ง)</h4>
                    <div class="val">${totalOverSpeedEvents}</div>
                </div>
                <div class="card">
                    <h4>Idling (ครั้ง)</h4>
                    <div class="val">${totalIdlingEvents}</div>
                </div>
                <div class="card">
                    <h4>Critical Events</h4>
                    <div class="val">${totalCriticalEvents}</div>
                    <small>(เบรก/ออกตัว กระชาก)</small>
                </div>
                <div class="card">
                    <h4>พื้นที่ห้ามจอด (ครั้ง)</h4>
                    <div class="val">${totalForbiddenEvents}</div>
                </div>
            </div>

            <!-- 1. Over Speed -->
            <h3>1. การใช้ความเร็วเกินกำหนด (Top 10 Over Speed by Duration)</h3>
            <table>
                <tr>
                    <th style="width: 10%">No.</th>
                    <th style="width: 50%">ทะเบียนรถ/ชื่อรถ</th>
                    <th style="width: 20%">จำนวนครั้ง</th>
                    <th style="width: 20%">รวมเวลา</th>
                </tr>
                ${topOverSpeed.length > 0 ? topOverSpeed.map((item, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td class="text-left">${item.car}</td>
                    <td>${item.count}</td>
                    <td class="warning">${formatSecondsToText(item.duration)}</td>
                </tr>`).join('') : '<tr><td colspan="4">ไม่มีข้อมูลความเร็วเกินกำหนด</td></tr>'}
            </table>

            <!-- 2. Idling -->
            <h3>2. การจอดไม่ดับเครื่อง (Top 10 Idling by Duration)</h3>
            <table>
                <tr>
                    <th style="width: 10%">No.</th>
                    <th style="width: 50%">ทะเบียนรถ/ชื่อรถ</th>
                    <th style="width: 20%">จำนวนครั้ง</th>
                    <th style="width: 20%">รวมเวลา</th>
                </tr>
                ${topIdling.length > 0 ? topIdling.map((item, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td class="text-left">${item.car}</td>
                    <td>${item.count}</td>
                    <td class="warning">${formatSecondsToText(item.duration)}</td>
                </tr>`).join('') : '<tr><td colspan="4">ไม่มีข้อมูลจอดไม่ดับเครื่อง</td></tr>'}
            </table>

            <div class="page-break"></div>

            <!-- 3. Critical Events -->
            <h2>3. เหตุการณ์วิกฤต (Critical Safety Events)</h2>
            
            <h3>3.1 Sudden Brake (เบรกกะทันหัน)</h3>
            <table>
                <tr>
                    <th style="width: 10%">No.</th>
                    <th style="width: 30%">ทะเบียนรถ</th>
                    <th style="width: 20%">เวลาที่เกิดเหตุ</th>
                    <th style="width: 40%">สถานที่ (ตำบล/อำเภอ)</th>
                </tr>
                ${listSudden.length > 0 ? listSudden.map((row, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td class="text-left">${row['ชื่อรถ'] || row['ทะเบียนรถ']}</td>
                    <td>${row['วันที่บันทึก'] ? row['วันที่บันทึก'].split(' ')[1] : '-'}</td>
                    <td class="text-left">${row['ตำบล'] || '-'} ${row['อำเภอ'] || '-'}</td>
                </tr>`).join('') : '<tr><td colspan="4">ไม่มีข้อมูลเบรกกะทันหัน</td></tr>'}
            </table>

            <h3>3.2 Harsh Start (ออกตัวกระชาก)</h3>
            <table>
                <tr>
                    <th style="width: 10%">No.</th>
                    <th style="width: 30%">ทะเบียนรถ</th>
                    <th style="width: 20%">เวลาที่เกิดเหตุ</th>
                    <th style="width: 40%">สถานที่ (ตำบล/อำเภอ)</th>
                </tr>
                ${listHarsh.length > 0 ? listHarsh.map((row, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td class="text-left">${row['ชื่อรถ'] || row['ทะเบียนรถ']}</td>
                    <td>${row['วันที่บันทึก'] ? row['วันที่บันทึก'].split(' ')[1] : '-'}</td>
                    <td class="text-left">${row['ตำบล'] || '-'} ${row['อำเภอ'] || '-'}</td>
                </tr>`).join('') : '<tr><td colspan="4">ไม่มีข้อมูลออกตัวกระชาก</td></tr>'}
            </table>

            <!-- 4. Forbidden Parking -->
            <h3>4. รายงานพื้นที่ห้ามจอด (Prohibited Parking Area Report)</h3>
            <table>
                <tr>
                    <th style="width: 10%">No.</th>
                    <th style="width: 30%">ทะเบียนรถ</th>
                    <th style="width: 30%">สถานีห้ามจอด</th>
                    <th style="width: 15%">จำนวนครั้ง</th>
                    <th style="width: 15%">รวมเวลา</th>
                </tr>
                ${topForbidden.length > 0 ? topForbidden.map((item, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td class="text-left">${item.car}</td>
                    <td class="text-left">${item.location}</td>
                    <td>${item.count}</td>
                    <td class="warning">${formatSecondsToText(item.duration)}</td>
                </tr>`).join('') : '<tr><td colspan="5">ไม่มีข้อมูลจอดในพื้นที่ห้ามจอด</td></tr>'}
            </table>
        </body>
        </html>
        `;

        // Generate PDF
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true,
            margin: { top: '20px', bottom: '20px', left: '20px', right: '20px' }
        });
        console.log(`   ✅ PDF Report Generated: ${pdfPath}`);

        // =================================================================
        // STEP 2.5: Zip & Email
        // =================================================================
        console.log('📧 Step 8: Zipping CSVs & Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        const csvsToZip = allFiles.filter(f => f.startsWith('Converted_') && f.endsWith('.csv'));

        if (csvsToZip.length > 0 || fs.existsSync(pdfPath)) {
            const zipName = `DTC_Report_Data_${todayStr.replace(/\//g, '-')}.zip`;
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
                subject: `รายงานสรุปพฤติกรรมการขับขี่ (Fleet Safety Report) - ${todayStr}`,
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
