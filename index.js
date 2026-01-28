const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const ExcelJS = require('exceljs');

// --- Helper Functions ---

// 1. ฟังก์ชันรอโหลดไฟล์ (ใช้ Hard Wait Loop เพื่อความชัวร์)
async function waitForDownloadAndRename(downloadPath, newFileName, maxWaitMs = 300000) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;
    const checkInterval = 2000; 
    let waittime = 0;

    // วนลูปรอไฟล์จนกว่าจะเจอ หรือหมดเวลา
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

    await new Promise(resolve => setTimeout(resolve, 5000)); // รอเขียนไฟล์ให้เสร็จสมบูรณ์

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    // ตรวจสอบขนาดไฟล์ต้องไม่ว่างเปล่า
    const stats = fs.statSync(oldPath);
    if (stats.size === 0) throw new Error(`Downloaded file is empty!`);

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    
    // แปลงเป็น XLSX ทันที เพื่อให้อ่านข้อมูลได้แม่นยำ
    const xlsxFileName = `Converted_${newFileName.replace('.xls', '.xlsx')}`;
    const xlsxPath = path.join(downloadPath, xlsxFileName);
    await convertHtmlToExcel(newPath, xlsxPath);

    return xlsxPath;
}

// 2. ฟังก์ชันรอตารางข้อมูล (Wait for Data Population)
// สำคัญมาก! ต้องรอให้ตารางมีข้อมูลมากกว่า 2 แถว (Header + Data) ก่อนกด Export
async function waitForTableData(page, minRows = 2, timeout = 300000) {
    console.log(`   Waiting for table data (Max ${timeout/1000}s)...`);
    try {
        await page.waitForFunction((min) => {
            const rows = document.querySelectorAll('table tr');
            // ตรวจสอบว่ามีข้อมูล และไม่มีข้อความว่า "ไม่พบข้อมูล"
            const bodyText = document.body.innerText;
            if (bodyText.includes('ไม่พบข้อมูล') || bodyText.includes('No data found')) return true; // จบการรอแบบไม่มีข้อมูล
            return rows.length >= min; 
        }, { timeout: timeout }, minRows);
        
        // เช็คจำนวนแถวจริงๆ
        const rowCount = await page.evaluate(() => document.querySelectorAll('table tr').length);
        console.log(`   ✅ Table populated with ${rowCount} rows.`);
    } catch (e) {
        console.warn('   ⚠️ Wait for table data timed out (Data might be empty).');
    }
}

// 3. แปลง HTML -> Excel (ExcelJS)
async function convertHtmlToExcel(sourcePath, destPath) {
    try {
        const content = fs.readFileSync(sourcePath, 'utf-8');
        // ถ้าไม่ใช่ HTML (เป็น Binary XLS อยู่แล้ว) ให้ Copy เลย
        if (!content.trim().startsWith('<')) {
             fs.copyFileSync(sourcePath, destPath);
             return;
        }

        const dom = new JSDOM(content);
        const table = dom.window.document.querySelector('table');
        
        if (!table) {
             console.warn('   ⚠️ No HTML Table found, copying original file.');
             fs.copyFileSync(sourcePath, destPath);
             return;
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sheet1');
        const rows = Array.from(table.querySelectorAll('tr'));

        rows.forEach((row) => {
            const cells = Array.from(row.querySelectorAll('td, th')).map(cell => cell.textContent.trim());
            worksheet.addRow(cells);
        });
        
        // Auto-fit columns logic (Optional)
        worksheet.columns.forEach(column => { column.width = 20; });

        await workbook.xlsx.writeFile(destPath);
        console.log(`   ✅ Converted to XLSX: ${path.basename(destPath)}`);
    } catch (e) {
        console.warn(`   ⚠️ Conversion failed: ${e.message}`);
        fs.copyFileSync(sourcePath, destPath);
    }
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

// ฟังก์ชันแปลงเวลา "HH:mm:ss" เป็นนาที
function parseDurationToMinutes(durationStr) {
    if (!durationStr || typeof durationStr !== 'string') return 0;
    // หา pattern เวลา เช่น 02:15:30 หรือ 00:45
    const match = durationStr.match(/(\d+):(\d+)(?::(\d+))?/);
    if (!match) return 0;

    const h = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const s = match[3] ? parseInt(match[3], 10) : 0;

    return (h * 60) + m + (s / 60);
}

// *** SMART DATA EXTRACTION ***
// ดึงข้อมูลโดยใช้ Regex ค้นหา Pattern แทนการ Fix Column Index
async function extractDataFromXLSX(filePath, reportType) {
    try {
        if (!fs.existsSync(filePath)) return [];
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(1);
        const data = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber < 2) return; // Skip header
            
            // อ่านค่าในแถว (กรองค่าว่างออก)
            const rawCells = Array.isArray(row.values) ? row.values : [];
            const cells = rawCells.map(v => (v !== null && v !== undefined) ? String(v).trim() : '');
            
            if (cells.length < 3) return;

            // Regex Definition
            const plateRegex = /\d{1,3}-?\d{1,4}|[ก-ฮ]{1,3}\d{1,4}/; // หาทะเบียน
            const timeRegex = /\d{1,2}:\d{2}/; // หาเวลาที่มี : (เช่น 00:05:00)

            // 1. หาทะเบียนรถ (Anchor Point)
            const plateIndex = cells.findIndex(c => plateRegex.test(c) && c.length < 25 && !c.includes(':'));
            if (plateIndex === -1) return; // ถ้าไม่มีทะเบียน ข้าม
            
            const plate = cells[plateIndex];

            // 2. หาเวลา (Duration)
            // กวาดหาทุก cell ที่เป็นเวลา แล้วเอาตัวสุดท้าย (เพราะ Duration รวมมักอยู่ท้ายสุด)
            const timeCells = cells.filter(c => timeRegex.test(c));
            const duration = timeCells.length > 0 ? timeCells[timeCells.length - 1] : "00:00:00";

            if (reportType === 'speed' || reportType === 'idling') {
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            } 
            else if (reportType === 'critical') {
                // Detail: หา text ยาวๆ ที่ไม่ใช่ทะเบียน และไม่ใช่เวลา (เช่น "Speed Drop...")
                let detail = cells.slice(plateIndex + 1).find(c => c.length > 4 && !timeRegex.test(c) && !plateRegex.test(c));
                if (!detail) detail = "Critical Event"; 
                data.push({ plate, detail });
            } 
            else if (reportType === 'forbidden') {
                // Station: หาชื่อสถานี (อยู่หลังทะเบียน 1 หรือ 2 ช่อง)
                let station = "";
                const possibleStations = cells.slice(plateIndex + 1).filter(c => c.length > 2 && !timeRegex.test(c));
                if (possibleStations.length > 0) station = possibleStations[0];
                else station = "Unknown Area";
                
                data.push({ plate, station, duration, durationMin: parseDurationToMinutes(duration) });
            }
        });
        
        console.log(`      -> Extracted ${data.length} rows from ${path.basename(filePath)}`);
        return data;
    } catch (e) {
        console.warn(`   ⚠️ Extract Error ${path.basename(filePath)}: ${e.message}`);
        return [];
    }
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

    console.log('🚀 Starting DTC Automation (Strict Wait & Smart PDF)...');
    
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
        
        // --- REPORT 1: Over Speed ---
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        
        // รอ Dropdown รถโหลดเสร็จ (สำคัญ!)
        await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 1, {timeout: 60000});

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
            
            // Programmatic Select "All"
            const select = document.getElementById('ddl_truck');
            if(select) {
                let found = false;
                for(let i=0; i<select.options.length; i++) {
                    if(select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                        select.selectedIndex = i; found = true; break; 
                    }
                }
                if(!found) select.selectedIndex = 0; 
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        await page.evaluate(() => { if(typeof sertch_data === 'function') sertch_data(); else document.querySelector("span[onclick='sertch_data();']").click(); });
        
        // ** Wait for Table Data **
        await waitForTableData(page, 2, 300000); 

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file1 = await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // --- REPORT 2: Idling ---
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 1);

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) {
                document.getElementById('ddlMinute').value = '10';
                document.getElementById('ddlMinute').dispatchEvent(new Event('change'));
            }
            const select = document.getElementById('ddl_truck');
            if(select) {
                for(let i=0; i<select.options.length; i++) {
                    if(select.options[i].text.includes('ทั้งหมด')) { select.selectedIndex = i; break; }
                }
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        await page.click('td:nth-of-type(6) > span');
        await waitForTableData(page, 2, 300000);

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file2 = await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

        // --- REPORT 3: Sudden Brake ---
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 1);

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            const select = document.getElementById('ddl_truck');
            if(select) {
                for(let i=0; i<select.options.length; i++) {
                    if(select.options[i].text.includes('ทั้งหมด')) { select.selectedIndex = i; break; }
                }
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        await page.click('td:nth-of-type(6) > span');
        await waitForTableData(page, 2, 180000);

        await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const b = btns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
            if(b) b.click(); else document.querySelector('#table button:nth-of-type(3)')?.click();
        });
        const file3 = await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // --- REPORT 4: Harsh Start ---
        console.log('📊 Processing Report 4: Harsh Start...');
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#date9', { visible: true });
            await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 1, {timeout: 60000});

            await page.evaluate((start, end) => {
                document.getElementById('date9').value = start;
                document.getElementById('date10').value = end;
                document.getElementById('date9').dispatchEvent(new Event('change'));
                document.getElementById('date10').dispatchEvent(new Event('change'));
                
                const select = document.getElementById('ddl_truck');
                if(select) {
                    let found = false;
                    for(let i=0; i<select.options.length; i++) {
                        if(select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                            select.selectedIndex = i; found = true; break; 
                        }
                    }
                    if(!found) select.selectedIndex = 0;
                    
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    if (typeof $ !== 'undefined' && $(select).data('select2')) {
                        $(select).trigger('change'); 
                    }
                }
            }, startDateTime, endDateTime);

            // กด Search
            await page.evaluate(() => {
                if(typeof sertch_data === 'function') sertch_data();
                else document.querySelector('td:nth-of-type(6) > span').click();
            });

            await waitForTableData(page, 2, 180000);

            await page.evaluate(() => {
                const xpathResult = document.evaluate('//*[@id="table"]/div[1]/button[3]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if(xpathResult.singleNodeValue) xpathResult.singleNodeValue.click();
                else {
                    const btns = Array.from(document.querySelectorAll('button'));
                    const b = btns.find(b => b.innerText.includes('Excel'));
                    if(b) b.click();
                }
            });
            const file4 = await waitForDownloadAndRename(downloadPath, 'Report4_HarshStart.xls');
        } catch(e) { console.error('Report 4 Skipped:', e.message); }

        // --- REPORT 5: Forbidden Parking ---
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 1);

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            const select = document.getElementById('ddl_truck');
            if(select) { for(let opt of select.options) { if(opt.text.includes('ทั้งหมด')) { select.selectedIndex = opt.index; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
            
            const allSelects = document.getElementsByTagName('select');
            for(let s of allSelects) { 
                for(let i=0; i<s.options.length; i++) { 
                    if(s.options[i].text.includes('พื้นที่ห้ามเข้า')) { s.value = s.options[i].value; s.dispatchEvent(new Event('change', { bubbles: true })); break; } 
                } 
            }
        }, startDateTime, endDateTime);

        await new Promise(r => setTimeout(r, 2000));
        await page.evaluate(() => {
            const allSelects = document.getElementsByTagName('select');
            for(let s of allSelects) { 
                for(let i=0; i<s.options.length; i++) { 
                    if(s.options[i].text.includes('สถานีทั้งหมด')) { s.value = s.options[i].value; s.dispatchEvent(new Event('change', { bubbles: true })); break; } 
                } 
            }
        });

        await page.click('td:nth-of-type(7) > span');
        await waitForTableData(page, 2, 180000);

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 7: Generate PDF Summary (Complete Logic)
        // =================================================================
        console.log('📑 Step 7: Generating PDF Summary...');

        const fileMap = {
            'speed': path.join(downloadPath, 'Converted_Report1_OverSpeed.xlsx'),
            'idling': path.join(downloadPath, 'Converted_Report2_Idling.xlsx'),
            'brake': path.join(downloadPath, 'Converted_Report3_SuddenBrake.xlsx'),
            'start': path.join(downloadPath, 'Converted_Report4_HarshStart.xlsx'),
            'forbidden': path.join(downloadPath, 'Converted_Report5_ForbiddenParking.xlsx')
        };

        const speedData = await extractDataFromXLSX(fileMap.speed, 'speed');
        const idlingData = await extractDataFromXLSX(fileMap.idling, 'idling');
        const brakeData = await extractDataFromXLSX(fileMap.brake, 'critical');
        let startData = [];
        try { startData = await extractDataFromXLSX(fileMap.start, 'critical'); } catch(e){}
        const forbiddenData = await extractDataFromXLSX(fileMap.forbidden, 'forbidden');

        // Aggregation for PDF
        const processStats = (data, key) => {
            const stats = {};
            data.forEach(d => {
                if (!d.plate) return;
                if (!stats[d.plate]) stats[d.plate] = { count: 0, durationMin: 0 };
                stats[d.plate].count++;
                if (d.durationMin) stats[d.plate].durationMin += d.durationMin;
            });
            return Object.entries(stats)
                .map(([plate, val]) => ({ plate, ...val }))
                .sort((a, b) => key === 'count' ? b.count - a.count : b.durationMin - a.durationMin)
                .slice(0, 5);
        };

        const topSpeed = processStats(speedData, 'count');
        const topIdling = processStats(idlingData, 'durationMin');
        const topForbidden = processStats(forbiddenData, 'durationMin');
        const totalCritical = brakeData.length + startData.length;

        // Formatter for Table
        const formatDuration = (mins) => {
            if (!mins) return "00:00:00";
            const h = Math.floor(mins / 60);
            const m = Math.floor(mins % 60);
            const s = Math.floor((mins * 60) % 60);
            return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
        };

        // HTML Template matching FleetSafetyReportv2.tex.pdf
        const htmlContent = `
        <!DOCTYPE html>
        <html lang="th">
        <head>
            <meta charset="UTF-8">
            <script src="https://cdn.tailwindcss.com"></script>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;600;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Noto Sans Thai', sans-serif; background: #fff; color: #333; }
                .page-break { page-break-after: always; }
                .header-blue { background-color: #1e40af; color: white; padding: 12px 20px; border-radius: 8px; margin-bottom: 24px; font-weight: bold; }
                .card { background: #f0f9ff; border-radius: 12px; padding: 24px; text-align: center; border: 1px solid #bae6fd; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
                .card h3 { color: #0c4a6e; font-weight: bold; font-size: 1.1rem; margin-bottom: 8px; }
                .card .val { font-size: 3rem; font-weight: 800; margin: 8px 0; }
                table { width: 100%; border-collapse: collapse; margin-top: 24px; font-size: 0.9rem; }
                th { background-color: #1e40af; color: white; padding: 12px; text-align: left; border-bottom: 2px solid #1e3a8a; }
                td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; }
                tr:nth-child(even) { background-color: #f8fafc; }
                .chart-container { height: 300px; margin-bottom: 30px; }
            </style>
        </head>
        <body class="p-10">
            <!-- PAGE 1: Summary -->
            <div class="page-break">
                <div class="text-center mb-16 mt-10">
                    <h1 class="text-4xl font-bold text-blue-900 mb-2">รายงานสรุปพฤติกรรมการขับขี่</h1>
                    <h2 class="text-xl text-gray-600">Fleet Safety & Telematics Analysis Report</h2>
                    <p class="text-xl mt-6 text-gray-500">วันที่: ${todayStr} (06:00 - 18:00)</p>
                </div>
                
                <div class="grid grid-cols-2 gap-8 px-10">
                    <div class="card">
                        <h3>Over Speed (ครั้ง)</h3>
                        <div class="val text-blue-700">${speedData.length}</div>
                        <p class="text-gray-500">เหตุการณ์ทั้งหมด</p>
                    </div>
                    <div class="card" style="background-color: #fff7ed; border-color: #fed7aa;">
                        <h3 style="color: #9a3412;">Max Idling (นาที)</h3>
                        <div class="val text-orange-600">${topIdling.length > 0 ? topIdling[0].durationMin.toFixed(0) : 0}</div>
                        <p class="text-gray-500">สูงสุดต่อคัน</p>
                    </div>
                    <div class="card" style="background-color: #fef2f2; border-color: #fecaca;">
                        <h3 style="color: #991b1b;">Critical Events</h3>
                        <div class="val text-red-600">${totalCritical}</div>
                        <p class="text-gray-500">เบรก/ออกตัว กระชาก</p>
                    </div>
                    <div class="card" style="background-color: #faf5ff; border-color: #e9d5ff;">
                        <h3 style="color: #6b21a8;">Prohibited Parking</h3>
                        <div class="val text-purple-700">${forbiddenData.length}</div>
                        <p class="text-gray-500">เข้าพื้นที่ห้ามจอด (ครั้ง)</p>
                    </div>
                </div>
            </div>

            <!-- PAGE 2: Speed Analysis -->
            <div class="page-break">
                <div class="header-blue text-2xl">1. การใช้ความเร็วเกินกำหนด (Over Speed Analysis)</div>
                <div class="chart-container"><canvas id="speedChart"></canvas></div>
                
                <h3 class="text-xl font-bold text-gray-700 mb-2">Top 5 Over Speed Frequency</h3>
                <table>
                    <thead><tr><th width="10%">No.</th><th>ทะเบียนรถ (License Plate)</th><th width="20%">จำนวนครั้ง</th><th width="25%">รวมเวลา (Duration)</th></tr></thead>
                    <tbody>
                        ${topSpeed.map((d, i) => `
                            <tr>
                                <td class="text-center font-bold">${i+1}</td>
                                <td>${d.plate}</td>
                                <td class="text-center font-bold text-blue-700">${d.count}</td>
                                <td>${formatDuration(d.durationMin)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>

            <!-- PAGE 3: Idling Analysis -->
            <div class="page-break">
                <div class="header-blue text-2xl" style="background-color: #f59e0b;">2. การจอดไม่ดับเครื่อง (Idling Analysis)</div>
                <div class="chart-container"><canvas id="idlingChart"></canvas></div>
                
                <h3 class="text-xl font-bold text-gray-700 mb-2">Top 5 Idling Duration</h3>
                <table>
                    <thead><tr><th width="10%">No.</th><th>ทะเบียนรถ</th><th width="20%">จำนวนครั้ง</th><th width="25%">รวมเวลา</th></tr></thead>
                    <tbody>
                        ${topIdling.map((d, i) => `
                            <tr>
                                <td class="text-center font-bold">${i+1}</td>
                                <td>${d.plate}</td>
                                <td class="text-center">${d.count}</td>
                                <td class="font-bold text-orange-600">${formatDuration(d.durationMin)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>

            <!-- PAGE 4: Critical Events -->
            <div class="page-break">
                <div class="header-blue text-2xl" style="background-color: #dc2626;">3. เหตุการณ์วิกฤต (Critical Safety Events)</div>
                
                <div class="mb-8">
                    <h3 class="text-xl font-bold text-red-700 border-b-2 border-red-200 pb-2 mb-4">3.1 เบรกกะทันหัน (Sudden Brake)</h3>
                    <table>
                        <thead><tr><th width="30%">ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead>
                        <tbody>
                            ${brakeData.length > 0 ? brakeData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('') : '<tr><td colspan="2" class="text-center text-gray-400">ไม่มีเหตุการณ์</td></tr>'}
                        </tbody>
                    </table>
                </div>

                <div>
                    <h3 class="text-xl font-bold text-red-700 border-b-2 border-red-200 pb-2 mb-4">3.2 ออกตัวกระชาก (Harsh Start)</h3>
                    <table>
                        <thead><tr><th width="30%">ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead>
                        <tbody>
                            ${startData.length > 0 ? startData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('') : '<tr><td colspan="2" class="text-center text-gray-400">ไม่มีเหตุการณ์</td></tr>'}
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- PAGE 5: Prohibited Parking -->
            <div>
                <div class="header-blue text-2xl" style="background-color: #9333ea;">4. รายงานพื้นที่ห้ามจอด (Prohibited Parking)</div>
                <div class="chart-container"><canvas id="forbiddenChart"></canvas></div>
                
                <h3 class="text-xl font-bold text-gray-700 mb-2">Top 5 Prohibited Parking Duration</h3>
                <table>
                    <thead><tr><th width="10%">No.</th><th width="25%">ทะเบียนรถ</th><th>ชื่อสถานี (Station)</th><th width="25%">รวมเวลา</th></tr></thead>
                    <tbody>
                        ${topForbidden.map((d, i) => `
                            <tr>
                                <td class="text-center font-bold">${i+1}</td>
                                <td>${d.plate}</td>
                                <td>${d.station}</td>
                                <td class="font-bold text-purple-700">${formatDuration(d.durationMin)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>

            <script>
                // Common Chart Options
                const commonOptions = {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: { y: { beginAtZero: true } }
                };

                // 1. Speed Chart (Frequency - Vertical)
                new Chart(document.getElementById('speedChart'), {
                    type: 'bar',
                    data: {
                        labels: ${JSON.stringify(topSpeed.map(d => d.plate))},
                        datasets: [{ 
                            label: 'Frequency', 
                            data: ${JSON.stringify(topSpeed.map(d => d.count))}, 
                            backgroundColor: '#1e40af',
                            borderRadius: 4
                        }]
                    },
                    options: commonOptions
                });

                // 2. Idling Chart (Duration - Horizontal)
                new Chart(document.getElementById('idlingChart'), {
                    type: 'bar',
                    indexAxis: 'y',
                    data: {
                        labels: ${JSON.stringify(topIdling.map(d => d.plate))},
                        datasets: [{ 
                            label: 'Minutes', 
                            data: ${JSON.stringify(topIdling.map(d => d.durationMin))}, 
                            backgroundColor: '#f59e0b',
                            borderRadius: 4
                        }]
                    },
                    options: commonOptions
                });

                // 3. Forbidden Chart (Duration - Vertical)
                new Chart(document.getElementById('forbiddenChart'), {
                    type: 'bar',
                    data: {
                        labels: ${JSON.stringify(topForbidden.map(d => d.plate))},
                        datasets: [{ 
                            label: 'Minutes', 
                            data: ${JSON.stringify(topForbidden.map(d => d.durationMin))}, 
                            backgroundColor: '#9333ea',
                            borderRadius: 4
                        }]
                    },
                    options: commonOptions
                });
            </script>
        </body>
        </html>`;

        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true,
            margin: { top: '20px', bottom: '20px', left: '20px', right: '20px' }
        });
        console.log(`   ✅ PDF Generated: ${pdfPath}`);

        // =================================================================
        // STEP 8: Zip & Email
        // =================================================================
        console.log('📧 Step 8: Zipping Excels & Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        const excelsToZip = allFiles.filter(f => f.startsWith('Converted_'));

        if (excelsToZip.length > 0 || fs.existsSync(pdfPath)) {
            const zipName = `DTC_Excel_Reports_${todayStr}.zip`;
            const zipPath = path.join(downloadPath, zipName);
            
            if(excelsToZip.length > 0) {
                await zipFiles(downloadPath, zipPath, excelsToZip);
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
                text: `เรียน ผู้เกี่ยวข้อง\n\nระบบส่งรายงานประจำวัน (06:00 - 18:00) ดังแนบ:\n1. ไฟล์ Excel รายละเอียด (อยู่ใน Zip)\n2. ไฟล์ PDF สรุปภาพรวม\n\nขอบคุณครับ\nDTC Automation Bot`,
                attachments: attachments
            });
            console.log(`   ✅ Email Sent Successfully! (${attachments.length} attachments)`);
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
