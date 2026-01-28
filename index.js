const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');
const ExcelJS = require('exceljs');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 5 นาที (300 วินาที) เผื่อไฟล์ใหญ่มาก
    for (let i = 0; i < 300; i++) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => 
            (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
            !f.endsWith('.crdownload') && 
            !f.startsWith('DTC_Completed_') &&
            !f.startsWith('Converted_')
        );
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    await new Promise(resolve => setTimeout(resolve, 3000));

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    // ตรวจสอบขนาดไฟล์เบื้องต้น
    const stats = fs.statSync(oldPath);
    if (stats.size === 0) throw new Error(`Downloaded file ${downloadedFile} is empty!`);
    
    // แสดงขนาดไฟล์ (KB)
    const fileSizeInKB = (stats.size / 1024).toFixed(2);
    console.log(`   ✅ Found File: ${downloadedFile} (Size: ${fileSizeInKB} KB)`);
    
    if (stats.size < 100) console.warn(`   ⚠️ Warning: File ${downloadedFile} seems too small (${stats.size} bytes).`);

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    
    // แปลงเป็น XLSX
    const xlsxFileName = `Converted_${newFileName.replace('.xls', '.xlsx')}`;
    const xlsxPath = path.join(downloadPath, xlsxFileName);
    await convertHtmlToExcel(newPath, xlsxPath);

    return xlsxPath;
}

// ฟังก์ชันรอให้ตารางมีข้อมูล
async function waitForTableData(page, minRows = 2, timeout = 300000) {
    console.log(`   Waiting for table data (Max ${timeout/1000}s)...`);
    try {
        await page.waitForFunction((min) => {
            const rows = document.querySelectorAll('table tr');
            return rows.length >= min; 
        }, { timeout: timeout }, minRows);
        console.log('   ✅ Table data populated.');
    } catch (e) {
        console.warn('   ⚠️ Wait for table data timed out (Data might be empty).');
    }
}

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
             fs.copyFileSync(sourcePath, destPath);
             return;
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('DMS Report');
        const rows = Array.from(table.querySelectorAll('tr'));

        rows.forEach((row, rowIndex) => {
            const cells = Array.from(row.querySelectorAll('td, th'));
            const rowData = cells.map(cell => cell.textContent.replace(/<[^>]*>/g, '').trim());
            const excelRow = worksheet.addRow(rowData);

            // จัดรูปแบบ
            excelRow.eachCell((cell, colNumber) => {
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                // Header Styling
                if (rowIndex === 0 || cells[colNumber-1].tagName === 'TH') {
                    cell.font = { bold: true };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
            });
        });

        // Auto-fit Columns (Simplified)
        worksheet.columns.forEach(column => { column.width = 20; });

        await workbook.xlsx.writeFile(destPath);
        console.log(`   ✅ Converted to XLSX: ${path.basename(destPath)}`);
    } catch (e) {
        console.warn(`   ⚠️ Conversion failed, keeping original: ${e.message}`);
        fs.copyFileSync(sourcePath, destPath);
    }
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

function parseDurationToMinutes(durationStr) {
    if (!durationStr || typeof durationStr !== 'string') return 0;
    const cleanStr = durationStr.replace(/[^\d:]/g, ''); 
    if (!cleanStr.includes(':')) return 0;
    const parts = cleanStr.split(':').map(Number);
    if (parts.length === 3) return (parts[0] * 60) + parts[1] + (parts[2] / 60);
    if (parts.length === 2) return (parts[0] * 60) + parts[1];
    return 0;
}

async function extractDataFromXLSX(filePath, reportType) {
    try {
        if (!fs.existsSync(filePath)) return [];
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(1);
        const data = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber < 2) return; 
            const cells = Array.isArray(row.values) ? row.values.slice(1) : []; 
            if (cells.length < 4) return;

            const plateRegex = /\d{1,3}-?\d{1,4}|[ก-ฮ]{1,3}\d{1,4}/;
            const plateVal = cells.find(c => plateRegex.test(String(c)));
            
            if (!plateVal) return;
            const plate = String(plateVal).trim();

            if (reportType === 'speed') {
                const duration = String(cells[cells.length - 1]); 
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            } else if (reportType === 'idling') {
                const duration = String(cells[cells.length - 1]);
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            } else if (reportType === 'critical') {
                const detailIndex = cells.findIndex(c => String(c).includes(plate)) + 1;
                const detail = cells[detailIndex] ? String(cells[detailIndex]) : 'Event';
                data.push({ plate, detail });
            } else if (reportType === 'forbidden') {
                const duration = String(cells[cells.length - 1]);
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            }
        });
        console.log(`      -> Extracted ${data.length} records from ${path.basename(filePath)}`);
        return data;
    } catch (e) {
        console.warn(`   ⚠️ Extract Error: ${e.message}`);
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

    console.log('🚀 Starting DTC Automation (Verified Version)...');
    
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
        
        // =================================================================
        // REPORT 1: Over Speed
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        
        // รอ Dropdown รถ
        await page.waitForFunction(() => {
            const s = document.getElementById('ddl_truck');
            return s && s.options.length > 1; 
        }, { timeout: 60000 });

        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            
            // เลือก "ทั้งหมด" (วนลูปหา)
            const select = document.getElementById('ddl_truck');
            if(select) {
                for(let i=0; i<select.options.length; i++) {
                    if(select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                        select.selectedIndex = i; break; 
                    }
                }
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        await page.evaluate(() => { if(typeof sertch_data === 'function') sertch_data(); else document.querySelector("span[onclick='sertch_data();']").click(); });
        
        // *** Hard Wait 5 mins (For Wait For Table Data) ***
        await waitForTableData(page, 2, 300000); 

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file1 = await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // =================================================================
        // REPORT 2: Idling
        // =================================================================
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.waitForFunction(() => {
            const s = document.getElementById('ddl_truck');
            return s && s.options.length > 1; 
        }, { timeout: 60000 });

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '10';
            
            const select = document.getElementById('ddl_truck');
            if(select) {
                for(let i=0; i<select.options.length; i++) {
                    if(select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                        select.selectedIndex = i; break; 
                    }
                }
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        await page.click('td:nth-of-type(6) > span');
        
        // Wait for Data
        await waitForTableData(page, 2, 180000);

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file2 = await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

        // =================================================================
        // REPORT 3: Sudden Brake
        // =================================================================
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.waitForFunction(() => {
            const s = document.getElementById('ddl_truck');
            return s && s.options.length > 1; 
        }, { timeout: 60000 });

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            
            const select = document.getElementById('ddl_truck');
            if(select) {
                for(let i=0; i<select.options.length; i++) {
                    if(select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                        select.selectedIndex = i; break; 
                    }
                }
                select.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        await page.click('td:nth-of-type(6) > span');
        
        // Wait for Data
        await waitForTableData(page, 2, 180000);

        await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const b = btns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
            if(b) b.click(); else document.querySelector('#table button:nth-of-type(3)')?.click();
        });
        const file3 = await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // =================================================================
        // REPORT 4: Harsh Start
        // =================================================================
        console.log('📊 Processing Report 4: Harsh Start...');
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            await page.waitForSelector('#date9', { visible: true });
            
            // รอ Dropdown โหลด
            await page.waitForFunction(() => {
                const s = document.getElementById('ddl_truck');
                return s && s.options.length > 1; 
            }, { timeout: 60000 });

            console.log('   Setting Report 4 Conditions...');

            await page.evaluate((start, end) => {
                document.getElementById('date9').value = start;
                document.getElementById('date10').value = end;
                document.getElementById('date9').dispatchEvent(new Event('change'));
                document.getElementById('date10').dispatchEvent(new Event('change'));
                
                // Programmatic Select: All Trucks
                const select = document.getElementById('ddl_truck');
                if(select) {
                    for(let i=0; i<select.options.length; i++) {
                        if(select.options[i].text.includes('ทั้งหมด') || select.options[i].text.toLowerCase().includes('all')) {
                            select.selectedIndex = i; break; 
                        }
                    }
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    if (typeof $ !== 'undefined' && $(select).data('select2')) {
                        $(select).trigger('change'); 
                    }
                }
            }, startDateTime, endDateTime);

            // กด Search
            console.log('   Clicking Search Report 4...');
            await page.evaluate(() => {
                if(typeof sertch_data === 'function') sertch_data();
                else document.querySelector('td:nth-of-type(6) > span').click();
            });

            // Wait for Data
            await waitForTableData(page, 2, 180000);

            // กด Export
            console.log('   Clicking Export Report 4...');
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

        } catch (error) {
            console.error('❌ Report 4 Failed:', error.message);
        }

        // =================================================================
        // REPORT 5: Forbidden Parking
        // =================================================================
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.waitForFunction(() => {
            const s = document.getElementById('ddl_truck');
            return s && s.options.length > 1; 
        }, { timeout: 60000 });

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            
            const select = document.getElementById('ddl_truck');
            if(select) { 
                for(let opt of select.options) { 
                    if(opt.text.includes('ทั้งหมด')) { select.selectedIndex = opt.index; break; } 
                } 
                select.dispatchEvent(new Event('change', { bubbles: true })); 
            }
            
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
        
        // Wait for Data
        await waitForTableData(page, 2, 180000);

        await page.evaluate(() => document.getElementById('btnexport').click());
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 7: Generate PDF Summary
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

        // Aggregation & PDF Generation
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

        // HTML Content
        const htmlContent = `
        <!DOCTYPE html>
        <html lang="th">
        <head>
            <meta charset="UTF-8">
            <script src="https://cdn.tailwindcss.com"></script>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;600;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Noto Sans Thai', sans-serif; background: #fff; }
                .page-break { page-break-after: always; }
                .header-blue { background-color: #1e40af; color: white; padding: 10px 20px; border-radius: 8px; margin-bottom: 20px; }
                .card { background: #eff6ff; border-radius: 12px; padding: 20px; text-align: center; }
                .card h3 { color: #1e40af; font-weight: bold; font-size: 1.2rem; }
                .card .val { font-size: 2.5rem; font-weight: bold; margin: 10px 0; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th { background-color: #1e40af; color: white; padding: 10px; text-align: left; }
                td { padding: 10px; border-bottom: 1px solid #e5e7eb; }
                tr:nth-child(even) { background-color: #eff6ff; }
            </style>
        </head>
        <body class="p-8">
            <div class="page-break">
                <div class="text-center mb-10">
                    <h1 class="text-3xl font-bold text-blue-800">รายงานสรุปพฤติกรรมการขับขี่</h1>
                    <h2 class="text-xl text-gray-600">Fleet Safety & Telematics Analysis Report</h2>
                    <p class="text-lg mt-2">วันที่: ${todayStr} (06:00 - 18:00)</p>
                </div>
                <div class="grid grid-cols-2 gap-6 mt-10">
                    <div class="card"><h3>Over Speed (ครั้ง)</h3><div class="val text-blue-800">${speedData.length}</div></div>
                    <div class="card bg-orange-50"><h3>Max Idling (นาที)</h3><div class="val text-orange-500">${topIdling.length > 0 ? topIdling[0].durationMin.toFixed(0) : 0}</div></div>
                    <div class="card bg-red-50"><h3>Critical Events</h3><div class="val text-red-600">${totalCritical}</div></div>
                    <div class="card bg-purple-50"><h3>Prohibited</h3><div class="val text-purple-600">${forbiddenData.length}</div></div>
                </div>
            </div>

            <div class="page-break">
                <div class="header-blue"><h2 class="text-2xl">1. การใช้ความเร็วเกินกำหนด</h2></div>
                <div class="h-64 mb-6"><canvas id="speedChart"></canvas></div>
                <table><thead><tr><th>ทะเบียนรถ</th><th>จำนวนครั้ง</th><th>รวมเวลา (นาที)</th></tr></thead>
                <tbody>${topSpeed.map(d => `<tr><td>${d.plate}</td><td>${d.count}</td><td>${d.durationMin.toFixed(2)}</td></tr>`).join('')}</tbody></table>
            </div>

            <div class="page-break">
                <div class="header-blue" style="background-color: #f59e0b;"><h2 class="text-2xl">2. การจอดไม่ดับเครื่อง</h2></div>
                <div class="h-64 mb-6"><canvas id="idlingChart"></canvas></div>
                <table><thead><tr><th>ทะเบียนรถ</th><th>รวมเวลา (นาที)</th></tr></thead>
                <tbody>${topIdling.map(d => `<tr><td>${d.plate}</td><td>${d.durationMin.toFixed(2)}</td></tr>`).join('')}</tbody></table>
            </div>

            <div class="page-break">
                <div class="header-blue"><h2 class="text-2xl">3. เหตุการณ์วิกฤต</h2></div>
                <h3 class="text-xl mt-4">เบรกกะทันหัน</h3>
                <table><thead><tr><th>ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead><tbody>${brakeData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('')}</tbody></table>
                <h3 class="text-xl mt-8">ออกตัวกระชาก</h3>
                <table><thead><tr><th>ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead><tbody>${startData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('')}</tbody></table>
            </div>

            <div>
                <div class="header-blue"><h2 class="text-2xl">4. รายงานพื้นที่ห้ามจอด</h2></div>
                <div class="h-64 mb-6"><canvas id="forbiddenChart"></canvas></div>
                <table><thead><tr><th>ทะเบียนรถ</th><th>สถานี</th><th>รวมเวลา (นาที)</th></tr></thead>
                <tbody>${topForbidden.map(d => `<tr><td>${d.plate}</td><td>-</td><td>${d.durationMin.toFixed(2)}</td></tr>`).join('')}</tbody></table>
            </div>

            <script>
                const chartConfig = (id, label, labels, data, color) => new Chart(document.getElementById(id), {
                    type: 'bar', data: { labels, datasets: [{ label, data, backgroundColor: color }] }, options: { maintainAspectRatio: false }
                });
                chartConfig('speedChart', 'Count', ${JSON.stringify(topSpeed.map(d=>d.plate))}, ${JSON.stringify(topSpeed.map(d=>d.count))}, '#1e40af');
                chartConfig('idlingChart', 'Minutes', ${JSON.stringify(topIdling.map(d=>d.plate))}, ${JSON.stringify(topIdling.map(d=>d.durationMin))}, '#f59e0b');
                chartConfig('forbiddenChart', 'Minutes', ${JSON.stringify(topForbidden.map(d=>d.plate))}, ${JSON.stringify(topForbidden.map(d=>d.durationMin))}, '#9333ea');
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
            if (fs.existsSync(pdfPath)) attachments.push({ filename: 'Summary_Report.pdf', path: pdfPath });

            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงาน DTC Report ประจำวันที่ ${todayStr}`,
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
