const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 5 นาที (300 วินาที)
    for (let i = 0; i < 300; i++) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => 
            (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
            !f.endsWith('.crdownload') && 
            !f.startsWith('DTC_Completed_')
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
    
    const stats = fs.statSync(oldPath);
    console.log(`   Found File: ${downloadedFile} (Size: ${stats.size} bytes)`);
    
    if (stats.size === 0) {
        throw new Error(`Downloaded file ${downloadedFile} is empty!`);
    }

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    console.log(`   ✅ Renamed to: ${finalFileName}`);
    return newPath;
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

function parseDurationToMinutes(durationStr) {
    if (!durationStr || !durationStr.includes(':')) return 0;
    const parts = durationStr.split(':').map(Number);
    if (parts.length === 3) return (parts[0] * 60) + parts[1] + (parts[2] / 60);
    if (parts.length === 2) return (parts[0] * 60) + parts[1];
    return 0;
}

function extractDataFromReport(filePath, reportType) {
    try {
        const content = fs.readFileSync(filePath, 'utf-8');
        const dom = new JSDOM(content);
        const rows = Array.from(dom.window.document.querySelectorAll('table tr'));
        const data = [];
        for (let i = 1; i < rows.length; i++) {
            const cells = Array.from(rows[i].querySelectorAll('td')).map(td => td.textContent.trim());
            if (cells.length < 3) continue;
            if (reportType === 'speed') {
                const plate = cells.find(c => c.match(/\d{1,3}-\d{4}/)) || cells[1]; 
                const duration = cells[cells.length - 1]; 
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            } 
            else if (reportType === 'idling') {
                const plate = cells.find(c => c.match(/\d{1,3}-\d{4}/)) || cells[1];
                const duration = cells[cells.length - 1];
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            }
            else if (reportType === 'critical') {
                const plate = cells.find(c => c.match(/\d{1,3}-\d{4}/)) || cells[1];
                const detail = cells[2] || 'Unknown';
                data.push({ plate, detail });
            }
            else if (reportType === 'forbidden') {
                const plate = cells.find(c => c.match(/\d{1,3}-\d{4}/)) || cells[1];
                const station = cells[2] || ''; 
                const duration = cells[cells.length - 1];
                data.push({ plate, station, duration, durationMin: parseDurationToMinutes(duration) });
            }
        }
        return data;
    } catch (e) {
        console.warn(`   ⚠️ Failed to parse ${path.basename(filePath)}: ${e.message}`);
        return [];
    }
}

// --- Main Script ---

(async () => {
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing DTC_USERNAME or DTC_PASSWORD secrets.');
        process.exit(1);
    }

    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (Report 4: Type "ทั้งหมด" Fix)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(1800000); // 30 mins
    page.setDefaultTimeout(1800000);
    
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

        // Report 1: Over Speed
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        await new Promise(r => setTimeout(r, 2000));
        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            var select = document.getElementById('ddl_truck'); 
            for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } 
            select.dispatchEvent(new Event('change', { bubbles: true }));
        }, startDateTime, endDateTime);
        await page.evaluate(() => { if(typeof sertch_data === 'function') sertch_data(); else document.querySelector("span[onclick='sertch_data();']").click(); });
        await new Promise(r => setTimeout(r, 300000)); 
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        const file1 = await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // Report 2: Idling
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 2000));
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
        await new Promise(r => setTimeout(r, 300000));
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        const file2 = await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

        // Report 3: Sudden Brake
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 2000));
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);
        await page.click('td:nth-of-type(6) > span');
        
        console.log('   ⏳ Waiting 3 mins...'); 
        await new Promise(r => setTimeout(r, 180000)); 

        await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const b = btns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
            if (b) b.click(); else document.querySelector('#table button:nth-of-type(3)')?.click();
        });
        const file3 = await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // =================================================================
        // REPORT 4: Harsh Start (FIXED: Type "ทั้งหมด" in #s2id_autogen2)
        // =================================================================
        console.log('📊 Processing Report 4: Harsh Start...');
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
            
            // Debug 1
            await page.screenshot({ path: path.join(downloadPath, 'report4_01_loaded.png') });
            
            // รอวันที่
            await page.waitForSelector('#date9', { visible: true, timeout: 60000 });

            console.log('   Setting Report 4 Conditions (ID: s2id_ddl_truck)...');
            
            // 1. ตั้งวันที่
            await page.evaluate((start, end) => {
                document.getElementById('date9').value = start;
                document.getElementById('date10').value = end;
                document.getElementById('date9').dispatchEvent(new Event('change'));
                document.getElementById('date10').dispatchEvent(new Event('change'));
            }, startDateTime, endDateTime);

            // 2. เลือกรถ (เปิด Dropdown -> หา input #s2id_autogen2 -> พิมพ์ "ทั้งหมด" -> Enter)
            const select2Container = await page.$('#s2id_ddl_truck');
            if (select2Container) {
                console.log('   Found #s2id_ddl_truck, interacting with Select2...');
                await select2Container.click(); // เปิด Dropdown
                
                try {
                    // รอ animation เปิด Dropdown และรอ input โผล่
                    // ใช้ ID #s2id_autogen2 ตาม Recording
                    const searchInputSelector = '#s2id_autogen2';
                    await page.waitForSelector(searchInputSelector, { visible: true, timeout: 5000 });
                    
                    console.log('   Found Select2 Search Input (#s2id_autogen2), typing "ทั้งหมด"...');
                    await page.type(searchInputSelector, 'ทั้งหมด');
                    await new Promise(r => setTimeout(r, 1000));
                    await page.keyboard.press('Enter');
                    console.log('   Select2: Typed "ทั้งหมด" and pressed Enter.');
                } catch (e) {
                    console.warn('   ⚠️ Select2 interaction failed:', e.message);
                    // Fallback: ถ้าหา input ไม่เจอ ลองกด ArrowDown + Enter
                    console.log('   Fallback: Trying ArrowDown + Enter...');
                    await page.keyboard.press('ArrowDown');
                    await page.keyboard.press('Enter');
                }
            } else {
                console.log('   #s2id_ddl_truck not found, trying standard #ddl_truck...');
                await page.evaluate(() => {
                    const select = document.getElementById('ddl_truck');
                    if (select) {
                        select.selectedIndex = 0;
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                });
            }

            // Debug 2
            await page.screenshot({ path: path.join(downloadPath, 'report4_02_before_search.png') });

            // 3. กดค้นหา
            console.log('   Clicking Search Report 4...');
            await page.evaluate(() => {
                if (typeof sertch_data === 'function') {
                    sertch_data();
                } else {
                    document.querySelector('td:nth-of-type(6) > span').click();
                }
            });

            // รอ 3 นาที (180s) พร้อม Log
            console.log('   ⏳ Waiting 3 mins for Report 4 data...');
            for (let i = 1; i <= 3; i++) {
                await new Promise(r => setTimeout(r, 60000)); // รอ 1 นาที
                console.log(`      ... Passed ${i} minute(s)`);
            }

            // Debug 3
            await page.screenshot({ path: path.join(downloadPath, 'report4_03_after_wait.png') });

            // กด Export
            console.log('   Clicking Export Report 4...');
            await page.evaluate(() => {
                const xpathResult = document.evaluate('//*[@id="table"]/div[1]/button[3]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                const btn = xpathResult.singleNodeValue;
                if (btn) btn.click();
                else {
                    const allBtns = Array.from(document.querySelectorAll('button'));
                    const excelBtn = allBtns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
                    if (excelBtn) excelBtn.click();
                    else throw new Error("Cannot find Export button for Report 4");
                }
            });

            const file4 = await waitForDownloadAndRename(downloadPath, 'Report4_HarshStart.xls');

        } catch (error) {
            console.error('❌ Report 4 Failed:', error.message);
            await page.screenshot({ path: path.join(downloadPath, 'report4_error_snapshot.png') });
            throw error; 
        }

        // Report 5: Forbidden
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        await new Promise(r => setTimeout(r, 2000));
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            var select = document.getElementById('ddl_truck'); 
            if (select) { for (let opt of select.options) { if (opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) { for(var i=0; i<s.options.length; i++) { if(s.options[i].text.includes('พื้นที่ห้ามเข้า')) { s.value = s.options[i].value; s.dispatchEvent(new Event('change', { bubbles: true })); break; } } }
        }, startDateTime, endDateTime);
        await new Promise(r => setTimeout(r, 2000));
        await page.evaluate(() => {
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) { for(var i=0; i<s.options.length; i++) { if(s.options[i].text.includes('สถานีทั้งหมด')) { s.value = s.options[i].value; s.dispatchEvent(new Event('change', { bubbles: true })); break; } } }
        });
        await page.click('td:nth-of-type(7) > span');
        await new Promise(r => setTimeout(r, 300000));
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 7: Generate PDF Summary
        // =================================================================
        console.log('📑 Step 7: Generating PDF Summary...');

        const speedData = extractDataFromReport(file1, 'speed');
        const idlingData = extractDataFromReport(file2, 'idling');
        const brakeData = extractDataFromReport(file3, 'critical');
        // Check if Report 4 file exists before reading
        let startData = [];
        try {
            startData = extractDataFromReport(path.join(downloadPath, 'Report4_HarshStart.xls'), 'critical');
        } catch(e) { console.warn("Skipping Report 4 data in PDF due to missing file"); }
        
        const forbiddenData = extractDataFromReport(file5, 'forbidden');

        // Aggregation Logic (Top 5)
        const speedStats = {};
        speedData.forEach(d => {
            if (!speedStats[d.plate]) speedStats[d.plate] = { count: 0, durationMin: 0 };
            speedStats[d.plate].count++;
            speedStats[d.plate].durationMin += d.durationMin;
        });
        const topSpeed = Object.entries(speedStats)
            .map(([plate, data]) => ({ plate, ...data }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 5);

        const idlingStats = {};
        idlingData.forEach(d => {
            if (!idlingStats[d.plate]) idlingStats[d.plate] = { durationMin: 0 };
            idlingStats[d.plate].durationMin += d.durationMin;
        });
        const topIdling = Object.entries(idlingStats)
            .map(([plate, data]) => ({ plate, ...data }))
            .sort((a, b) => b.durationMin - a.durationMin)
            .slice(0, 5);

        const forbiddenStats = {};
        forbiddenData.forEach(d => {
            if (!forbiddenStats[d.plate]) forbiddenStats[d.plate] = { durationMin: 0, count: 0 };
            forbiddenStats[d.plate].durationMin += d.durationMin;
            forbiddenStats[d.plate].count++;
        });
        const topForbidden = Object.entries(forbiddenStats)
            .map(([plate, data]) => ({ plate, ...data }))
            .sort((a, b) => b.durationMin - a.durationMin)
            .slice(0, 5);

        const totalCritical = brakeData.length + startData.length;

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
                    <div class="card">
                        <h3>Over Speed (ครั้ง)</h3>
                        <div class="val text-blue-800">${speedData.length}</div>
                        <p class="text-sm text-gray-500">เหตุการณ์ทั้งหมด</p>
                    </div>
                    <div class="card" style="background-color: #fff7ed;">
                        <h3 style="color: #f59e0b;">Max Idling (นาที)</h3>
                        <div class="val text-orange-500">${topIdling.length > 0 ? topIdling[0].durationMin.toFixed(0) : 0}</div>
                        <p class="text-sm text-gray-500">สูงสุดต่อคัน</p>
                    </div>
                    <div class="card" style="background-color: #fef2f2;">
                        <h3 style="color: #dc2626;">Critical Events</h3>
                        <div class="val text-red-600">${totalCritical}</div>
                        <p class="text-sm text-gray-500">เบรก/ออกตัว กระชาก</p>
                    </div>
                    <div class="card" style="background-color: #f3e8ff;">
                        <h3 style="color: #9333ea;">Prohibited Parking</h3>
                        <div class="val text-purple-600">${forbiddenData.length}</div>
                        <p class="text-sm text-gray-500">เข้าพื้นที่ห้ามจอด (ครั้ง)</p>
                    </div>
                </div>
            </div>
            <div class="page-break">
                <div class="header-blue"><h2 class="text-2xl">1. การใช้ความเร็วเกินกำหนด (Over Speed Analysis)</h2></div>
                <div class="h-64 mb-6"><canvas id="speedChart"></canvas></div>
                <table>
                    <thead><tr><th>No.</th><th>ทะเบียนรถ</th><th>จำนวนครั้ง</th><th>รวมเวลา (นาที)</th></tr></thead>
                    <tbody>
                        ${topSpeed.map((d, i) => `<tr><td>${i+1}</td><td>${d.plate}</td><td>${d.count}</td><td>${d.durationMin.toFixed(2)}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>
            <div class="page-break">
                <div class="header-blue" style="background-color: #f59e0b;"><h2 class="text-2xl">2. การจอดไม่ดับเครื่อง (Idling Analysis)</h2></div>
                <div class="h-64 mb-6"><canvas id="idlingChart"></canvas></div>
                <table>
                    <thead><tr><th>No.</th><th>ทะเบียนรถ</th><th>รวมเวลา (นาที)</th></tr></thead>
                    <tbody>
                        ${topIdling.map((d, i) => `<tr><td>${i+1}</td><td>${d.plate}</td><td>${d.durationMin.toFixed(2)}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>
            <div class="page-break">
                <div class="header-blue" style="background-color: #dc2626;"><h2 class="text-2xl">3. เหตุการณ์วิกฤต (Critical Safety Events)</h2></div>
                <h3 class="text-xl font-bold mt-4 mb-2">3.1 เบรกกะทันหัน (Sudden Brake)</h3>
                <table>
                    <thead><tr><th>ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead>
                    <tbody>
                        ${brakeData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('')}
                    </tbody>
                </table>
                <h3 class="text-xl font-bold mt-8 mb-2">3.2 ออกตัวกระชาก (Harsh Start)</h3>
                <table>
                    <thead><tr><th>ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead>
                    <tbody>
                        ${startData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>
            <div>
                <div class="header-blue" style="background-color: #9333ea;"><h2 class="text-2xl">4. รายงานพื้นที่ห้ามจอด (Prohibited Parking)</h2></div>
                <div class="h-64 mb-6"><canvas id="forbiddenChart"></canvas></div>
                <table>
                    <thead><tr><th>No.</th><th>ทะเบียนรถ</th><th>สถานี</th><th>รวมเวลา (นาที)</th></tr></thead>
                    <tbody>
                        ${topForbidden.map((d, i) => `<tr><td>${i+1}</td><td>${d.plate}</td><td>-</td><td>${d.durationMin.toFixed(2)}</td></tr>`).join('')}
                    </tbody>
                </table>
            </div>
            <script>
                new Chart(document.getElementById('speedChart'), {
                    type: 'bar',
                    data: {
                        labels: ${JSON.stringify(topSpeed.map(d => d.plate))},
                        datasets: [{ label: 'Frequency', data: ${JSON.stringify(topSpeed.map(d => d.count))}, backgroundColor: '#1e40af' }]
                    },
                    options: { maintainAspectRatio: false }
                });
                new Chart(document.getElementById('idlingChart'), {
                    type: 'bar',
                    indexAxis: 'y',
                    data: {
                        labels: ${JSON.stringify(topIdling.map(d => d.plate))},
                        datasets: [{ label: 'Duration (Min)', data: ${JSON.stringify(topIdling.map(d => d.durationMin))}, backgroundColor: '#f59e0b' }]
                    },
                    options: { maintainAspectRatio: false }
                });
                new Chart(document.getElementById('forbiddenChart'), {
                    type: 'bar',
                    data: {
                        labels: ${JSON.stringify(topForbidden.map(d => d.plate))},
                        datasets: [{ label: 'Duration (Min)', data: ${JSON.stringify(topForbidden.map(d => d.durationMin))}, backgroundColor: '#9333ea' }]
                    },
                    options: { maintainAspectRatio: false }
                });
            </script>
        </body>
        </html>
        `;

        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
        const pdfPath = path.join(downloadPath, 'Fleet_Safety_Analysis_Report.pdf');
        await page.pdf({
            path: pdfPath,
            format: 'A4',
            printBackground: true,
            margin: { top: '20px', bottom: '20px', left: '20px', right: '20px' }
        });
        console.log(`   ✅ PDF Generated: ${pdfPath}`);

        // Step 8: Send Email
        console.log('📧 Step 8: Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        
        const filesToSend = allFiles.filter(file => 
            file.startsWith('DTC_Completed_') || file.endsWith('.pdf')
        );
        
        const attachments = filesToSend.map(file => {
            const filePath = path.join(downloadPath, file);
            return { filename: file, path: filePath };
        });

        if (attachments.length > 0) {
            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงาน DTC Report (5 ฉบับ + สรุป PDF) ประจำวันที่ ${todayStr}`,
                text: `เรียน ผู้เกี่ยวข้อง,\n\nระบบส่งไฟล์รายงานจำนวน ${attachments.length} ฉบับ ดังแนบ\n(ข้อมูลช่วงเวลา 06:00 - 18:00)\n\nประกอบด้วย:\n1. Excel Reports (5 files)\n2. PDF Summary Report\n\nขอบคุณครับ\nDTC Automation Bot`,
                attachments: attachments
            });
            console.log(`   ✅ Email Sent Successfully! (${attachments.length} files)`);
        } else {
            console.warn('⚠️ No files found to send!');
        }

        console.log('🧹 Cleanup...');
        // fs.rmSync(downloadPath, { recursive: true, force: true });
        console.log('   ✅ Cleanup Complete.');

    } catch (err) {
        console.error('❌ Fatal Error:', err);
        await page.screenshot({ path: path.join(downloadPath, 'error_screenshot.png') });
        process.exit(1);
    } finally {
        await browser.close();
        console.log('🏁 Browser Closed.');
    }
})();
