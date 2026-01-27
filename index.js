const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { JSDOM } = require('jsdom');
const archiver = require('archiver');

// --- Helper Functions ---

// ฟังก์ชันรอโหลดไฟล์แบบเดิม (รอเวลาที่กำหนดแล้วค่อยเช็คไฟล์)
async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Processing file: ${newFileName}...`);
    let downloadedFile = null;

    // เช็คไฟล์ใน folder (หลังจากที่ Hard Wait ข้างนอกเสร็จแล้ว)
    const files = fs.readdirSync(downloadPath);
    downloadedFile = files.find(f => 
        (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
        !f.endsWith('.crdownload') && 
        !f.startsWith('DTC_Completed_')
    );

    if (!downloadedFile) {
        // ถ้าหาไม่เจอ ให้โอกาสอีกนิด (Retry Loop สั้นๆ)
        for(let i=0; i<30; i++) { // รอเพิ่มสูงสุด 30 วิ
            await new Promise(r => setTimeout(r, 1000));
            const retryFiles = fs.readdirSync(downloadPath);
            downloadedFile = retryFiles.find(f => (f.endsWith('.xls') || f.endsWith('.xlsx')) && !f.endsWith('.crdownload') && !f.startsWith('DTC_Completed_'));
            if(downloadedFile) break;
        }
    }

    if (!downloadedFile) {
        throw new Error(`File not found for ${newFileName}`);
    }

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    const stats = fs.statSync(oldPath);
    if (stats.size === 0) throw new Error(`Downloaded file ${downloadedFile} is empty!`);

    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    fs.renameSync(oldPath, newPath);
    console.log(`   ✅ Saved as: ${finalFileName}`);
    return newPath;
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

function extractDataFromReport(filePath, reportType) {
    try {
        if (!fs.existsSync(filePath)) return [];
        const content = fs.readFileSync(filePath, 'utf-8');
        const dom = new JSDOM(content);
        const rows = Array.from(dom.window.document.querySelectorAll('table tr'));
        const data = [];

        rows.forEach((row) => {
            const cells = Array.from(row.querySelectorAll('td')).map(td => td.textContent.trim());
            if (cells.length < 4) return; 

            const plateRegex = /\d{1,3}-?\d{1,4}|[ก-ฮ]{1,3}\d{1,4}/;
            const plateIndex = cells.findIndex(c => plateRegex.test(c) && c.length < 20);
            if (plateIndex === -1) return;

            const plate = cells[plateIndex];

            if (reportType === 'speed') {
                const duration = cells[cells.length - 1]; 
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            } 
            else if (reportType === 'idling') {
                const duration = cells[cells.length - 1];
                data.push({ plate, duration, durationMin: parseDurationToMinutes(duration) });
            }
            else if (reportType === 'critical') { 
                let detail = cells.find((c, i) => i > plateIndex && c.length > 5 && !c.includes(':')) || 'Critical Event';
                data.push({ plate, detail });
            }
            else if (reportType === 'forbidden') {
                const station = cells[plateIndex + 1] || 'Unknown Station'; 
                const duration = cells[cells.length - 1];
                data.push({ plate, station, duration, durationMin: parseDurationToMinutes(duration) });
            }
        });
        return data;
    } catch (e) {
        console.warn(`   ⚠️ Data extract error ${path.basename(filePath)}: ${e.message}`);
        return [];
    }
}

function zipExcelFiles(sourceDir, outPath, filesToZip) {
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
        console.error('❌ Error: Missing DTC Secrets.');
        process.exit(1);
    }

    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (Hard Wait Mode)...');
    
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
        console.log(`🕒 Time Settings: ${startDateTime} to ${endDateTime}`);

        // =================================================================
        // REPORT 1: Over Speed
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#speed_max', { visible: true });
        
        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            const select = document.getElementById('ddl_truck');
            if(select) { select.selectedIndex = 0; select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);

        await page.evaluate(() => { if(typeof sertch_data === 'function') sertch_data(); else document.querySelector("span[onclick='sertch_data();']").click(); });
        
        // Hard Wait 3 Mins
        console.log('   ⏳ Waiting 3 mins...');
        await new Promise(r => setTimeout(r, 180000));

        // Export
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        const file1 = await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // =================================================================
        // REPORT 2: Idling
        // =================================================================
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '10';
            const select = document.getElementById('ddl_truck');
            if(select) { select.selectedIndex = 0; select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);

        await page.click('td:nth-of-type(6) > span');
        
        // Hard Wait 3 Mins
        console.log('   ⏳ Waiting 3 mins...');
        await new Promise(r => setTimeout(r, 180000));

        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        const file2 = await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');

        // =================================================================
        // REPORT 3: Sudden Brake
        // =================================================================
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            const select = document.getElementById('ddl_truck');
            if(select) { select.selectedIndex = 0; select.dispatchEvent(new Event('change', { bubbles: true })); }
        }, startDateTime, endDateTime);

        await page.click('td:nth-of-type(6) > span');
        
        // Hard Wait 3 Mins
        console.log('   ⏳ Waiting 3 mins...');
        await new Promise(r => setTimeout(r, 180000));

        await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const b = btns.find(b => b.innerText.includes('Excel') || b.title === 'Excel');
            if(b) b.click(); else {
                const fallback = document.querySelector('#table button:nth-of-type(3)');
                if(fallback) fallback.click();
            }
        });
        
        const file3 = await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');

        // =================================================================
        // REPORT 4: Harsh Start (Robust Logic)
        // =================================================================
        console.log('📊 Processing Report 4: Harsh Start...');
        try {
            await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
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
                if (select) {
                    select.selectedIndex = 0; 
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }, startDateTime, endDateTime);

            await page.evaluate(() => {
                if(typeof sertch_data === 'function') sertch_data();
                else document.querySelector('td:nth-of-type(6) > span').click();
            });

            // Hard Wait 3 Mins
            console.log('   ⏳ Waiting 3 mins...');
            await new Promise(r => setTimeout(r, 180000));

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

        } catch (e) {
            console.error('❌ Report 4 Failed (Continuing without it):', e.message);
        }

        // =================================================================
        // REPORT 5: Forbidden Parking
        // =================================================================
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            const select = document.getElementById('ddl_truck');
            if(select) { for(let opt of select.options) { if(opt.text.includes('ทั้งหมด')) { select.value = opt.value; break; } } select.dispatchEvent(new Event('change', { bubbles: true })); }
            
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
        
        // Hard Wait 3 Mins
        console.log('   ⏳ Waiting 3 mins...');
        await new Promise(r => setTimeout(r, 180000));

        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        const file5 = await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');

        // =================================================================
        // STEP 7: Generate PDF Summary
        // =================================================================
        console.log('📑 Step 7: Generating PDF Summary...');

        const fileMap = {
            'speed': path.join(downloadPath, 'Report1_OverSpeed.xls'),
            'idling': path.join(downloadPath, 'Report2_Idling.xls'),
            'brake': path.join(downloadPath, 'Report3_SuddenBrake.xls'),
            'start': path.join(downloadPath, 'Report4_HarshStart.xls'),
            'forbidden': path.join(downloadPath, 'Report5_ForbiddenParking.xls')
        };

        const speedData = extractDataFromReport(fileMap.speed, 'speed');
        const idlingData = extractDataFromReport(fileMap.idling, 'idling');
        const brakeData = extractDataFromReport(fileMap.brake, 'critical');
        const startData = extractDataFromReport(fileMap.start, 'critical');
        const forbiddenData = extractDataFromReport(fileMap.forbidden, 'forbidden');

        // Aggregation Logic
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
                    <div class="card" style="background-color: #fff7ed;"><h3 style="color: #f59e0b;">Max Idling (นาที)</h3><div class="val text-orange-500">${topIdling.length > 0 ? topIdling[0].durationMin.toFixed(0) : 0}</div></div>
                    <div class="card" style="background-color: #fef2f2;"><h3 style="color: #dc2626;">Critical Events</h3><div class="val text-red-600">${totalCritical}</div></div>
                    <div class="card" style="background-color: #f3e8ff;"><h3 style="color: #9333ea;">Prohibited Parking</h3><div class="val text-purple-600">${forbiddenData.length}</div></div>
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
                <div class="header-blue" style="background-color: #dc2626;"><h2 class="text-2xl">3. เหตุการณ์วิกฤต</h2></div>
                <h3 class="text-xl font-bold mt-4">เบรกกะทันหัน</h3>
                <table><thead><tr><th>ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead><tbody>${brakeData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('')}</tbody></table>
                <h3 class="text-xl font-bold mt-8">ออกตัวกระชาก</h3>
                <table><thead><tr><th>ทะเบียนรถ</th><th>รายละเอียด</th></tr></thead><tbody>${startData.slice(0, 10).map(d => `<tr><td>${d.plate}</td><td>${d.detail}</td></tr>`).join('')}</tbody></table>
            </div>

            <div>
                <div class="header-blue" style="background-color: #9333ea;"><h2 class="text-2xl">4. รายงานพื้นที่ห้ามจอด</h2></div>
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
        // STEP 8: Zip Excels & Email (Zip Excel Only, Attach PDF Separately)
        // =================================================================
        console.log('📧 Step 8: Zipping Excels & Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        const excelsToZip = allFiles.filter(f => f.startsWith('DTC_Completed_'));

        if (excelsToZip.length > 0 || fs.existsSync(pdfPath)) {
            const zipName = `DTC_Excel_Reports_${todayStr}.zip`;
            const zipPath = path.join(downloadPath, zipName);
            
            if(excelsToZip.length > 0) {
                await zipExcelFiles(downloadPath, zipPath, excelsToZip);
            }

            const attachments = [];
            if (fs.existsSync(zipPath)) {
                attachments.push({ filename: zipName, path: zipPath });
            }
            if (fs.existsSync(pdfPath)) {
                attachments.push({ filename: 'Summary_Report.pdf', path: pdfPath });
            }

            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงาน DTC Report ประจำวันที่ ${todayStr}`,
                text: `เรียน ผู้เกี่ยวข้อง,\n\nระบบส่งรายงานประจำวัน (06:00 - 18:00) ดังแนบ:\n1. ไฟล์ Excel รายละเอียด (อยู่ใน Zip)\n2. ไฟล์ PDF สรุปภาพรวม\n\nขอบคุณครับ\nDTC Automation Bot`,
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
