const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 60 วินาที
    for (let i = 0; i < 60; i++) {
        const files = fs.readdirSync(downloadPath);
        // หาไฟล์ Excel (.xls, .xlsx) ที่ไม่ใช่ไฟล์ชั่วคราว (.crdownload) และไม่ใช่ไฟล์ที่เราเพิ่งเปลี่ยนชื่อไป (Report_*)
        downloadedFile = files.find(f => (f.endsWith('.xls') || f.endsWith('.xlsx')) && !f.endsWith('.crdownload') && !f.startsWith('Report_'));
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    const oldPath = path.join(downloadPath, downloadedFile);
    const newPath = path.join(downloadPath, newFileName);
    
    // ลบไฟล์ปลายทางถ้ามีอยู่แล้ว
    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    
    fs.renameSync(oldPath, newPath);
    console.log(`   ✅ Saved as: ${newFileName}`);
    return newPath;
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
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

    console.log('🚀 Starting DTC Automation (Updated Step 2)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);
    
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    
    await page.setViewport({ width: 1920, height: 1080 });
    await page.emulateTimezone('Asia/Bangkok');

    try {
        // =================================================================
        // STEP 1: LOGIN
        // =================================================================
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        
        console.log('   Clicking Login...');
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        // =================================================================
        // STEP 2: REPORT 1 - Over Speed (Updated Code)
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        
        // 2️⃣ Step 2: Go to Report Page...
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        
        // 3️⃣ Step 3: Fill Form
        console.log('   Step 3: Fill Form...');
        
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        
        // รอสักนิดเพื่อให้ตัวเลือกใน Dropdown โหลดมาครบ
        await new Promise(r => setTimeout(r, 2000));

        // คำนวณเวลา 06:00 - 18:00 ของวันนี้ (แทนที่สูตรเดิมใน Snippet)
        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`   Setting Time: ${startDateTime} to ${endDateTime}`);

        await page.evaluate((start, end) => {
            // Speed (Command 8)
            document.getElementById('speed_max').value = '55';
            
            // Date Formula (แก้ไขให้ใช้เวลา 06:00 - 18:00 ตามที่รับค่ามา)
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            
            // Trigger Events
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));

            // Options (Command 13)
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            
            // --- Select Truck (UI.Vision Command 14) ---
            var selectElement = document.getElementById('ddl_truck'); 
            var options = selectElement.options; 
            for (var i = 0; i < options.length; i++) { 
                if (options[i].text.includes('ทั้งหมด')) { 
                    selectElement.value = options[i].value; 
                    break; 
                } 
            } 
            var event = new Event('change', { bubbles: true }); 
            selectElement.dispatchEvent(event);
        }, startDateTime, endDateTime);

        // 4️⃣ Step 4: Search
        console.log('   Step 4: Search...');
        await page.evaluate(() => {
            if(typeof sertch_data === 'function') sertch_data();
            else document.querySelector("span[onclick='sertch_data();']").click();
        });

        // 5️⃣ Step 5: Wait 120s (Hard Wait)
        // หมายเหตุ: ปรับลดลงเหลือ 60s ถ้าข้อมูลไม่เยอะมาก หรือใช้ waitForSelector ตามด้านล่างจะเร็วกว่า
        console.log('   Step 5: Waiting for Data Loading...');
        try {
            await page.waitForSelector('#btnexport', { visible: true, timeout: 300000 }); // รอสูงสุด 5 นาที
            // รอเพิ่มอีกนิดเพื่อให้ข้อมูลโหลดสมบูรณ์จริงๆ หลังปุ่มขึ้น
            await new Promise(r => setTimeout(r, 5000)); 
        } catch(e) {
            console.warn('   ⚠️ Warning: Export button wait timed out');
        }
        console.log('   ✅ Data Loaded.');

        // 6️⃣ Step 6: Export & Download
        console.log('   Step 6: Exporting...');
        
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        // ใช้ Helper Function แทน Loop ใน Snippet เพื่อเปลี่ยนชื่อไฟล์และจัดการ Error
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');


        // =================================================================
        // STEP 3-6: Other Reports (Placeholder for Puppeteer Replay)
        // =================================================================
        // ... พื้นที่สำหรับวาง Code Report 2-5 ...


        // =================================================================
        // STEP 7: Generate PDF (Placeholder)
        // =================================================================
        console.log('📑 Generating PDF Summary (Pending)...');


        // =================================================================
        // STEP 8: Send Email
        // =================================================================
        console.log('📧 Step 8: Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        const attachments = allFiles.map(file => ({
            filename: file,
            path: path.join(downloadPath, file)
        }));

        if (attachments.length > 0) {
            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงาน DTC Report ประจำวันที่ ${todayStr} (06:00 - 18:00)`,
                text: 'เรียน ผู้เกี่ยวข้อง,\n\nระบบได้ทำการดึงรายงานและแนบไฟล์มาพร้อมกับอีเมลฉบับนี้\n\nขอบคุณครับ\nDTC Automation Bot',
                attachments: attachments
            });
            console.log('   ✅ Email Sent Successfully!');
        } else {
            console.warn('⚠️ No files to send!');
        }

        // =================================================================
        // STEP 9: Cleanup Files
        // =================================================================
        console.log('🧹 Step 9: Cleaning up files...');
        const filesToDelete = fs.readdirSync(downloadPath);
        for (const file of filesToDelete) {
            try {
                fs.unlinkSync(path.join(downloadPath, file));
            } catch (err) { }
        }
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
