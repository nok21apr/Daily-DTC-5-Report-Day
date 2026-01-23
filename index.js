const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 120 วินาที (เผื่อเน็ตช้า)
    for (let i = 0; i < 120; i++) {
        const files = fs.readdirSync(downloadPath);
        // หาไฟล์ Excel (.xls, .xlsx) ที่ไม่ใช่ไฟล์ชั่วคราว (.crdownload) และไม่ใช่ไฟล์ที่เราเพิ่งเปลี่ยนชื่อไป (Report_*)
        downloadedFile = files.find(f => (f.endsWith('.xls') || f.endsWith('.xlsx')) && !f.endsWith('.crdownload') && !f.startsWith('Report_'));
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    // รออีกนิดเพื่อให้มั่นใจว่าเขียนไฟล์เสร็จสมบูรณ์ 100%
    await new Promise(resolve => setTimeout(resolve, 2000));

    const oldPath = path.join(downloadPath, downloadedFile);
    const newPath = path.join(downloadPath, newFileName);
    
    // ตรวจสอบขนาดไฟล์ก่อนเปลี่ยนชื่อ
    const stats = fs.statSync(oldPath);
    console.log(`   Original File: ${downloadedFile} (Size: ${stats.size} bytes)`);
    
    if (stats.size === 0) {
        throw new Error(`Downloaded file ${downloadedFile} is empty (0 bytes)!`);
    }

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

    console.log('🚀 Starting DTC Automation (Direct File Attachment Mode)...');
    
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
        // STEP 2: REPORT 1 - Over Speed
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        
        // Go to Report Page
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        
        // Fill Form
        console.log('   Filling Form...');
        
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        
        // รอสักนิดเพื่อให้ตัวเลือกใน Dropdown โหลดมาครบ
        await new Promise(r => setTimeout(r, 2000));

        // คำนวณเวลา 06:00 - 18:00 ของวันนี้
        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`   Setting Time: ${startDateTime} to ${endDateTime}`);

        await page.evaluate((start, end) => {
            // Speed (Command 8)
            document.getElementById('speed_max').value = '55';
            
            // Date Formula
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            
            // Trigger Events
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));

            // Options
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            
            // --- Select Truck ---
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

        // Search
        console.log('   Searching...');
        await page.evaluate(() => {
            if(typeof sertch_data === 'function') sertch_data();
            else document.querySelector("span[onclick='sertch_data();']").click();
        });

        // Wait for Data Loading
        console.log('   Waiting for Data Loading...');
        try {
            await page.waitForSelector('#btnexport', { visible: true, timeout: 300000 }); // รอสูงสุด 5 นาที
            await new Promise(r => setTimeout(r, 5000)); 
        } catch(e) {
            console.warn('   ⚠️ Warning: Export button wait timed out');
        }
        console.log('   ✅ Data Loaded.');

        // Export & Download
        console.log('   Exporting...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        // ใช้ Helper Function เพื่อเปลี่ยนชื่อไฟล์
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');


        // =================================================================
        // STEP 3-6: Other Reports (Placeholder)
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
        
        // อ่านไฟล์ทั้งหมดที่มีในโฟลเดอร์ตอนนี้
        const allFiles = fs.readdirSync(downloadPath);
        
        // กรองเฉพาะไฟล์ที่เราต้องการส่ง (ตัดไฟล์ระบบทิ้งถ้ามี)
        const validFiles = allFiles.filter(file => file.endsWith('.xls') || file.endsWith('.xlsx') || file.endsWith('.pdf'));
        
        const attachments = validFiles.map(file => {
            const filePath = path.join(downloadPath, file);
            const stats = fs.statSync(filePath);
            console.log(`   Attaching: ${file} (${stats.size} bytes)`);
            return {
                filename: file,
                path: filePath
            };
        });

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
