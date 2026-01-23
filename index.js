const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

// ฟังก์ชันรอการดาวน์โหลดและเปลี่ยนชื่อไฟล์ (สำคัญมากสำหรับ 5 รายงาน)
async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 60 วินาที
    for (let i = 0; i < 60; i++) {
        const files = fs.readdirSync(downloadPath);
        // หาไฟล์ Excel (.xls, .xlsx) ที่ไม่ใช่ไฟล์ชั่วคราว (.crdownload)
        downloadedFile = files.find(f => (f.endsWith('.xls') || f.endsWith('.xlsx')) && !f.endsWith('.crdownload') && !f.startsWith('Report_'));
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    const oldPath = path.join(downloadPath, downloadedFile);
    const newPath = path.join(downloadPath, newFileName); // เปลี่ยนชื่อไฟล์เพื่อไม่ให้ทับกัน
    
    // ลบไฟล์ปลายทางถ้ามีอยู่แล้ว
    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    
    fs.renameSync(oldPath, newPath);
    console.log(`   ✅ Saved as: ${newFileName}`);
    return newPath;
}

// ฟังก์ชันวันที่ YYYY-MM-DD
function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

// --- Main Script ---

(async () => {
    // 1. ตรวจสอบ Secrets
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing DTC_USERNAME or DTC_PASSWORD secrets.');
        process.exit(1);
    }

    // 2. เตรียมโฟลเดอร์ Downloads
    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (5 Reports)...');
    
    const browser = await puppeteer.launch({
        headless: true, // หรือ "new"
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    // ตั้งค่า Download Path
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    
    await page.setViewport({ width: 1920, height: 1080 });
    await page.emulateTimezone('Asia/Bangkok');

    try {
        // =================================================================
        // STEP 1: LOGIN (ใช้ร่วมกันทุกรายงาน)
        // =================================================================
        console.log('🔑 Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'networkidle2' });
        await page.waitForSelector('#txtname', { visible: true });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        
        await Promise.all([
            page.click('#btnLogin'),
            page.waitForNavigation({ waitUntil: 'networkidle2' })
        ]);
        console.log('   Login Success.');

        // =================================================================
        // STEP 2: REPORT 1 - Over Speed (แก้เวลาเป็น 06:00 - 18:00)
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        
        // ไปหน้ารายงาน
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_other_status.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });

        // เลือกทะเบียน "ทั้งหมด"
        await page.waitForSelector('#ddl_truck');
        await page.evaluate(() => {
            const select = document.getElementById('ddl_truck');
            for (let opt of select.options) {
                if (opt.text.includes('ทั้งหมด') || opt.text.toLowerCase().includes('all')) {
                    select.value = opt.value;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
        });

        // ตั้งเวลา 06:00 - 18:00
        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;

        console.log(`   Setting time: ${startDateTime} to ${endDateTime}`);
        
        // Clear และพิมพ์ค่าใหม่
        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', startDateTime);
        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', endDateTime);

        // กดค้นหา
        console.log('   Searching...');
        await page.click('td:nth-of-type(5) > span'); // ปุ่มค้นหา
        // รอข้อมูลโหลด (ปรับตามความเหมาะสม)
        await new Promise(r => setTimeout(r, 60000)); 

        // กด Export
        console.log('   Exporting Report 1...');
        const btnExportSelector = '#btnexport'; // เช็ค Selector ให้ชัวร์
        await page.waitForSelector(btnExportSelector, { visible: true });
        await page.click(btnExportSelector);

        // รอโหลดและเปลี่ยนชื่อไฟล์เป็น Report1_OverSpeed.xls
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');


        // =================================================================
        // STEP 3: REPORT 2 (พื้นที่สำหรับแปะโค้ด)
        // =================================================================
        console.log('📊 Processing Report 2...');
        // --- เริ่มต้นแปะโค้ด Puppeteer Replay สำหรับรายงานที่ 2 ตรงนี้ ---
        // ตัวอย่าง:
        // await page.goto('URL_OF_REPORT_2');
        // ... โค้ดเลือกเงื่อนไข ...
        // ... โค้ดกด Export ...
        // -----------------------------------------------------------
        
        // รอโหลดและเปลี่ยนชื่อไฟล์ (Uncomment เมื่อมีโค้ดแล้ว)
        // await waitForDownloadAndRename(downloadPath, 'Report2_Name.xls');


        // =================================================================
        // STEP 4: REPORT 3
        // =================================================================
        console.log('📊 Processing Report 3...');
        // --- แปะโค้ด Report 3 ตรงนี้ ---
        
        
        // await waitForDownloadAndRename(downloadPath, 'Report3_Name.xls');


        // =================================================================
        // STEP 5: REPORT 4
        // =================================================================
        console.log('📊 Processing Report 4...');
        // --- แปะโค้ด Report 4 ตรงนี้ ---
        
        
        // await waitForDownloadAndRename(downloadPath, 'Report4_Name.xls');


        // =================================================================
        // STEP 6: REPORT 5
        // =================================================================
        console.log('📊 Processing Report 5...');
        // --- แปะโค้ด Report 5 ตรงนี้ ---
        
        
        // await waitForDownloadAndRename(downloadPath, 'Report5_Name.xls');


        // =================================================================
        // STEP 7: Generate PDF (ทำภายหลัง)
        // =================================================================
        console.log('📑 Generating PDF Summary (Pending implementation)...');
        // ตรงนี้เราจะเขียน Logic อ่านไฟล์ Excel ทั้ง 5 ไฟล์ แล้วสร้าง PDF
        // ตาม Prompt ที่คุณเตรียมไว้


        // =================================================================
        // STEP 8: Send Email (ทำภายหลัง)
        // =================================================================
        console.log('📧 Sending Email (Pending implementation)...');
        // แนบไฟล์ PDF และ Excel ทั้ง 5 ไฟล์ส่งเมล


    } catch (err) {
        console.error('❌ Fatal Error:', err);
        await page.screenshot({ path: path.join(downloadPath, 'error_screenshot.png') });
        process.exit(1);
    } finally {
        await browser.close();
        console.log('🏁 Browser Closed.');
    }
})();
