const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    for (let i = 0; i < 60; i++) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => (f.endsWith('.xls') || f.endsWith('.xlsx')) && !f.endsWith('.crdownload') && !f.startsWith('Report_'));
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    const oldPath = path.join(downloadPath, downloadedFile);
    const newPath = path.join(downloadPath, newFileName);
    
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
    // เคลียร์ไฟล์เก่าก่อนเริ่มทำงานเสมอ
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (Full Flow)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
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


        // Step 2: Report 1 (Over Speed)
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_other_status.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true });
        
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

        const todayStr = getTodayFormatted();
        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', `${todayStr} 06:00`);
        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', `${todayStr} 18:00`);

        console.log('   Searching Report 1...');
        await page.click('td:nth-of-type(5) > span');
        await new Promise(r => setTimeout(r, 60000)); // รอโหลดข้อมูล

        console.log('   Exporting Report 1...');
        await page.waitForSelector('#btnexport', { visible: true });
        await page.click('#btnexport');
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // Step 3-6: Other Reports (Placeholder)
        // ... (ใส่ Code สำหรับ Report 2-5 ตรงนี้ และใช้ waitForDownloadAndRename ตามลำดับ) ...

        // Step 7: Generate PDF (Placeholder)
        console.log('📑 Generating PDF Summary (Pending)...');
        // TODO: ใส่ Logic สร้าง PDF ตรงนี้ และ save ไฟล์เป็น 'Summary_Report.pdf' ลงใน downloadPath

        // Step 8: Send Email
        console.log('📧 Step 8: Sending Email...');
        
        // อ่านรายชื่อไฟล์ทั้งหมดใน folder downloads เพื่อแนบไปกับเมล์
        const allFiles = fs.readdirSync(downloadPath);
        const attachments = allFiles.map(file => ({
            filename: file,
            path: path.join(downloadPath, file)
        }));

        if (attachments.length === 0) {
            console.warn('⚠️ No files to send!');
        } else {
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
        }

        // Step 9: Cleanup Files
        console.log('🧹 Step 9: Cleaning up files...');
        const filesToDelete = fs.readdirSync(downloadPath);
        for (const file of filesToDelete) {
            try {
                fs.unlinkSync(path.join(downloadPath, file));
                console.log(`   Deleted: ${file}`);
            } catch (err) {
                console.error(`   Failed to delete ${file}:`, err.message);
            }
        }
        console.log('   ✅ Cleanup Complete.');

    } catch (err) {
        console.error('❌ Fatal Error:', err);
        // ถ่ายรูปตอน Error เก็บไว้ (จะถูก Upload ขึ้น GitHub Artifacts)
        await page.screenshot({ path: path.join(downloadPath, 'error_screenshot.png') });
        process.exit(1);
    } finally {
        await browser.close();
        console.log('🏁 Browser Closed.');
    }
})();
