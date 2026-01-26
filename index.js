const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 120 วินาที
    for (let i = 0; i < 120; i++) {
        const files = fs.readdirSync(downloadPath);
        
        // LOGIC: หาไฟล์ .xls/.xlsx ที่ไม่ใช่ไฟล์ชั่วคราว และไม่มี Prefix "DTC_Completed_"
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

    // รอให้ไฟล์เขียนเสร็จสมบูรณ์
    await new Promise(resolve => setTimeout(resolve, 3000));

    const oldPath = path.join(downloadPath, downloadedFile);
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    // ตรวจสอบขนาดไฟล์
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

    console.log('🚀 Starting DTC Automation (Updated Reports 3, 4, 5 Conditions)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    // Timeout รวม 30 นาที
    page.setDefaultNavigationTimeout(1800000);
    page.setDefaultTimeout(1800000);
    
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
        
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`🕒 Global Time Settings: ${startDateTime} to ${endDateTime}`);

        // =================================================================
        // STEP 2: REPORT 1 - Over Speed
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            
            // เลือกทะเบียน "ทั้งหมด"
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


        // =================================================================
        // STEP 3: REPORT 2 - Idling
        // =================================================================
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '10';

            // เลือกทะเบียน "ทั้งหมด"
            var selectElement = document.getElementById('ddl_truck'); 
            if (selectElement) {
                var options = selectElement.options; 
                for (var i = 0; i < options.length; i++) { 
                    if (options[i].text.includes('ทั้งหมด')) { selectElement.value = options[i].value; break; } 
                } 
                selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        console.log('   Searching Report 2...');
        await page.click('td:nth-of-type(6) > span');

        console.log('   ⏳ Waiting 5 mins...');
        await new Promise(resolve => setTimeout(resolve, 300000));

        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        console.log('   Exporting Report 2...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');


        // =================================================================
        // STEP 4: REPORT 3 - Sudden Brake (เบรกกะทันหัน)
        // =================================================================
        console.log('📊 Processing Report 3: Sudden Brake...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_hd.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true }); // รอ Dropdown
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));

            // เลือกทะเบียน "ทั้งหมด" (Updated)
            var selectElement = document.getElementById('ddl_truck'); 
            if (selectElement) {
                var options = selectElement.options; 
                for (var i = 0; i < options.length; i++) { 
                    if (options[i].text.includes('ทั้งหมด')) { selectElement.value = options[i].value; break; } 
                } 
                selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        console.log('   Searching Report 3...');
        await page.click('td:nth-of-type(6) > span');

        console.log('   ⏳ Waiting 2 mins...');
        await new Promise(resolve => setTimeout(resolve, 120000));

        console.log('   Exporting Report 3...');
        await page.evaluate(() => {
            const buttons = Array.from(document.querySelectorAll('button'));
            const excelBtn = buttons.find(b => b.innerText.includes('Excel') || b.getAttribute('title') === 'Excel' || b.getAttribute('aria-label') === 'Excel');
            if (excelBtn) excelBtn.click();
            else {
                const fallback = document.querySelector('#table button:nth-of-type(3)');
                if (fallback) fallback.click();
            }
        });
        
        await waitForDownloadAndRename(downloadPath, 'Report3_SuddenBrake.xls');


        // =================================================================
        // STEP 5: REPORT 4 - Harsh Start (ออกตัวกระชาก)
        // =================================================================
        console.log('📊 Processing Report 4: Harsh Start...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_ha.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true }); // รอ Dropdown
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));

            // เลือกทะเบียน "ทั้งหมด" (Updated)
            var selectElement = document.getElementById('ddl_truck'); 
            if (selectElement) {
                var options = selectElement.options; 
                for (var i = 0; i < options.length; i++) { 
                    if (options[i].text.includes('ทั้งหมด')) { selectElement.value = options[i].value; break; } 
                } 
                selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }, startDateTime, endDateTime);

        console.log('   Searching Report 4...');
        await page.click('td:nth-of-type(6) > span');

        console.log('   ⏳ Waiting 2 mins...');
        await new Promise(resolve => setTimeout(resolve, 120000));

        console.log('   Exporting Report 4...');
        await page.evaluate(() => {
            const buttons = Array.from(document.querySelectorAll('button'));
            const excelBtn = buttons.find(b => b.innerText.includes('Excel') || b.getAttribute('title') === 'Excel' || b.getAttribute('aria-label') === 'Excel');
            if (excelBtn) excelBtn.click();
            else {
                const fallback = document.querySelector('#table button:nth-of-type(3)');
                if (fallback) fallback.click();
            }
        });
        
        await waitForDownloadAndRename(downloadPath, 'Report4_HarshStart.xls');


        // =================================================================
        // STEP 6: REPORT 5 - Forbidden Parking (พื้นที่ห้ามจอด/เข้าสถานี)
        // =================================================================
        console.log('📊 Processing Report 5: Forbidden Parking...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_Instation.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate((start, end) => {
            // 1. วันที่
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));

            // 2. เลือกทะเบียน "ทั้งหมด" (Updated)
            var truckSelect = document.getElementById('ddl_truck'); 
            if (truckSelect) {
                for (var i = 0; i < truckSelect.options.length; i++) { 
                    if (truckSelect.options[i].text.includes('ทั้งหมด')) { truckSelect.value = truckSelect.options[i].value; break; } 
                } 
                truckSelect.dispatchEvent(new Event('change', { bubbles: true }));
            }

            // 3. เลือกประเภทสถานี "พื้นที่ห้ามเข้า" (Updated)
            // ค้นหา Select Element ทุกตัว เพื่อหาตัวที่มี Option นี้
            var allSelects = document.getElementsByTagName('select');
            for(var s of allSelects) {
                for(var i=0; i<s.options.length; i++) {
                    if(s.options[i].text.includes('พื้นที่ห้ามเข้า')) {
                        s.value = s.options[i].value;
                        s.dispatchEvent(new Event('change', { bubbles: true }));
                        break;
                    }
                }
            }
        }, startDateTime, endDateTime);

        // รอสักครู่เพื่อให้ Dropdown สถานีโหลดใหม่ตามประเภท
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate(() => {
            // 4. เลือกสถานี "สถานีทั้งหมด" (Updated)
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

        console.log('   Searching Report 5...');
        await page.click('td:nth-of-type(7) > span');

        console.log('   ⏳ Waiting 5 mins...');
        await new Promise(resolve => setTimeout(resolve, 300000));

        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        console.log('   Exporting Report 5...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        await waitForDownloadAndRename(downloadPath, 'Report5_ForbiddenParking.xls');


        // =================================================================
        // STEP 7: Generate PDF (Pending)
        // =================================================================
        console.log('📑 Generating PDF Summary (Pending)...');


        // =================================================================
        // STEP 8: Send Email
        // =================================================================
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
                subject: `รายงาน DTC Report (5 ฉบับ) ประจำวันที่ ${todayStr}`,
                text: `เรียน ผู้เกี่ยวข้อง,\n\nระบบส่งไฟล์รายงานจำนวน ${attachments.length} ฉบับ ดังแนบ\n(ข้อมูลช่วงเวลา 06:00 - 18:00)\n\nขอบคุณครับ\nDTC Automation Bot`,
                attachments: attachments
            });
            console.log(`   ✅ Email Sent Successfully! (${attachments.length} files)`);
        } else {
            console.warn('⚠️ No "DTC_Completed_" files found to send!');
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
