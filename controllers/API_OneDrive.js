const fs = require('fs');
const fsPromises = fs.promises;
const path = require('path');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

// -------------------------------------------------------------
// MSAL & Graph Client Setup
// -------------------------------------------------------------
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};
const cca = new ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
    const authResponse = await cca.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default'],
    });
    return Client.init({
        authProvider: (done) => done(null, authResponse.accessToken)
    });
}

function getFormattedDate() {
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    return `${day}-${month}-${year}`;
}

// -------------------------------------------------------------
// Helper: สร้าง/ตรวจสอบ โฟลเดอร์ ปลายทางเพียงครั้งเดียว
// -------------------------------------------------------------
async function ensureOneDriveFolder(client, targetUser, folderPath) {
    const folders = folderPath.split('/').filter(f => f.trim() !== '');
    let currentPath = '';

    for (const folder of folders) {
        currentPath = currentPath ? `${currentPath}/${folder}` : folder;
        try {
            await client.api(`/users/${targetUser}/drive/root:/${currentPath}`).get();
        } catch (err) {
            if (err.statusCode === 404) {
                const parentPath = currentPath.includes('/') 
                    ? `root:/${currentPath.substring(0, currentPath.lastIndexOf('/'))}:` 
                    : 'root';
                
                await client.api(`/users/${targetUser}/drive/${parentPath}/children`).post({
                    name: folder,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'replace'
                });
            } else {
                throw err;
            }
        }
    }
}

// -------------------------------------------------------------
// Helper: อัปโหลดไฟล์ขนาดใหญ่ (> 4MB) แบบ Upload Session
// -------------------------------------------------------------
async function uploadLargeFile(graphClient, targetUser, targetOneDriveFolder, fileName, filePath) {
    const stats = await fsPromises.stat(filePath);
    const fileSize = stats.size;

    const sessionUrl = `/users/${targetUser}/drive/root:/${targetOneDriveFolder}/${fileName}:/createUploadSession`;
    const uploadSession = await graphClient.api(sessionUrl).post({
        item: { 
            '@microsoft.graph.conflictBehavior': 'replace', 
            name: fileName 
        }
    });

    const chunkSize = 320 * 1024 * 10; // Chunk size: 3.2 MB
    const fileStream = fs.createReadStream(filePath, { highWaterMark: chunkSize });
    let start = 0;
    let result = null;

    for await (const chunk of fileStream) {
        const end = start + chunk.length - 1;
        const res = await fetch(uploadSession.uploadUrl, {
            method: 'PUT',
            headers: {
                'Content-Length': chunk.length,
                'Content-Range': `bytes ${start}-${end}/${fileSize}`
            },
            body: chunk
        });

        if (res.status === 200 || res.status === 201) {
            result = await res.json();
        }
        start += chunk.length;
    }
    return result;
}

// -------------------------------------------------------------
// Helper: ประมวลผลไฟล์เดี่ยว + ลบไฟล์ทันทีหลังอัปโหลดสำเร็จ
// -------------------------------------------------------------
async function processSingleFile(client, targetUser, targetOneDriveFolder, localFolderPath, fileName) {
    const filePath = path.join(localFolderPath, fileName);
    try {
        const stats = await fsPromises.stat(filePath);

        if (stats.isDirectory()) return null;

        let uploadRes = null;
        let isSuccess = false;

        if (stats.size <= 4 * 1024 * 1024) {
            // ไฟล์ขนาด <= 4MB
            const fileBuffer = await fsPromises.readFile(filePath);
            uploadRes = await client
                .api(`/users/${targetUser}/drive/root:/${targetOneDriveFolder}/${fileName}:/content`)
                .put(fileBuffer);

            if (uploadRes && uploadRes.id) {
                isSuccess = true;
            }
        } else {
            // ไฟล์ขนาด > 4MB
            uploadRes = await uploadLargeFile(client, targetUser, targetOneDriveFolder, fileName, filePath);
            if (uploadRes && uploadRes.id) {
                isSuccess = true;
            }
        }

        // *** ถ้าอัปโหลดสำเร็จ ให้สั่งลบไฟล์ออกจาก เครื่อง Local ทันที ***
        if (isSuccess) {
            await fsPromises.unlink(filePath);
            return { fileName, size: stats.size, id: uploadRes.id, status: 'Success & Deleted' };
        } else {
            return { fileName, status: 'Failed', error: 'Upload finished but no file ID returned' };
        }

    } catch (err) {
        // หากอัปโหลดหรือลบไฟล์ไม่สำเร็จ จะไม่ทำให้ไฟล์อื่นล่ม และไฟล์เดิมจะยังอยู่บนเครื่อง
        console.error(`Error processing ${fileName}:`, err.message);
        return { fileName, status: 'Failed', error: err.message };
    }
}

// -------------------------------------------------------------
// Helper: ประมวลผลอัปโหลดแบบ Parallel Batches (Concurrency Control)
// -------------------------------------------------------------
async function uploadFilesInParallel(client, targetUser, targetOneDriveFolder, localFolderPath, allowedExtensions, concurrency = 5) {
    const files = await fsPromises.readdir(localFolderPath);
    
    // กรองไฟล์ตามนามสกุลที่กำหนด
    const validFiles = files.filter(fileName => {
        const ext = path.extname(fileName).toLowerCase();
        return allowedExtensions.includes(ext);
    });

    const results = [];
    
    // แบ่งกลุ่มทำทีละ Batch พร้อมกันตามจำนวน concurrency
    for (let i = 0; i < validFiles.length; i += concurrency) {
        const batch = validFiles.slice(i, i + concurrency);
        const batchPromises = batch.map(fileName => 
            processSingleFile(client, targetUser, targetOneDriveFolder, localFolderPath, fileName)
        );
        
        const batchResults = await Promise.all(batchPromises);
        results.push(...batchResults.filter(r => r !== null));
    }

    return results;
}

// -------------------------------------------------------------
// 1. Endpoint: อัปโหลด Excel
// -------------------------------------------------------------
exports.uploadExcelFiles = async (req, res) => {
    try {
        const { folderName, date } = req.body;

        if (!folderName) {
            return res.status(400).json({ error: 'กรุณาระบุ folderName ใน Body (เช่น { "folderName": "MSC" })' });
        }

        const client = await getGraphClient();
        const targetUser = process.env.TARGET_USER;
        const dateFolder = date || getFormattedDate();

        const localFolderPath = path.join(process.env.EXCEL_BASE_PATH, folderName);
        const targetOneDriveFolder = `WICE x INET/Online/${dateFolder}/${folderName}`;
        const allowedExtensions = ['.xlsx', '.xls', '.csv'];

        if (!fs.existsSync(localFolderPath)) {
            return res.status(400).json({ error: 'ไม่พบโฟลเดอร์ต้นทางฝั่ง Local', path: localFolderPath });
        }

        await ensureOneDriveFolder(client, targetUser, targetOneDriveFolder);

        const results = await uploadFilesInParallel(
            client, 
            targetUser, 
            targetOneDriveFolder, 
            localFolderPath, 
            allowedExtensions, 
            5
        );

        const successCount = results.filter(r => r.status.startsWith('Success')).length;
        const failedCount = results.filter(r => r.status === 'Failed').length;

        return res.status(200).json({ 
            message: 'อัปโหลดและลบไฟล์ Excel ต้นทางเรียบร้อยแล้ว', 
            localPathUsed: localFolderPath,
            targetOneDrivePath: targetOneDriveFolder,
            summary: {
                total: results.length,
                successDeleted: successCount,
                failedKeepLocal: failedCount
            },
            details: results 
        });

    } catch (err) {
        console.error('Upload Excel Main Error:', err);
        return res.status(500).json({ error: 'เกิดข้อผิดพลาดในกระบวนการอัปโหลด', message: err.message });
    }
};

// -------------------------------------------------------------
// 2. Endpoint: อัปโหลด รูปภาพ & ไฟล์ PDF
// -------------------------------------------------------------
exports.uploadImageFiles = async (req, res) => {
    try {
        const { folderName, date } = req.body;

        if (!folderName) {
            return res.status(400).json({ error: 'กรุณาระบุ folderName ใน Body (เช่น { "folderName": "MSC" })' });
        }

        const client = await getGraphClient();
        const targetUser = process.env.TARGET_USER;
        const dateFolder = date || getFormattedDate();

        const localFolderPath = path.join(process.env.IMAGE_BASE_PATH, folderName);
        const targetOneDriveFolder = `WICE x INET/Online/${dateFolder}/Image/${folderName}`;
        const allowedExtensions = ['.png', '.jpg', '.jpeg', '.webp', '.pdf'];

        if (!fs.existsSync(localFolderPath)) {
            return res.status(400).json({ error: 'ไม่พบโฟลเดอร์ต้นทางฝั่ง Local', path: localFolderPath });
        }

        await ensureOneDriveFolder(client, targetUser, targetOneDriveFolder);

        const results = await uploadFilesInParallel(
            client, 
            targetUser, 
            targetOneDriveFolder, 
            localFolderPath, 
            allowedExtensions, 
            5
        );

        const successCount = results.filter(r => r.status.startsWith('Success')).length;
        const failedCount = results.filter(r => r.status === 'Failed').length;

        return res.status(200).json({ 
            message: 'อัปโหลดและลบไฟล์รูปภาพ/PDF ต้นทางเรียบร้อยแล้ว', 
            localPathUsed: localFolderPath,
            targetOneDrivePath: targetOneDriveFolder,
            summary: {
                total: results.length,
                successDeleted: successCount,
                failedKeepLocal: failedCount
            },
            details: results 
        });

    } catch (err) {
        console.error('Upload Image Main Error:', err);
        return res.status(500).json({ error: 'เกิดข้อผิดพลาดในกระบวนการอัปโหลด', message: err.message });
    }
};