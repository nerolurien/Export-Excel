// npm i express exceljs
// Fungsi Express di sini:
// Membuat server HTTP kecil
// Menangani route /export saat tombol ditekan
// Menyusun Excel menggunakan exceljs
// Mengirim file Excel sebagai download

const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

app.use(express.static(__dirname)); // untuk akses file index.html dan logo

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/export', async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Laporan');

    // Tambahkan logo
    const logoPath = path.join(__dirname, 'logo.png');
    if (fs.existsSync(logoPath)) {
        const imageId = workbook.addImage({
            filename: logoPath,
            extension: 'png',
        });

        worksheet.addImage(imageId, {
            tl: { col: 0, row: 0 },
            ext: { width: 80, height: 80 },
        });
    }

    // Kop surat
    worksheet.mergeCells('C1:F1');
    worksheet.getCell('C1').value = 'PT. Rafid Teknologi Nusantara';
    worksheet.getCell('C1').font = { size: 14, bold: true };
    worksheet.getCell('C1').alignment = { horizontal: 'center' };

    worksheet.mergeCells('C2:F2');
    worksheet.getCell('C2').value = 'Jl. Inovasi No.1, Kabupaten Bogor, Indonesia';
    worksheet.getCell('C2').alignment = { horizontal: 'center' };

    worksheet.mergeCells('C3:F3');
    worksheet.getCell('C3').value = 'Telepon: (021) 12345678 | Email: info@rafidtech.co.id';
    worksheet.getCell('C3').alignment = { horizontal: 'center' };

    // Garis pembatas
    for (let col = 1; col <= 6; col++) {
        worksheet.getRow(4).getCell(col).border = {
            bottom: { style: 'thin' },
        };
    }

    // Header
    const headers = ['ID', 'Nama Barang', 'Jumlah', 'Harga'];
    worksheet.getRow(6).values = [null, ...headers];
    worksheet.getRow(6).font = { bold: true };
    worksheet.getRow(6).alignment = { horizontal: 'center' };

    // Dummy data
    const data = [
        [1, 'Laptop', 10, 15000000],
        [2, 'Mouse', 25, 75000],
        [3, 'Keyboard', 15, 200000],
        [4, 'Monitor', 7, 1750000],
    ];

    data.forEach((item, index) => {
        worksheet.getRow(7 + index).values = [null, ...item];
    });

    // Styling tabel
    const endRow = 6 + data.length;
    for (let i = 6; i <= endRow; i++) {
        for (let j = 2; j <= 5; j++) {
            const cell = worksheet.getRow(i).getCell(j);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
    }

    // Lebar kolom
    worksheet.getColumn(2).width = 12;
    worksheet.getColumn(3).width = 25;
    worksheet.getColumn(4).width = 10;
    worksheet.getColumn(5).width = 15;

    // Simpan dan kirim ke client
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Laporan-Inventaris.xlsx');

    await workbook.xlsx.write(res);
    res.end();
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
