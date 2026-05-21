const express = require('express');
const multer = require('multer');
const cors = require('cors');
const PDFMerger = require('pdf-merger-js');
const pdfParse = require('pdf-parse');
const docx = require('docx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const pdfTableExtractor = require('pdf-table-extractor');
const PDFParser = require("pdf2json");
const axios = require('axios');
const FormData = require('form-data');
const XLSX = require('xlsx');


const app = express();
const PORT = process.env.PORT || 8989;

// Middleware
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Pastikan folder 'uploads' ada saat aplikasi dinyalakan oleh PM2
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)){
    fs.mkdirSync(uploadDir);
}

// Konfigurasi Multer untuk handle multi-upload (Maksimal 50 file PDF)
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, 'uploads/'),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage: storage });

// ==========================================
// 1. FITUR: COMPILE / MERGE PDF (MENGGUNAKAN PDF-LIB) - PASTI BERHASIL
// ==========================================
app.post('/api/pdf/merge', upload.array('pdfs', 50), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'Tidak ada file yang diunggah.' });
        }

        // Buat satu dokumen PDF kosong yang baru
        const mergedPdf = await PDFDocument.create();

        // Urutkan file berdasarkan nama asli agar susunannya teratur
        const sortedFiles = req.files.sort((a, b) => a.originalname.localeCompare(b.originalname));

        // Proses membaca dan menggabungkan halaman satu per satu
        for (const file of sortedFiles) {
            const fileBytes = fs.readFileSync(file.path);
            const srcPdf = await PDFDocument.load(fileBytes);
            
            // Ambil semua halaman dari file PDF saat ini
            const copiedPages = await mergedPdf.copyPages(srcPdf, srcPdf.getPageIndices());
            
            // Masukkan halaman-halaman tersebut ke dokumen PDF baru
            copiedPages.forEach(page => mergedPdf.addPage(page));
        }

        // Simpan dokumen hasil penggabungan menjadi data bytes
        const mergedPdfBytes = await mergedPdf.save();

        const outputPath = path.join(uploadDir, `merged_${Date.now()}.pdf`);
        fs.writeFileSync(outputPath, mergedPdfBytes);

        // Langsung bersihkan file mentah yang diupload di folder uploads
        req.files.forEach(file => { 
            if (fs.existsSync(file.path)) fs.unlinkSync(file.path); 
        });

        // Kirim hasil file gabungan ke browser pengguna
        res.download(outputPath, 'hasil_gabungan.pdf', () => {
            if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
        });

    } catch (error) {
        console.error("Error pada Proses Merge PDF:", error);
        
        // Hapus sisa file temporary jika di tengah jalan terjadi error agar Drive D bersih
        if (req.files) {
            req.files.forEach(file => { 
                if (fs.existsSync(file.path)) fs.unlinkSync(file.path); 
            });
        }
        res.status(500).json({ error: 'Gagal menggabungkan PDF: ' + error.message });
    }
});

// ==========================================
// 3. FITUR KHUSUS: PDF TO EXCEL (HIGH-PRECISION COORDINATE TABLE ENGINE)
// ==========================================
app.post('/api/pdf/to-excel', upload.single('pdf'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan.' });

        // Menggunakan PDFParser (pdf2json) yang sudah terbukti sakti di fitur Word om
        const pdfParser = new PDFParser();

        pdfParser.on("pdfParser_dataReady", async (pdfData) => {
            try {
                const excelRows = [];

                // 🔄 LOOPING MEMBACA SEMUA HALAMAN PDF
                pdfData.Pages.forEach((page) => {
                    const texts = page.Texts;

                    // 1. Kelompokkan data berdasarkan koordinat Y (Baris)
                    const rowsMap = {};
                    texts.forEach(t => {
                        // Membulatkan koordinat Y agar teks yang agak miring tetap dianggap satu baris
                        const yKey = Math.round(t.y * 2) / 2; 
                        if (!rowsMap[yKey]) rowsMap[yKey] = [];
                        
                        const txt = decodeURIComponent(t.R[0].T).trim();
                        if (txt !== '') {
                            rowsMap[yKey].push({ x: t.x, text: txt });
                        }
                    });

                    // 2. Urutkan baris dari atas ke bawah
                    const sortedYKeys = Object.keys(rowsMap).sort((a, b) => parseFloat(a) - parseFloat(b));

                    // 3. PECAH TEKS MENJADI KOLOM TERPISAH BERDASARKAN KOORDINAT X (KIRI-KANAN)
                    sortedYKeys.forEach(yKey => {
                        // Urutkan teks di baris tersebut dari kiri ke kanan (Kolom)
                        const rowItems = rowsMap[yKey].sort((a, b) => a.x - b.x);
                        
                        // Ambil teksnya saja dan masukkan ke array array [] untuk cell Excel
                        const columnsData = rowItems.map(item => item.text);

                        if (columnsData.length > 0) {
                            excelRows.push(columnsData); // Masuk sebagai baris baru di Excel
                        }
                    });
                });

                if (excelRows.length === 0) {
                    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
                    return res.status(400).json({ error: 'Gagal mengekstrak tabel dari PDF.' });
                }

                // 4. Proses pembuatan File Excel (.xlsx) murni
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(excelRows);
                
                // Atur lebar kolom otomatis agar tulisan data siswanya tidak terpotong (wch: lebar karakter)
                const maxCols = Math.max(...excelRows.map(r => r.length));
                ws['!cols'] = Array(maxCols).fill({ wch: 25 }); 

                XLSX.utils.book_append_sheet(wb, ws, 'Hasil Konversi');

                const outputPath = path.join(uploadDir, `converted_${Date.now()}.xlsx`);
                XLSX.writeFile(wb, outputPath);

                // Hapus file master PDF lama di lokal
                if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

                // Kirim file Excel yang sudah kotak-kotak rapi ke browser user
                return res.download(outputPath, 'hasil_konversi.xlsx', () => {
                    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
                });

            } catch (innerErr) {
                console.error("Error internal Excel generator:", innerErr);
                if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
                return res.status(500).json({ error: 'Gagal merender struktur tabel Excel.' });
            }
        });

        pdfParser.on("pdfParser_dataError", errData => {
            console.error("Error parser pdf2json Excel:", errData.parserError);
            if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
            return res.status(500).json({ error: 'Gagal menganalisis koordinat PDF.' });
        });

        pdfParser.loadPDF(req.file.path);

    } catch (error) {
        console.error("Error global Excel Convert:", error.message);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: 'Server error pada komponen konversi Excel.' });
    }
});

const CloudmersiveConvertApiClient = require('cloudmersive-convert-api-client');

// ==========================================
// 2. FITUR KHUSUS: PDF TO WORD (PRO HIGH-PRECISION CLOUD CONVERT)
// ==========================================
app.post('/api/pdf/to-word', upload.single('pdf'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan.' });

        // Inisialisasi API Client Cloudmersive
        const defaultClient = CloudmersiveConvertApiClient.ApiClient.instance;
        const Apikey = defaultClient.authentications['Apikey'];
        
        // 🔑 MASUKKAN API KEY GRATAN KAMU DI SINI OM
        Apikey.apiKey = 'fb310ea5-c700-459e-905a-5e5750e3b7f0'; 

        const apiInstance = new CloudmersiveConvertApiClient.ConvertDocumentApi();
        const inputFile = fs.readFileSync(req.file.path); // Membaca file PDF lokal

        const outputPath = path.join(uploadDir, `converted_${Date.now()}.docx`);

        // Panggil fungsi konversi presisi tinggi (PDF to DOCX)
        apiInstance.convertDocumentPdfToDocx(inputFile, (error, data, response) => {
            // Hapus file master PDF yang diupload agar storage server bersih
            if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

            if (error) {
                console.error("Error dari Cloudmersive API:", error);
                return res.status(500).json({ error: 'Gagal melakukan konversi presisi tinggi via cloud.' });
            }

            // Simpan buffer hasil konversi dari cloud menjadi file Word lokal
            fs.writeFileSync(outputPath, data);

            // Kirim file .docx yang sudah jadi dan rapi ke browser user
            return res.download(outputPath, 'hasil_konversi.docx', () => {
                if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
            });
        });

    } catch (error) {
        console.error("Error global Cloud Convert:", error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: 'Server error pada komponen konversi cloud.' });
    }
});

// Taruh ini di bagian paling atas server.js bersama require lainnya
const { PDFDocument, PDFName, PDFRawStream, PDFNumber } = require('pdf-lib');

// ==========================================
// 4. FITUR: COMPRESS PDF (REAL IMAGE DOWNSAMPLE) - FIXED VERSION
// ==========================================
app.post('/api/pdf/compress', upload.single('pdf'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan.' });

        const dataBuffer = fs.readFileSync(req.file.path);
        const pdfDoc = await PDFDocument.load(dataBuffer);
        const pages = pdfDoc.getPages();

        // Looping setiap halaman untuk mencari objek gambar internal
        for (const page of pages) {
            const { node } = page;
            const resources = node.Resources();
            if (!resources) continue;

            const xObjects = resources.get(PDFName.of('XObject'));
            if (!xObjects) continue;

            const xObjectMap = xObjects.entries();
            for (const [name, ref] of xObjectMap) {
                const object = pdfDoc.context.lookup(ref);
                
                // Cek apakah objek tersebut adalah gambar (Subtype Image)
                if (object instanceof PDFRawStream && object.dict.get(PDFName.of('Subtype')) === PDFName.of('Image')) {
                    try {
                        // Ambil data byte gambar asli
                        const imageBytes = object.contents;
                        
                        // Kompres gambar menggunakan Sharp (turunkan quality ke 60%)
                        const compressedImageBuffer = await sharp(imageBytes)
                            .jpeg({ quality: 60, progressive: true })
                            .toBuffer();

                        // Masukkan kembali gambar yang sudah dikompres ke objek PDF tersebut
                        object.contents = new Uint8Array(compressedImageBuffer);
                        object.dict.set(PDFName.of('Length'), PDFNumber.of(compressedImageBuffer.length));
                    } catch (imgError) {
                        // Skip jika ada format gambar internal PDF yang tidak disupport sharp
                        continue;
                    }
                }
            }
        }

        // Simpan PDF dengan kompresi objek biner tambahan
        const compressedPdfBytes = await pdfDoc.save({ useObjectStreams: true });

        const outputPath = path.join(uploadDir, `compressed_${Date.now()}.pdf`);
        fs.writeFileSync(outputPath, compressedPdfBytes);

        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

        res.download(outputPath, 'hasil_compress.pdf', () => {
            if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
        });

    } catch (error) {
        console.error("Error pada proses Compress PDF:", error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: 'Gagal mengompres PDF: ' + error.message });
    }
});

// Ganti angka 3000 yang ada di dalam app.listen menjadi variabel PORT
app.listen(PORT, () => console.log(`Backend Server running on port ${PORT}`));