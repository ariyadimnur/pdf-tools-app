document.addEventListener('DOMContentLoaded', () => {
    // State manajemen untuk menampung file sementara
    let selectedFiles = {
        merge: [],
        convert: [],
        compress: []
    };

    const backendUrl = 'http://localhost:8989/api';

    // --- LOGIKA TAB SWITCHING ---
    const navButtons = document.querySelectorAll('.nav-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            navButtons.forEach(b => b.classList.remove('active'));
            tabContents.forEach(t => t.classList.remove('active'));

            btn.classList.add('active');
            const targetTab = btn.getAttribute('data-tab');
            document.getElementById(`tab-${targetTab}`).classList.add('active');
        });
    });

    // --- SETUP DRAG & DROP DAN TRIGGER INPUT FILE ---
    setupDropzone('merge', true);
    setupDropzone('convert', false);
    setupDropzone('compress', false);

    function setupDropzone(type, isMultiple) {
        const dropzone = document.getElementById(`dropzone-${type}`);
        const fileInput = document.getElementById(`file-${type}`);

        dropzone.addEventListener('click', () => fileInput.click());

        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        });

        dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
            handleFileSelection(e.dataTransfer.files, type, isMultiple);
        });

        fileInput.addEventListener('change', (e) => {
            handleFileSelection(e.target.files, type, isMultiple);
        });
    }

    // --- HANDLER MANAJEMEN FILE ---
    function handleFileSelection(files, type, isMultiple) {
        const validFiles = Array.from(files).filter(file => file.type === 'application/pdf');
        
        if (isMultiple) {
            // Batasi maksimal 50 file
            if (selectedFiles[type].length + validFiles.length > 50) {
                alert('Maksimal penggabungan adalah 50 file PDF.');
                return;
            }
            selectedFiles[type] = [...selectedFiles[type], ...validFiles];
        } else {
            if (validFiles.length > 0) {
                selectedFiles[type] = [validFiles[0]]; // Hanya ambil 1 file utama
            }
        }

        renderPreviews(type);
    }

    function renderPreviews(type) {
        const container = document.getElementById(`preview-${type}`);
        const actionBtn = document.getElementById(`btn-${type}`);
        container.innerHTML = '';

        selectedFiles[type].forEach((file, index) => {
            const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <div class="file-info">
                    <i class="fa-solid fa-file-pdf"></i>
                    <div>
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${sizeInMB} MB</div>
                    </div>
                </div>
                <button class="remove-file" data-index="${index}"><i class="fa-solid fa-trash-can"></i></button>
            `;
            container.appendChild(item);
        });

        // Event listener untuk tombol hapus item file
        container.querySelectorAll('.remove-file').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const index = btn.getAttribute('data-index');
                selectedFiles[type].splice(index, 1);
                renderPreviews(type);
            });
        });

        // Aktifkan / matikan tombol eksekusi utama
        actionBtn.disabled = selectedFiles[type].length === 0;
    }

    // --- PROSES SUBMIT / HIT API KE BACKEND ---
    const overlay = document.getElementById('loading-overlay');
    const loadingText = document.getElementById('loading-text');

    // 1. Aksi Gabungkan PDF
    document.getElementById('btn-merge').addEventListener('click', async () => {
        const formData = new FormData();
        selectedFiles.merge.forEach(file => formData.append('pdfs', file));

        showLoading('Menggabungkan dokumen PDF...');

        try {
            const response = await fetch(`${backendUrl}/pdf/merge`, { method: 'POST', body: formData });
            if (!response.ok) throw new Error('Gagal memproses penggabungan.');
            
            await downloadBlob(response, 'hasil_gabungan.pdf');
            selectedFiles.merge = [];
            renderPreviews('merge');
        } catch (err) {
            alert(err.message);
        } finally {
            hideLoading();
        }
    });

    // 2. Aksi Convert PDF (Word / Excel)
    document.getElementById('btn-convert').addEventListener('click', async () => {
        const targetFormat = document.querySelector('input[name="convert-target"]:checked').value;
        const formData = new FormData();
        formData.append('pdf', selectedFiles.convert[0]);

        showLoading(`Mengonversi PDF ke format ${targetFormat.toUpperCase()}...`);

        // Menentukan endpoint dinamis berdasarkan opsi radio yang dipilih
        const endpoint = targetFormat === 'word' ? '/pdf/to-word' : '/pdf/to-excel';
        const extension = targetFormat === 'word' ? '.docx' : '.xlsx';

        try {
            const response = await fetch(`${backendUrl}${endpoint}`, { method: 'POST', body: formData });
            if (!response.ok) throw new Error('Gagal mengonversi file.');

            await downloadBlob(response, `hasil_konversi${extension}`);
            selectedFiles.convert = [];
            renderPreviews('convert');
        } catch (err) {
            alert(err.message);
        } finally {
            hideLoading();
        }
    });

    // 3. Aksi Compress PDF (Tambahkan kode ini di bagian bawah sebelum HELPER FUNCTIONS)
    document.getElementById('btn-compress').addEventListener('click', async () => {
        if (selectedFiles.compress.length === 0) return;

        const formData = new FormData();
        formData.append('pdf', selectedFiles.compress[0]);

        showLoading('Sedang memperkecil ukuran file PDF...');

        try {
            const response = await fetch(`${backendUrl}/pdf/compress`, { 
                method: 'POST', 
                body: formData 
            });
            
            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.error || 'Gagal mengompres file.');
            }

            await downloadBlob(response, 'hasil_compress.pdf');
            selectedFiles.compress = [];
            renderPreviews('compress');
        } catch (err) {
            alert(err.message);
        } finally {
            hideLoading();
        }
    });

    // --- HELPER FUNCTIONS ---
    function showLoading(text) {
        loadingText.innerText = text;
        overlay.classList.add('active');
    }

    function hideLoading() {
        overlay.classList.remove('active');
    }

    async function downloadBlob(response, defaultFilename) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = defaultFilename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    }
});