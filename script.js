// Mengambil elemen dari DOM (pastikan DOM sudah siap)
document.addEventListener('DOMContentLoaded', () => {
    const myForm = document.getElementById('myForm');
    const myInput = document.getElementById('myInput');
    const lokasiInput = document.getElementById('lokasiInput');
    const jamInput = document.getElementById('jamInput');
    const catatanInput = document.getElementById('catatanInput');


    const resultDiv = document.getElementById('resultDiv');
    const averageResult = document.getElementById('averageResult');
    const warningNote = document.getElementById('warningNote');

    // Set default time to current time
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    jamInput.value = hours + ':' + minutes;

    // Verifikasi form sudah terbaca
    if (!myForm || !myInput) {
        console.error('Form atau input tidak ditemukan! myForm:', myForm, 'myInput:', myInput);
    } else {
        console.log('✓ Form dan input berhasil dimuat');
    }

    // Kalau elemen tidak ada, stop di sini.
    if (!myForm || !myInput) return;

    // Array untuk menyimpan semua data
    let dataArray = [];

    // Counter untuk tracking cell mana yang akan diisi
    let cellCounter = 1;

    // Menambah event listener saat form di-submit
    myForm.addEventListener('submit', function (e) {
        e.preventDefault();
        e.stopPropagation();

        const rawValue = myInput.value.trim();
        if (rawValue === '') {
            myInput.focus();
            return;
        }

        // Validasi angka (desibel)
        const value = Number(rawValue.replace(',', '.'));
        if (Number.isNaN(value)) {
            alert('Nilai dB harus berupa angka.');
            myInput.focus();
            return;
        }

        // Masukkan ke cell dan array jika belum penuh
        if (cellCounter <= 10) {
            const cellId = 'cell' + cellCounter;
            const cellEl = document.getElementById(cellId);
            if (cellEl) {
                cellEl.innerText = value;
                console.log('Saved to cell:', cellId, 'Value:', value);
            }

            dataArray[cellCounter - 1] = value;
            cellCounter++;

            // Jika semua 10 cell sudah terisi, hitung rata-rata
            if (cellCounter === 11 && dataArray.length === 10 && dataArray.every(v => typeof v === 'number' && !Number.isNaN(v))) {
                hitungRataRata();
            }
        } else {
            alert('Semua 10 titik sudah terisi! Silakan klik "Ulang Pengukuran" untuk memulai lagi.');
        }

        // Kosongkan input dan lanjut input berikutnya
        myInput.value = '';
        myInput.focus();
    });

    // Event listener untuk edit cell
    const editableCells = document.querySelectorAll('.editable-cell');
    editableCells.forEach((cell) => {
        cell.addEventListener('click', function () {
            if (this.innerText === '-') return;

            const currentValue = this.innerText;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentValue;
            input.className = 'edit-input';

            this.innerText = '';
            this.appendChild(input);
            input.focus();
            input.select();

            const saveEdit = () => {
                const newRawValue = input.value.trim();
                if (newRawValue === '') {
                    this.innerText = currentValue;
                    return;
                }

                const newValue = Number(newRawValue.replace(',', '.'));
                if (Number.isNaN(newValue)) {
                    alert('Nilai dB harus berupa angka.');
                    input.focus();
                    return;
                }

                this.innerText = newValue;

                const cellNumber = parseInt(this.id.replace('cell', ''), 10);
                dataArray[cellNumber - 1] = newValue;

                if (dataArray.length === 10 && dataArray.every(val => typeof val === 'number' && !Number.isNaN(val))) {
                    hitungRataRata();
                }
            };

            input.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') saveEdit();
            });

            input.addEventListener('blur', saveEdit);
        });
    });

    const NAB_BY_DURATION_HOURS = {
        8: 85,
        4: 88,
        2: 91,
        1: 94,
        0.5: 97,
        0.25: 100,
        0.125: 103,
        0.0625: 106,
        0.03125: 109
    };

    function getSelectedDurationNab() {
        const durasiSelect = document.getElementById('durasiSelect');
        const durasiVal = durasiSelect ? Number(durasiSelect.value) : 8;
        const nab = NAB_BY_DURATION_HOURS[durasiVal];
        return {
            durationHours: durasiVal,
            nab: typeof nab === 'number' ? nab : 85
        };
    }

    // Fungsi untuk menghitung rata-rata
    function hitungRataRata() {
        const validData = dataArray.filter(val => typeof val === 'number' && !Number.isNaN(val));
        if (validData.length !== 10) return;

        const total = validData.reduce((sum, num) => sum + num, 0);
        const rataRata = total / 10;

        const { durationHours, nab } = getSelectedDurationNab();

        averageResult.innerText = `${rataRata.toFixed(2)} dB`;

        if (rataRata > nab) {
            averageResult.classList.add('high-value');
            warningNote.style.display = 'block';
        } else {
            averageResult.classList.remove('high-value');
            warningNote.style.display = 'none';
        }

        resultDiv.style.display = 'block';
    }

    function resetTabel() {
        for (let i = 1; i <= 10; i++) {
            document.getElementById('cell' + i).innerText = '-';
        }
        dataArray = [];
        cellCounter = 1;
        resultDiv.style.display = 'none';
        averageResult.innerText = '';
        myInput.focus();
    }

    // Fungsi export ke Excel
    function exportToExcel() {
        if (lokasiInput.value.trim() === '') {
            alert('Mohon isi lokasi terlebih dahulu!');
            lokasiInput.focus();
            return;
        }

        if (jamInput.value.trim() === '') {
            alert('Mohon isi jam terlebih dahulu!');
            jamInput.focus();
            return;
        }

        const currentData = [];
        for (let i = 1; i <= 10; i++) {
            currentData.push(document.getElementById('cell' + i).innerText);
        }

        const rataRata = averageResult.innerText;
        const lokasi = lokasiInput.value;
        const jam = jamInput.value;

        const now = new Date();
        const tanggal = now.toLocaleDateString('id-ID', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });

        const { durationHours, nab } = getSelectedDurationNab();
        const catatanText = catatanInput.value.trim();

        const workbook = XLSX.utils.book_new();

        const sheetData = [['Lokasi', 'Tanggal', 'Jam', 'Durasi Jam (h)', 'NAB (dBA)', 'Titik 1', 'Titik 2', 'Titik 3', 'Titik 4', 'Titik 5', 'Titik 6', 'Titik 7', 'Titik 8', 'Titik 9', 'Titik 10', 'Rata-rata', 'Catatan']];
        
        const row = [
            lokasi,
            tanggal,
            jam,
            durationHours ?? '',
            nab ?? '',
            ...currentData,
            rataRata,
            catatanText
        ];

        sheetData.push(row);

        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        worksheet['!cols'] = [
            { wch: 15 },
            { wch: 12 },
            { wch: 10 },
            { wch: 14 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 10 },
            { wch: 12 },
            { wch: 12 }
        ];

        XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Pengukuran');

        // Nama file mengikuti input user: nama lokasi + tanggal pengukuran
        // Catatan: tidak memakai history yang lama.
        const sanitizeForFileName = (str) => String(str || '')
            .trim()
            .replace(/[\\/:*?"<>|]+/g, '-')
            .replace(/\s+/g, '_');

        const tanggalFilePart = sanitizeForFileName(tanggal).replace(/[\.-]+/g, '-');
        const namaFile = `Hasil_Kebisingan_${sanitizeForFileName(lokasi)}_${tanggalFilePart}.xlsx`;

        XLSX.writeFile(workbook, namaFile);


        // reset dilakukan manual via tombol "Ulang Pengukuran" (tidak otomatis saat export)
        // lokasi tidak diriset di sini


        alert('Data berhasil diekspor! File: data_history.xlsx\nTabel sudah direset.');
    }

    // Helper function to convert image to base64 (dari URL)
    function imageToBase64(src) {

        return new Promise((resolve, reject) => {
            const img = new Image();
            img.crossOrigin = 'anonymous';
            img.onload = () => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.width = img.width;
                canvas.height = img.height;
                ctx.drawImage(img, 0, 0);
                const dataURL = canvas.toDataURL('image/png');
                resolve(dataURL);
            };
            img.onerror = () => reject(new Error(`Failed to load image: ${src}`));
            img.src = src;
        });
    }

    // Fungsi export PDF - menggunakan jsPDF native (lebih andal)
    async function exportToPdf() {
        console.log('=== exportToPdf called ===');
        
        // Cek apakah library yang dibutuhkan sudah dimuat
        if (typeof window.jspdf === 'undefined' || !window.jspdf.jsPDF) {
            console.error('jsPDF library not loaded! window.jspdf:', window.jspdf);
            alert('Library PDF belum dimuat. Silakan refresh halaman dan coba lagi.');
            return;
        }
        console.log('jsPDF library loaded successfully');

        const rataRataText = averageResult.innerText || '';
        if (!rataRataText.trim()) {
            alert('Hitung dulu rata-rata kebisingan sebelum mengunduh PDF.');
            return;
        }

        const lokasi = lokasiInput.value.trim();
        if (!lokasi) {
            alert('Mohon isi lokasi terlebih dahulu!');
            lokasiInput.focus();
            return;
        }

        if (jamInput.value.trim() === '') {
            alert('Mohon isi jam terlebih dahulu!');
            jamInput.focus();
            return;
        }

        // Tampilkan loading indicator
        const exportPdfBtn = document.getElementById('exportPdfBtn');
        const originalText = exportPdfBtn.innerText;
        exportPdfBtn.innerText = '⏳ Memproses...';
        exportPdfBtn.disabled = true;

        const now = new Date();
        const tanggal = now.toLocaleDateString('id-ID', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
        const jam = jamInput.value;
        const { durationHours, nab } = getSelectedDurationNab();
        const catatanText = catatanInput.value.trim();

        // Ambil data titik pengukuran
        const titikData = [];
        for (let i = 1; i <= 10; i++) {
            titikData.push(document.getElementById('cell' + i).innerText);
        }
        console.log('Data titik:', titikData);

        const isHighValue = averageResult.classList.contains('high-value');

        try {
            // Load logos first
            console.log('Loading logos...');
            let gniLogoData = null;
            let hseLogoData = null;
            
            try {
                // Gunakan URL absolut relatif terhadap halaman saat ini (biar path tidak salah saat runtime)
                const gniLogoUrl = new URL('gnilogo.png', window.location.href).toString();
                gniLogoData = await imageToBase64(gniLogoUrl);
                console.log('GNI logo loaded from:', gniLogoUrl);

            } catch (e) {
                console.warn('Could not load GNI logo:', e.message);
            }
            
            try {
                const hseLogoUrl = new URL('hselogo.png', window.location.href).toString();
                hseLogoData = await imageToBase64(hseLogoUrl);
                console.log('HSE logo loaded from:', hseLogoUrl);

            } catch (e) {
                console.warn('Could not load HSE logo:', e.message);
            }

            const { jsPDF } = window.jspdf;
            console.log('Creating jsPDF instance...');
            const doc = new jsPDF({
                orientation: 'p',
                unit: 'mm',
                format: 'a4'
            });
            console.log('jsPDF instance created');

            // Margin dan layout
            const margin = 15;
            const pageWidth = doc.internal.pageSize.getWidth();
            const pageHeight = doc.internal.pageSize.getHeight();
            const contentWidth = pageWidth - (margin * 2);
            let y = margin + 5;

            console.log('Page dimensions:', pageWidth, 'x', pageHeight, 'mm');

            // === HEADER dengan LOGO ===
            console.log('Drawing header with logos...');
            const gniLogoWidth = 20;  // mm
            const gniLogoHeight = 20; // mm
            const hseLogoWidth = 20;  // mm
            const hseLogoHeight = 20; // mm
            const logoY = margin;

            // Fallback: coba ambil logo langsung dari HTML jika base64 gagal
            const topGniImg = document.querySelector('.top-logo.left-logo') || document.querySelector('img[alt="GNI"]');
            const topHseImg = document.querySelector('.top-logo.right-logo') || document.querySelector('img[alt="HSE"]');

            if (!gniLogoData && topGniImg) {
                try {
                    gniLogoData = await imageToBase64(topGniImg.src);
                    console.log('Fallback GNI logo loaded from DOM src');
                } catch (e) {
                    console.warn('Fallback GNI logo load failed:', e);
                }
            }

            if (!hseLogoData && topHseImg) {
                try {
                    hseLogoData = await imageToBase64(topHseImg.src);
                    console.log('Fallback HSE logo loaded from DOM src');
                } catch (e) {
                    console.warn('Fallback HSE logo load failed:', e);
                }
            }

            if (gniLogoData) {
                try {
                    doc.addImage(gniLogoData, 'PNG', margin, logoY, gniLogoWidth, gniLogoHeight);
                } catch (e) {
                    console.warn('Failed to add GNI logo:', e);
                }
            }

            if (hseLogoData) {
                try {
                    doc.addImage(hseLogoData, 'PNG', pageWidth - margin - hseLogoWidth, logoY, hseLogoWidth, hseLogoHeight);
                } catch (e) {
                    console.warn('Failed to add HSE logo:', e);
                }
            }

            const headerBgY = y;
            const headerBgH = 18;
            const headerLeft = margin + (gniLogoData ? gniLogoWidth + 4 : 0);
            const headerRight = pageWidth - margin - (hseLogoData ? hseLogoWidth + 4 : 0);
            const headerBgW = Math.max(0, headerRight - headerLeft);
            if (headerBgW > 0) {
                doc.setFillColor(24, 90, 157);
                doc.rect(headerLeft, headerBgY, headerBgW, headerBgH, 'F');
            }

            doc.setTextColor(255, 255, 255);
            doc.setFontSize(14);
            doc.setFont('helvetica', 'bold');
            doc.text('HASIL PENGUKURAN KEBISINGAN', pageWidth / 2, y + 12, { align: 'center' });

            y += 22;

            // === INFORMASI UMUM ===
            console.log('Drawing info section...');
            
            // Tentukan tinggi info box berdasarkan jumlah baris lokasi
            const lokasiTextX = margin + 23;
            const lokasiMaxWidth = contentWidth / 3 - 10;
            const lokasiLineHeight = 3.5;
            
            // Wrap teks lokasi berdasarkan 3 kata per baris
            const wrapByWords = (text, wordsPerLine) => {
                const words = text.split(/\s+/);
                const lines = [];
                for (let i = 0; i < words.length; i += wordsPerLine) {
                    lines.push(words.slice(i, i + wordsPerLine).join(' '));
                }
                return lines;
            };
            const lokasiLines = wrapByWords(lokasi, 3);
            const infoBoxHeight = 20 + Math.max(0, (lokasiLines.length - 1) * lokasiLineHeight);
            
            doc.setFillColor(240, 244, 248);
            doc.rect(margin, y, contentWidth, infoBoxHeight, 'F');

            doc.setTextColor(0, 0, 0);
            doc.setFontSize(9);
            doc.setFont('helvetica', 'normal');

            const infoY = y + 6;
            const colW = contentWidth / 3;

            // Row 1
            doc.setFont('helvetica', 'bold');
            doc.text('Lokasi:', margin + 3, infoY);

            doc.setFont('helvetica', 'normal');
            // Wrap teks lokasi ketika mencapai 3 kata
            
            // Gambar baris lokasi
            lokasiLines.forEach((line, idx) => {
                const lineY = infoY + (idx * lokasiLineHeight);
                doc.text(line, lokasiTextX, lineY);
            });


            doc.setFont('helvetica', 'bold');
            doc.text('Tanggal:', margin + 3 + colW, infoY);
            doc.setFont('helvetica', 'normal');
            doc.text(tanggal, margin + 23 + colW, infoY);

            doc.setFont('helvetica', 'bold');
            doc.text('Jam:', margin + 3 + colW * 2, infoY);
            doc.setFont('helvetica', 'normal');
            doc.text(jam, margin + 23 + colW * 2, infoY);

            // Row 2 (harus geser mengikuti tinggi lokasi agar tidak menimpa)
            const infoRow1Height = lokasiLines.length * lokasiLineHeight;
            const infoY2 = infoY + 7 + Math.max(0, infoRow1Height - 3.5);

            doc.setFont('helvetica', 'bold');
            doc.text('Durasi:', margin + 3, infoY2);
            doc.setFont('helvetica', 'normal');
            doc.text(durationHours + ' jam', margin + 23, infoY2);

            doc.setFont('helvetica', 'bold');
            doc.text('NAB:', margin + 3 + colW, infoY2);
            doc.setFont('helvetica', 'normal');
            doc.text(nab + ' dBA', margin + 23 + colW, infoY2);

            y += infoBoxHeight + 5;

            // === TABEL DATA PENGUKURAN ===
            console.log('Drawing table...');
            doc.setTextColor(0, 0, 0);
            doc.setFontSize(11);
            doc.setFont('helvetica', 'bold');
            doc.text('DATA PENGUKURAN (10 TITIK)', margin, y);
            y += 5;

            // Header tabel
            doc.setFillColor(24, 90, 157);
            doc.rect(margin, y, contentWidth, 7, 'F');

            doc.setTextColor(255, 255, 255);
            doc.setFontSize(8);
            doc.setFont('helvetica', 'bold');

            const colWidth = contentWidth / 5;
            for (let i = 0; i < 5; i++) {
                const x = margin + (i * colWidth);
                doc.text('Titik ' + (i + 1), x + colWidth / 2, y + 5, { align: 'center' });
            }
            y += 7;

            // Baris 1 (Titik 1-5)
            doc.setTextColor(0, 0, 0);
            doc.setFontSize(10);
            doc.setFont('helvetica', 'bold');
            for (let i = 0; i < 5; i++) {
                const x = margin + (i * colWidth);
                doc.setDrawColor(200, 200, 200);
                doc.rect(x, y, colWidth, 9);
                doc.text(String(titikData[i]), x + colWidth / 2, y + 6, { align: 'center' });
            }
            y += 9;

            // Baris 2 (Titik 6-10)
            for (let i = 5; i < 10; i++) {
                const col = i - 5;
                const x = margin + (col * colWidth);
                doc.setDrawColor(200, 200, 200);
                doc.rect(x, y, colWidth, 9);
                doc.text(String(titikData[i]), x + colWidth / 2, y + 6, { align: 'center' });
            }
            y += 14;

            // === HASIL RATA-RATA ===
            console.log('Drawing result section...');
            const resultBoxHeight = 20;
            const resultColorRGB = isHighValue ? [220, 53, 69] : [40, 167, 69];
            const resultBgRGB = isHighValue ? [255, 235, 238] : [232, 245, 233];

            doc.setFillColor(resultBgRGB[0], resultBgRGB[1], resultBgRGB[2]);
            doc.rect(margin, y, contentWidth, resultBoxHeight, 'F');

            // Border kiri berwarna
            doc.setFillColor(resultColorRGB[0], resultColorRGB[1], resultColorRGB[2]);
            doc.rect(margin, y, 3, resultBoxHeight, 'F');

            doc.setTextColor(51, 51, 51);
            doc.setFontSize(10);
            doc.setFont('helvetica', 'normal');
            doc.text('HASIL RATA-RATA', pageWidth / 2, y + 6, { align: 'center' });

            doc.setTextColor(resultColorRGB[0], resultColorRGB[1], resultColorRGB[2]);
            doc.setFontSize(18);
            doc.setFont('helvetica', 'bold');
            doc.text(averageResult.innerText, pageWidth / 2, y + 13, { align: 'center' });

            doc.setFontSize(8);
            doc.setFont('helvetica', 'normal');
            if (isHighValue) {
                doc.setTextColor(220, 53, 69);
                doc.text('NILAI KEBISINGAN MELEBIHI NAB', pageWidth / 2, y + 18, { align: 'center' });
            } else {
                doc.setTextColor(40, 167, 69);
                doc.text('NILAI KEBISINGAN AMAN (DI BAWAH NAB)', pageWidth / 2, y + 18, { align: 'center' });
            }

            y += resultBoxHeight + 5;

            // === CATATAN ===
            if (catatanText) {
                console.log('Drawing notes section...');
                const noteBoxHeight = 15;
                doc.setFillColor(255, 243, 205);
                doc.rect(margin, y, contentWidth, noteBoxHeight, 'F');

                doc.setFillColor(255, 193, 7);
                doc.rect(margin, y, 3, noteBoxHeight, 'F');

                doc.setTextColor(133, 100, 4);
                doc.setFontSize(9);
                doc.setFont('helvetica', 'bold');
                doc.text('CATATAN:', pageWidth / 2, y + 5, { align: 'center' });

                doc.setFont('helvetica', 'normal');
                doc.setFontSize(8);
                const noteLines = doc.splitTextToSize(catatanText, contentWidth - 10);
                let noteY = y + 9;
                noteLines.forEach((line, index) => {
                    if (index < 2) {
                        doc.text(line, pageWidth / 2, noteY + (index * 4), { align: 'center' });
                    }
                });

                y += noteBoxHeight + 5;
            } else {
                // Tambahkan spasi jika tidak ada catatan
                y += 5;
            }

            // === FOOTER ===
            console.log('Drawing footer...');
            doc.setTextColor(153, 153, 153);
            doc.setFontSize(7);
            doc.setFont('helvetica', 'italic');
            doc.text('Dokumen ini dibuat secara otomatis oleh Sistem Pengukuran Kebisingan', pageWidth / 2, y, { align: 'center' });
            y += 4;
            doc.text(now.toLocaleDateString('id-ID') + ' - ' + now.toLocaleTimeString('id-ID'), pageWidth / 2, y, { align: 'center' });

            // === NAMA FILE ===
            const namaFile = `Hasil_Kebisingan_${lokasi.replace(/\s+/g, '_')}_${now.toISOString().split('T')[0]}.pdf`;

            // Simpan PDF
            console.log('Saving PDF as:', namaFile);
            doc.save(namaFile);
            console.log('PDF saved successfully!');

            // Kembalikan tombol
            exportPdfBtn.innerText = originalText;
            exportPdfBtn.disabled = false;

        } catch (err) {
            console.error('Error generating PDF:', err);
            console.error('Error stack:', err.stack);
            
            if (exportPdfBtn) {
                exportPdfBtn.innerText = originalText;
                exportPdfBtn.disabled = false;
            }
            alert('Gagal membuat PDF: ' + err.message);
        }
    }

    document.getElementById('exportBtn').addEventListener('click', exportToExcel);
    document.getElementById('exportPdfBtn').addEventListener('click', exportToPdf);

    document.getElementById('resetBtn').addEventListener('click', function () {
        const confirmReset = confirm('Apakah Anda yakin ingin mereset pengukuran? Data yang ada akan dihapus.');
        if (confirmReset) {
            resetTabel();
        }
    });

});

