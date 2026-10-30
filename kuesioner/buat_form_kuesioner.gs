/**
 * Mengisi Google Form "Kuesioner Minat dan Motivasi Belajar Keamanan Siber"
 * dan membuat Spreadsheet responsnya di folder Drive "Pengabdian".
 *
 * Cara pakai: buka form di FORM_ID, klik menu titik tiga > Apps Script, hapus isi Code.gs,
 * tempel seluruh isi file ini, simpan, pilih fungsi buatKuesioner, klik Run,
 * lalu izinkan akses Forms, Drive dan Sheets. Link form akan muncul di Execution log.
 * Semua pertanyaan lama di form dihapus dulu, jadi script aman dijalankan ulang.
 */
var FORM_ID = '1Dy7l_W6mD-FQ4U5sbMDEUBW7WGRs3KOzkSb3alveDMw';
var FOLDER_ID = '16xWyhXBHibrNYeaqnhmcDcFI8bhEBZon';

var SKALA = ['Sangat Tidak Setuju', 'Tidak Setuju', 'Netral', 'Setuju', 'Sangat Setuju'];

function buatKuesioner() {
  var form = FormApp.openById(FORM_ID);
  form.getItems().forEach(function (item) { form.deleteItem(item); });
  form.setTitle('Kuesioner Minat dan Motivasi Belajar Keamanan Siber - Cybersecurity Forge-Day 2026');
  form.setDescription(
    'Kuesioner ini merupakan bagian dari evaluasi kegiatan pengabdian kepada masyarakat ' +
    '"Edukasi Keamanan Siber Untuk Pelajar Menggunakan Media Simulasi Capture The Flag (CTF)" ' +
    'oleh FILKOM Universitas Brawijaya bersama SMKN 2, 4, dan 5 Malang.\n\n' +
    'Tidak ada jawaban benar atau salah. Jawablah sesuai dengan kondisi Anda yang sebenarnya. ' +
    'Jawaban Anda hanya digunakan untuk keperluan evaluasi kegiatan dan tidak memengaruhi nilai apa pun.');
  form.setCollectEmail(true);
  form.setProgressBar(true);
  form.setConfirmationMessage('Terima kasih telah mengisi kuesioner. Semangat terus belajar keamanan siber!');

  // Bagian 1. Identitas
  form.addSectionHeaderItem().setTitle('Bagian 1. Identitas Peserta');
  form.addTextItem().setTitle('Nama lengkap').setRequired(true);
  form.addMultipleChoiceItem().setTitle('Asal sekolah')
    .setChoiceValues(['SMKN 2 Malang', 'SMKN 4 Malang', 'SMKN 5 Malang']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('Apakah Anda pernah mengikuti kegiatan atau kompetisi CTF sebelum kegiatan ini?')
    .setChoiceValues(['Pernah', 'Belum pernah']).setRequired(true);

  // Bagian 2. Minat sebelum dan sesudah kegiatan (indikator target 80%)
  form.addPageBreakItem().setTitle('Bagian 2. Minat terhadap Bidang Keamanan Siber')
    .setHelpText('Bandingkan minat Anda sebelum dan sesudah mengikuti rangkaian kegiatan.');
  form.addScaleItem().setTitle('Seberapa besar minat Anda terhadap bidang keamanan siber SEBELUM mengikuti kegiatan ini?')
    .setBounds(1, 5).setLabels('Sangat rendah', 'Sangat tinggi').setRequired(true);
  form.addScaleItem().setTitle('Seberapa besar minat Anda terhadap bidang keamanan siber SESUDAH mengikuti kegiatan ini?')
    .setBounds(1, 5).setLabels('Sangat rendah', 'Sangat tinggi').setRequired(true);
  form.addMultipleChoiceItem()
    .setTitle('Dibandingkan sebelum mengikuti kegiatan, minat Anda terhadap bidang keamanan siber saat ini:')
    .setChoiceValues(['Jauh lebih tinggi', 'Lebih tinggi', 'Sama saja', 'Lebih rendah']).setRequired(true);

  // Bagian 3. Motivasi belajar, diadaptasi dari model ARCS (Keller). Aspek Kepuasan tidak ditanyakan di sini
  // karena dianalisis secara kualitatif dari umpan balik pada form pendataan sertifikat.
  form.addPageBreakItem().setTitle('Bagian 3. Motivasi Belajar')
    .setHelpText('Berikan tanggapan Anda terhadap setiap pernyataan berikut.');
  grid(form, 'A. Perhatian (Attention)', [
    'Materi dan tantangan CTF membuat saya tertarik untuk terus mengikuti pelatihan.',
    'Format kompetisi CTF membuat saya lebih fokus belajar dibandingkan metode ceramah.',
    'Variasi kategori soal (web, kriptografi, forensik, reverse engineering, dll.) membuat saya penasaran untuk mencoba.'
  ]);
  grid(form, 'B. Relevansi (Relevance)', [
    'Materi keamanan siber yang diberikan relevan dengan jurusan atau bidang yang saya pelajari di sekolah.',
    'Saya melihat manfaat keterampilan keamanan siber untuk karier saya di masa depan.',
    'Materi pelatihan berkaitan dengan hal yang saya temui dalam penggunaan internet sehari-hari.'
  ]);
  grid(form, 'C. Kepercayaan Diri (Confidence)', [
    'Setelah pelatihan, saya yakin dapat menyelesaikan soal CTF tingkat dasar.',
    'Tingkat kesulitan soal sesuai dengan kemampuan saya.',
    'Saya yakin dapat mempelajari topik keamanan siber lebih lanjut secara mandiri.'
  ]);

  // Bagian 4. Minat lanjutan
  form.addPageBreakItem().setTitle('Bagian 4. Minat Lanjutan');
  grid(form, 'Rencana setelah kegiatan', [
    'Saya ingin mengikuti kegiatan atau kompetisi CTF serupa di kemudian hari.',
    'Saya ingin bergabung dengan klub atau ekstrakurikuler keamanan siber jika ada di sekolah.',
    'Saya akan berlatih di platform CTF secara mandiri setelah kegiatan ini.',
    'Saya tertarik menekuni bidang keamanan siber sebagai pilihan studi lanjut atau karier.'
  ]);

  // Spreadsheet respons (dibuat sekali saja)
  var ssUrl;
  try {
    ssUrl = SpreadsheetApp.openById(form.getDestinationId()).getUrl();
  } catch (e) {
    var ss = SpreadsheetApp.create('Kuesioner Minat dan Motivasi (Responses)');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(FOLDER_ID));
    ssUrl = ss.getUrl();
  }

  Logger.log('Link untuk peserta: ' + form.getPublishedUrl());
  Logger.log('Spreadsheet respons: ' + ssUrl);
}

function grid(form, title, rows) {
  form.addGridItem().setTitle(title).setRows(rows).setColumns(SKALA).setRequired(true);
}
