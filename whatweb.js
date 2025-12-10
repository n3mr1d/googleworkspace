const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const xlsx = require("xlsx");
const path = require("path");

// Inisialisasi client WhatsApp
const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: {
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-accelerated-2d-canvas",
      "--no-first-run",
      "--no-zygote",
      "--disable-gpu",
    ],
  },
});

// Path file Excel
const pathExcel = path.join(process.cwd(), "recipientsWa.xlsx");

// Fungsi untuk membaca data dari Excel
const readContactsFromExcel = () => {
  try {
    console.log("Membaca file Excel...");
    const workbook = xlsx.readFile(pathExcel);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    console.log(data);
    // Validasi data
    const contacts = data
      .map((row, index) => {
        if (!row.nama || !row.nomor) {
          console.warn(`⚠ Baris ${index + 2}: Data tidak lengkap, dilewati`);
          return null;
        }

        // Format nomor telepon
        let nomor = String(row.nomor).replace(/\D/g, ""); // Hapus karakter non-digit

        // Tambahkan kode negara jika belum ada
        if (!nomor.startsWith("62")) {
          if (nomor.startsWith("0")) {
            nomor = "62" + nomor.substring(1);
          } else {
            nomor = "62" + nomor;
          }
        }

        return {
          nama: row.nama,
          nomor: nomor,
        };
      })
      .filter((contact) => contact !== null);

    console.log(`✓ Berhasil membaca ${contacts.length} kontak dari Excel\n`);
    return contacts;
  } catch (error) {
    console.error("✗ Error membaca file Excel:", error.message);
    console.log(
      "\nPastikan file 'recipientsWa.xlsx' ada di folder yang sama dengan script ini.",
    );
    console.log("Format Excel yang dibutuhkan:");
    console.log("┌──────────┬─────────────────┐");
    console.log("│   nama   │     nomor       │");
    console.log("├──────────┼─────────────────┤");
    console.log("│  Budi    │  081234567890   │");
    console.log("│  Ani     │  628123456789   │");
    console.log("└──────────┴─────────────────┘");
    process.exit(1);
  }
};

// Template pesan
const getMessageTemplate = (nama) => {
  return `Dear ${nama},

This is a friendly reminder for your upcoming HR Mentorship & Interview Simulation Session at QarirGenerator. The session will focus on enhancing your interview skills, improving communication clarity, and building confidence for real hiring stages, guided directly by HR professional Merina.

📅 Date: Thursday, 11 December 2025
⏰ Time: 20.30 – 22.00 WIB (90 minutes)
💻 Platform: Google Meet
🔗 Link: https://meet.google.com/euz-dnrz-qmy

Kindly reply to this message to confirm your attendance.
Thank you. --Dantca Bot Qarirgenerator `;
};

// Event: Generate QR Code
client.on("qr", (qr) => {
  console.log("╔════════════════════════════════════════╗");
  console.log("║  Scan QR code ini dengan WhatsApp      ║");
  console.log("╚════════════════════════════════════════╝\n");
  qrcode.generate(qr, { small: true });
});

// Event: Loading
client.on("loading_screen", (percent, message) => {
  console.log("Loading:", percent, "%", message);
});

// Event: Authenticated
client.on("authenticated", () => {
  console.log("✓ Autentikasi berhasil!");
});

// Event: Client siap
client.on("ready", async () => {
  console.log("\n╔════════════════════════════════════════╗");
  console.log("║     Client WhatsApp siap!              ║");
  console.log("╚════════════════════════════════════════╝\n");

  // Baca kontak dari Excel
  const contacts = readContactsFromExcel();

  if (contacts.length === 0) {
    console.log("Tidak ada kontak yang akan dikirimi pesan.");
    return;
  }

  console.log("Mulai mengirim pesan...\n");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

  let successCount = 0;
  let failedCount = 0;

  // Kirim pesan ke setiap kontak
  for (let i = 0; i < contacts.length; i++) {
    const contact = contacts[i];
    const progress = `[${i + 1}/${contacts.length}]`;

    try {
      const chatId = `${contact.nomor}@c.us`;
      const message = getMessageTemplate(contact.nama);

      // Cek apakah nomor terdaftar di WhatsApp
      const isRegistered = await client.isRegisteredUser(chatId);

      if (!isRegistered) {
        console.log(
          `${progress} ⚠ ${contact.nama} (${contact.nomor}) tidak terdaftar di WhatsApp`,
        );
        failedCount++;
        continue;
      }

      await client.sendMessage(chatId, message);
      console.log(
        `${progress} ✓ Berhasil → ${contact.nama} (${contact.nomor})`,
      );
      successCount++;

      // Delay random 3-5 detik antar pesan untuk menghindari spam detection
      if (i < contacts.length - 1) {
        const delay = Math.floor(Math.random() * 2000) + 3000;
        console.log(`    ⏳ Menunggu ${(delay / 1000).toFixed(1)} detik...\n`);
        await new Promise((resolve) => setTimeout(resolve, delay));
      }
    } catch (error) {
      console.error(
        `${progress} ✗ Gagal → ${contact.nama}: ${error.message}\n`,
      );
      failedCount++;
    }
  }

  // Summary
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("\n╔════════════════════════════════════════╗");
  console.log("║           RINGKASAN PENGIRIMAN         ║");
  console.log("╠════════════════════════════════════════╣");
  console.log(`║  Total kontak    : ${contacts.length.toString().padEnd(18)}║`);
  console.log(`║  Berhasil        : ${successCount.toString().padEnd(18)}║`);
  console.log(`║  Gagal           : ${failedCount.toString().padEnd(18)}║`);
  console.log("╚════════════════════════════════════════╝\n");

  console.log("Semua pesan telah diproses!");
  console.log("Tekan Ctrl+C untuk keluar.\n");
});

// Event: Authentikasi gagal
client.on("auth_failure", (msg) => {
  console.error("✗ Autentikasi gagal:", msg);
});

// Event: Client terputus
client.on("disconnected", (reason) => {
  console.log("⚠ Client terputus:", reason);
  console.log("Silakan restart aplikasi untuk menghubungkan kembali.");
});

// Event: Error
client.on("error", (error) => {
  console.error("✗ Error:", error);
});

// Inisialisasi client
console.log("╔════════════════════════════════════════╗");
console.log("║   WhatsApp Bulk Message Sender         ║");
console.log("╚════════════════════════════════════════╝\n");
console.log("Menginisialisasi WhatsApp Web...\n");

client.initialize();
