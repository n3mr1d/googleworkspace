const fs = require("fs").promises;
const path = require("path");
const process = require("process");
const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");
const xlsx = require("xlsx");
const readline = require("readline");

// ==================== CONFIGURATION ====================
const CONFIG = {
  SCOPES: ["https://www.googleapis.com/auth/gmail.send"],
  TOKEN_PATH: path.join(process.cwd(), "token.json"),
  CREDENTIALS_PATH: path.join(process.cwd(), "credentials.json"),
  SENDER_EMAIL: "academic@qarirgenerator.com",
  MSG: "Could you resend your career mapping? I couldn’t find the attachment.",
  // Rate limiting & batch settings
  BATCH_SIZE: 10, // Kirim per batch
  DELAY_BETWEEN_EMAILS: 1000, // 1 detik antar email
  DELAY_BETWEEN_BATCHES: 5000, // 5 detik antar batch
  MAX_RETRIES: 3, // Maksimal retry jika gagal

  // Logging
  LOG_FILE: path.join(process.cwd(), "email_log.json"),
  ERROR_LOG: path.join(process.cwd(), "error_log.json"),
};

// ==================== EMAIL DATA ====================
const EMAIL_CAMPAIGN = {
  subject: "Important: Career Mapping (Nadya)",
  recipients: [
    { email: "sashnadya@gmail.com", nama: "Nadya Sashafiana", group: "1" },
  ],
};

// ==================== UTILITIES ====================
class Logger {
  static async log(type, data) {
    const timestamp = new Date().toISOString();
    const logEntry = { timestamp, type, ...data };

    const logFile = type === "error" ? CONFIG.ERROR_LOG : CONFIG.LOG_FILE;

    try {
      let logs = [];
      try {
        const content = await fs.readFile(logFile, "utf8");
        logs = JSON.parse(content);
      } catch (err) {
        // File tidak ada, buat baru
      }

      logs.push(logEntry);
      await fs.writeFile(logFile, JSON.stringify(logs, null, 2));
    } catch (err) {
      console.error("Failed to write log:", err.message);
    }
  }

  static async getStats() {
    try {
      const content = await fs.readFile(CONFIG.LOG_FILE, "utf8");
      const logs = JSON.parse(content);

      const stats = {
        total: logs.length,
        success: logs.filter((l) => l.type === "success").length,
        failed: logs.filter((l) => l.type === "error").length,
        lastRun: logs[logs.length - 1]?.timestamp || "N/A",
      };

      return stats;
    } catch (err) {
      return { total: 0, success: 0, failed: 0, lastRun: "N/A" };
    }
  }
}

class ProgressBar {
  constructor(total) {
    this.total = total;
    this.current = 0;
  }

  update(current) {
    this.current = current;
    const percentage = Math.floor((current / this.total) * 100);
    const filled = Math.floor(percentage / 2);
    const empty = 50 - filled;

    const bar = "█".repeat(filled) + "░".repeat(empty);
    process.stdout.write(
      `\r[${bar}] ${percentage}% (${current}/${this.total})`,
    );
  }

  complete() {
    process.stdout.write("\n");
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ==================== EMAIL TEMPLATE ====================
class EmailTemplate {
  static create(nama) {
    return `
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>QarirGenerator</title>
    <style>
      /* RESET STYLES */
      body,
      table,
      td,
      a {
        -webkit-text-size-adjust: 100%;
        -ms-text-size-adjust: 100%;
      }
      table,
      td {
        mso-table-lspace: 0pt;
        mso-table-rspace: 0pt;
      }
      img {
        -ms-interpolation-mode: bicubic;
        border: 0;
        height: auto;
        line-height: 100%;
        outline: none;
        text-decoration: none;
      }
      table {
        border-collapse: collapse !important;
      }
      body {
        height: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        width: 100% !important;
        background-color: #f4f4f7;
        font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
      }

      /* RESPONSIVE STYLES */
      @media screen and (max-width: 600px) {
        .email-container {
          width: 100% !important;
        }
        .column-stack {
          display: block !important;
          width: 100% !important;
          padding-bottom: 20px;
        }
        .mobile-center {
          text-align: center !important;
        }
        .mobile-padding {
          padding-left: 20px !important;
          padding-right: 20px !important;
        }
        .banner-td {
          padding: 30px 20px !important;
        }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f4f4f7">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="center" style="padding: 20px">
          <!-- Main Container -->
          <table
            border="0"
            cellpadding="0"
            cellspacing="0"
            width="600"
            class="email-container"
            style="
              background-color: #ffffff;
              border-radius: 12px;
              overflow: hidden;
              box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            "
          >
            <!-- BANNER SECTION -->
            <tr>
              <td
                align="center"
                class="banner-td"
                style="
                  background-color: #ff9500;
                  background-image: url(&quot;https://qarirgenerator.com/assets/gambar_perusahaan/QarirGenerator_Banner.png&quot;);
                  background-repeat: no-repeat;
                  background-size: cover;
                  background-position: center center;
                  padding: 113px 10px;
                  width: 100%;
                  height: 100%;
                "
              >
                <h1
                  style="
                    color: #ffffff;
                    font-size: 32px;
                    margin: 0;
                    font-weight: 800;
                    letter-spacing: 1px;
                    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.5);
                    line-height: 1.2;
                  "
                >
                  QARIR GENERATOR
                </h1>
              </td>
            </tr>
            <!-- END BANNER SECTION -->

            <!-- GREETING SECTION -->
            <tr>
              <td
                class="mobile-padding"
                style="
                  padding: 40px 40px 20px 40px;
                  color: #333333;
                  text-align: center;
                "
              >
    
                <p
                  style="
                    font-size: 16px;
                    line-height: 1.6;
                    color: #555555;
                    margin-bottom: 0;
                    text-align: left;
                  "
                >
                  Dear <span style="font-weight: bolder">${nama}</span>,
                </p>
                <p
                  style="
                    font-size: 16px;
                    line-height: 1.6;
                    color: #555555;
                    margin-top: 15px;
                    text-align: left;
                  "
                >
                 ${CONFIG.MSG}  --<strong>QarirGenerator</strong>
                </p>
              </td>
            </tr>



            <!-- FOOTER -->
            <tr>
              <td
                style="
                  background-color: #ff9500;
                  padding: 30px;
                  text-align: center;
                "
              >
                <p
                  style="
                    color: white;
                    margin: 0;
                    font-size: 14px;
                    font-weight: bold;
                  "
                >
                  The Academic Team - QarirGenerator
                </p>
                <p style="margin: 10px 0 0 0; color: white; font-size: 12px">
                  Questions? Contact us at
                  <span style="color: white">academic@qarirgenerator.com</span>
                </p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>`;
  }
}

// ==================== AUTHENTICATION ====================
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(CONFIG.TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

async function saveCredentials(client) {
  const content = await fs.readFile(CONFIG.CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(CONFIG.TOKEN_PATH, payload);
}

async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    console.log("✓ Menggunakan token yang tersimpan");
    return client;
  }
  console.log("⚠ Token tidak ditemukan. Membuka browser untuk login...");
  client = await authenticate({
    scopes: CONFIG.SCOPES,
    keyfilePath: CONFIG.CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
    console.log("✓ Token berhasil disimpan");
  }
  return client;
}

// ==================== EMAIL SENDER ====================
class EmailSender {
  constructor(auth) {
    this.gmail = google.gmail({ version: "v1", auth });
  }

  makeBody(to, from, subject, message) {
    const str = [
      `To: ${to}`,
      `From: ${from}`,
      `Subject: ${subject}`,
      "MIME-Version: 1.0",
      "Content-Type: text/html; charset=UTF-8",
      "",
      message,
    ].join("\n");

    return Buffer.from(str)
      .toString("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");
  }

  async sendSingleEmail(recipient, subject, htmlBody, retries = 0) {
    try {
      const rawMessage = this.makeBody(
        recipient.email,
        CONFIG.SENDER_EMAIL,
        subject,
        htmlBody,
      );

      const res = await this.gmail.users.messages.send({
        userId: "me",
        requestBody: { raw: rawMessage },
      });

      await Logger.log("success", {
        email: recipient.email,
        nama: recipient.nama,
        messageId: res.data.id,
      });

      return { success: true, messageId: res.data.id };
    } catch (error) {
      if (retries < CONFIG.MAX_RETRIES) {
        console.log(
          `\n⚠ Retry ${retries + 1}/${CONFIG.MAX_RETRIES} untuk ${recipient.email}`,
        );
        await sleep(2000);
        return this.sendSingleEmail(recipient, subject, htmlBody, retries + 1);
      }

      await Logger.log("error", {
        email: recipient.email,
        nama: recipient.nama,
        error: error.message,
      });

      return { success: false, error: error.message };
    }
  }

  async sendBulk(recipients, subject, templateFunction) {
    console.log(`\n${"=".repeat(60)}`);
    console.log(`🚀 MEMULAI PENGIRIMAN BULK EMAIL`);
    console.log(`${"=".repeat(60)}`);
    console.log(`📧 Total Penerima: ${recipients.length}`);
    console.log(`📦 Ukuran Batch: ${CONFIG.BATCH_SIZE}`);
    console.log(`⏱️  Delay antar email: ${CONFIG.DELAY_BETWEEN_EMAILS}ms`);
    console.log(`⏱️  Delay antar batch: ${CONFIG.DELAY_BETWEEN_BATCHES}ms`);
    console.log(`${"=".repeat(60)}\n`);

    const progressBar = new ProgressBar(recipients.length);
    const results = {
      success: [],
      failed: [],
    };

    // Kirim dalam batch
    for (let i = 0; i < recipients.length; i += CONFIG.BATCH_SIZE) {
      const batch = recipients.slice(i, i + CONFIG.BATCH_SIZE);
      const batchNumber = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
      const totalBatches = Math.ceil(recipients.length / CONFIG.BATCH_SIZE);

      console.log(
        `\n📦 Batch ${batchNumber}/${totalBatches} (${batch.length} email)`,
      );

      for (const recipient of batch) {
        const htmlBody = templateFunction(recipient.nama, recipient);
        const result = await this.sendSingleEmail(recipient, subject, htmlBody);

        if (result.success) {
          results.success.push(recipient.email);
          console.log(`  ✓ ${recipient.email}`);
        } else {
          results.failed.push({ email: recipient.email, error: result.error });
          console.log(`  ✗ ${recipient.email} - ${result.error}`);
        }

        progressBar.update(results.success.length + results.failed.length);
        await sleep(CONFIG.DELAY_BETWEEN_EMAILS);
      }

      // Delay antar batch (kecuali batch terakhir)
      if (i + CONFIG.BATCH_SIZE < recipients.length) {
        console.log(
          `\n⏳ Menunggu ${CONFIG.DELAY_BETWEEN_BATCHES}ms sebelum batch berikutnya...`,
        );
        await sleep(CONFIG.DELAY_BETWEEN_BATCHES);
      }
    }

    progressBar.complete();
    return results;
  }
}

// ==================== MAIN FUNCTION ====================
async function main() {
  try {
    // Show previous stats
    const stats = await Logger.getStats();
    if (stats.total > 0) {
      console.log("\n📊 STATISTIK SEBELUMNYA:");
      console.log(`   Total email dikirim: ${stats.total}`);
      console.log(`   Berhasil: ${stats.success} ✓`);
      console.log(`   Gagal: ${stats.failed} ✗`);
      console.log(`   Terakhir dijalankan: ${stats.lastRun}\n`);
    }

    // Konfirmasi
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });

    const confirm = await new Promise((resolve) => {
      rl.question(
        `\n⚠️  Akan mengirim email ke ${EMAIL_CAMPAIGN.recipients.length} penerima.\nLanjutkan? (yes/no): `,
        (answer) => {
          rl.close();
          resolve(
            answer.toLowerCase() === "yes" || answer.toLowerCase() === "y",
          );
        },
      );
    });

    if (!confirm) {
      console.log("\n❌ Pengiriman dibatalkan.");
      return;
    }

    // Authorize & Send
    console.log("\n🔐 Melakukan autentikasi...");
    const auth = await authorize();

    const sender = new EmailSender(auth);
    const results = await sender.sendBulk(
      EMAIL_CAMPAIGN.recipients,
      EMAIL_CAMPAIGN.subject,
      EmailTemplate.create,
    );

    // Summary
    console.log(`\n${"=".repeat(60)}`);
    console.log(`📊 RINGKASAN PENGIRIMAN`);
    console.log(`${"=".repeat(60)}`);
    console.log(`✓ Berhasil: ${results.success.length}`);
    console.log(`✗ Gagal: ${results.failed.length}`);
    console.log(
      `📈 Success Rate: ${((results.success.length / EMAIL_CAMPAIGN.recipients.length) * 100).toFixed(2)}%`,
    );

    if (results.failed.length > 0) {
      console.log(`\n❌ Email yang gagal:`);
      results.failed.forEach((fail) => {
        console.log(`   - ${fail.email}: ${fail.error}`);
      });
    }

    console.log(`\n📁 Log tersimpan di: ${CONFIG.LOG_FILE}`);
    console.log(`${"=".repeat(60)}\n`);
  } catch (error) {
    console.error("\n❌ ERROR:", error.message);
    await Logger.log("fatal_error", {
      error: error.message,
      stack: error.stack,
    });
    process.exit(1);
  }
}

// ==================== RUN ====================
if (require.main === module) {
  main();
}

module.exports = { EmailSender, EmailTemplate, Logger };
