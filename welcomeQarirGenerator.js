const fs = require("fs").promises;
const path = require("path");
const process = require("process");
const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");
const xlsx = require("xlsx");
const readline = require("readline");
const waktu = {
  linkMeet: "https://meet.google.com/vpb-mawu-gtq",
  Date: "Wednesday, 3 Dec 2025",
  Time: "8:00 – 9:30 PM WIB",
  Platform: "Google Meet",
};
// ==================== CONFIGURATION ====================
const CONFIG = {
  SCOPES: ["https://www.googleapis.com/auth/gmail.send"],
  TOKEN_PATH: path.join(process.cwd(), "token.json"),
  CREDENTIALS_PATH: path.join(process.cwd(), "credentials.json"),
  SENDER_EMAIL: "academic@qarirgenerator.com",

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
  subject: "Invitation: Onboarding + Career Mapping Session",
  recipients: [
    { email: "safeffendi@gmail.com", nama: "Safrie Effendi" },
    { email: "risantianjalasari@gmail.com", nama: "Risanti Anjalasari" },
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
    <title>QarirGenerator - Onboarding Session</title>
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
                  background-image: url(&quot;https://enfileup.prtcl.icu/storage/files/ab6955db-dab5-4aa6-8646-3a36c4556ecb.png&quot;);
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
                  WELCOME TO<br />QARIR GENERATOR
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
                <h2
                  style="font-size: 22px; margin-bottom: 15px; color: #1f2937"
                >
                  Your Journey Begins Here 🎯
                </h2>
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
                  It is with profound joy and anticipation that we welcome you
                  to <strong>QarirGenerator</strong>, a place where ambition
                  meets opportunity, and where every individual is given the
                  tools to transform potential into achievement.
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
                  By choosing to begin your journey with us, you have taken an
                  important step towards shaping not only your career, but also
                  your story of global impact.
                </p>
              </td>
            </tr>

            <!-- SESSION DETAILS BOX -->
               <!-- EVENT DETAILS -->
            <tr>
              <td style="padding: 10px 20px">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f9fafb; border-radius: 8px; border: 1px solid #e5e7eb;">
                  <tr>
                    <td class="mobile-padding" style="padding: 25px">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
  <td class="mobile-padding" style="padding: 25px">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <!-- DATE -->
        <td class="column-stack mobile-center" align="center" width="33%" valign="top">
          <img  width="22" style="margin-bottom:6px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAAsTAAALEwEAmpwYAAACkElEQVR4nO2Yv2/TQBTHLST6X7TpViYYWrNjUx+NaKmqpi0qdGRg5sfAEImJtkhkZcGOE4mSRk1a1KFSnYWECcVCRLLBpgOwUilpxHropRCBYzs55+Ifkr/SVzpdnPfu47xnP4VhYsWKRaySUk2WKtUfZaX6fe/43VzQcYgFCcuVGu5YqX0LOg6xukn/OOg4tvolzIy3ObbY5tlWm7+K/7U1sfXzQV3uG4dtnfFsqXltZorxcPifdklbXG9i2CM9PFkc9hTONDDA+Z23D9a4fKknMeyRAniIUxgcwKZs/lqZnOhJrCQmiAGI43BskwDALXEC7xcPukn3dw9wZTLhAYA8DkMDAH7mw+TNTnJICuvGFW8ldEgYhwoANBokhzsIhrXXJm4QxqECEKSZGICPCICqmziMZmIAPQbAoQDI5Pf+M+19NQbQYwAcagA16k2sRh0gE/USylAGeJEv4ke5bfzg7RN8+80GnpPnO4Y17Em1Av6gfQ4nwOPcczyfTWFeQq5e2VnH11+hpdAA1LUv+J78sO/Beyyi7XQ6fWEoABp+erRFfviuhc1AAeT3u46Hs8rpOk4UFgMpIWhYqOfhAdBJqpAa8x0AnjZu5WGV27WceCPlO8CGfJ8egITyvgMsZFdcD9xPlmbWfW/ipLxAD0BEZ5EG4ETB/v/Sum62RgWwtnOXHoCENHsAzSyNCgBmG1pNzItCzhbgo3YyVdeN01EAwGBG7yk0u8w46ZNhjKu6WVA1s0kTAKbK1dd3hn+RSeir44ts1IKpcuhfIItuMUEKpkre6zAnomdM0IKRmJOELXIAYbPvOO2nOFFY5EVkDHDXjcDLxknTL6cvwmAGsw082+ENC4b1+d7sMlxj9+XfIRmdY/8OZn4AAAAASUVORK5CYII=" alt="calendar-plus">
          <p style="margin:8px 0 0 0; font-weight:bold; color:#ff9500; font-size:14px;">DATE</p>
          <p style="margin:5px 0 0 0; color:#555; font-size:14px;">${waktu.Date}</p>
        </td>

        <!-- TIME -->
        <td class="column-stack mobile-center" align="center" width="33%" valign="top">
       <img  width="22" style="margin-bottom:6px" src="data:image/png;base64,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" alt="time-machine">
<p style="margin:8px 0 0 0; font-weight:bold; color:#ff9500; font-size:14px;">TIME</p>
          <p style="margin:5px 0 0 0; color:#555; font-size:14px;">${waktu.Time}</p>
        </td>

        <!-- PLATFORM -->
        <td class="column-stack mobile-center" align="center" width="33%" valign="top">
<img width="22" style="margin-bottom:6px"  src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAADI0lEQVR4nO3Zb0gTcRgH8EUE1es56UW9iuqFd3Ob1iJlG5lu7rfdzRCEkMQ2CQ0NLXpjaqhNKhFCTS0QMSxK78wCGTPusDIdvgnsVcJCNCTNYRqIWzxxLZNy/zw73cV94fvqXtzvw8HD3XMymRQpUqRIEUMgN3e3r07pXGUx9z8vg1GrDHYVRrT7BEdMa1O754tVk34WA8HKJI3BuGaPoIiZk6mwUKFaEhTCYhBg8Lzwh5HtstBEgZkm0nkjuPquqQVF+IO9GeosaADtRxT5FNEkmCnSyxuxfZCkW3+fJavPdsBMEx4OESzxmTdipyAmmsAQTXxcR8QICYfYCQiiSBOiyMU/ETFAIiG2G2KmyDIzTXzfiIgCiYb4ObXK1ctCQ74x2B0zRbaEBkSBxILg+qU8bVZIxLwbh9JH+onIiAiQrpzLba+yzk65ss/5BiyFgX6rHdb63FKwMmTMm/Pk5C9Mt1dC4J2DV/3DKRER3kEcrE4V2O7r53lDDjZMew81zEC01jOLwDf+0aywiPFnStDVqgG/rgZjm2FJlJC+HiVoqoMIroYWvV9UkFUGg+ZO5W8A/qtpzXoQDeTrSwzK7iZvQOBignxyYZB3WxUSgYsF8v4FDpn14RG4GCAuWgfHa8IDcLFAbD25oKzWiB9S6CqCzJ5sUNemih+CaBJMTyyQ0qAVPwRxX3i9BGgbT4kfgrhSJKS3Gv4DCB2s4UEGKKvWh0Bai8EvSgjXMw9NkHwjOATS753m/9KYWPxmWFEyMqW4NOpLLB0LJJZ6YK2K0rEVRcnbuSNXPAsXe4ag8nU1r+YM5EY8nPExAk3dCdB1ZPB/jZfbWW+Cg4Vo1Tg7ZqPfhH+ze62Q0W2cEBySVNe0LCQEcZh+K/9P3VghuLNRUATiRjNN8F8+xCOE1zooXiGbXtDFM2RTK1O5g/0QLxBEE/yX2HIHQ8U0teqblgR/In22iL8VEE3kI4pMC3ldcWEYlzvY5WiQYzWtk4JCKHJU0160tR89CReYwwkOplFuZwfldtYdqker2roQRboFKGWmiQpd5/m9W0JIkSJFihTZNuUHjEU4SvaMY7IAAAAASUVORK5CYII=" alt="google-meet--v1">
<p style="margin:8px 0 0 0; font-weight:bold; color:#ff9500; font-size:14px;">PLATFORM</p>
          <p style="margin:5px 0 0 0; color:#555; font-size:14px;">${waktu.Platform}</p>
        </td>
      </tr>
    </table>
  </td>
</tr>

            <!-- WHAT TO EXPECT SECTION -->
            <tr>
              <td class="mobile-padding" style="padding: 30px 40px">
                <h3
                  style="
                    font-size: 18px;
                    color: #111;
                    border-bottom: 2px solid #4f46e5;
                    display: inline-block;
                    padding-bottom: 5px;
                  "
                >
                  🌟 What to Expect in This Session
                </h3>
                <ul
                  style="
                    font-size: 16px;
                    color: #555;
                    line-height: 1.8;
                    padding-left: 20px;
                    margin-top: 20px;
                  "
                >
                  <li>
                    Discover how QarirGenerator connects you with
                    <strong
                      >mentors, recruiters, and global opportunities</strong
                    >.
                  </li>
                  <li>
                    Understand the
                    <strong>academic framework and mentorship structure</strong>
                    that will support your growth.
                  </li>
                  <li>
                    Gain clarity on how to make the most of the
                    <strong
                      >resources, networks, and international pathways</strong
                    >
                    offered to you.
                  </li>
                  <li>
                    Meet the team that will be accompanying you throughout your
                    journey.
                  </li>
                </ul>
              </td>
            </tr>

            <!-- PREPARATION REQUIREMENTS -->
            <tr>
              <td class="mobile-padding" style="padding: 0 40px 30px 40px">
                <h3
                  style="
                    font-size: 18px;
                    color: #111;
                    border-bottom: 2px solid #ff9500;
                    display: inline-block;
                    padding-bottom: 5px;
                  "
                >
                  📋 Before the Session
                </h3>
                <p
                  style="
                    font-size: 16px;
                    color: #555;
                    line-height: 1.8;
                    margin-top: 20px;
                  "
                >
                  Please prepare and send the following to
                  <strong>academic@qarirgenerator.com</strong>:
                <ep>
                <ol
                  style="
                    font-size: 16px;
                    color: #555;
                    line-height: 1.8;
                    padding-left: 20px;
                    margin-top: 10px;
                  "
                >
                  <li>Your updated <strong>CV</strong> (in PDF format)</li>
                  <li>The link to your <strong>LinkedIn profile</strong></li>
                </ol>
                <p
                  style="
                    font-size: 14px;
                    color: #666;
                    line-height: 1.6;
                    margin-top: 15px;
                    font-style: italic;
                  "
                >
                  These materials will help us better understand your
                  professional background and tailor the mentorship process to
                  your goals.
                </p>
              </td>
            </tr>
 <tr>
              <td align="center" style="padding-bottom: 40px">
                <table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td
                      align="center"
                      style="border-radius: 50px; background-color: #ff9e16"
                    >
                      <a
href="${waktu.linkMeet}"
                        style="
                          font-size: 18px;
                          font-family: Helvetica, Arial, sans-serif;
                          color: #ffffff;
                          text-decoration: none;
                          border-radius: 50px;
                          padding: 16px 40px;
                          display: inline-block;
                          font-weight: bold;
                          box-shadow: 0 4px 6px rgba(79, 70, 229, 0.3);
                        "
                      >
                        JOIN MEETING NOW
                      </a>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

           
            <tr>
              <td align="center" style="padding-bottom: 40px">
                <table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td
                      align="center"
                      style="border-radius: 50px; background-color: #4cca59"
                    >
                      <a
                        href="https://wa.me/6285123887483/?text=Hi%20im%20${nama}%20ready%20to%20join%20the%20qarirGenerator%20and%20confirm%20my%20attendance!"
                        style="
                          font-size: 18px;
                          font-family: Helvetica, Arial, sans-serif;
                          color: #ffffff;
                          text-decoration: none;
                          border-radius: 50px;
                          padding: 16px 40px;
                          display: inline-block;
                          font-weight: bold;
                          box-shadow: 0 4px 6px rgba(79, 70, 229, 0.3);
                        "
                      >
                        CONFIRM ATTENDANCE
                      </a>
                    </td>
                  </tr>
                </table>
                <p style="font-size: 12px; color: #999; margin-top: 15px">
                  Please confirm your attendance or request an alternative time
                </p>
              </td>
            </tr>

            <!-- CLOSING MESSAGE -->
            <tr>
              <td
                class="mobile-padding"
                style="padding: 0 40px 30px 40px; text-align: center"
              >
                <p
                  style="
                    font-size: 16px;
                    line-height: 1.6;
                    color: #555555;
                    font-style: italic;
                  "
                >
                  "Onboarding is not merely about information; it is about
                  transformation. Come with an open mind, a willingness to
                  learn, and the courage to dream beyond boundaries."
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
</html>

`;
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
