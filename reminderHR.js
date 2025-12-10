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
  IMPORTANT_NOTE: "14:00 - 14.15",
  // Rate limiting & batch settings
  BATCH_SIZE: 10,
  DELAY_BETWEEN_EMAILS: 1000,
  DELAY_BETWEEN_BATCHES: 5000,
  MAX_RETRIES: 3,
  OPEN_TIME: "14:00 - 14.15",
  END_TIME: "17.40 - 17.55 ",
  // Logging
  LOG_FILE: path.join(process.cwd(), "email_log.json"),
  ERROR_LOG: path.join(process.cwd(), "error_log.json"),
  CAMPAIGN_HISTORY: path.join(process.cwd(), "campaign_history.json"),

  // Excel Import
  EXCEL_PATH: path.join(process.cwd(), "recipients.xlsx"),
};

// ==================== CAMPAIGN CONFIGURATION ====================
const CAMPAIGN = {
  SESSION_TITLE:
    "Invitation: HR Mentorship and Interview Simulation Category 1 DSC + AI",
  TRAINER_NAME: "Merina",
  PLATFORM_NAME: "Google Meet",
  PLATFORM_LINK: "meet.google.com/zxm-qnui-kgp",
  ACCESS_PIN: "461 900 337#",
  DATE: "Sunday, December 7 2025",
  TIME: "14:00 – 17.25",
  STYLE: "Email",
  AUDIENCE: "Students",

  CALL_TO_ACTION: "JOIN MEETING",

  GROUPS: [
    {
      group_name: "Group 1",
      time_slot: "14.15  - 14.40 WIB",
      members: ["Nur Afni", "Fransiscus", "Hartanto", "Ario"],
      mentor: "Merina",
      color: "#3b82f6",
    },
    {
      group_name: "Group 2",
      time_slot: "14.40 - 15.05 WIB",
      members: ["Ifan", "Kholidin", "Arli", "Louisa"],
      mentor: "Merina",
      color: "#10b981",
    },
    {
      group_name: "Group 3",
      time_slot: "15.05 - 15.30 WIB",
      members: ["Zamroni", "Husnul", "Maysa", "Kamal"],
      mentor: "Merina",
      color: "#fb542b",
    },
    {
      group_name: "Break",
      time_slot: "15.30 - 16.00 WIB",
      members: [""],
      mentor: "Merina",
      color: "#ffff00",
    },
    {
      group_name: "Group 4",
      time_slot: "16.00-16.25 WIB",
      members: ["Jonathan", "Kevin", "Krina", "Lazuardi"],
      mentor: "Merina",
      color: "#10b981",
    },
    {
      group_name: "Group 5",
      time_slot: "16.25-16.50 WIB",
      members: ["Milzon", "Arif", "Petra", "Nabih"],
      mentor: "Merina",
      color: "#fb542b",
    },
    {
      group_name: "Group 6",
      time_slot: "16.50-17.15 WIB",
      members: ["Narendra", "Nikodemus", "Nurwidy", "Sulianto"],
      mentor: "Merina",
      color: "#ffff00",
    },
    {
      group_name: "Group 7",
      time_slot: "17.15-17.40 WIB",
      members: ["Rizki", "Rizqi", "Vira", "Darshan"],
      mentor: "Merina",
      color: "#10b981",
    },
  ],

  RECIPIENTS: [{ email: "osama.work54@gmail.com", nama: "Osama", group: "1" }],

  CUSTOM_FIELDS: {
    company_name: "QarirGenerator",
    footer_text: "The Academic Team - QarirGenerator",
    support_email: "academic@qarirgenerator.com",
  },
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
        // File doesn't exist, create new
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

  static async saveCampaign(campaignData) {
    try {
      let history = [];
      try {
        const content = await fs.readFile(CONFIG.CAMPAIGN_HISTORY, "utf8");
        history = JSON.parse(content);
      } catch (err) {
        // File doesn't exist
      }

      history.push({
        timestamp: new Date().toISOString(),
        ...campaignData,
      });

      await fs.writeFile(
        CONFIG.CAMPAIGN_HISTORY,
        JSON.stringify(history, null, 2),
      );
    } catch (err) {
      console.error("Failed to save campaign history:", err.message);
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

// ==================== EXCEL IMPORTER ====================
class ExcelImporter {
  static async importRecipients(filePath) {
    try {
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(worksheet);

      console.log(`✓ Imported ${data.length} recipients from Excel`);
      return data.map((row) => ({
        email: row.email || row.Email,
        nama: row.nama || row.Nama || row.name || row.Name,
        group: String(row.group || row.Group || "1"),
      }));
    } catch (err) {
      console.error("❌ Failed to import Excel:", err.message);
      return [];
    }
  }
}

// ==================== TEMPLATE ENGINE ====================
class TemplateEngine {
  static replaceVariables(template, variables) {
    let result = template;
    for (const [key, value] of Object.entries(variables)) {
      const regex = new RegExp(`{{${key}}}`, "g");
      result = result.replace(regex, value);
    }
    return result;
  }

  static getGroupInfo(groupNumber) {
    return CAMPAIGN.GROUPS.find((g) => g.group_name === `Group ${groupNumber}`);
  }

  // ==================== EMAIL TEMPLATE ====================
  static createEmailTemplate(nama, customData = {}) {
    const group = customData.group || "1";
    const groupInfo = this.getGroupInfo(group);

    const variables = {
      RECIPIENT_NAME: nama,
      SESSION_TITLE: CAMPAIGN.SESSION_TITLE,
      TRAINER_NAME: CAMPAIGN.TRAINER_NAME,
      PLATFORM_NAME: CAMPAIGN.PLATFORM_NAME,
      PLATFORM_LINK: CAMPAIGN.PLATFORM_LINK,
      ACCESS_PIN: CAMPAIGN.ACCESS_PIN,
      DATE: CAMPAIGN.DATE,
      TIME: CAMPAIGN.TIME,
      CALL_TO_ACTION: CAMPAIGN.CALL_TO_ACTION,
      COMPANY_NAME: CAMPAIGN.CUSTOM_FIELDS.company_name,
      FOOTER_TEXT: CAMPAIGN.CUSTOM_FIELDS.footer_text,
      SUPPORT_EMAIL: CAMPAIGN.CUSTOM_FIELDS.support_email,
      WHATSAPP_MESSAGE: encodeURIComponent(
        `Hi, I'm ${nama} ready to join the session and confirm my attendance!`,
      ),
    };

    const htmlTemplate = `
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{{COMPANY_NAME}} - {{SESSION_TITLE}}</title>
    <style>
      body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
      table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
      img { -ms-interpolation-mode: bicubic; border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; }
      table { border-collapse: collapse !important; }
      body { height: 100% !important; margin: 0 !important; padding: 0 !important; width: 100% !important; background-color: #f4f4f7; font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; }
      a{
color:white;
text-decoration:none;
}
      @media screen and (max-width: 600px) {
        .email-container { width: 100% !important; }
        .column-stack { display: block !important; width: 100% !important; padding-bottom: 20px; }
        .mobile-center { text-align: center !important; }
        .mobile-padding { padding-left: 20px !important; padding-right: 20px !important; }
        .banner-td { padding: 30px 20px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f4f4f7">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="center" style="padding: 20px">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="email-container" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);">
            

            <tr>
              <td align="center" class="banner-td" style="background-color: #ff9500; background-image: url('https://enfileup.prtcl.icu/storage/files/ab6955db-dab5-4aa6-8646-3a36c4556ecb.png'); background-repeat: no-repeat; background-size: cover; background-position: center center; padding: 113px 10px; width: 100%; height: 100%;">
                <h1 style="color: #ffffff; font-size: 32px; margin: 0; font-weight: 800; letter-spacing: 1px; text-shadow: 0 2px 4px rgba(0, 0, 0, 0.5); line-height: 1.2;">
                {{SESSION_TITLE}}
                </h1>
              </td>
            </tr>

            <!-- GREETING -->
            <tr>
              <td class="mobile-padding" style="padding: 40px 40px 20px 40px; color: #333333; text-align: center;">
                <h2 style="font-size: 22px; margin-bottom: 15px; color: #1f2937">
                  Ready to Ace Your Interview? 🎯
                </h2>
                <p style="font-size: 16px; line-height: 1.6; color: #555555; margin-bottom: 0;">
                  Hi <strong>{{RECIPIENT_NAME}}</strong>, we're excited to invite you to a dedicated mentorship session guided by <strong>{{TRAINER_NAME}}</strong>. This session is designed to strengthen your interview performance.
                </p>
              </td>
            </tr>

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
          <p style="margin:5px 0 0 0; color:#555; font-size:14px;">{{DATE}}</p>
        </td>

        <!-- TIME -->
        <td class="column-stack mobile-center" align="center" width="33%" valign="top">
       <img  width="22" style="margin-bottom:6px" src="data:image/png;base64,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" alt="time-machine">
<p style="margin:8px 0 0 0; font-weight:bold; color:#ff9500; font-size:14px;">TIME</p>
          <p style="margin:5px 0 0 0; color:#555; font-size:14px;">{{TIME}}</p>
        </td>

        <!-- PLATFORM -->
        <td class="column-stack mobile-center" align="center" width="33%" valign="top">
<img width="22" style="margin-bottom:6px"  src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAACXBIWXMAAAsTAAALEwEAmpwYAAADI0lEQVR4nO3Zb0gTcRgH8EUE1es56UW9iuqFd3Ob1iJlG5lu7rfdzRCEkMQ2CQ0NLXpjaqhNKhFCTS0QMSxK78wCGTPusDIdvgnsVcJCNCTNYRqIWzxxLZNy/zw73cV94fvqXtzvw8HD3XMymRQpUqRIEUMgN3e3r07pXGUx9z8vg1GrDHYVRrT7BEdMa1O754tVk34WA8HKJI3BuGaPoIiZk6mwUKFaEhTCYhBg8Lzwh5HtstBEgZkm0nkjuPquqQVF+IO9GeosaADtRxT5FNEkmCnSyxuxfZCkW3+fJavPdsBMEx4OESzxmTdipyAmmsAQTXxcR8QICYfYCQiiSBOiyMU/ETFAIiG2G2KmyDIzTXzfiIgCiYb4ObXK1ctCQ74x2B0zRbaEBkSBxILg+qU8bVZIxLwbh9JH+onIiAiQrpzLba+yzk65ss/5BiyFgX6rHdb63FKwMmTMm/Pk5C9Mt1dC4J2DV/3DKRER3kEcrE4V2O7r53lDDjZMew81zEC01jOLwDf+0aywiPFnStDVqgG/rgZjm2FJlJC+HiVoqoMIroYWvV9UkFUGg+ZO5W8A/qtpzXoQDeTrSwzK7iZvQOBignxyYZB3WxUSgYsF8v4FDpn14RG4GCAuWgfHa8IDcLFAbD25oKzWiB9S6CqCzJ5sUNemih+CaBJMTyyQ0qAVPwRxX3i9BGgbT4kfgrhSJKS3Gv4DCB2s4UEGKKvWh0Bai8EvSgjXMw9NkHwjOATS753m/9KYWPxmWFEyMqW4NOpLLB0LJJZ6YK2K0rEVRcnbuSNXPAsXe4ag8nU1r+YM5EY8nPExAk3dCdB1ZPB/jZfbWW+Cg4Vo1Tg7ZqPfhH+ze62Q0W2cEBySVNe0LCQEcZh+K/9P3VghuLNRUATiRjNN8F8+xCOE1zooXiGbXtDFM2RTK1O5g/0QLxBEE/yX2HIHQ8U0teqblgR/In22iL8VEE3kI4pMC3ldcWEYlzvY5WiQYzWtk4JCKHJU0160tR89CReYwwkOplFuZwfldtYdqker2roQRboFKGWmiQpd5/m9W0JIkSJFihTZNuUHjEU4SvaMY7IAAAAASUVORK5CYII=" alt="google-meet--v1">
<p style="margin:8px 0 0 0; font-weight:bold; color:#ff9500; font-size:14px;">PLATFORM</p>
          <p style="margin:5px 0 0 0; color:#555; font-size:14px;">{{PLATFORM_NAME}}</p>
        </td>
      </tr>
    </table>
  </td>
</tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- SESSION FLOW -->
            <tr>
              <td class="mobile-padding" style="padding: 30px 40px 20px 40px">
                <h3 style="font-size: 18px; color: #111; border-bottom: 2px solid #ff9500; display: inline-block; padding-bottom: 5px;">
                  📋 Session Flow
                </h3>
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top: 20px;">
                  <tr>
                    <td style="padding: 15px; background-color: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 8px; margin-bottom: 10px;">
                      <p style="margin: 0 0 5px 0; font-weight: bold; color: #92400e;">⏰ ${CONFIG.OPEN_TIME} WIB</p>
                      <p style="margin: 0; font-size: 14px; color: #78350f;">
                        <strong>Introduction for All</strong><br/>
                        Interview tips, common mistakes, and answer structuring guidance.
                      </p>
                    </td>
                  </tr>
               ${CAMPAIGN.GROUPS.map((g, idx) => {
                 const isCurrent = group === String(idx + 1);
                 return `
                <tr>
                  <td style="padding: 15px; background-color: ${isCurrent ? "#dbeafe" : "#f3f4f6"}; border-left: 4px solid ${g.color}; border-radius: 8px; margin-bottom: 10px;">
                    <p style="margin: 0 0 5px 0; font-weight: bold; color: #1e40af;">⏰ ${g.time_slot}</p>
                    <p style="margin: 0; font-size: 14px; color: #1e3a8a;">
                      <strong>${g.group_name} Interview Simulation</strong><br/>
                      Members: ${g.members.join(", ")}
                      ${isCurrent ? "<br/><strong>👉 YOUR GROUP</strong>" : ""}
                    </p>
                  </td>
                </tr>
              `;
               }).join("")}
                  <tr>
                    <td style="padding: 15px; background-color: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 8px;">
                      <p style="margin: 0 0 5px 0; font-weight: bold; color: #92400e;">⏰ ${CONFIG.END_TIME} WIB</p>
                      <p style="margin: 0; font-size: 14px; color: #78350f;">
                        <strong>Closing for All</strong><br/>
                        Shared reflection and key takeaways.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
  <tr><td style="padding:20px 40px 30px 40px">
    <table width="100%" style="background:#fff7d6; border-left:4px solid #fbbf24; border-radius:8px">
      <tr><td style="padding:18px">
        <p style="margin:0; font-weight:700; color:#92400e;">📌 Required Preparation</p>
        <ul style="font-size:14px; color:#7a4b0b; line-height:1.6; margin-top:10px;">
          <li>Send your latest CV before the session</li>
          <li>Share your updated LinkedIn profile</li>
          <li>Prepare roles/companies you aim to apply for</li>
          <li>Highlight any areas you want focused feedback on</li>
        </ul>
      </td></tr>
    </table>
  </td></tr>

</tr>
</table>
</td>
</tr>
<tr>
<td align="center">
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td align="center" style="border-radius: 50px; background-color: #4285f4; box-shadow: 0 4px 12px rgba(66, 133, 244, 0.3);">
<a href="https://forms.gle/hGtTKdk2Mv755vdQA" target="_blank" style="font-size: 16px; font-family: Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 50px; padding: 16px 40px; display: inline-block; font-weight: bold;">
📝 SUBMIT CV & LINKEDIN PROFILE
</a>
</td>
</tr>
</table>
</td>
</tr>
            <!-- CTA BUTTON -->
            <tr>
              <td align="center" style="padding: 30px 40px">
                <table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="center" style="border-radius: 50px; background: linear-gradient(135deg, #ff9500 0%, #ff6b00 100%);">
                      <a href="{{PLATFORM_LINK}}" target="_blank" style="font-size: 18px; font-family: Helvetica, Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 50px; padding: 16px 40px; display: inline-block; font-weight: bold; box-shadow: 0 4px 6px rgba(255, 149, 0, 0.3);">
                        {{CALL_TO_ACTION}}
                      </a>
                    </td>
                  </tr>
                </table>
                <p style="font-size: 12px; color: #999; margin-top: 15px">
                  📞 PIN: {{ACCESS_PIN}} | ⏰ Join 5 minutes early!
                </p>
              </td>
            </tr>
 
            <!-- CTA BUTTON -->
<tr>
  <td align="center" style="padding: 30px 40px">
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="center" style="border-radius: 50px; background-color:#25d366;">
          <a href="https://wa.me/6285123887483?text={{WHATSAPP_MESSAGE}}" target="_blank"
            style="font-size: 18px; font-family: Helvetica, Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 50px; padding: 16px 40px; display: inline-block; font-weight: bold; box-shadow: 0 4px 6px rgba(15, 157, 88, 0.3);">
            Confirm Attendance via WhatsApp
          </a>
        </td>
      </tr>
    </table>
    <p style="font-size: 12px; color: #666; margin-top: 15px">
      Tap the button and send the message to confirm your attendance! 🚀
    </p>
  </td>
</tr>

            <!-- Important Note -->
<tr>
  <td style="padding:20px 40px 10px 40px">
    <table width="100%" style="background:#eef6ff; border-left:4px solid red; border-radius:8px">
      <tr>
        <td style="padding:18px;">
          <p style="margin:0; font-weight:700; font-size:14px; color:#1e40af;">
            🔔 Important Note About Attendance
          </p>
          <p style="margin:10px 0 0 0; font-size:14px; line-height:1.6; color:#1e3a8a;">
            To ensure you receive the maximum benefit from this Mentorship Session, please join on time at the scheduled start ${CONFIG.IMPORTANT_NOTE}. 
            Late arrivals may miss crucial instructions or not be admitted once the session begins.
          </p>
        </td>
      </tr>
    </table>
  </td>
</tr>

            <!-- FOOTER -->
            <tr>
              <td  style="background-color: #ff9500; padding: 30px; text-align: center;">
                <p style="color: white; margin: 0; font-size: 14px; font-weight: bold;">
                  {{FOOTER_TEXT}}
                </p>
                <p style="margin: 10px 0 0 0; color: white; font-size: 12px;">
                  Questions? Contact us at <span style="color:white;">{{SUPPORT_EMAIL}}</span>
                </p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>`;

    return this.replaceVariables(htmlTemplate, variables);
  }

  // ==================== TEMPLATE SELECTOR ====================
  static getTemplate(style = "Email") {
    const templates = {
      Email: this.createEmailTemplate,

      Friendly: this.createWhatsAppTemplate,
      Formal: this.createEmailTemplate,
    };

    return templates[style] || this.createEmailTemplate;
  }
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
          `\n⚠ Retry ${retries + 1}/${CONFIG.MAX_RETRIES} for ${recipient.email}`,
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
    console.log(`🚀 STARTING BULK EMAIL CAMPAIGN`);
    console.log(`${"=".repeat(60)}`);
    console.log(`📧 Total Recipients: ${recipients.length}`);
    console.log(`📦 Batch Size: ${CONFIG.BATCH_SIZE}`);
    console.log(`⏱️ Delay between emails: ${CONFIG.DELAY_BETWEEN_EMAILS}ms`);
    console.log(`⏱️ Delay between batches: ${CONFIG.DELAY_BETWEEN_BATCHES}ms`);
    console.log(`${"=".repeat(60)}\n`);

    const progressBar = new ProgressBar(recipients.length);
    const results = { success: [], failed: [] };

    for (let i = 0; i < recipients.length; i += CONFIG.BATCH_SIZE) {
      const batch = recipients.slice(i, i + CONFIG.BATCH_SIZE);
      const batchNumber = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
      const totalBatches = Math.ceil(recipients.length / CONFIG.BATCH_SIZE);

      console.log(
        `\n📦 Batch ${batchNumber}/${totalBatches} (${batch.length} emails)`,
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

      if (i + CONFIG.BATCH_SIZE < recipients.length) {
        console.log(
          `\n⏳ Waiting ${CONFIG.DELAY_BETWEEN_BATCHES}ms before next batch...`,
        );
        await sleep(CONFIG.DELAY_BETWEEN_BATCHES);
      }
    }

    progressBar.complete();
    return results;
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
    console.log("✓ Using saved credentials");
    return client;
  }

  console.log("⚠ Token not found. Opening browser for authentication...");

  client = await authenticate({
    scopes: CONFIG.SCOPES,
    keyfilePath: CONFIG.CREDENTIALS_PATH,
  });

  if (client.credentials) {
    await saveCredentials(client);
    console.log("✓ Token saved successfully");
  }

  return client;
}

// ==================== MAIN FUNCTION ====================
async function main() {
  const auth = await authorize();
  const emailSender = new EmailSender(auth);

  // Load recipients from Excel (optional)
  let recipients = CAMPAIGN.RECIPIENTS;
  if (await fs.stat(CONFIG.EXCEL_PATH).catch(() => false)) {
    recipients = await ExcelImporter.importRecipients(CONFIG.EXCEL_PATH);
  }

  const templateFunction = TemplateEngine.getTemplate(CAMPAIGN.STYLE).bind(
    TemplateEngine,
  );

  const subject = CAMPAIGN.SESSION_TITLE;

  const results = await emailSender.sendBulk(
    recipients,
    subject,
    templateFunction,
  );

  console.log("\n🎯 Campaign Finished!");
  console.log(`✔ Success: ${results.success.length}`);
  console.log(`❌ Failed: ${results.failed.length}`);

  await Logger.saveCampaign({
    session: CAMPAIGN.SESSION_TITLE,
    total: recipients.length,
    success: results.success.length,
    failed: results.failed.length,
  });
}

main().catch(console.error);
