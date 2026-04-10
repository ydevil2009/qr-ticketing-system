require("dotenv").config();

const express = require("express");
const bodyParser = require("body-parser");
const QRCode = require("qrcode");
const { v4: uuidv4 } = require("uuid");
const path = require("path");
const PDFDocument = require("pdfkit");
const multer = require("multer");
const ExcelJS = require("exceljs");
const axios = require("axios");
const { v2: cloudinary } = require("cloudinary");
const { createClient } = require("@supabase/supabase-js");

const app = express();

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

const PORT = Number(process.env.PORT) || 5500;

const BREVO_API_KEY = process.env.BREVO_API_KEY;
const SENDER_EMAIL = process.env.SENDER_EMAIL;

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const SUPABASE_EXPORTS_BUCKET = process.env.SUPABASE_EXPORTS_BUCKET || "exports";

const CLOUDINARY_CLOUD_NAME = process.env.CLOUDINARY_CLOUD_NAME;
const CLOUDINARY_API_KEY = process.env.CLOUDINARY_API_KEY;
const CLOUDINARY_API_SECRET = process.env.CLOUDINARY_API_SECRET;
const CLOUDINARY_FOLDER =
  process.env.CLOUDINARY_FOLDER || "ticket-payment-screenshots";

if (!BREVO_API_KEY || !SENDER_EMAIL) {
  console.warn("⚠️ Missing BREVO_API_KEY or SENDER_EMAIL");
}

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  console.warn("⚠️ Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY");
}

if (!CLOUDINARY_CLOUD_NAME || !CLOUDINARY_API_KEY || !CLOUDINARY_API_SECRET) {
  console.warn("⚠️ Missing Cloudinary credentials");
}

const supabase = createClient(
  SUPABASE_URL || "https://example.supabase.co",
  SUPABASE_SERVICE_ROLE_KEY || "missing-key",
  {
    auth: { persistSession: false, autoRefreshToken: false },
  }
);

cloudinary.config({
  cloud_name: CLOUDINARY_CLOUD_NAME,
  api_key: CLOUDINARY_API_KEY,
  api_secret: CLOUDINARY_API_SECRET,
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 8 * 1024 * 1024 },
});

function getBaseUrl() {
  return String(process.env.BASE_URL || `http://localhost:${PORT}`).replace(
    /\/$/,
    ""
  );
}

function formatEventName(name = "") {
  return String(name).trim().toUpperCase();
}

function formatTimestamp(dateInput) {
  if (!dateInput) return "-";
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) return String(dateInput);
  return date.toLocaleString();
}

function htmlEscape(value = "") {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function listTickets() {
  const { data, error } = await supabase
    .from("tickets")
    .select("*")
    .order("created_at", { ascending: true });

  if (error) throw error;
  return data || [];
}

async function getTicketById(id) {
  const { data, error } = await supabase
    .from("tickets")
    .select("*")
    .eq("id", id)
    .maybeSingle();

  if (error) throw error;
  return data;
}

async function getEventPassCount(eventName) {
  const normalized = formatEventName(eventName);

  const { count, error } = await supabase
    .from("tickets")
    .select("*", { count: "exact", head: true })
    .eq("event_name", normalized);

  if (error) throw error;
  return count || 0;
}

async function createTicket(ticket) {
  const { data, error } = await supabase
    .from("tickets")
    .insert([ticket])
    .select("*")
    .single();

  if (error) throw error;
  return data;
}

async function updateTicket(id, values) {
  const { data, error } = await supabase
    .from("tickets")
    .update(values)
    .eq("id", id)
    .select("*")
    .single();

  if (error) throw error;
  return data;
}

async function uploadImageToCloudinary(file) {
  try {
    const dataUri = `data:${file.mimetype};base64,${file.buffer.toString(
      "base64"
    )}`;

    const result = await cloudinary.uploader.upload(dataUri, {
      folder: CLOUDINARY_FOLDER,
      resource_type: "image",
    });

    return {
      url: result.secure_url,
      publicId: result.public_id,
    };
  } catch (err) {
    console.error("CLOUDINARY ERROR:", err);
    throw err;
  }
}

function pdfToBuffer(doc) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
    doc.end();
  });
}

async function createTicketPdfBuffer({
  eventName,
  passLabel,
  name,
  email,
  contact,
  passType,
  id,
  verifyUrl,
}) {
  const qrBuffer = await QRCode.toBuffer(verifyUrl, {
    type: "png",
    width: 300,
    margin: 2,
    errorCorrectionLevel: "H",
  });

  const doc = new PDFDocument({ size: "A4", margin: 50 });

  doc.rect(30, 30, 535, 782).stroke("#333333");
  doc.fontSize(26).fillColor("black").text(passLabel, { align: "center" });
  doc.moveDown();
  doc.fontSize(14).fillColor("black").text(`Event: ${eventName}`);
  doc.text(`Name: ${name}`);
  doc.text(`Email: ${email}`);
  doc.text(`Contact: ${contact}`);
  doc.text(`Pass Type: ${passType}`);
  doc.text(`Ticket ID: ${id}`);
  doc.text(`Verification URL: ${verifyUrl}`, { width: 500 });
  doc.moveDown(2);
  doc.image(qrBuffer, 180, 260, { width: 200, height: 200 });
  doc.moveDown(15);
  doc
    .fontSize(14)
    .fillColor("black")
    .text("Show this QR code at entry.", { align: "center" });
  doc
    .fontSize(11)
    .fillColor("gray")
    .text("This ticket can only be assigned once.", { align: "center" });

  return pdfToBuffer(doc);
}

async function sendMailWithBrevo({ to, name, passLabel, pdfBuffer }) {
  try {
    const response = await axios.post(
      "https://api.brevo.com/v3/smtp/email",
      {
        sender: {
          name: "LITTLANE (Nexora)",
          email: SENDER_EMAIL,
        },
        to: [{ email: to, name }],
        subject: passLabel,
        htmlContent: `
          <div style="font-family:Arial,sans-serif;padding:30px;color:#111;line-height:1.7;">
            <h2 style="margin:0 0 18px 0;">Hello ${htmlEscape(name)}</h2>
            <h1 style="margin:0 0 10px 0;font-size:30px;">PROM NIGHT 2026 ✨</h1>
            <h3 style="margin:0 0 22px 0;font-weight:normal;color:#444;">A Night of Glamour, Music & Memories</h3>
            <p style="margin:8px 0;"><strong>Your Pass:</strong> ${htmlEscape(passLabel)}</p>
            <p style="margin:8px 0;"><strong>Date:</strong> 25 April 2026</p>
            <p style="margin:8px 0;"><strong>Time:</strong> 7:00 PM Onwards</p>
            <p style="margin:8px 0 18px 0;"><strong>Venue:</strong> Tantrumss Hinjewadi</p>
            <p style="margin:20px 0 8px 0;"><strong>Important:</strong></p>
            <ul style="margin-top:8px;padding-left:22px;">
              <li>Carry this pass for entry</li>
              <li>Pass is non-transferable</li>
              <li>Follow event guidelines</li>
            </ul>
            <p style="margin-top:22px;">Your PDF pass is attached with this email.</p>
            <p style="margin-top:30px;">Regards,<br><strong>LITTLANE ENTERTAINMENT</strong></p>
          </div>
        `,
        attachment: [
          {
            name: `${passLabel}.pdf`,
            content: pdfBuffer.toString("base64"),
          },
        ],
      },
      {
        headers: {
          "api-key": BREVO_API_KEY,
          "Content-Type": "application/json",
          Accept: "application/json",
        },
      }
    );

    return { success: true, messageId: response.data?.messageId || "" };
  } catch (err) {
    return {
      success: false,
      error: err.response?.data?.message || err.message,
    };
  }
}

async function buildExcelBuffer(tickets) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Tickets");

  sheet.columns = [
    { header: "No", key: "no", width: 8 },
    { header: "Ticket ID", key: "id", width: 40 },
    { header: "Event Name", key: "event_name", width: 28 },
    { header: "Pass Label", key: "pass_label", width: 30 },
    { header: "Name", key: "name", width: 22 },
    { header: "Email", key: "email", width: 30 },
    { header: "Contact", key: "contact", width: 18 },
    { header: "Pass Type", key: "pass_type", width: 20 },
    { header: "Payment Screenshot", key: "payment_ss_url", width: 45 },
    { header: "Verify URL", key: "verify_url", width: 45 },
    { header: "Assigned", key: "assigned", width: 12 },
    { header: "Assigned Time", key: "assigned_time", width: 24 },
    { header: "Mail Sent", key: "mail_sent", width: 12 },
    { header: "Mail Status", key: "mail_status", width: 35 },
    { header: "Created At", key: "created_at", width: 24 },
  ];

  sheet.getRow(1).font = { bold: true };

  tickets.forEach((t, index) => {
    sheet.addRow({
      no: index + 1,
      id: t.id || "",
      event_name: t.event_name || "",
      pass_label: t.pass_label || "",
      name: t.name || "",
      email: t.email || "",
      contact: t.contact || "",
      pass_type: t.pass_type || "",
      payment_ss_url: t.payment_ss_url || "",
      verify_url: t.verify_url || "",
      assigned: t.assigned ? "Yes" : "No",
      assigned_time: t.assigned_time ? formatTimestamp(t.assigned_time) : "",
      mail_sent: t.mail_sent ? "Yes" : "No",
      mail_status: t.mail_status || "",
      created_at: t.created_at ? formatTimestamp(t.created_at) : "",
    });
  });

  return workbook.xlsx.writeBuffer();
}

async function uploadExcelToSupabase(tickets) {
  const buffer = await buildExcelBuffer(tickets);
  const filePath = "ticket-data/latest.xlsx";

  const { error } = await supabase.storage
    .from(SUPABASE_EXPORTS_BUCKET)
    .upload(filePath, buffer, {
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      upsert: true,
    });

  if (error) throw error;

  const { data } = supabase.storage
    .from(SUPABASE_EXPORTS_BUCKET)
    .getPublicUrl(filePath);

  return { buffer, publicUrl: data.publicUrl, path: filePath };
}

function renderHomePage() {
  return `
    <html>
      <head>
        <title>Ticket Portal</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      </head>
      <body style="margin:0;background:#000;color:#fff;font-family:Arial,sans-serif;min-height:100vh;display:flex;justify-content:center;">
        <div style="width:100%;max-width:980px;padding:40px 20px 70px 20px;text-align:center;">
          <h1 style="font-size:72px;margin:20px 0 40px 0;font-weight:800;">🎟️ Ticket Portal</h1>
          <form method="POST" action="/generate" enctype="multipart/form-data"
            style="width:100%;max-width:650px;margin:0 auto;text-align:left;background:linear-gradient(90deg,#111,#161616);padding:40px 44px 46px 44px;border-radius:30px;box-shadow:0 0 30px rgba(255,255,255,0.03);">

            <label style="display:block;font-size:32px;margin-bottom:14px;">🎉 EVENT NAME:-</label>
            <input name="eventName" required placeholder="PROM NIGHT"
              style="width:100%;padding:18px 20px;margin:0 0 30px 0;border-radius:6px;border:1px solid #777;font-size:22px;box-sizing:border-box;">

            <label style="display:block;font-size:32px;margin-bottom:14px;">👤 NAME:-</label>
            <input name="name" required
              style="width:100%;padding:18px 20px;margin:0 0 30px 0;border-radius:6px;border:1px solid #777;font-size:22px;box-sizing:border-box;">

            <label style="display:block;font-size:32px;margin-bottom:14px;">✉️ EMAIL:-</label>
            <input name="email" type="email" required
              style="width:100%;padding:18px 20px;margin:0 0 30px 0;border-radius:6px;border:1px solid #777;font-size:22px;box-sizing:border-box;">

            <label style="display:block;font-size:32px;margin-bottom:14px;">📞 CONTACT:-</label>
            <input name="contact" required
              style="width:100%;padding:18px 20px;margin:0 0 30px 0;border-radius:6px;border:1px solid #777;font-size:22px;box-sizing:border-box;">

            <label style="display:block;font-size:32px;margin-bottom:14px;">🎫 PASS TYPE:-</label>
            <input name="passType" required placeholder="VIP / GENERAL / COUPLE"
              style="width:100%;padding:18px 20px;margin:0 0 30px 0;border-radius:6px;border:1px solid #777;font-size:22px;box-sizing:border-box;">

            <label style="display:block;font-size:32px;margin-bottom:18px;">📸 PAYMENT SS:-</label>
            <input name="paymentSS" type="file" accept="image/*" required
              style="width:100%;margin:0 0 42px 0;color:white;font-size:18px;">

            <button type="submit"
              style="width:100%;padding:22px 20px;font-size:28px;border:none;border-radius:6px;cursor:pointer;background:#e9e9e9;color:#111;font-weight:500;">
              Generate Ticket
            </button>
          </form>

          <div style="margin-top:55px;"><a href="/resend-all" style="font-size:28px;color:#56a8ff;">Send New QR to Everyone</a></div>
          <div style="margin-top:22px;"><a href="/scan" style="font-size:24px;color:#56a8ff;">Open Scanner</a></div>
          <div style="margin-top:22px;"><a href="/tickets" style="font-size:24px;color:#56a8ff;">View All Tickets</a></div>
          <div style="margin-top:22px;"><a href="/download-excel" style="font-size:24px;color:#56a8ff;">Download Excel</a></div>
        </div>
      </body>
    </html>
  `;
}

app.get("/healthz", (req, res) => {
  res.status(200).send("OK");
});

app.get("/test-route", (req, res) => {
  res.status(200).send("VERIFY ROUTE DEPLOYED");
});

app.get("/scan", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "scan.html"));
});

app.get("/scan.html", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "scan.html"));
});

app.get("/", (req, res) => {
  res.send(renderHomePage());
});

app.post("/generate", upload.single("paymentSS"), async (req, res) => {
  try {
    const { eventName, name, email, contact, passType } = req.body;

    if (!eventName || !name || !email || !contact || !passType || !req.file) {
      return res
        .status(400)
        .send("Please fill all details and upload payment screenshot.");
    }

    const normalizedEventName = formatEventName(eventName);
    const passNumber = (await getEventPassCount(normalizedEventName)) + 1;
    const passLabel = `${normalizedEventName} PASS ${passNumber}`;
    const id = uuidv4();
    const verifyUrl = `${getBaseUrl()}/verify/${id}`;

    const uploadedImage = await uploadImageToCloudinary(req.file);

    let ticket = await createTicket({
      id,
      event_name: normalizedEventName,
      pass_label: passLabel,
      pass_number: passNumber,
      name,
      email,
      contact,
      pass_type: passType,
      payment_ss_url: uploadedImage.url,
      payment_ss_public_id: uploadedImage.publicId,
      verify_url: verifyUrl,
      assigned: false,
      assigned_time: null,
      mail_sent: false,
      mail_status: "Pending",
    });

    const qrPreview = await QRCode.toDataURL(verifyUrl, {
      width: 300,
      margin: 2,
      errorCorrectionLevel: "H",
    });

    const pdfBuffer = await createTicketPdfBuffer({
      eventName: normalizedEventName,
      passLabel,
      name,
      email,
      contact,
      passType,
      id,
      verifyUrl,
    });

    const mailResult = await sendMailWithBrevo({
      to: email,
      name,
      passLabel,
      pdfBuffer,
    });

    ticket = await updateTicket(id, {
      mail_sent: mailResult.success,
      mail_status: mailResult.success
        ? "Sent"
        : `Failed: ${mailResult.error || "Unknown error"}`,
    });

    const allTickets = await listTickets();
    await uploadExcelToSupabase(allTickets);

    res.send(`
      <html>
        <body style="text-align:center;background:#111;color:white;padding-top:50px;font-family:Arial;">
          <h1>✅ Ticket Generated</h1>
          <h2>${htmlEscape(name)}</h2>
          <p>${htmlEscape(passLabel)}</p>
          <img src="${qrPreview}" width="250" />
          <p>Ticket created successfully.</p>
          <p>Mail Status: ${htmlEscape(ticket.mail_status || "Unknown")}</p>
          <br>
          <a href="/" style="color:#4da3ff;">Go Back</a>
          <br><br>
          <a href="/tickets" style="color:#4da3ff;">View All Tickets</a>
        </body>
      </html>
    `);
  } catch (err) {
    console.error("GENERATE ERROR:", err);
    res.status(500).send(`<body>Error: ${htmlEscape(err.message)}</body>`);
  }
});

app.get("/verify/:id", async (req, res) => {
  try {
    const ticket = await getTicketById(req.params.id);

    if (!ticket) {
      return res.send(`
        <html>
          <body style="text-align:center;background:#111;color:white;padding-top:50px;font-family:Arial;">
            <h1>❌ Invalid Ticket</h1>
            <p>Ticket not found.</p>
          </body>
        </html>
      `);
    }

    const detailsBox = `
      <div style="display:inline-block;text-align:left;background:#1a1a1a;padding:25px;border-radius:12px;min-width:340px;">
        <p><strong>🎉 EVENT:-</strong> ${htmlEscape(ticket.event_name || "")}</p>
        <p><strong>🎟 PASS LABEL:-</strong> ${htmlEscape(ticket.pass_label || "")}</p>
        <p><strong>👤 NAME:-</strong> ${htmlEscape(ticket.name || "")}</p>
        <p><strong>✉️ EMAIL:-</strong> ${htmlEscape(ticket.email || "")}</p>
        <p><strong>📞 CONTACT:-</strong> ${htmlEscape(ticket.contact || "")}</p>
        <p><strong>🎫 PASS TYPE:-</strong> ${htmlEscape(ticket.pass_type || "")}</p>
        ${
          ticket.assigned_time
            ? `<p><strong>⏱ ASSIGNED AT:-</strong> ${htmlEscape(
                formatTimestamp(ticket.assigned_time)
              )}</p>`
            : ""
        }
        <p><strong>📨 MAIL SENT:-</strong> ${ticket.mail_sent ? "Yes" : "No"}</p>
        <p><strong>📄 MAIL STATUS:-</strong> ${htmlEscape(ticket.mail_status || "")}</p>
        <p><strong>📅 CREATED AT:-</strong> ${htmlEscape(formatTimestamp(ticket.created_at))}</p>
        <p><strong>📸 PAYMENT SS:-</strong></p>
        ${
          ticket.payment_ss_url
            ? `<img src="${htmlEscape(
                ticket.payment_ss_url
              )}" width="280" style="border-radius:8px;border:1px solid #444;">`
            : "<p>No screenshot found</p>"
        }
      </div>
    `;

    if (ticket.assigned) {
      return res.send(`
        <html>
          <body style="text-align:center;background:#111;color:white;padding-top:40px;font-family:Arial;">
            <h1>⚠️ Already Assigned</h1>
            ${detailsBox}
          </body>
        </html>
      `);
    }

    res.send(`
      <html>
        <body style="text-align:center;background:#111;color:white;padding-top:40px;font-family:Arial;">
          <h1>🎫 Ticket Found</h1>
          ${detailsBox}
          <br><br>
          <form method="POST" action="/assign/${ticket.id}">
            <button style="padding:12px 25px;font-size:16px;cursor:pointer;background:#4CAF50;color:white;border:none;border-radius:8px;">
              ✅ Assign Entry
            </button>
          </form>
        </body>
      </html>
    `);
  } catch (err) {
    console.error("VERIFY ERROR:", err);
    res.status(500).send(`<body>Error: ${htmlEscape(err.message)}</body>`);
  }
});

app.post("/assign/:id", async (req, res) => {
  try {
    const ticket = await getTicketById(req.params.id);

    if (!ticket) {
      return res.send("Invalid Ticket");
    }

    if (ticket.assigned) {
      return res.send(`
        <html>
          <body style="text-align:center;background:#111;color:white;padding-top:50px;font-family:Arial;">
            <h1>⚠️ Already Assigned</h1>
            <p>${htmlEscape(ticket.name || "")}</p>
            <p>${htmlEscape(formatTimestamp(ticket.assigned_time))}</p>
            <a href="/verify/${ticket.id}" style="color:#4da3ff;">Go Back</a>
          </body>
        </html>
      `);
    }

    const updated = await updateTicket(ticket.id, {
      assigned: true,
      assigned_time: new Date().toISOString(),
    });

    const allTickets = await listTickets();
    await uploadExcelToSupabase(allTickets);

    res.send(`
      <html>
        <body style="text-align:center;background:#111;color:white;padding-top:50px;font-family:Arial;">
          <h1>✅ Entry Assigned</h1>
          <p><strong>${htmlEscape(updated.name || "")}</strong></p>
          <p>⏱ ${htmlEscape(formatTimestamp(updated.assigned_time))}</p>
          <br><br>
          <a href="/verify/${updated.id}" style="color:#4da3ff;">Refresh</a>
        </body>
      </html>
    `);
  } catch (err) {
    console.error("ASSIGN ERROR:", err);
    res.status(500).send(`<body>Error: ${htmlEscape(err.message)}</body>`);
  }
});

app.get("/resend-all", async (req, res) => {
  try {
    const tickets = await listTickets();

    if (tickets.length === 0) {
      return res.send("No tickets found");
    }

    for (const existingTicket of tickets) {
      const newId = uuidv4();
      const verifyUrl = `${getBaseUrl()}/verify/${newId}`;

      const pdfBuffer = await createTicketPdfBuffer({
        eventName: existingTicket.event_name,
        passLabel: existingTicket.pass_label,
        name: existingTicket.name,
        email: existingTicket.email,
        contact: existingTicket.contact,
        passType: existingTicket.pass_type,
        id: newId,
        verifyUrl,
      });

      const mailResult = await sendMailWithBrevo({
        to: existingTicket.email,
        name: existingTicket.name,
        passLabel: existingTicket.pass_label,
        pdfBuffer,
      });

      await updateTicket(existingTicket.id, {
        id: newId,
        verify_url: verifyUrl,
        assigned: false,
        assigned_time: null,
        mail_sent: mailResult.success,
        mail_status: mailResult.success
          ? "Sent"
          : `Failed: ${mailResult.error || "Unknown error"}`,
      });
    }

    const updatedTickets = await listTickets();
    await uploadExcelToSupabase(updatedTickets);

    res.send(`
      <html>
        <body style="text-align:center;background:#111;color:white;padding-top:50px;font-family:Arial;">
          <h1>✅ Sent new QR to everyone</h1>
          <br>
          <a href="/tickets" style="color:#4da3ff;">View All Tickets</a>
          <br><br>
          <a href="/download-excel" style="color:#4da3ff;">Download Excel</a>
        </body>
      </html>
    `);
  } catch (err) {
    console.error("RESEND-ALL ERROR:", err);
    res.status(500).send(`<body>Error: ${htmlEscape(err.message)}</body>`);
  }
});

app.get("/tickets", async (req, res) => {
  try {
    const tickets = await listTickets();

    const rows = tickets
      .map(
        (t, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${htmlEscape(t.event_name || "")}</td>
        <td>${htmlEscape(t.pass_label || "")}</td>
        <td>${htmlEscape(t.name || "")}</td>
        <td>${htmlEscape(t.email || "")}</td>
        <td>${htmlEscape(t.contact || "")}</td>
        <td>${htmlEscape(t.pass_type || "")}</td>
        <td>${htmlEscape(t.id || "")}</td>
        <td>${t.assigned ? "Yes" : "No"}</td>
        <td>${htmlEscape(
          t.assigned_time ? formatTimestamp(t.assigned_time) : "-"
        )}</td>
        <td>${t.mail_sent ? "Yes" : "No"}</td>
        <td>${htmlEscape(t.mail_status || "-")}</td>
        <td>${htmlEscape(t.created_at ? formatTimestamp(t.created_at) : "-")}</td>
        <td><a href="${htmlEscape(
          t.payment_ss_url || "#"
        )}" target="_blank" style="color:#56a8ff;">View</a></td>
        <td><a href="/verify/${encodeURIComponent(
          t.id
        )}" target="_blank" style="color:#56a8ff;">Open</a></td>
      </tr>
    `
      )
      .join("");

    res.send(`
      <html>
        <head><title>All Tickets</title></head>
        <body style="margin:0;background:#000;color:#fff;font-family:Arial,sans-serif;padding:30px;">
          <h1 style="text-align:center;">📊 All Ticket Data</h1>
          <div style="text-align:center;margin-bottom:20px;">
            <a href="/" style="color:#56a8ff;font-size:20px;">Back to Portal</a>
            <br><br>
            <a href="/download-excel" style="color:#56a8ff;font-size:20px;">Download Excel</a>
          </div>
          <div style="overflow:auto;">
            <table border="1" cellspacing="0" cellpadding="12" style="width:100%;border-collapse:collapse;background:#111;color:white;min-width:1600px;">
              <thead style="background:#222;">
                <tr>
                  <th>No.</th>
                  <th>Event Name</th>
                  <th>Pass Label</th>
                  <th>Name</th>
                  <th>Email</th>
                  <th>Contact</th>
                  <th>Pass Type</th>
                  <th>Ticket ID</th>
                  <th>Assigned</th>
                  <th>Assigned Time</th>
                  <th>Mail Sent</th>
                  <th>Mail Status</th>
                  <th>Created At</th>
                  <th>Payment SS</th>
                  <th>Verify</th>
                </tr>
              </thead>
              <tbody>
                ${rows || '<tr><td colspan="15" style="text-align:center;">No data found</td></tr>'}
              </tbody>
            </table>
          </div>
        </body>
      </html>
    `);
  } catch (err) {
    console.error("TICKETS ERROR:", err);
    res.status(500).send(`<body>Error: ${htmlEscape(err.message)}</body>`);
  }
});

app.get("/download-excel", async (req, res) => {
  try {
    const tickets = await listTickets();
    const { buffer } = await uploadExcelToSupabase(tickets);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="ticket-data.xlsx"'
    );
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("DOWNLOAD EXCEL ERROR:", err);
    res
      .status(500)
      .send(`<body>Error downloading Excel: ${htmlEscape(err.message)}</body>`);
  }
});

app.use((req, res) => {
  res.status(404).send(`
    <html>
      <body style="font-family:Arial,sans-serif;background:#111;color:white;padding:40px;">
        <h1>Not Found</h1>
        <p>Route not found: ${htmlEscape(req.method)} ${htmlEscape(req.originalUrl)}</p>
      </body>
    </html>
  `);
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`🚀 RUNNING ON ${getBaseUrl()}`);
  console.log(`✅ TEST ROUTE: ${getBaseUrl()}/test-route`);
});
