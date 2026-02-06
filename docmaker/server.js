#!/usr/bin/env node
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const PizZip = require("pizzip");
const JSZip = require("jszip");
const nodemailer = require("nodemailer");
const path = require("path");
const fs = require("fs");
const { promisify } = require("util");
const os = require("os");
const { spawn } = require("child_process");
const libre = require("libreoffice-convert");
const mammoth = require("mammoth");
const puppeteer = require("puppeteer");
const crypto = require("crypto");

const upload = multer();
const app = express();
const PORT = process.env.PORT || 3000;
const convertAsync = libre.convertAsync ? libre.convertAsync : promisify(libre.convert);
const OUTPUT_PDF = path.join(__dirname, "saida_pdf");

app.use(express.static(path.join(__dirname, "public")));

const XML_PATH = "word/document.xml";
const LIST_SEPARATOR = ", ";
const LIST_CONJUNCTION = "e";
const DEFAULT_EMAIL_SUBJECT = "Documento gerado";
const DEFAULT_EMAIL_MESSAGE = "Segue em anexo o documento gerado automaticamente.";
const JOB_TTL_MS = 1000 * 60 * 30;
const JOB_CLEANUP_INTERVAL_MS = 1000 * 60 * 5;
const jobs = new Map();

function createJobId() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return crypto.randomBytes(16).toString("hex");
}

function touchJob(job) {
  job.updatedAt = Date.now();
}

function createJob(type) {
  const now = Date.now();
  const job = {
    id: createJobId(),
    type,
    status: "running",
    total: 0,
    current: 0,
    message: "",
    error: "",
    emailTotal: 0,
    emailSent: 0,
    emailFailed: 0,
    paused: false,
    pausePromise: null,
    resume: null,
    result: null,
    createdAt: now,
    updatedAt: now,
  };
  jobs.set(job.id, job);
  return job;
}

function pauseJob(job) {
  if (!job || job.status !== "running") return;
  job.paused = true;
  job.status = "paused";
  if (!job.pausePromise) {
    job.pausePromise = new Promise((resolve) => {
      job.resume = resolve;
    });
  }
  touchJob(job);
}

function resumeJob(job) {
  if (!job || !job.paused) return;
  job.paused = false;
  job.status = "running";
  if (job.resume) job.resume();
  job.pausePromise = null;
  job.resume = null;
  touchJob(job);
}

async function waitIfPaused(job) {
  if (!job || !job.paused) return;
  if (!job.pausePromise) {
    job.pausePromise = new Promise((resolve) => {
      job.resume = resolve;
    });
  }
  await job.pausePromise;
}

function cleanupJobs() {
  const now = Date.now();
  jobs.forEach((job, id) => {
    if (now - job.updatedAt > JOB_TTL_MS) {
      jobs.delete(id);
    }
  });
}

setInterval(cleanupJobs, JOB_CLEANUP_INTERVAL_MS).unref();

function normalizeXml(xml) {
  // remove espaços entre chaves e colapsa sequências para evitar delimitadores quebrados
  return xml
    .replace(/\{\s+/g, "{")
    .replace(/\s+\}/g, "}")
    .replace(/\{ {0,}\{+/g, "{{")
    .replace(/\} {0,}\}+/g, "}}")
    .replace(/\{+/g, (m) => (m.length >= 2 ? "{{" : "{"))
    .replace(/\}+/g, (m) => (m.length >= 2 ? "}}" : "}"));
}

function fixSplitPlaceholders(xml) {
  // Encontra blocos entre {{ e }} mesmo atravessando tags, remove tags internas e recompõe o placeholder
  let result = "";
  let cursor = 0;
  while (true) {
    const start = xml.indexOf("{{", cursor);
    if (start === -1) break;
    const end = xml.indexOf("}}", start + 2);
    if (end === -1) break;
    const chunk = xml.slice(start, end + 2);
    const cleaned = chunk.replace(/<[^>]+>/g, "");
    const match = cleaned.match(/{{\s*([^}]+?)\s*}}/);
    const placeholder = match ? match[1].trim() : cleaned.replace(/[{}]/g, "").trim();
    result += xml.slice(cursor, start) + `{{${placeholder}}}`;
    cursor = end + 2;
  }
  result += xml.slice(cursor);
  return result;
}

function sanitizeTemplateBuffer(templateBuffer) {
  const zip = new PizZip(templateBuffer);
  const xml = zip.file(XML_PATH).asText();
  const cleaned = fixSplitPlaceholders(normalizeXml(xml));
  zip.file(XML_PATH, cleaned);
  return zip.generate({ type: "nodebuffer" });
}

function extractPlaceholders(templateBuffer) {
  // Usa o XML saneado para detectar placeholders reais, sem tags XML expostas
  const cleaned = sanitizeTemplateBuffer(templateBuffer);
  const zip = new PizZip(cleaned);
  const xml = zip.file(XML_PATH).asText();
  const plain = xml.replace(/<[^>]+>/g, "");
  const regex = /{{\s*([^}]+?)\s*}}/g;
  const placeholders = new Set();
  let match;
  while ((match = regex.exec(plain)) !== null) {
    placeholders.add(String(match[1] || "").trim());
  }
  return Array.from(placeholders);
}

function readColumns(excelBuffer) {
  const wb = xlsx.read(excelBuffer, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const header = rows[0] || [];
  return header.map((c) => String(c || "").trim()).filter(Boolean);
}

function normalizeRowKeys(row) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const trimmedKey = String(key || "").trim();
    if (!trimmedKey) return;
    normalized[trimmedKey] = value;
  });
  return normalized;
}

function readRows(excelBuffer) {
  const wb = xlsx.read(excelBuffer, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  // Mantem linhas em branco parciais; so filtra se estiver tudo vazio
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: "", blankrows: true }).map(normalizeRowKeys);
  return rows.filter((row) => Object.values(row).some((v) => String(v || "").trim() !== ""));
}

function sanitizeCellValue(rawValue) {
  let value = rawValue ?? "";
  if (typeof value === "string") {
    value = value.replace(/\"/g, "");
  }
  return String(value);
}

function getRowValue(row, columnName) {
  if (!row || !columnName) return "";
  if (Object.prototype.hasOwnProperty.call(row, columnName)) {
    return row[columnName];
  }
  const target = String(columnName || "").trim().toLowerCase();
  if (!target) return "";
  const key = Object.keys(row).find((k) => String(k || "").trim().toLowerCase() === target);
  return key ? row[key] : "";
}

function isPdfBuffer(buffer) {
  if (!buffer) return false;
  const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  if (buf.length < 4) return false;
  return buf.slice(0, 4).toString() === "%PDF";
}

function normalizePdfFilename(filename) {
  const base = sanitizeFilename(filename || "documento.pdf");
  if (base.toLowerCase().endsWith(".pdf")) return base;
  if (base.toLowerCase().endsWith(".docx")) {
    return base.replace(/\.docx$/i, ".pdf");
  }
  return `${base}.pdf`;
}

function preparePdfAttachment(attachment, context = {}) {
  if (!attachment || !attachment.buffer) return null;
  const buffer = Buffer.isBuffer(attachment.buffer)
    ? attachment.buffer
    : Buffer.from(attachment.buffer);
  if (!isPdfBuffer(buffer)) {
    const ctx = Object.entries(context)
      .map(([key, value]) => `${key}=${value}`)
      .join(" ");
    console.warn("Anexo ignorado: buffer não é PDF.", ctx);
    return null;
  }
  return {
    filename: normalizePdfFilename(attachment.filename),
    buffer,
  };
}

function ensureUniqueAttachmentName(usedNamesByRecipient, recipient, filename) {
  const key = String(recipient || "").toLowerCase();
  if (!key) return filename;
  const set = usedNamesByRecipient.get(key) || new Set();
  let finalName = filename;
  let counter = 2;
  const base = filename.replace(/\.pdf$/i, "");
  while (set.has(finalName.toLowerCase())) {
    finalName = `${base}_${counter}.pdf`;
    counter += 1;
  }
  set.add(finalName.toLowerCase());
  usedNamesByRecipient.set(key, set);
  return finalName;
}

function buildListPrefix(basePrefix, position, total) {
  const prefix = basePrefix ? String(basePrefix) : "";
  if (total <= 1 || position === 0) {
    return prefix;
  }
  if (position === total - 1) {
    const needsSpaceAfter = prefix && !/^\s/.test(prefix) ? " " : "";
    return ` ${LIST_CONJUNCTION}${needsSpaceAfter}${prefix}`;
  }
  return `${LIST_SEPARATOR}${prefix}`;
}

function resolveRowValues(mapping, row) {
  const items = mapping.map((item, index) => {
    const key = String(item.placeholder || "").replace(/^\{\{|\}\}$/g, "");
    const value = sanitizeCellValue(getRowValue(row, item.column));
    const trimmed = value.trim();
    const hasValue = trimmed !== "";
    const prefix = item.prefix ? String(item.prefix) : "";
    const group = item.group ? String(item.group).trim() : "";
    return {
      index,
      key,
      value,
      trimmed,
      hasValue,
      prefix,
      group,
    };
  });

  const seenByGroup = new Map();
  items.forEach((item) => {
    if (!item.group || !item.hasValue) return;
    const normalized = item.trimmed.toLowerCase();
    if (!seenByGroup.has(item.group)) {
      seenByGroup.set(item.group, new Set());
    }
    const groupSet = seenByGroup.get(item.group);
    if (groupSet.has(normalized)) {
      item.hasValue = false;
      return;
    }
    groupSet.add(normalized);
  });

  const groups = new Map();
  items.forEach((item) => {
    if (!item.group) return;
    if (!groups.has(item.group)) {
      groups.set(item.group, []);
    }
    groups.get(item.group).push(item);
  });

  const positions = new Map();
  groups.forEach((groupItems) => {
    const present = groupItems.filter((item) => item.hasValue);
    present.sort((a, b) => a.index - b.index);
    const total = present.length;
    present.forEach((item, pos) => {
      positions.set(item.index, { pos, total });
    });
  });

  const values = {};
  items.forEach((item) => {
    if (!item.hasValue) {
      values[item.key] = "";
      return;
    }
    let prefix = item.prefix;
    const posInfo = item.group ? positions.get(item.index) : null;
    if (posInfo) {
      prefix = buildListPrefix(prefix, posInfo.pos, posInfo.total);
    }
    values[item.key] = `${prefix}${item.value}`;
  });

  return values;
}

function sanitizeFilename(name) {
  const base = String(name || "arquivo").trim() || "arquivo";
  return base.replace(/[\\/:*?"<>|]/g, "_");
}

function parseBoolean(value) {
  if (typeof value === "boolean") return value;
  const text = String(value || "").trim().toLowerCase();
  return ["1", "true", "yes", "on", "sim"].includes(text);
}

function normalizeSmtpConfig(body) {
  return {
    host: String(body?.smtpHost || "").trim(),
    port: String(body?.smtpPort || "").trim(),
    user: String(body?.smtpUser || "").trim(),
    pass: String(body?.smtpPass || ""),
    from: String(body?.smtpFrom || "").trim(),
    secure: parseBoolean(body?.smtpSecure),
  };
}

function normalizeOutlookFrom(body) {
  return String(body?.outlookFrom || "").trim();
}

function parseRecipients(raw) {
  const text = String(raw || "");
  if (!text.trim()) return [];
  const matches = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi);
  if (!matches) return [];
  return matches.map((item) => item.trim()).filter(Boolean);
}

function normalizeEmailAttachments(item) {
  if (Array.isArray(item.attachments) && item.attachments.length) {
    return item.attachments;
  }
  if (item && item.filename && item.buffer) {
    return [{ filename: item.filename, buffer: item.buffer }];
  }
  return [];
}

function formatAttachmentNames(attachments) {
  return attachments
    .map((att) => att.filename)
    .filter(Boolean)
    .join(", ");
}

function queueEmailAttachment(map, recipient, attachment) {
  const key = String(recipient || "").toLowerCase();
  if (!key) return;
  const entry = map.get(key) || { to: recipient, attachments: [] };
  entry.attachments.push(attachment);
  map.set(key, entry);
}

function parsePowerShellJson(raw) {
  let text = String(raw || "").trim();
  if (!text) return [];
  text = text.replace(/^\uFEFF/, "");
  const tryParse = (input) => {
    const parsed = JSON.parse(input);
    return Array.isArray(parsed) ? parsed : [parsed];
  };
  try {
    return tryParse(text);
  } catch (err) {
    const normalized = text.replace(/,\s*([}\]])/g, "$1");
    return tryParse(normalized);
  }
}

function buildEmailReportBuffer(results) {
  const rows = results.map((item, idx) => ({
    linha: idx + 1,
    email: item.to || "",
    arquivo: item.filename || "",
    status: item.status || "",
    erro: item.error || "",
  }));
  const sheet = xlsx.utils.json_to_sheet(rows);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, sheet, "Envio");
  return xlsx.write(workbook, { bookType: "xlsx", type: "buffer" });
}

function escapePowerShellString(value) {
  return String(value || "").replace(/'/g, "''");
}

function runPowerShell(script) {
  return new Promise((resolve, reject) => {
    const child = spawn("powershell", ["-NoProfile", "-NonInteractive", "-Command", script], {
      windowsHide: true,
    });
    let stderr = "";
    let stdout = "";
    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) return resolve(stdout);
      const msg = stderr.trim() || `PowerShell exit code ${code}`;
      return reject(new Error(msg));
    });
  });
}

async function convertWithMsOffice(buffer) {
  if (process.platform !== "win32") return null;
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "docx2pdf-"));
  const inputPath = path.join(tempDir, "input.docx");
  const outputPath = path.join(tempDir, "output.pdf");
  try {
    fs.writeFileSync(inputPath, buffer);
    const psInput = escapePowerShellString(inputPath);
    const psOutput = escapePowerShellString(outputPath);
    const script = `
$ErrorActionPreference = 'Stop'
$inputPath = '${psInput}'
$outputPath = '${psOutput}'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
  $doc = $word.Documents.Open($inputPath, $false, $true)
  $doc.SaveAs([ref]$outputPath, [ref]17)
  $doc.Close($false)
} finally {
  $word.Quit()
  if ($doc) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null }
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
`;
    await runPowerShell(script);
    if (!fs.existsSync(outputPath)) return null;
    const pdf = fs.readFileSync(outputPath);
    return pdf.length ? pdf : null;
  } catch (err) {
    console.error("Falha ao converter para PDF (Microsoft Office):", err.message || err);
    return null;
  } finally {
    try {
      if (fs.rmSync) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } else if (fs.rmdirSync) {
        fs.rmdirSync(tempDir, { recursive: true });
      }
    } catch (err) {
      console.warn("Falha ao limpar temp de conversão:", err.message || err);
    }
  }
}

async function convertWithLibreOffice(buffer) {
  try {
    const pdf = await convertAsync(buffer, ".pdf", undefined);
    return pdf && pdf.length ? pdf : null;
  } catch (err) {
    console.error("Falha ao converter para PDF (LibreOffice):", err.message || err);
    return null;
  }
}

async function convertWithPuppeteer(buffer) {
  try {
    const { value: html = "" } = await mammoth.convertToHtml({ buffer });
    const browser = await puppeteer.launch({ headless: "new" });
    const page = await browser.newPage();
    const styledHtml = `<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <style>
      @page { size: A4; margin: 2.5cm 2cm; }
      body { font-family: "Times New Roman", serif; font-size: 12pt; line-height: 1.2; }
    </style>
  </head>
  <body>${html}</body>
</html>`;
    await page.setContent(styledHtml, { waitUntil: "networkidle0" });
    const pdf = await page.pdf({ format: "A4", printBackground: true });
    await browser.close();
    return pdf;
  } catch (err) {
    console.error("Falha ao converter para PDF (puppeteer):", err.message || err);
    return null;
  }
}

async function docxBufferToPdf(buffer) {
  const msOfficePdf = await convertWithMsOffice(buffer);
  if (msOfficePdf) {
    return msOfficePdf;
  }
  const pdf = await convertWithLibreOffice(buffer);
  if (pdf) {
    return pdf;
  }
  console.warn("Gerando PDF via Puppeteer; o layout pode divergir do DOCX.");
  return convertWithPuppeteer(buffer);
}

async function runGenerateJob(job, input) {
  const rows = input.rows || [];
  job.total = rows.length;
  job.current = 0;
  job.message = "Gerando documentos.";
  touchJob(job);
  try {
    const templateBuffer = input.templateBuffer;
    const mapping = input.mapping || [];
    const placeholders = input.placeholders || [];
    const nameTemplate = input.nameTemplate || "output_{index}_{primary}.docx";
    const primaryPlaceholder = input.primaryPlaceholder || "";
    const emailColumn = input.emailColumn || "";
    const emailSubject = input.emailSubject || "";
    const emailMessage = input.emailMessage || "";
    const emailMode = String(input.emailMode || "smtp").toLowerCase();
    const smtpConfig = input.smtpConfig || {};
    const outlookFrom = input.outlookFrom || "";

    const zipDocx = new JSZip();
    const zipPdf = new JSZip();
    const usedNames = new Set();
    const emailQueue = new Map();
    const usedNamesByRecipient = new Map();

    for (let idx = 0; idx < rows.length; idx += 1) {
      await waitIfPaused(job);
      const row = rows[idx];
      const data = {};
      placeholders.forEach((ph) => {
        data[ph] = "";
      });
      const resolved = resolveRowValues(mapping, row);
      Object.entries(resolved).forEach(([key, value]) => {
        data[key] = value;
      });

      const primaryMap = mapping.find((m) => m.placeholder === primaryPlaceholder) || mapping[0];
      const primaryCol = primaryMap?.column || "linha";
      const primaryValue = getRowValue(row, primaryCol);
      const ctx = { ...row, index: idx + 1, primary: primaryValue || idx + 1 };
      let filename = nameTemplate;
      try {
        filename = filename.replace(/{([^}]+)}/g, (_, key) => String(ctx[key] ?? ""));
      } catch (e) {
        /* noop */
      }
      if (!filename.toLowerCase().endsWith(".docx")) filename += ".docx";
      let finalName = filename;
      let counter = 2;
      while (usedNames.has(finalName)) {
        finalName = filename.replace(/\.docx$/i, `_${counter}.docx`);
        counter += 1;
      }
      usedNames.add(finalName);

      const docBuffer = renderDoc(templateBuffer, data, placeholders);
      zipDocx.file(finalName, docBuffer);

      const pdfName = finalName.replace(/\.docx$/i, ".pdf");
      const pdfBuffer = await docxBufferToPdf(docBuffer);
      const pdfValid = pdfBuffer && isPdfBuffer(pdfBuffer);
      if (pdfValid) {
        zipPdf.file(pdfName, pdfBuffer);
        try {
          fs.mkdirSync(OUTPUT_PDF, { recursive: true });
          fs.writeFileSync(path.join(OUTPUT_PDF, pdfName), pdfBuffer);
        } catch (err) {
          console.error("Falha ao salvar PDF em disco:", err.message || err);
        }
      } else if (pdfBuffer) {
        console.warn(`PDF inválido gerado na linha ${idx + 1}.`);
      }

      if (emailColumn) {
        const recipients = parseRecipients(getRowValue(row, emailColumn));
        const uniqueRecipients = Array.from(
          new Map(recipients.map((addr) => [addr.toLowerCase(), addr])).values()
        );
        if (uniqueRecipients.length) {
          if (!pdfValid) {
            console.warn("PDF não gerado; email não énviado para", uniqueRecipients.join(", "));
          } else {
            const attachment = preparePdfAttachment(
              { filename: pdfName, buffer: pdfBuffer },
              { row: idx + 1 }
            );
            uniqueRecipients.forEach((recipient) => {
              if (!attachment) return;
              const uniqueName = ensureUniqueAttachmentName(
                usedNamesByRecipient,
                recipient,
                attachment.filename
              );
              queueEmailAttachment(emailQueue, recipient, { filename: uniqueName, buffer: attachment.buffer });
            });
          }
        }
      }

      job.current = idx + 1;
      job.message = `Processando linha ${idx + 1} de ${rows.length}.`;
      touchJob(job);
    }

    job.message = "Finalizando arquivos.";
    touchJob(job);
    const zipDocxBuffer = await zipDocx.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const pdfHasFiles = Object.keys(zipPdf.files || {}).length > 0;
    const zipPdfBuffer = pdfHasFiles
      ? await zipPdf.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
      : null;

    job.result = {
      docxZipBuffer: zipDocxBuffer,
      pdfZipBuffer: zipPdfBuffer,
      pdfGenerated: pdfHasFiles,
    };
    job.status = "done";
    job.message = "Concluido.";
    touchJob(job);

    const emailsToSend = Array.from(emailQueue.values());
    if (emailsToSend.length) {
      const sendPromise =
        emailMode === "outlook"
          ? sendEmailsOutlook(emailsToSend, emailSubject, emailMessage, outlookFrom)
          : sendEmailsSmtp(emailsToSend, emailSubject, emailMessage, smtpConfig);
      sendPromise.catch((e) => {
        console.error("Falha ao enviar emails:", e.message || e);
      });
    }
  } catch (err) {
    job.status = "error";
    job.error = err.message || "Erro ao gerar";
    touchJob(job);
  }
}

async function runSendJob(job, input) {
  const rows = input.rows || [];
  job.total = rows.length;
  job.current = 0;
  job.message = "Gerando PDFs.";
  touchJob(job);
  try {
    const templateBuffer = input.templateBuffer;
    const mapping = input.mapping || [];
    const placeholders = input.placeholders || [];
    const primaryPlaceholder = input.primaryPlaceholder || "";
    const emailColumn = input.emailColumn || "";
    const emailSubject = input.emailSubject || "";
    const emailMessage = input.emailMessage || "";
    const emailMode = String(input.emailMode || "smtp").toLowerCase();
    const smtpConfig = input.smtpConfig || {};
    const outlookFrom = input.outlookFrom || "";
    const nameTemplate = input.nameTemplate || "output_{index}_{primary}.docx";

    const emailQueue = new Map();
    const usedNamesByRecipient = new Map();
    let pdfFailures = 0;
    let recipientCount = 0;

    for (let idx = 0; idx < rows.length; idx += 1) {
      await waitIfPaused(job);
      const row = rows[idx];
      const data = {};
      placeholders.forEach((ph) => {
        data[ph] = "";
      });
      const resolved = resolveRowValues(mapping, row);
      Object.entries(resolved).forEach(([key, value]) => {
        data[key] = value;
      });

      const primaryMap = mapping.find((m) => m.placeholder === primaryPlaceholder) || mapping[0];
      const primaryCol = primaryMap?.column || "linha";
      const primaryValue = getRowValue(row, primaryCol);
      const ctx = { ...row, index: idx + 1, primary: primaryValue || idx + 1 };
      let filename = nameTemplate;
      try {
        filename = filename.replace(/{([^}]+)}/g, (_, key) => String(ctx[key] ?? ""));
      } catch (e) {
        /* noop */
      }
      if (!filename.toLowerCase().endsWith(".docx")) filename += ".docx";

      const recipients = parseRecipients(getRowValue(row, emailColumn));
      const uniqueRecipients = Array.from(
        new Map(recipients.map((addr) => [addr.toLowerCase(), addr])).values()
      );
      if (!uniqueRecipients.length) {
        job.current = idx + 1;
        touchJob(job);
        continue;
      }
      recipientCount += uniqueRecipients.length;
      const buffer = renderDoc(templateBuffer, data, placeholders);
      const pdfBuffer = await docxBufferToPdf(buffer);
      const pdfValid = pdfBuffer && isPdfBuffer(pdfBuffer);
      if (!pdfValid) {
        pdfFailures += uniqueRecipients.length;
        job.current = idx + 1;
        touchJob(job);
        continue;
      }
      const attachName = filename.replace(/\.docx$/i, ".pdf");
      const attachment = preparePdfAttachment({ filename: attachName, buffer: pdfBuffer }, { row: idx + 1 });
      uniqueRecipients.forEach((recipient) => {
        if (!attachment) return;
        const uniqueName = ensureUniqueAttachmentName(usedNamesByRecipient, recipient, attachment.filename);
        queueEmailAttachment(emailQueue, recipient, { filename: uniqueName, buffer: attachment.buffer });
      });

      job.current = idx + 1;
      job.message = `Processando linha ${idx + 1} de ${rows.length}.`;
      touchJob(job);
    }

    const emailsToSend = Array.from(emailQueue.values());
    if (!emailsToSend.length) {
      let msg = "Nenhum destinatário encontrado.";
      if (recipientCount && pdfFailures) {
        msg = "PDF não foi gerado. Verifique o LibreOffice no servidor.";
      } else if (recipientCount) {
        msg = "Nenhum anexo gerado para envio.";
      }
      job.status = "error";
      job.error = msg;
      touchJob(job);
      return;
    }

    job.emailTotal = emailsToSend.length;
    job.emailSent = 0;
    job.emailFailed = 0;
    job.message = "Enviando emails.";
    touchJob(job);
    await waitIfPaused(job);

    let result;
    if (emailMode === "outlook") {
      result = await sendEmailsOutlook(emailsToSend, emailSubject, emailMessage, outlookFrom);
      job.emailSent = result.sent;
      job.emailFailed = result.failed;
      job.message = "Envio finalizado.";
      touchJob(job);
    } else {
      result = await sendEmailsSmtp(
        emailsToSend,
        emailSubject,
        emailMessage,
        smtpConfig,
        ({ processed, total, sent, failed }) => {
          job.emailTotal = total;
          job.emailSent = sent;
          job.emailFailed = failed;
          job.message = `Enviando emails: ${processed} de ${total}.`;
          touchJob(job);
        }
      );
    }

    const reportBuffer =
      result && result.results && result.results.length ? buildEmailReportBuffer(result.results) : null;
    job.result = {
      sent: result?.sent || 0,
      failed: result?.failed || 0,
      errors: result?.errors || [],
      reportBuffer,
    };
    job.status = "done";
    job.message = "Concluido.";
    touchJob(job);
  } catch (err) {
    job.status = "error";
    job.error = err.message || "Erro ao enviar emails";
    touchJob(job);
  }
}

function buildSmtpOptions(config) {
  const host = config.host || process.env.SMTP_HOST;
  const port = Number(config.port || process.env.SMTP_PORT || 587);
  const user = config.user || process.env.SMTP_USER;
  const pass = config.pass || process.env.SMTP_PASS;
  if (!host || !user || !pass) return null;
  const secure = typeof config.secure === "boolean" ? config.secure : port === 465;
  const from = config.from || process.env.SMTP_FROM || user;
  return {
    host,
    port,
    secure,
    auth: { user, pass },
    from,
  };
}

async function sendEmailsSmtp(queue, subject, message, smtpConfig, onProgress) {
  const options = buildSmtpOptions(smtpConfig || {});
  if (!options) {
    console.warn("SMTP não configurado; emails não énviados.");
    const results = queue.map((item) => {
      const attachments = normalizeEmailAttachments(item);
      return {
        to: item.to,
        filename: formatAttachmentNames(attachments),
        status: "failed",
        error: "SMTP não configurado",
      };
    });
    return { sent: 0, failed: queue.length, errors: ["SMTP não configurado"], results };
  }
  const transporter = nodemailer.createTransport(options);
  const from = options.from;
  let sent = 0;
  let processed = 0;
  const errors = [];
  const results = [];
  for (const item of queue) {
    const attachments = normalizeEmailAttachments(item);
    const filenameList = formatAttachmentNames(attachments);
    if (!attachments.length) {
      const errorMsg = "Sem anexos para enviar";
      errors.push(`Erro para ${item.to}: ${errorMsg}`);
      results.push({ to: item.to, filename: filenameList, status: "failed", error: errorMsg });
      processed += 1;
      if (typeof onProgress === "function") {
        const failed = processed - sent;
        onProgress({ processed, total: queue.length, sent, failed });
      }
      continue;
    }
    const mail = {
      from,
      to: item.to,
      subject: subject || DEFAULT_EMAIL_SUBJECT,
      text: message || DEFAULT_EMAIL_MESSAGE,
      attachments: attachments.map((att) => ({
        filename: att.filename,
        content: att.buffer,
        contentType: "application/pdf",
      })),
    };
    try {
      await transporter.sendMail(mail);
      console.log(`Email enviado para ${item.to} (${filenameList || "sem anexos"})`);
      sent += 1;
      results.push({ to: item.to, filename: filenameList, status: "sent", error: "" });
    } catch (err) {
      console.error(`Falha ao enviar para ${item.to}:`, err.message || err);
      const msg = err.message || err;
      errors.push(`Erro para ${item.to}: ${msg}`);
      results.push({ to: item.to, filename: filenameList, status: "failed", error: String(msg) });
    } finally {
      processed += 1;
      if (typeof onProgress === "function") {
        const failed = processed - sent;
        onProgress({ processed, total: queue.length, sent, failed });
      }
    }
  }
  return { sent, failed: queue.length - sent, errors, results };
}

async function sendEmailsOutlook(queue, subject, message, outlookFrom) {
  if (process.platform !== "win32") {
    return { sent: 0, failed: queue.length, errors: ["Outlook disponivel apenas no Windows."] };
  }
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-mail-"));
  const fromAddress = String(outlookFrom || "").trim();
  try {
    const payload = queue.map((item, idx) => {
      const attachments = normalizeEmailAttachments(item);
      const prepared = attachments.map((att, attIdx) => {
        const filename = sanitizeFilename(att.filename || `documento_${idx + 1}_${attIdx + 1}.pdf`);
        const attachmentPath = path.join(tempDir, `att_${idx + 1}_${attIdx + 1}.pdf`);
        fs.writeFileSync(attachmentPath, att.buffer);
        return { attachmentPath, filename };
      });
      return {
        to: item.to,
        subject: subject || DEFAULT_EMAIL_SUBJECT,
        body: message || DEFAULT_EMAIL_MESSAGE,
        attachments: prepared,
      };
    });

    const jsonPath = path.join(tempDir, "queue.json");
    fs.writeFileSync(jsonPath, JSON.stringify(payload), "utf8");

    const psJsonPath = escapePowerShellString(jsonPath);
    const psFrom = escapePowerShellString(fromAddress);
    const script = `
$ErrorActionPreference = 'Stop'
$items = Get-Content -Path '${psJsonPath}' -Raw | ConvertFrom-Json
$fromAddress = '${psFrom}'
$report = @()
$skipSend = $false
$created = $false
try {
  $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
  $outlook = New-Object -ComObject Outlook.Application
  $created = $true
}
try {
  $account = $null
  if ($fromAddress) {
    $account = $outlook.Session.Accounts | Where-Object { $_.SmtpAddress -ieq $fromAddress } | Select-Object -First 1
    if (-not $account) {
      $available = $outlook.Session.Accounts | ForEach-Object { $_.SmtpAddress } | Where-Object { $_ }
      $list = ($available -join ", ")
      $errorMsg = "Conta Outlook não éncontrada: $fromAddress. Disponíveis: $list"
      foreach ($item in $items) {
        $fileList = ''
        if ($item.attachments) { $fileList = ($item.attachments | ForEach-Object { $_.filename }) -join ', ' }
        $report += [PSCustomObject]@{ to = $item.to; status = 'failed'; error = $errorMsg; filename = $fileList }
      }
      $skipSend = $true
    }
  }
      foreach ($item in $items) {
    if ($skipSend) { break }
    if (-not $item.to) {
      $report += [PSCustomObject]@{ to = $item.to; status = 'skipped'; error = 'Destinatário vazio'; filename = '' }
      continue
    }
    $attachments = @()
    if ($item.attachments) { $attachments = $item.attachments }
    $fileList = ($attachments | ForEach-Object { $_.filename }) -join ', '
    if (-not $attachments -or $attachments.Count -eq 0) {
      $report += [PSCustomObject]@{ to = $item.to; status = 'failed'; error = 'Sem anexos'; filename = $fileList }
      continue
    }
    $mail = $outlook.CreateItem(0)
    if ($account) { $mail.SendUsingAccount = $account }
    $mail.Subject = $item.subject
    $mail.Body = $item.body
    foreach ($att in $attachments) { $null = $mail.Attachments.Add($att.attachmentPath, 1, [System.Type]::Missing, $att.filename) }
    $recipients = $mail.Recipients
    $addresses = ($item.to -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($addr in $addresses) { $null = $recipients.Add($addr) }
    $null = $recipients.ResolveAll()
    try {
      $mail.Send()
      $report += [PSCustomObject]@{ to = $item.to; status = 'sent'; filename = $fileList }
    } catch {
      $errMsg = $_.Exception.Message
      $report += [PSCustomObject]@{ to = $item.to; status = 'failed'; error = $errMsg; filename = $fileList }
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
  }
} finally {
  if ($created) { $outlook.Quit() }
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
}
$report | ConvertTo-Json -Depth 4
`;
    const output = await runPowerShell(script);
    let report = [];
    try {
      report = parsePowerShellJson(output);
    } catch (err) {
      throw new Error(`Falha ao ler relatório do Outlook: ${err.message || err}`);
    }
    const sent = report.filter((r) => r.status === "sent").length;
    const failed = report.filter((r) => r.status === "failed").length;
    const errors = report.filter((r) => r.error).map((r) => `${r.to}: ${r.error}`);
    const skipped = report.filter((r) => r.status === "skipped").length;
    const results = report.map((r) => ({
      to: r.to,
      filename: r.filename,
      status: r.status,
      error: r.error || "",
    }));
    return { sent, failed: failed + skipped, errors, results };
  } catch (err) {
    console.error("Falha ao enviar via Outlook:", err.message || err);
    const results = queue.map((item) => {
      const attachments = normalizeEmailAttachments(item);
      return {
        to: item.to,
        filename: formatAttachmentNames(attachments),
        status: "failed",
        error: err.message || String(err),
      };
    });
    return { sent: 0, failed: queue.length, errors: [err.message || String(err)], results };
  } finally {
    try {
      if (fs.rmSync) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } else if (fs.rmdirSync) {
        fs.rmdirSync(tempDir, { recursive: true });
      }
    } catch (err) {
      console.warn("Falha ao limpar temp de email:", err.message || err);
    }
  }
}

function escapeXml(val) {
  return String(val || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function preserveTextSpaces(xml) {
  return xml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (match, attrs, text) => {
    if (!text) return match;
    if (!/^\s|\s$/.test(text)) return match;
    if (attrs && /xml:space="preserve"/.test(attrs)) return match;
    const newAttrs = attrs ? `${attrs} xml:space="preserve"` : ' xml:space="preserve"';
    return `<w:t${newAttrs}>${text}</w:t>`;
  });
}

function renderDoc(templateBuffer, data, placeholders) {
  // Substituição manual apenas (evita erros de delimitador do Docxtemplater)
  const cleaned = sanitizeTemplateBuffer(templateBuffer);
  const zip = new PizZip(cleaned);
  let xml = zip.file(XML_PATH).asText();
  Object.entries(data).forEach(([key, val]) => {
    const safeKey = String(key || "").trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const simple = new RegExp(`{{\\s*${safeKey}\\s*}}`, "g");
    xml = xml.replace(simple, escapeXml(val));
  });
  placeholders.forEach((ph) => {
    const safeKey = String(ph || "").trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const simple = new RegExp(`{{\\s*${safeKey}\\s*}}`, "g");
    xml = xml.replace(simple, "");
  });
  xml = preserveTextSpaces(xml);
  zip.file(XML_PATH, xml);
  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

app.post(
  "/inspect",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "excel", maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const template = req.files?.template?.[0];
      if (!template) {
        return res.status(400).json({ error: "Envie o template DOCX." });
      }
      const placeholders = extractPlaceholders(template.buffer);
      let columns = [];
      if (req.files?.excel?.[0]) {
        columns = readColumns(req.files.excel[0].buffer);
      }
      res.json({ placeholders, columns });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message || "Erro ao inspecionar" });
    }
  }
);

app.post(
  "/generate-job",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "excel", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const template = req.files?.template?.[0];
      const excel = req.files?.excel?.[0];
      const mapping = JSON.parse(req.body.mapping || "[]");
      const nameTemplate = req.body.nameTemplate || "output_{index}_{primary}.docx";
      const primaryPlaceholder = req.body.primaryPlaceholder || "";
      const emailColumn = req.body.emailColumn || "";
      const emailSubject = req.body.emailSubject || "";
      const emailMessage = req.body.emailMessage || "";
      const emailMode = String(req.body.emailMode || "smtp").toLowerCase();
      const smtpConfig = normalizeSmtpConfig(req.body);
      const outlookFrom = normalizeOutlookFrom(req.body);

      if (!template) return res.status(400).json({ error: "Template faltando." });
      if (!excel) return res.status(400).json({ error: "Planilha faltando." });
      if (!mapping.length) return res.status(400).json({ error: "Mapping faltando." });

      const rows = readRows(excel.buffer);
      const placeholders = extractPlaceholders(template.buffer);
      const job = createJob("generate");
      job.total = rows.length;
      job.message = "Fila criada.";
      touchJob(job);

      runGenerateJob(job, {
        templateBuffer: template.buffer,
        rows,
        mapping,
        placeholders,
        nameTemplate,
        primaryPlaceholder,
        emailColumn,
        emailSubject,
        emailMessage,
        emailMode,
        smtpConfig,
        outlookFrom,
      });

      res.json({ jobId: job.id, total: job.total });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message || "Erro ao iniciar geração" });
    }
  }
);

app.post(
  "/send-emails-job",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "excel", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const template = req.files?.template?.[0];
      const excel = req.files?.excel?.[0];
      const mapping = JSON.parse(req.body.mapping || "[]");
      const primaryPlaceholder = req.body.primaryPlaceholder || "";
      const emailColumn = req.body.emailColumn || "";
      const emailSubject = req.body.emailSubject || "";
      const emailMessage = req.body.emailMessage || "";
      const emailMode = String(req.body.emailMode || "smtp").toLowerCase();
      const smtpConfig = normalizeSmtpConfig(req.body);
      const outlookFrom = normalizeOutlookFrom(req.body);
      const nameTemplate = req.body.nameTemplate || "output_{index}_{primary}.docx";

      if (!template) return res.status(400).json({ error: "Template faltando." });
      if (!excel) return res.status(400).json({ error: "Planilha faltando." });
      if (!mapping.length) return res.status(400).json({ error: "Mapping faltando." });
      if (!emailColumn) return res.status(400).json({ error: "Defina a coluna de email." });

      const rows = readRows(excel.buffer);
      const placeholders = extractPlaceholders(template.buffer);
      const job = createJob("send");
      job.total = rows.length;
      job.message = "Fila criada.";
      touchJob(job);

      runSendJob(job, {
        templateBuffer: template.buffer,
        rows,
        mapping,
        placeholders,
        primaryPlaceholder,
        emailColumn,
        emailSubject,
        emailMessage,
        emailMode,
        smtpConfig,
        outlookFrom,
        nameTemplate,
      });

      res.json({ jobId: job.id, total: job.total });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message || "Erro ao iniciar envio" });
    }
  }
);

app.get("/jobs/:id", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job não éncontrado." });
  res.json({
    id: job.id,
    type: job.type,
    status: job.status,
    total: job.total,
    current: job.current,
    message: job.message || "",
    error: job.error || "",
    emailTotal: job.emailTotal || 0,
    emailSent: job.emailSent || 0,
    emailFailed: job.emailFailed || 0,
    resultReady: job.status === "done",
  });
});

app.get("/jobs/:id/result", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job não éncontrado." });
  if (job.status !== "done") return res.status(400).json({ error: "Resultado indisponivel." });
  if (job.type === "generate") {
    if (!job.result?.docxZipBuffer) {
      return res.status(400).json({ error: "Resultado incompleto." });
    }
    return res.json({
      docxZip: job.result.docxZipBuffer.toString("base64"),
      pdfZip: job.result.pdfZipBuffer ? job.result.pdfZipBuffer.toString("base64") : null,
      pdfGenerated: Boolean(job.result.pdfGenerated),
    });
  }
  const reportBuffer = job.result?.reportBuffer || null;
  return res.json({
    sent: job.result?.sent || 0,
    failed: job.result?.failed || 0,
    errors: job.result?.errors || [],
    reportXlsx: reportBuffer ? reportBuffer.toString("base64") : null,
  });
});

app.post("/jobs/:id/pause", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job não éncontrado." });
  if (job.status === "done" || job.status === "error") {
    return res.json({ status: job.status });
  }
  pauseJob(job);
  return res.json({ status: job.status });
});

app.post("/jobs/:id/resume", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job não éncontrado." });
  if (job.status === "done" || job.status === "error") {
    return res.json({ status: job.status });
  }
  resumeJob(job);
  return res.json({ status: job.status });
});

app.post(
  "/generate",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "excel", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const template = req.files?.template?.[0];
      const excel = req.files?.excel?.[0];
      const mapping = JSON.parse(req.body.mapping || "[]");
      const nameTemplate = req.body.nameTemplate || "output_{index}_{primary}.docx";
      const primaryPlaceholder = req.body.primaryPlaceholder || "";
      const emailColumn = req.body.emailColumn || "";
      const emailSubject = req.body.emailSubject || "";
      const emailMessage = req.body.emailMessage || "";
      const emailMode = String(req.body.emailMode || "smtp").toLowerCase();
      const smtpConfig = normalizeSmtpConfig(req.body);
      const outlookFrom = normalizeOutlookFrom(req.body);

      if (!template) return res.status(400).json({ error: "Template faltando." });
      if (!excel) return res.status(400).json({ error: "Planilha faltando." });
      if (!mapping.length) return res.status(400).json({ error: "Mapping faltando." });

      const rows = readRows(excel.buffer);
      const placeholders = extractPlaceholders(template.buffer);
      const zipDocx = new JSZip();
      const zipPdf = new JSZip();
      const usedNames = new Set();
      const emailQueue = new Map();

      for (let idx = 0; idx < rows.length; idx += 1) {
        const row = rows[idx];
        const data = {};
        // Preenche todos os placeholders encontrados com vazio para evitar erros de tag não resolvida
        placeholders.forEach((ph) => {
          data[ph] = "";
        });
        const resolved = resolveRowValues(mapping, row);
        Object.entries(resolved).forEach(([key, value]) => {
          data[key] = value;
        });

        const primaryMap = mapping.find((m) => m.placeholder === primaryPlaceholder) || mapping[0];
        const primaryCol = primaryMap?.column || "linha";
        const primaryValue = getRowValue(row, primaryCol);
        const ctx = { ...row, index: idx + 1, primary: primaryValue || idx + 1 };
        let filename = nameTemplate;
        try {
          filename = filename.replace(/{([^}]+)}/g, (_, key) => String(ctx[key] ?? ""));
        } catch (e) {
          /* noop */
        }
        if (!filename.toLowerCase().endsWith(".docx")) filename += ".docx";
        // Evita sobrescrever nomes duplicados
        let finalName = filename;
        let counter = 2;
        while (usedNames.has(finalName)) {
          finalName = filename.replace(/\.docx$/i, `_${counter}.docx`);
          counter += 1;
        }
        usedNames.add(finalName);

        const docBuffer = renderDoc(template.buffer, data, placeholders);
        zipDocx.file(finalName, docBuffer);

        const pdfName = finalName.replace(/\.docx$/i, ".pdf");
        const pdfBuffer = await docxBufferToPdf(docBuffer);
        const pdfValid = pdfBuffer && isPdfBuffer(pdfBuffer);
        if (pdfValid) {
          zipPdf.file(pdfName, pdfBuffer);
          try {
            fs.mkdirSync(OUTPUT_PDF, { recursive: true });
            fs.writeFileSync(path.join(OUTPUT_PDF, pdfName), pdfBuffer);
          } catch (err) {
            console.error("Falha ao salvar PDF em disco:", err.message || err);
          }
        } else if (pdfBuffer) {
          console.warn(`PDF inválido gerado na linha ${idx + 1}.`);
        }

        // Envio de email opcional (apenas PDF)
        if (emailColumn) {
          const recipients = parseRecipients(getRowValue(row, emailColumn));
          const uniqueRecipients = Array.from(
            new Map(recipients.map((addr) => [addr.toLowerCase(), addr])).values()
          );
          if (uniqueRecipients.length) {
            if (!pdfValid) {
              console.warn("PDF não gerado; email não énviado para", uniqueRecipients.join(", "));
            } else {
              const attachment = preparePdfAttachment({ filename: pdfName, buffer: pdfBuffer }, { row: idx + 1 });
              uniqueRecipients.forEach((recipient) => {
                if (attachment) {
                  queueEmailAttachment(emailQueue, recipient, attachment);
                }
              });
            }
          }
        }
      }

      const zipDocxBuffer = await zipDocx.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
      const pdfHasFiles = Object.keys(zipPdf.files || {}).length > 0;
      const zipPdfBuffer = pdfHasFiles
        ? await zipPdf.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
        : null;

      res.json({
        docxZip: zipDocxBuffer.toString("base64"),
        pdfZip: zipPdfBuffer ? zipPdfBuffer.toString("base64") : null,
        pdfGenerated: pdfHasFiles,
      });

      // Envia emails em background (melhor esforco)
      const emailsToSend = Array.from(emailQueue.values());
      if (emailsToSend.length) {
        const sendPromise =
          emailMode === "outlook"
            ? sendEmailsOutlook(emailsToSend, emailSubject, emailMessage, outlookFrom)
            : sendEmailsSmtp(emailsToSend, emailSubject, emailMessage, smtpConfig);
        sendPromise.catch((e) => {
          console.error("Falha ao enviar emails:", e.message || e);
        });
      }
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message || "Erro ao gerar" });
    }
  }
);

// API para apenas enviar os emails (sem retornar ZIP)
app.post(
  "/send-emails",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "excel", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const template = req.files?.template?.[0];
      const excel = req.files?.excel?.[0];
      const mapping = JSON.parse(req.body.mapping || "[]");
      const primaryPlaceholder = req.body.primaryPlaceholder || "";
      const emailColumn = req.body.emailColumn || "";
      const emailSubject = req.body.emailSubject || "";
      const emailMessage = req.body.emailMessage || "";
      const emailMode = String(req.body.emailMode || "smtp").toLowerCase();
      const smtpConfig = normalizeSmtpConfig(req.body);
      const outlookFrom = normalizeOutlookFrom(req.body);

      if (!template) return res.status(400).json({ error: "Template faltando." });
      if (!excel) return res.status(400).json({ error: "Planilha faltando." });
      if (!mapping.length) return res.status(400).json({ error: "Mapping faltando." });
      if (!emailColumn) return res.status(400).json({ error: "Defina a coluna de email." });

      const rows = readRows(excel.buffer);
      const placeholders = extractPlaceholders(template.buffer);
      const emailQueue = new Map();
      const usedNamesByRecipient = new Map();
      let pdfFailures = 0;
      let recipientCount = 0;

      for (let idx = 0; idx < rows.length; idx += 1) {
        const row = rows[idx];
        const data = {};
        placeholders.forEach((ph) => {
          data[ph] = "";
        });
        const resolved = resolveRowValues(mapping, row);
        Object.entries(resolved).forEach(([key, value]) => {
          data[key] = value;
        });

        const primaryMap = mapping.find((m) => m.placeholder === primaryPlaceholder) || mapping[0];
        const primaryCol = primaryMap?.column || "linha";
        const primaryValue = getRowValue(row, primaryCol);
        const ctx = { ...row, index: idx + 1, primary: primaryValue || idx + 1 };
        let filename = req.body.nameTemplate || "output_{index}_{primary}.docx";
        try {
          filename = filename.replace(/{([^}]+)}/g, (_, key) => String(ctx[key] ?? ""));
        } catch (e) {
          /* noop */
        }
        if (!filename.toLowerCase().endsWith(".docx")) filename += ".docx";

        const recipients = parseRecipients(getRowValue(row, emailColumn));
        const uniqueRecipients = Array.from(
          new Map(recipients.map((addr) => [addr.toLowerCase(), addr])).values()
        );
        if (!uniqueRecipients.length) continue;
        recipientCount += uniqueRecipients.length;
        const buffer = renderDoc(template.buffer, data, placeholders);
        const pdfBuffer = await docxBufferToPdf(buffer);
        const pdfValid = pdfBuffer && isPdfBuffer(pdfBuffer);
        if (!pdfValid) {
          pdfFailures += uniqueRecipients.length;
          continue;
        }
        const attachName = filename.replace(/\.docx$/i, ".pdf");
        const attachment = preparePdfAttachment({ filename: attachName, buffer: pdfBuffer }, { row: idx + 1 });
        uniqueRecipients.forEach((recipient) => {
          if (!attachment) return;
          const uniqueName = ensureUniqueAttachmentName(usedNamesByRecipient, recipient, attachment.filename);
          queueEmailAttachment(emailQueue, recipient, { filename: uniqueName, buffer: attachment.buffer });
        });
      }

      const emailsToSend = Array.from(emailQueue.values());
      if (!emailsToSend.length) {
        let msg = "Nenhum destinatário encontrado.";
        if (recipientCount && pdfFailures) {
          msg = "PDF não foi gerado. Verifique o LibreOffice no servidor.";
        } else if (recipientCount) {
          msg = "Nenhum anexo gerado para envio.";
        }
        return res.status(400).json({ error: msg });
      }

      const result =
        emailMode === "outlook"
          ? await sendEmailsOutlook(emailsToSend, emailSubject, emailMessage, outlookFrom)
          : await sendEmailsSmtp(emailsToSend, emailSubject, emailMessage, smtpConfig);
      const reportBuffer = result.results && result.results.length
        ? buildEmailReportBuffer(result.results)
        : null;
      res.json({
        status: "ok",
        sent: result.sent,
        failed: result.failed,
        errors: result.errors,
        reportXlsx: reportBuffer ? reportBuffer.toString("base64") : null,
      });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message || "Erro ao enviar emails" });
    }
  }
);

function startServer(port = PORT) {
  const server = app.listen(port, () => {
    console.log(`Rodando em http://localhost:${port}`);
  });

  server.on("error", (err) => {
    if (err.code === "EADDRINUSE") {
      console.error(`Porta ${port} em uso. Defina PORT para outro valor (ex.: PORT=3001 npm start) ou finalize o processo que ocupa a porta.`);
    } else {
      console.error("Erro ao subir o servidor:", err);
    }
    process.exit(1);
  });

  return server;
}

if (require.main === module) {
  startServer();
}

module.exports = { app, startServer };
