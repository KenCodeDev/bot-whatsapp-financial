require("dotenv").config();
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require("@whiskeysockets/baileys");
const P = require("pino");
const Groq = require("groq-sdk");
const Database = require("better-sqlite3");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");

// ======================================================
// GROQ API KEY (LANGSUNG DIHARDCODE)
// ======================================================
const GROQ_API_KEY = "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";

// ======================================================
// DATABASE
// ======================================================
const db = new Database("finance.db");

db.prepare(
    `CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT,
        type TEXT,
        amount INTEGER,
        category TEXT,
        note TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`
).run();

db.prepare(
    `CREATE TABLE IF NOT EXISTS chat_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT,
        role TEXT,
        content TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`
).run();

db.prepare(
    `CREATE TABLE IF NOT EXISTS budgets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT,
        amount INTEGER,
        period TEXT DEFAULT 'monthly',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(user_id, period)
    )`
).run();

db.prepare(`DELETE FROM chat_history WHERE created_at < datetime('now', '-7 days')`).run();

function rupiah(number) {
    return new Intl.NumberFormat("id-ID", {
        style: "currency",
        currency: "IDR",
        minimumFractionDigits: 0,
    }).format(number);
}

// ======================================================
// MEMORY CHAT
// ======================================================
function saveChatMessage(userId, role, content) {
    try {
        const stmt = db.prepare(`INSERT INTO chat_history (user_id, role, content) VALUES (?, ?, ?)`);
        stmt.run(userId, role, content);
    } catch (err) {
        console.error("Gagal simpan chat history:", err);
    }
}

function getChatHistory(userId, limit = 10) {
    const stmt = db.prepare(`SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?`);
    const rows = stmt.all(userId, limit);
    return rows.reverse();
}

// ======================================================
// AI DENGAN MEMORY
// ======================================================
let groq;

async function aiWithMemory(userId, userMessage, options = {}) {
    const historyLimit = options.historyLimit || 10;
    const history = getChatHistory(userId, historyLimit);

    const messages = [
        {
            role: "system",
            content: `
Kamu adalah Finance Bot By Keniii, asisten keuangan pribadi yang santai dan seru. Tugasmu adalah membantu pengguna mencatat transaksi keuangan mereka dengan cara yang natural, kayak teman chat. Kamu ga terlalu formal, boleh lucu, dan jangan monoton.
Kamu diciptakan oleh Kenichi Ichi (Keniii, berasal dari Indonesia) untuk membantu orang-orang dalam mengelola keuangan mereka dengan cara yang menyenangkan dan mudah dipahami.
Pakai Aku dan Kamu untuk ngobrol, jangan pernah pakai "saya/anda". Ingat konteks percakapan sebelumnya untuk memberikan jawaban yang relevan.

Kepribadian:
- santai
- natural
- seru
- seperti teman chat
- tidak terlalu formal
- boleh lucu
- jangan monoton
- INGAT KONTEKS PERCAKAPAN SEBELUMNYA

Tugas:
- memahami bahasa Indonesia santai
- memahami typo manusia
- memahami slang Indonesia
- memahami percakapan casual
- membantu pencatatan keuangan

Kemungkinan intent:
1. add_transaction
2. check_balance
3. history
4. export_pdf
5. export_excel
6. financial_advice
7. casual_chat
8. delete_transaction
9. set_budget
10. check_budget
11. unknown

Aturan uang:
- jt = juta
- rb = ribu
- k = ribu
- ceban = 10000

ATURAN PENTING:
- jangan bikin nominal 0
- jangan nebak nominal
- jangan asal simpan transaksi
- jika informasi transaksi belum lengkap, tetap gunakan "intent": "add_transaction" dan "missing_field"

Contoh untuk delete_transaction:
User: "hapus transaksi terakhir"
Output:
{
  "intent": "delete_transaction",
  "reply": "✅ Transaksi terakhir berhasil dihapus!"
}

Contoh untuk set_budget:
User: "atur budget bulanan 2 juta"
Output:
{
  "intent": "set_budget",
  "amount": 2000000,
  "reply": "💰 Budget bulanan berhasil disetel ke Rp2.000.000. Awas jangan jajan terus ya!"
}

Contoh untuk check_budget:
User: "budget gue masih sisa berapa?"
Output:
{
  "intent": "check_budget",
  "reply": null
}

Contoh untuk financial_advice:
User: "investasi apa yang paling oke sekarang?"
Output:
{
  "intent": "financial_advice",
  "reply": "📈 Untuk pemula, reksa dana atau emas batangan lebih aman. Saham agak riskan tapi potensi besar. Mau coba yang mana? 😄"
}

User: "tips nabung buat anak kuliah"
Output:
{
  "intent": "financial_advice",
  "reply": "🎓 Coba metode 50/30/20: 50% kebutuhan, 30% keinginan, 20% tabung. Mulai dari 20rb sehari aja, setahun bisa 7jt loh!"
}

User: "yang paling okee apa?" (dalam konteks investasi)
Output:
{
  "intent": "financial_advice",
  "reply": "Kalau kamu pemula, mending mulai dari reksa dana pasar uang. Minimal 100rb udah bisa. Mau saya jelasin lebih detail? 🚀"
}

Contoh transaksi lengkap:
{
  "intent": "add_transaction",
  "type": "expense",
  "amount": 25000,
  "category": "food",
  "note": "ngopi fore",
  "reply": "☕ Ngopi mulu yaa 😭 pengeluaran berhasil dicatat"
}

Contoh tidak lengkap:
{
  "intent": "add_transaction",
  "type": "expense",
  "missing_field": "amount",
  "reply": "Kepake berapa tuh?"
}

Contoh casual:
{
  "intent": "casual_chat",
  "reply": "Haiii jugaaa 😭"
}

Contoh cek saldo:
{
  "intent": "check_balance"
}

Contoh history:
{
  "intent": "history"
}

Contoh export PDF:
{
  "intent": "export_pdf"
}

Contoh export Excel:
{
  "intent": "export_excel"
}

Selalu return JSON VALID.
`,
        },
        ...history.map((row) => ({ role: row.role, content: row.content })),
        { role: "user", content: userMessage },
    ];

    try {
        const completion = await groq.chat.completions.create({
            model: "llama-3.3-70b-versatile",
            temperature: 0.8,
            response_format: { type: "json_object" },
            messages: messages,
        });

        const result = JSON.parse(completion.choices[0].message.content);

        saveChatMessage(userId, "user", userMessage);

        if (
            result.reply &&
            (result.intent === "casual_chat" ||
                result.intent === "financial_advice" ||
                (result.intent === "add_transaction" && result.missing_field) ||
                result.intent === "unknown" ||
                result.intent === "delete_transaction" ||
                result.intent === "set_budget")
        ) {
            saveChatMessage(userId, "assistant", result.reply);
        }
        if (result.intent === "add_transaction" && !result.missing_field && result.reply) {
            saveChatMessage(userId, "assistant", result.reply);
        }

        return result;
    } catch (error) {
        console.error("Error AI:", error);
        return {
            intent: "unknown",
            reply: "😭 Aduh AI-ku error, coba lagi ya!",
        };
    }
}

// ======================================================
// FUNGSI PDF BALANCE
// ======================================================
async function showBalancePDF(userId, sock, jid) {
    const incomeStmt = db.prepare(`SELECT SUM(amount) as total FROM transactions WHERE user_id = ? AND type = 'income'`);
    const expenseStmt = db.prepare(`SELECT SUM(amount) as total FROM transactions WHERE user_id = ? AND type = 'expense'`);
    const income = incomeStmt.get(userId);
    const expense = expenseStmt.get(userId);

    const totalIncome = income.total || 0;
    const totalExpense = expense.total || 0;
    const saldo = totalIncome - totalExpense;

    const filePath = path.join(__dirname, `balance-${userId}.pdf`);
    const doc = new PDFDocument({ margin: 40, size: "A4" });
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);

    const GREEN = "#93C47D";
    const LIGHT_GREEN = "#D9EAD3";

    doc.rect(0, 0, doc.page.width, 80).fill(GREEN);
    doc.fillColor("white").fontSize(24).text("Saldo Report", 0, 30, { align: "center" });
    doc.moveDown(3);

    const yStart = 180;
    doc.fillColor("black").fontSize(14);

    doc.rect(40, yStart, 520, 35).fill(LIGHT_GREEN);
    doc.fillColor("black").fontSize(12).text("Total Pemasukan", 55, yStart + 10);
    doc.fontSize(14).text(rupiah(totalIncome), 400, yStart + 8, { width: 150, align: "right" });

    doc.rect(40, yStart + 40, 520, 35).fill(LIGHT_GREEN);
    doc.fillColor("black").fontSize(12).text("Total Pengeluaran", 55, yStart + 50);
    doc.fontSize(14).text(rupiah(totalExpense), 400, yStart + 48, { width: 150, align: "right" });

    doc.rect(40, yStart + 80, 520, 35).fill(LIGHT_GREEN);
    doc.fillColor("black").fontSize(12).text("Saldo", 55, yStart + 90);
    doc.fontSize(14).text(rupiah(saldo), 400, yStart + 88, { width: 150, align: "right" });

    doc.fontSize(9).fillColor("gray").text("Generated by AI Finance Assistant", 40, 780, { align: "center" });
    doc.end();

    stream.on("finish", async () => {
        await sock.sendMessage(jid, {
            document: fs.readFileSync(filePath),
            mimetype: "application/pdf",
            fileName: `balance-${userId}.pdf`,
        });
        fs.unlinkSync(filePath);
    });
}

// ======================================================
// FUNGSI EXPORT PDF LENGKAP
// ======================================================
async function exportPDF(userId, sock, jid) {
    const stmt = db.prepare(`SELECT * FROM transactions WHERE user_id = ? ORDER BY id DESC`);
    const transactions = stmt.all(userId);

    if (transactions.length === 0) {
        await sock.sendMessage(jid, { text: "😭 Belum ada transaksi" });
        return;
    }

    const filePath = path.join(__dirname, `report-${userId}.pdf`);
    const doc = new PDFDocument({ margin: 40, size: "A4" });
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);

    const GREEN = "#93C47D";
    const LIGHT_GREEN = "#D9EAD3";
    const BORDER = "#CCCCCC";

    doc.rect(0, 0, doc.page.width, 80).fill(GREEN);
    doc.fillColor("white").fontSize(24).text("Finance Report", 0, 30, { align: "center" });
    doc.moveDown(3);

    const tableTop = 120;
    const col1 = 40, col2 = 120, col3 = 240, col4 = 350, col5 = 470;

    doc.fillColor(LIGHT_GREEN).rect(40, tableTop, 520, 25).fill();
    doc.fillColor("black").fontSize(11);
    doc.text("Type", col1 + 5, tableTop + 7);
    doc.text("Amount", col2 + 5, tableTop + 7);
    doc.text("Category", col3 + 5, tableTop + 7);
    doc.text("Note", col4 + 5, tableTop + 7);
    doc.text("Date", col5 + 5, tableTop + 7);

    let y = tableTop + 25;
    let totalIncome = 0, totalExpense = 0;

    transactions.forEach((item, index) => {
        if (item.type === "income") totalIncome += item.amount;
        else totalExpense += item.amount;

        if (index % 2 === 0) {
            doc.fillColor("#F7F7F7").rect(40, y, 520, 25).fill();
        }
        doc.fillColor("black").fontSize(10);
        doc.text(item.type || "-", col1 + 5, y + 7);
        doc.text(rupiah(item.amount || 0), col2 + 5, y + 7);
        doc.text(item.category || "-", col3 + 5, y + 7);
        doc.text(item.note || "-", col4 + 5, y + 7, { width: 100 });
        doc.text(item.created_at || "-", col5 + 5, y + 7);

        doc.strokeColor(BORDER).rect(40, y, 520, 25).stroke();
        y += 25;

        if (y > 720) {
            doc.addPage();
            y = 50;
        }
    });

    const saldo = totalIncome - totalExpense;
    y += 30;
    doc.fillColor(LIGHT_GREEN).rect(40, y, 300, 100).fill();
    doc.strokeColor(BORDER).rect(40, y, 300, 100).stroke();
    doc.fillColor("black").fontSize(12);
    doc.text(`Total Pemasukan : ${rupiah(totalIncome)}`, 55, y + 20);
    doc.text(`Total Pengeluaran : ${rupiah(totalExpense)}`, 55, y + 45);
    doc.text(`Saldo : ${rupiah(saldo)}`, 55, y + 70);

    doc.fontSize(9).fillColor("gray").text("Generated by AI Finance Assistant", 40, 780, { align: "center" });
    doc.end();

    stream.on("finish", async () => {
        await sock.sendMessage(jid, {
            document: fs.readFileSync(filePath),
            mimetype: "application/pdf",
            fileName: `report-${userId}.pdf`,
        });
        fs.unlinkSync(filePath);
    });
}

// ======================================================
// FUNGSI EXPORT EXCEL
// ======================================================
async function exportExcel(userId, sock, jid) {
    const stmt = db.prepare(`SELECT * FROM transactions WHERE user_id = ? ORDER BY id DESC`);
    const transactions = stmt.all(userId);

    if (transactions.length === 0) {
        await sock.sendMessage(jid, { text: "😭 Belum ada transaksi" });
        return;
    }

    let totalIncome = 0, totalExpense = 0;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Finance Report");

    sheet.mergeCells("A1:E1");
    const titleCell = sheet.getCell("A1");
    titleCell.value = "Finance Report";
    titleCell.font = { size: 18, bold: true, color: { argb: "FFFFFF" } };
    titleCell.alignment = { vertical: "middle", horizontal: "center" };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "6AA84F" } };
    sheet.getRow(1).height = 30;

    sheet.columns = [
        { header: "Type", key: "type", width: 15 },
        { header: "Amount", key: "amount", width: 20 },
        { header: "Category", key: "category", width: 20 },
        { header: "Note", key: "note", width: 35 },
        { header: "Date", key: "date", width: 25 },
    ];

    const headerRow = sheet.getRow(2);
    headerRow.values = ["Type", "Amount", "Category", "Note", "Date"];
    headerRow.font = { bold: true, color: { argb: "FFFFFF" } };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "93C47D" } };
    headerRow.alignment = { horizontal: "center", vertical: "middle" };

    let rowIndex = 3;
    for (const item of transactions) {
        if (item.type === "income") totalIncome += item.amount;
        else totalExpense += item.amount;

        sheet.addRow({
            type: item.type,
            amount: item.amount,
            category: item.category || "-",
            note: item.note || "-",
            date: item.created_at,
        });

        const row = sheet.getRow(rowIndex);
        row.eachCell((cell) => {
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };
            cell.alignment = { vertical: "middle", horizontal: "center" };
        });
        rowIndex++;
    }

    const saldo = totalIncome - totalExpense;
    rowIndex += 2;
    sheet.getCell(`A${rowIndex}`).value = "Total Pemasukan";
    sheet.getCell(`B${rowIndex}`).value = rupiah(totalIncome);
    rowIndex++;
    sheet.getCell(`A${rowIndex}`).value = "Total Pengeluaran";
    sheet.getCell(`B${rowIndex}`).value = rupiah(totalExpense);
    rowIndex++;
    sheet.getCell(`A${rowIndex}`).value = "Saldo";
    sheet.getCell(`B${rowIndex}`).value = rupiah(saldo);

    for (let i = rowIndex - 2; i <= rowIndex; i++) {
        sheet.getCell(`A${i}`).font = { bold: true };
        sheet.getCell(`A${i}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "D9EAD3" } };
        sheet.getCell(`A${i}`).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
        };
        sheet.getCell(`B${i}`).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
        };
    }

    const filePath = path.join(__dirname, `report-${userId}.xlsx`);
    await workbook.xlsx.writeFile(filePath);
    await sock.sendMessage(jid, {
        document: fs.readFileSync(filePath),
        mimetype: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileName: `report-${userId}.xlsx`,
    });
    fs.unlinkSync(filePath);
}

// ======================================================
// FUNGSI DELETE TRANSAKSI TERAKHIR
// ======================================================
async function deleteLastTransaction(userId, sock, jid) {
    const stmt = db.prepare(`SELECT id FROM transactions WHERE user_id = ? ORDER BY id DESC LIMIT 1`);
    const last = stmt.get(userId);
    if (!last) {
        await sock.sendMessage(jid, { text: "😭 Ga ada transaksi yang bisa dihapus" });
        return;
    }
    db.prepare("DELETE FROM transactions WHERE id = ?").run(last.id);
    await sock.sendMessage(jid, { text: "✅ Transaksi terakhir berhasil dihapus!" });
}

// ======================================================
// FUNGSI BUDGET
// ======================================================
async function setBudget(userId, amount, sock, jid) {
    const stmt = db.prepare(`
        INSERT INTO budgets (user_id, amount, period)
        VALUES (?, ?, 'monthly')
        ON CONFLICT(user_id, period) DO UPDATE SET amount = excluded.amount
    `);
    stmt.run(userId, amount);
    await sock.sendMessage(jid, { text: `💰 Budget bulanan berhasil disetel ke ${rupiah(amount)}. Semoga cukup ya!` });
}

async function checkBudget(userId, sock, jid) {
    const budgetStmt = db.prepare(`SELECT amount FROM budgets WHERE user_id = ? AND period = 'monthly' LIMIT 1`);
    const budget = budgetStmt.get(userId);
    if (!budget) {
        await sock.sendMessage(jid, { text: "🤔 Kamu belum setting budget bulanan. Ketik 'atur budget <jumlah>' ya!" });
        return;
    }
    const expenseStmt = db.prepare(`
        SELECT COALESCE(SUM(amount),0) as total
        FROM transactions
        WHERE user_id = ? AND type = 'expense'
        AND strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now')
    `);
    const expense = expenseStmt.get(userId);
    const totalExpense = expense.total;
    const sisa = budget.amount - totalExpense;

    await sock.sendMessage(jid, {
        text: `📊 Budget Bulan Ini\n\n🎯 Budget : ${rupiah(budget.amount)}\n💸 Terpakai : ${rupiah(totalExpense)}\n💰 Sisa : ${rupiah(sisa > 0 ? sisa : 0)}\n${sisa < 0 ? "⚠️ Udah over budget nih!" : "✅ Masih aman, jaga terus ya!"}`,
    });
}

// ======================================================
// PENDING TRANSACTION
// ======================================================
const pendingTransaction = {};

// ======================================================
// MAIN BOT
// ======================================================
async function startBot() {
    // Inisialisasi Groq
    groq = new Groq({ apiKey: GROQ_API_KEY });

    const { state, saveCreds } = await useMultiFileAuthState("auth_info_baileys");
    const sock = makeWASocket({
        auth: state,
        printQRInTerminal: false,  // QR code tidak ditampilkan
        logger: P({ level: "silent" }),
        browser: ["Ubuntu", "Chrome", "20.0.04"],
        markOnlineOnConnect: true,
    });

    sock.ev.on("creds.update", saveCreds);

    sock.ev.on("connection.update", async (update) => {
        const { connection, lastDisconnect, qr } = update;
        if (connection === "close") {
            const shouldReconnect = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
            if (shouldReconnect) startBot();
            else console.log("Logged out, restart bot.");
        } else if (connection === "open") {
            console.log("✅ Connected to WhatsApp");
        }
    });

    // Pairing Code dengan format khusus
    if (!state.creds.registered) {
        const number = process.env.BOT_NUMBER;
        if (number) {
            setTimeout(async () => {
                try {
                    const code = await sock.requestPairingCode(number);
                    const first4 = code.slice(0, 4).split('').join(' ');
                    const rest = code.slice(4).split('').join(' ');
                    const formattedCode = `${first4} - ${rest}`;
                    
                    console.log("\x1b[1mPairing Code\x1b[0m");
                    console.log(`🔢 ${formattedCode}`);
                    console.log("\n\x1b[1mCara connect:\x1b[0m");
                    console.log("1. Buka Aplikasi WhatsApp");
                    console.log("2. Tekan titik 3 di pojok kanan atas");
                    console.log("3. Pilih perangkat tertaut");
                    console.log("4. Klik tombol Tautkan Perangkat");
                    console.log("5. Dibagian bawah, klik Tautkan dengan nomor telepon saja");
                    console.log("6. Masukkan kode pairing diatas");
                } catch (err) {
                    console.error("Gagal dapat pairing code:", err);
                }
            }, 3000);
        } else {
            console.warn("BOT_NUMBER tidak ada di .env, tidak bisa pairing code.");
        }
    }

    sock.ev.on("messages.upsert", async ({ messages }) => {
        const msg = messages[0];
        if (msg.key.remoteJid.includes("@g.us")) return; // Skip group messages
        if (!msg.message || msg.key.fromMe) return;
        const jid = msg.key.remoteJid;
        let text = "";
        if (msg.message.conversation) text = msg.message.conversation;
        else if (msg.message.extendedTextMessage) text = msg.message.extendedTextMessage.text;
        if (!text) return;
        if (text.startsWith("/")) return;

        const userId = jid;

        // Pending transaction
        if (pendingTransaction[userId]) {
            const pending = pendingTransaction[userId];
            const combinedMessage = pending.original_message + " " + text;
            delete pendingTransaction[userId];

            const result = await aiWithMemory(userId, combinedMessage, { historyLimit: 10 });

            if (result.intent === "add_transaction" && result.missing_field) {
                pendingTransaction[userId] = {
                    missing_field: result.missing_field,
                    original_message: combinedMessage,
                };
                await sock.sendMessage(jid, { text: result.reply });
                return;
            }

            if (result.intent === "add_transaction" && !result.missing_field) {
                const insertStmt = db.prepare(
                    `INSERT INTO transactions (user_id, type, amount, category, note) VALUES (?, ?, ?, ?, ?)`
                );
                insertStmt.run(userId, result.type, result.amount, result.category, result.note);
                await sock.sendMessage(jid, {
                    text: `${result.reply}\n\n💵 ${rupiah(result.amount)}\n🏷 ${result.category}`,
                });
                return;
            }
        }

        // Analisis AI
        const result = await aiWithMemory(userId, text);
        console.log("Intent:", result.intent, "| User:", userId);

        if (result.intent === "add_transaction" && result.missing_field) {
            pendingTransaction[userId] = {
                missing_field: result.missing_field,
                original_message: text,
            };
            await sock.sendMessage(jid, { text: result.reply });
            return;
        }

        if (result.intent === "add_transaction" && !result.missing_field) {
            const insertStmt = db.prepare(
                `INSERT INTO transactions (user_id, type, amount, category, note) VALUES (?, ?, ?, ?, ?)`
            );
            insertStmt.run(userId, result.type, result.amount, result.category, result.note);
            await sock.sendMessage(jid, {
                text: `${result.reply}\n\n💵 ${rupiah(result.amount)}\n🏷 ${result.category}`,
            });
            return;
        }

        if (result.intent === "check_balance") {
            if (result.reply) await sock.sendMessage(jid, { text: result.reply });
            await showBalancePDF(userId, sock, jid);
            return;
        }

        if (result.intent === "history") {
            if (result.reply) await sock.sendMessage(jid, { text: result.reply });
            await exportPDF(userId, sock, jid);
            return;
        }

        if (result.intent === "export_pdf") {
            if (result.reply) await sock.sendMessage(jid, { text: result.reply });
            await exportPDF(userId, sock, jid);
            return;
        }

        if (result.intent === "export_excel") {
            if (result.reply) await sock.sendMessage(jid, { text: result.reply });
            await exportExcel(userId, sock, jid);
            return;
        }

        if (result.intent === "delete_transaction") {
            if (result.reply) await sock.sendMessage(jid, { text: result.reply });
            await deleteLastTransaction(userId, sock, jid);
            return;
        }

        if (result.intent === "set_budget") {
            if (result.amount > 0) {
                await setBudget(userId, result.amount, sock, jid);
            } else {
                await sock.sendMessage(jid, { text: "😕 Nominal budget-nya berapa ya? Coba sebutin angka." });
            }
            return;
        }

        if (result.intent === "check_budget") {
            if (result.reply) await sock.sendMessage(jid, { text: result.reply });
            await checkBudget(userId, sock, jid);
            return;
        }

        if (result.intent === "financial_advice" || result.intent === "casual_chat") {
            await sock.sendMessage(jid, { text: result.reply });
            return;
        }

        await sock.sendMessage(jid, { text: result.reply || "😭 Aku belum ngerti" });
    });
}

startBot().catch(console.error);
console.log("🤖 Bot sedang berjalan...");
