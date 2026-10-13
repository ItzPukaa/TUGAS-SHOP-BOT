import {
    Client,
    GatewayIntentBits,
    REST,
    Routes,
    SlashCommandBuilder,
    MessageFlags,
    ActionRowBuilder,
    EmbedBuilder,
    StringSelectMenuBuilder,
    PermissionFlagsBits,
    ButtonBuilder,
    ButtonStyle,
    ChannelType,
    ModalBuilder,
    TextInputBuilder,
    TextInputStyle,
    AttachmentBuilder
} from "discord.js";
import { google } from "googleapis";
import fs from "fs";
import { fileURLToPath } from "url";

// Em local, o Node pode carregar .env; no Render, configura as variáveis no painel.
try { process.loadEnvFile(); } catch {}

// CONFIGURATION
const TOKEN = process.env.DISCORD_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const STAFF_ROLE_ID = process.env.STAFF_ROLE_ID;
const GUILD_ID = process.env.GUILD_ID;

if (!TOKEN) {
    console.error("Missing DISCORD_TOKEN in the .env file");
    process.exit(1);
}

// GOOGLE AUTH
// No Render, define GOOGLE_SERVICE_ACCOUNT_JSON com o conteúdo completo do credentials.json.
// Em testes locais, continua a aceitar credentials.json na pasta do projeto.
const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
let googleAuthOptions;
if (serviceAccountJson) {
    try {
        googleAuthOptions = { credentials: JSON.parse(serviceAccountJson) };
    } catch (error) {
        console.error("GOOGLE_SERVICE_ACCOUNT_JSON does not contain valid JSON.", error.message);
        process.exit(1);
    }
} else {
    googleAuthOptions = { keyFile: "credentials.json" };
}
const auth = new google.auth.GoogleAuth({
    ...googleAuthOptions,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
});

let sheetsClient;

// Dashboard persistente: uma única mensagem editada, sem spam.
const DASHBOARD_CHANNEL_NAME = "dashboard";
const DASHBOARD_MARKER = "TUGAS_SHOP_STOCK_DASHBOARD";
let dashboardMessageId = null;
let dashboardSpecies = null;
let dashboardColor = null;
let dashboardGender = "m";
let dashboardRefreshRunning = false;

// ===========================================================
// [PARTE 1] CONFIGURAÇÃO DO SHOP CATALOG / HOW TO BUY
// ===========================================================
// Nomes (parte do nome) dos canais onde o bot publica as mensagens informativas
const CATALOG_CHANNEL_NAME = "shop-catalog";
const HOWTOBUY_CHANNEL_NAME = "how-to-buy";
const TICKET_CHANNEL_NAME = "open-a-ticket";
// Textos usados no rodapé para o bot reconhecer a sua própria mensagem (e editá-la em vez de duplicar)
const CATALOG_MARKER = "TUGAS SHOP • Shop Catalog";
const HOWTOBUY_MARKER = "TUGAS SHOP • How to Buy";
const SHOP_COLOR = 0x2ecc71;
let catalogSignature = null;        // lista de espécies da última publicação (evita edições inúteis)
let catalogRefreshRunning = false;

// ===========================================================
// [PARTE 3] CONFIGURAÇÃO DOS FRIDGES / ORDERS
// ===========================================================
const TRADINGPOST_SHEET = "TradingPost";   // registo de cada order (cor, número, estado, staff)
const FRIDGES_SHEET = "Fridges";           // lista de cores/números de fridges (gerida por /addfridges, /editfridges, /delfridges)
const ORDERS_LOG_CHANNEL_NAME = "orders-logs"; // canal onde ficam TODAS as orders
const FRIDGE_DEFAULT_COUNT = 10;           // por defeito cada cor nova tem #01 … #10
const FRIDGE_MAX_NUMBER = 99;              // maior número permitido
const FRIDGE_MAX_PER_COLOR = 25;           // limite de opções de um menu do Discord
const FRIDGE_MAX_COLORS = 25;              // limite de opções de um menu do Discord

// Nomes dos campos do embed da order (a staff lê estes campos para saber o estado, mesmo após reiniciar o bot)
const F_CUSTOMER = "Customer";
const F_COLOR = "🧊 Fridge Color";
const F_NUMBER = "🧊 Fridge Number";
const F_PIN = "🔑 PIN";
const F_STATUS = "📌 Status";
const F_TICKET = "Customer ticket";
const F_CLAIM = "👤 Claimed by";           // INTERNO: só aparece no canal de logs (staff)
const PIN_PLACEHOLDER = "⏳ Not set yet";
const UNCLAIMED = "— Unclaimed";

// Imagem enviada na DM "order ready". Coloca o ficheiro em ./assets/tugas-banner.png (junto ao index.js).
// Se o ficheiro não existir, a DM é enviada na mesma, apenas sem imagem.
const DM_BANNER_FILE = fileURLToPath(new URL("./assets/tugas-banner.png", import.meta.url));
const DM_BANNER_NAME = "tugas-banner.png";
const PIN_PLACEHOLDER = "⏳ Not set yet";

async function initSheets() {
    const authClient = await auth.getClient();
    sheetsClient = google.sheets({ version: "v4", auth: authClient });
    await ensureSystemSheets();
}

// Layout antigo: Nome | Tek Pair/Bundle | Tek Extra | PayPal Pair/Bundle | PayPal Extra
// Layout novo:   Nome | Tek Pair | Tek Bundle | Tek Extra | PayPal Pair | PayPal Bundle | PayPal Extra
// Migra automaticamente (o preço antigo "Pair/Bundle" é copiado para Pair e Bundle; depois ajusta com /editdino).
async function migratePrecosLayout() {
    const res = await sheetsClient.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "Precos!A:G" });
    const rows = res.data.values || [];
    const header = rows[0] || [];
    const isOld = header.length > 0 && header.length <= 5 && String(header[1] || "").includes("Pair/Bundle");
    if (!isOld) return;

    const migrated = rows.slice(1)
        .filter(r => (r[0] || "").trim() !== "")
        .map(r => [r[0], r[1] ?? "", r[1] ?? "", r[2] ?? "", r[3] ?? "", r[3] ?? "", r[4] ?? ""]);

    await sheetsClient.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: "Precos!A2:G" });
    if (migrated.length) {
        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: SHEET_ID, range: `Precos!A2:G${migrated.length + 1}`, valueInputOption: "USER_ENTERED",
            resource: { values: migrated }
        });
    }
    console.log(`Precos sheet migrated to the new layout (${migrated.length} species).`);
}

async function ensureSystemSheets() {
    const res = await sheetsClient.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const titles = new Set((res.data.sheets || []).map(sh => sh.properties.title));
    const requests = [];
    for (const title of ["Precos", "Stock", TRADINGPOST_SHEET, FRIDGES_SHEET]) {
        if (!titles.has(title)) requests.push({ addSheet: { properties: { title } } });
    }
    if (requests.length) {
        await sheetsClient.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, resource: { requests } });
    }
    await migratePrecosLayout();
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: "Precos!A1:G1", valueInputOption: "USER_ENTERED",
        resource: { values: [["Name", "Tek Pair", "Tek Bundle", "Tek Extra", "PayPal Pair", "PayPal Bundle", "PayPal Extra"]] }
    });
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: "Stock!A1:C1", valueInputOption: "USER_ENTERED",
        resource: { values: [["Name", "Bundles", "Status"]] }
    });
    // [PARTE 3] Cabeçalho da folha TradingPost (o PIN NÃO é guardado aqui de propósito)
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `${TRADINGPOST_SHEET}!A1:H1`, valueInputOption: "USER_ENTERED",
        resource: { values: [["Date", "Order ID", "Customer ID", "Customer", "Fridge Color", "Fridge Number", "Status", "Claimed By"]] }
    });
    // [PARTE 3] Cabeçalho da folha Fridges (uma linha por fridge: cor + número)
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `${FRIDGES_SHEET}!A1:D1`, valueInputOption: "USER_ENTERED",
        resource: { values: [["Color", "Number", "Occupied By Order", "Customer"]] }
    });
}

// Cliente Discord (para slash commands só precisa de Guilds)
const client = new Client({
    intents: [GatewayIntentBits.Guilds]
});

/* -----------------------------------------------------------
   HELPERS
------------------------------------------------------------*/

// Nome de folha seguro para ranges (aspas simples, com escape)
function q(sheetName) {
    return `'${sheetName.replace(/'/g, "''")}'`;
}

const IGNORED_SHEETS = ["Stock", "Precos", TRADINGPOST_SHEET, FRIDGES_SHEET];

async function getSpeciesSheets() {
    const res = await sheetsClient.spreadsheets.get({ spreadsheetId: SHEET_ID });
    return res.data.sheets
        .map(s => s.properties.title)
        .filter(s => !IGNORED_SHEETS.includes(s));
}

async function getBundlesForSpecies(sheetName) {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${q(sheetName)}!A2:C`
    });

    const rows = (res.data.values || []).filter(r => (r[0] || "").trim() !== "");

    // Sem cores = sem bundles
    if (rows.length === 0) return 0;

    // Um bundle precisa de 1 macho + 1 fêmea de CADA cor,
    // por isso o limite é a cor com menos stock.
    let bundles = Infinity;

    for (const row of rows) {
        const male = parseInt(row[1] || "0", 10);
        const female = parseInt(row[2] || "0", 10);
        const pairs = Math.min(isNaN(male) ? 0 : male, isNaN(female) ? 0 : female);
        bundles = Math.min(bundles, pairs);
    }

    return bundles;
}

// 🔴 sem stock | 🟠 último bundle | 🟡 2 bundles | 🟢 3 ou mais
function stockBall(bundles) {
    if (bundles <= 0) return "🔴";
    if (bundles === 1) return "🟠";
    if (bundles === 2) return "🟡";
    return "🟢";
}

async function getStockData() {
    const species = await getSpeciesSheets();
    const data = [];
    for (const s of species) {
        data.push({ name: s, bundles: await getBundlesForSpecies(s) });
    }
    return data;
}

async function updateStockSheet() {
    const data = await getStockData();
    const values = [["Name", "Bundles", "Status"]];

    for (const d of data) {
        values.push([d.name, d.bundles, stockBall(d.bundles)]);
    }

    // Limpa primeiro para não ficarem linhas antigas
    await sheetsClient.spreadsheets.values.clear({
        spreadsheetId: SHEET_ID,
        range: "Stock!A:C"
    });

    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: "Stock!A1:C",
        valueInputOption: "USER_ENTERED",
        resource: { values }
    });

    if (client.isReady()) {
        await refreshDashboard();
        // [PARTE 1] Atualiza o catálogo público apenas se a lista de espécies mudou
        await refreshShopInfo(data.map(d => d.name));
    }
}

function isStaff(interaction) {
    return interaction.member?.roles?.cache?.has(STAFF_ROLE_ID);
}

// Devolve o sheetId numérico de uma folha (necessário para apagar)
async function getSheetId(sheetName) {
    const res = await sheetsClient.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheet = res.data.sheets.find(s => s.properties.title === sheetName);
    return sheet ? sheet.properties.sheetId : null;
}

// Apaga uma linha inteira (rowIndex é 0-based; 0 = cabeçalho)
async function deleteRow(sheetName, rowIndex) {
    const sheetId = await getSheetId(sheetName);
    if (sheetId === null) throw new Error(`Sheet "${sheetName}" not found.`);

    await sheetsClient.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        resource: {
            requests: [{
                deleteDimension: {
                    range: {
                        sheetId,
                        dimension: "ROWS",
                        startIndex: rowIndex,
                        endIndex: rowIndex + 1
                    }
                }
            }]
        }
    });
}

// Lista as cores (coluna A) de uma espécie
async function getColorsForSpecies(sheetName) {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${q(sheetName)}!A2:A`
    });
    return (res.data.values || []).map(r => r[0]).filter(Boolean);
}


/* -----------------------------------------------------------
   DASHBOARD DE STOCK (uma mensagem, dropdowns + botões)
------------------------------------------------------------*/
function dashboardMarker(embed) {
    return (embed?.footer?.text || "").includes(DASHBOARD_MARKER);
}

async function dashboardPayload() {
    const species = await getSpeciesSheets();
    if (!species.length) {
        const embed = new EmbedBuilder()
            .setTitle("📦 TUGAS SHOP — Stock")
            .setDescription("No dinos registered yet. Use `/adddino` to create the first species.")
            .setFooter({ text: DASHBOARD_MARKER });
        return { embeds: [embed], components: [] };
    }

    if (!dashboardSpecies || !species.includes(dashboardSpecies)) dashboardSpecies = species[0];
    const colors = await getColorStock(dashboardSpecies);
    if (!dashboardColor || !colors.some(c => c.color === dashboardColor)) dashboardColor = colors[0]?.color || null;
    if (!colors.some(c => c.color === dashboardColor)) dashboardGender = "m";

    const prices = await getPrices(dashboardSpecies);
    const lines = colors.map(c => {
        const selected = c.color === dashboardColor ? "➡️ " : "";
        return `${selected}**${c.color}** — ♂ ${c.male} · ♀ ${c.female}`;
    });
    let description = lines.length ? lines.join("\n") : "No colors yet. Use `/addcolor` to add a color with starting stock at zero.";
    if (description.length > 3800) description = description.slice(0, 3750) + "\n… (list shortened; use the menu to select a color)";

    const selectedColor = colors.find(c => c.color === dashboardColor);
    const selectedQty = selectedColor ? (dashboardGender === "m" ? selectedColor.male : selectedColor.female) : 0;
    const embed = new EmbedBuilder()
        .setTitle(`📦 TUGAS SHOP — ${dashboardSpecies}`)
        .setDescription(description)
        .addFields(
            { name: "Pair — Tek", value: prices?.tek?.pair == null ? "Not set" : `${prices.tek.pair} Tek`, inline: true },
            { name: "Bundle — Tek", value: prices?.tek?.bundle == null ? "Not set" : `${prices.tek.bundle} Tek`, inline: true },
            { name: "Extra — Tek", value: prices?.tek?.dino == null ? "Not set" : `${prices.tek.dino} Tek`, inline: true },
            { name: "Pair — PayPal", value: prices?.paypal?.pair == null ? "Not set" : `${prices.paypal.pair}€`, inline: true },
            { name: "Bundle — PayPal", value: prices?.paypal?.bundle == null ? "Not set" : `${prices.paypal.bundle}€`, inline: true },
            { name: "Extra — PayPal", value: prices?.paypal?.dino == null ? "Not set" : `${prices.paypal.dino}€`, inline: true },
            { name: "Adjust stock", value: selectedColor ? `Color: **${selectedColor.color}**\nGender: **${dashboardGender === "m" ? "♂ Male" : "♀ Female"}**\nCurrent quantity: **${selectedQty}**` : "Select a color to adjust the stock." }
        )
        .setFooter({ text: `${DASHBOARD_MARKER} • Auto-refreshes every 30 minutes and after stock changes` });

    const rows = [];
    rows.push(new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
            .setCustomId("dash_species")
            .setPlaceholder("Select a dino")
            .addOptions(species.slice(0, 25).map(name => ({ label: name.slice(0, 100), value: name, default: name === dashboardSpecies })))
    ));

    if (colors.length) {
        rows.push(new ActionRowBuilder().addComponents(
            new StringSelectMenuBuilder()
                .setCustomId("dash_color")
                .setPlaceholder("Select a color")
                .addOptions(colors.slice(0, 25).map(c => ({
                    label: c.color.slice(0, 100),
                    value: c.color,
                    description: `♂ ${c.male} · ♀ ${c.female}`.slice(0, 100),
                    default: c.color === dashboardColor
                })))
        ));
    }

    rows.push(new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("dash_gender_m").setLabel("♂ Male").setStyle(dashboardGender === "m" ? ButtonStyle.Success : ButtonStyle.Secondary).setDisabled(!selectedColor),
        new ButtonBuilder().setCustomId("dash_gender_f").setLabel("♀ Female").setStyle(dashboardGender === "f" ? ButtonStyle.Success : ButtonStyle.Secondary).setDisabled(!selectedColor)
    ));
    rows.push(new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("dash_add_1").setLabel("+1").setStyle(ButtonStyle.Success).setDisabled(!selectedColor),
        new ButtonBuilder().setCustomId("dash_add_5").setLabel("+5").setStyle(ButtonStyle.Success).setDisabled(!selectedColor),
        new ButtonBuilder().setCustomId("dash_add_10").setLabel("+10").setStyle(ButtonStyle.Success).setDisabled(!selectedColor)
    ));
    rows.push(new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("dash_sub_10").setLabel("-10").setStyle(ButtonStyle.Danger).setDisabled(!selectedColor || selectedQty < 10),
        new ButtonBuilder().setCustomId("dash_sub_5").setLabel("-5").setStyle(ButtonStyle.Danger).setDisabled(!selectedColor || selectedQty < 5),
        new ButtonBuilder().setCustomId("dash_sub_1").setLabel("-1").setStyle(ButtonStyle.Danger).setDisabled(!selectedColor || selectedQty < 1)
    ));

    return { embeds: [embed], components: rows };
}

async function refreshDashboard() {
    if (!client.isReady() || !sheetsClient || dashboardRefreshRunning) return;
    dashboardRefreshRunning = true;
    try {
        const guild = await client.guilds.fetch(GUILD_ID);
        await guild.channels.fetch();
        const channel = guild.channels.cache.find(c =>
            c.type === ChannelType.GuildText && c.name.toLowerCase().includes(DASHBOARD_CHANNEL_NAME));
        if (!channel) {
            console.warn("Dashboard channel not found. Create a channel with a name like ⛔│「dashboard」.");
            return;
        }

        const messages = await channel.messages.fetch({ limit: 100 });
        let message = dashboardMessageId ? messages.get(dashboardMessageId) : null;
        if (!message) message = messages.find(m => m.author.id === client.user.id && m.embeds.some(dashboardMarker));
        const payload = await dashboardPayload();
        if (message) {
            dashboardMessageId = message.id;
            await message.edit(payload);
        } else {
            message = await channel.send(payload);
            dashboardMessageId = message.id;
        }
    } catch (err) {
        console.error("Error refreshing the stock dashboard:", err);
    } finally {
        dashboardRefreshRunning = false;
    }
}

async function handleDashboardComponent(i) {
    if (!i.customId?.startsWith("dash_")) return false;
    if (!isStaff(i)) {
        await i.reply({ content: "Only staff can change or browse the stock dashboard.", flags: MessageFlags.Ephemeral });
        return true;
    }

    try {
        if (i.isStringSelectMenu()) {
            if (i.customId === "dash_species") {
                dashboardSpecies = i.values[0];
                const colors = await getColorStock(dashboardSpecies);
                dashboardColor = colors[0]?.color || null;
                dashboardGender = "m";
            } else if (i.customId === "dash_color") {
                dashboardColor = i.values[0];
            }
            await i.update(await dashboardPayload());
            return true;
        }

        if (i.customId === "dash_gender_m" || i.customId === "dash_gender_f") {
            dashboardGender = i.customId.endsWith("_m") ? "m" : "f";
            await i.update(await dashboardPayload());
            return true;
        }

        const match = i.customId.match(/^dash_(add|sub)_(1|5|10)$/);
        if (!match) {
            await i.reply({ content: "Unknown dashboard action.", flags: MessageFlags.Ephemeral });
            return true;
        }
        if (!dashboardSpecies || !dashboardColor) {
            await i.reply({ content: "Select a dino and a color first.", flags: MessageFlags.Ephemeral });
            return true;
        }

        await i.deferUpdate();
        const delta = Number(match[2]) * (match[1] === "add" ? 1 : -1);
        const res = await sheetsClient.spreadsheets.values.get({
            spreadsheetId: SHEET_ID,
            range: `${q(dashboardSpecies)}!A:C`
        });
        const rows = res.data.values || [];
        const rowIndex = rows.findIndex((r, idx) => idx > 0 && (r[0] || "").trim() === dashboardColor);
        if (rowIndex < 1) throw new Error(`Color ${dashboardColor} not found in ${dashboardSpecies}.`);

        const row = rows[rowIndex];
        const male = parseInt(row[1] || "0", 10) || 0;
        const female = parseInt(row[2] || "0", 10) || 0;
        const current = dashboardGender === "m" ? male : female;
        const next = current + delta;
        if (next < 0) throw new Error(`Not enough stock: only ${current} units available to remove.`);

        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: SHEET_ID,
            range: `${q(dashboardSpecies)}!A${rowIndex + 1}:C${rowIndex + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: { values: [[dashboardColor, dashboardGender === "m" ? next : male, dashboardGender === "f" ? next : female]] }
        });
        await updateStockSheet();
        await refreshDashboard();
    } catch (err) {
        console.error("Error in dashboard stock control:", err);
        if (i.deferred || i.replied) await i.followUp({ content: `🔴 ${err.message}`, flags: MessageFlags.Ephemeral }).catch(() => {});
        else await i.reply({ content: `🔴 ${err.message}`, flags: MessageFlags.Ephemeral }).catch(() => {});
    }
    return true;
}

/* -----------------------------------------------------------
   [PARTE 3] FOLHA TradingPost (registo de cada order)
------------------------------------------------------------*/
// Regista uma nova order (RAW para os IDs do Discord não perderem dígitos).
// A cor/número do fridge ficam vazios até a staff os escolher.
async function saveTradingPostOrder({ messageId, userId, username }) {
    const date = new Date().toLocaleString("en-GB", { timeZone: "Europe/Lisbon" });
    await sheetsClient.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: `${q(TRADINGPOST_SHEET)}!A:H`,
        valueInputOption: "RAW",
        resource: { values: [[date, messageId, userId, username, "", "", "Pending", ""]] }
    });
}

// Atualiza fridge/estado/staff de uma order (procura a linha pelo ID da mensagem da order)
async function updateTradingPostRow(messageId, { color, number, status, staff } = {}) {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${q(TRADINGPOST_SHEET)}!A:H`
    });
    const rows = res.data.values || [];
    const idx = rows.findIndex((r, n) => n > 0 && String(r[1] || "") === messageId);
    if (idx === -1) {
        console.warn(`TradingPost: order ${messageId} not found in the sheet.`);
        return;
    }
    const row = rows[idx];
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${q(TRADINGPOST_SHEET)}!E${idx + 1}:H${idx + 1}`,
        valueInputOption: "RAW",
        resource: { values: [[
            color ?? row[4] ?? "",
            number ?? row[5] ?? "",
            status ?? row[6] ?? "",
            staff ?? row[7] ?? ""
        ]] }
    });
}

/* -----------------------------------------------------------
   [PARTE 3] FOLHA Fridges (cores e números definidos pela staff)
------------------------------------------------------------*/
function padNum(n) {
    return String(n).padStart(2, "0");
}

// "1-10", "11", "1,2,5" → [1..10] / [11] / [1,2,5]; devolve null se for inválido
function parseNumbers(input) {
    const set = new Set();
    for (const token of input.split(/[,\s]+/).filter(Boolean)) {
        const m = token.match(/^#?(\d{1,2})(?:-#?(\d{1,2}))?$/);
        if (!m) return null;
        const a = parseInt(m[1], 10);
        const b = m[2] ? parseInt(m[2], 10) : a;
        if (a < 1 || b > FRIDGE_MAX_NUMBER || a > b) return null;
        for (let n = a; n <= b; n++) set.add(n);
    }
    return set.size ? [...set].sort((x, y) => x - y) : null;
}

// [1,2,3,5] → "#01–#03, #05"
function formatNumbers(nums) {
    const parts = [];
    let start = nums[0], prev = nums[0];
    for (let k = 1; k <= nums.length; k++) {
        const n = nums[k];
        if (n === prev + 1) { prev = n; continue; }
        parts.push(start === prev ? `#${padNum(start)}` : `#${padNum(start)}–#${padNum(prev)}`);
        start = prev = n;
    }
    return parts.join(", ");
}

function validFridgeColor(name) {
    return !!name && name.length <= 40 && !name.includes("|");
}

// Lê a folha Fridges → [{ color, numbers: [1,2,...], occupied: { 3: { order, customer } } }]
// Colunas: A=Cor | B=Número | C=ID da order que ocupa o fridge (vazio = livre) | D=Cliente
async function getFridges() {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${q(FRIDGES_SHEET)}!A2:D`
    });
    const map = new Map();
    for (const r of (res.data.values || [])) {
        const color = (r[0] || "").trim();
        const n = parseInt(String(r[1] ?? "").replace("#", ""), 10);
        if (!color || isNaN(n)) continue;
        const key = color.toLowerCase();
        if (!map.has(key)) map.set(key, { color, numbers: [], occupied: {} });
        const f = map.get(key);
        if (!f.numbers.includes(n)) f.numbers.push(n);
        const order = String(r[2] ?? "").trim();
        if (order) f.occupied[n] = { order, customer: String(r[3] ?? "").trim() };
    }
    return [...map.values()].map(f => ({ ...f, numbers: f.numbers.sort((a, b) => a - b) }));
}

// Reescreve a folha Fridges (uma linha por cor+número, mantendo a ocupação)
async function saveFridges(fridges) {
    const rows = fridges.flatMap(f => f.numbers.map(n => [
        f.color,
        padNum(n),
        f.occupied?.[n]?.order || "",
        f.occupied?.[n]?.customer || ""
    ]));
    await sheetsClient.spreadsheets.values.clear({
        spreadsheetId: SHEET_ID,
        range: `${q(FRIDGES_SHEET)}!A2:D`
    });
    if (rows.length) {
        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: SHEET_ID,
            range: `${q(FRIDGES_SHEET)}!A2:D${rows.length + 1}`,
            valueInputOption: "RAW",
            resource: { values: rows }
        });
    }
}

// O fridge está livre? (ou já é desta order, para permitir mudar/repetir a escolha)
function isFridgeFree(fridge, number, orderId = null) {
    const o = fridge.occupied?.[number];
    return !o || o.order === orderId;
}

// Marca o fridge como OCUPADO pela order (e liberta qualquer fridge anterior dessa mesma order)
async function assignFridge(orderId, customer, color, number) {
    const fridges = await getFridges();
    for (const f of fridges) {
        for (const n of Object.keys(f.occupied)) {
            if (f.occupied[n].order === orderId) delete f.occupied[n];
        }
    }
    const target = fridges.find(f => f.color === color);
    if (!target || !target.numbers.includes(number)) {
        throw new Error(`The fridge ${color} #${padNum(number)} no longer exists.`);
    }
    if (target.occupied[number]) {
        throw new Error(`The fridge ${color} #${padNum(number)} is already occupied by another order.`);
    }
    target.occupied[number] = { order: orderId, customer };
    await saveFridges(fridges);
}

// Liberta o(s) fridge(s) ocupado(s) por uma order
async function releaseFridge(orderId) {
    const fridges = await getFridges();
    let changed = false;
    for (const f of fridges) {
        for (const n of Object.keys(f.occupied)) {
            if (f.occupied[n].order === orderId) { delete f.occupied[n]; changed = true; }
        }
    }
    if (changed) await saveFridges(fridges);
}

/* -----------------------------------------------------------
   [PARTE 1] SHOP CATALOG + HOW TO BUY (mensagens informativas)
   Nota: nunca mostra quantidades, stock por cor nem preços.
------------------------------------------------------------*/
// Procura um canal de texto pelo nome (parte do nome)
function findTextChannel(guild, namePart) {
    return guild.channels.cache.find(c =>
        c.type === ChannelType.GuildText && c.name.toLowerCase().includes(namePart));
}

// Botão que abre o ticket diretamente (reutiliza o handler "t_open")
function openTicketRow() {
    return new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("t_open").setLabel("Open a ticket").setEmoji("🎫").setStyle(ButtonStyle.Success)
    );
}

// Mensagem do canal 📦│「shop-catalog」
function buildCatalogPayload(species, ticketChannel, howToChannel) {
    const ticketRef = ticketChannel ? `<#${ticketChannel.id}>` : "the ticket channel";
    const howToRef = howToChannel ? `<#${howToChannel.id}>` : "the How to Buy channel";

    // Lista só com os nomes das espécies (sem stock e sem preços)
    const sorted = [...species].sort((a, b) => a.localeCompare(b));
    let list = sorted.length
        ? sorted.map(n => `🦖 **${n}**`).join("\n")
        : "No species available right now. Check back soon!";
    if (list.length > 1800) list = list.slice(0, 1800).replace(/\n[^\n]*$/, "") + "\n…and more!";

    const header = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("🦕 TUGAS SHOP — Catalog")
        .setDescription(
            "Welcome! Here you can see what we offer and how our purchase system works.\n" +
            "Every order is made through a **private ticket** with our staff."
        );

    const speciesEmbed = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("🦖 Available species")
        .setDescription(`${list}\n\n*Availability and exact prices are shown while you build your order inside a ticket.*`);

    const types = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("🛒 Purchase types")
        .addFields(
            { name: "🦖 Pair", value: "1 male ♂ + 1 female ♀ of the **same color**. You choose the species and the color." },
            { name: "📦 Bundle", value: "1 male ♂ + 1 female ♀ of **every color** of a species. The full collection in a single purchase." },
            { name: "➕ Extra", value: "A **single dino** (male or female) in the color you choose. Perfect to complete a set or replace one." }
        );

    const payment = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("💳 Payment methods")
        .addFields(
            { name: "Tek", value: "Pay in-game with Tek.", inline: true },
            { name: "PayPal", value: "Pay in euros (€) via PayPal.", inline: true }
        );

    const start = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("🎫 Ready to order?")
        .setDescription(
            `Press the button below or go to ${ticketRef} to open your ticket.\n` +
            `Not sure how it works? Read the step-by-step guide in ${howToRef}.`
        )
        .setFooter({ text: CATALOG_MARKER });

    return { content: "", embeds: [header, speciesEmbed, types, payment, start], components: [openTicketRow()] };
}

// Mensagem do canal ❓│「how-to-buy」
function buildHowToBuyPayload(ticketChannel, catalogChannel) {
    const ticketRef = ticketChannel ? `<#${ticketChannel.id}>` : "the ticket channel";
    const catalogRef = catalogChannel ? `<#${catalogChannel.id}>` : "the Shop Catalog channel";

    const steps = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("❓ How to buy")
        .setDescription("Buying is simple and the bot guides you through every step. Here's how it works:")
        .addFields(
            { name: "1️⃣ Open a ticket", value: `Press **Open a ticket** below (or go to ${ticketRef}). A private channel is created that only you and our staff can see.` },
            { name: "2️⃣ Choose your payment method", value: "Pick **Tek** or **PayPal**. Prices depend on the method you choose." },
            { name: "3️⃣ Choose what to buy", value: `Pick a **Bundle** or a **Dino** (a Dino purchase can be a **Pair** or an **Extra**). See ${catalogRef} for what each one means.` },
            { name: "4️⃣ Pick species, color and gender", value: "Availability for each color is shown inside your ticket as you go." },
            { name: "5️⃣ Review your cart", value: "Added something by mistake? Use **Remove item** and pick the dino and color to take out just that entry — no need to clear the whole cart. Want more? Use **Add more purchases**." },
            { name: "6️⃣ Complete your order", value: "Press **Complete order**. Our staff is notified and will talk to you in your ticket to finish the payment." },
            { name: "7️⃣ We prepare it", value: "Once everything is confirmed, we place your order in a fridge and set its PIN." },
            { name: "8️⃣ Collect your order", value: "When it's ready you receive a **DM** with the fridge color, number and **PIN**. Go to the fridge and pick it up!" }
        );

    const tips = new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("💡 Good to know")
        .setDescription(
            "• Keep your **DMs open** so we can tell you when your order is ready.\n" +
            "• You can only have **one open ticket** at a time.\n" +
            "• Only the person who opened the ticket can edit the cart.\n" +
            "• Wait for staff to confirm your order before sending any payment."
        )
        .setFooter({ text: HOWTOBUY_MARKER });

    return { content: "", embeds: [steps, tips], components: [openTicketRow()] };
}

// Edita a mensagem do bot (se já existir) ou envia uma nova — evita duplicados
async function upsertInfoMessage(channel, marker, payload) {
    const messages = await channel.messages.fetch({ limit: 50 });
    const existing = messages.find(m =>
        m.author.id === client.user.id &&
        m.embeds.some(e => (e.footer?.text || "").includes(marker)));
    if (existing) return existing.edit(payload);
    return channel.send(payload);
}

// Publica/atualiza o catálogo (e, com force=true, também o how-to-buy)
async function refreshShopInfo(speciesList = null, force = false) {
    if (!client.isReady() || !sheetsClient || catalogRefreshRunning) return;
    catalogRefreshRunning = true;
    try {
        const species = speciesList || await getSpeciesSheets();
        const signature = [...species].sort().join("|");
        if (!force && signature === catalogSignature) return; // nada mudou

        const guild = await client.guilds.fetch(GUILD_ID);
        await guild.channels.fetch();
        const catalogChannel = findTextChannel(guild, CATALOG_CHANNEL_NAME);
        const howToChannel = findTextChannel(guild, HOWTOBUY_CHANNEL_NAME);
        const ticketChannel = findTextChannel(guild, TICKET_CHANNEL_NAME);

        if (catalogChannel) {
            await upsertInfoMessage(catalogChannel, CATALOG_MARKER, buildCatalogPayload(species, ticketChannel, howToChannel));
        } else {
            console.warn('Channel "📦│「shop-catalog」" not found. Create it or check the name/permissions.');
        }

        if (force) {
            if (howToChannel) {
                await upsertInfoMessage(howToChannel, HOWTOBUY_MARKER, buildHowToBuyPayload(ticketChannel, catalogChannel));
            } else {
                console.warn('Channel "❓│「how-to-buy」" not found. Create it or check the name/permissions.');
            }
        }
        catalogSignature = signature;
    } catch (err) {
        console.error("Error refreshing the shop catalog:", err);
    } finally {
        catalogRefreshRunning = false;
    }
}

/* -----------------------------------------------------------
   SLASH COMMANDS
------------------------------------------------------------*/
async function registerCommands() {
    const commands = [
        new SlashCommandBuilder()
            .setName("adddino")
            .setDescription("Add a new species")
            .addStringOption(o => o.setName("name").setDescription("Species name").setRequired(true))
            .addNumberOption(o => o.setName("pair").setDescription("Pair price (1♂ + 1♀ of one color) in Tek").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("bundle").setDescription("Bundle price (1♂ + 1♀ of EACH color) in Tek").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("extra").setDescription("Extra dino price in Tek").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("pair_paypal").setDescription("Pair price in PayPal (€)").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("bundle_paypal").setDescription("Bundle price in PayPal (€)").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("extra_paypal").setDescription("Extra dino price in PayPal (€)").setRequired(true).setMinValue(0)),

        new SlashCommandBuilder()
            .setName("addcolor")
            .setDescription("Add a color to a species")
            .addStringOption(o =>
                o.setName("species")
                    .setDescription("Species")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o => o.setName("color").setDescription("Color name").setRequired(true)),

        new SlashCommandBuilder()
            .setName("stock")
            .setDescription("View the stock of all species"),

        new SlashCommandBuilder()
            .setName("ticketpanel")
            .setDescription("Post the ticket panel in this channel"),

        new SlashCommandBuilder()
            .setName("addfridges")
            .setDescription("Add a fridge color, or add numbers to an existing color")
            .addStringOption(o => o.setName("color").setDescription("Fridge color").setRequired(true).setAutocomplete(true))
            .addStringOption(o => o.setName("numbers").setDescription("Numbers to add, e.g. 1-10, 11 or 1,2,5 (default: 1-10)")),

        new SlashCommandBuilder()
            .setName("editfridges")
            .setDescription("Rename a fridge color, or change one of its numbers")
            .addStringOption(o => o.setName("color").setDescription("Fridge color").setRequired(true).setAutocomplete(true))
            .addStringOption(o => o.setName("new_name").setDescription("New color name"))
            .addStringOption(o => o.setName("number").setDescription("Number to change").setAutocomplete(true))
            .addIntegerOption(o => o.setName("new_number").setDescription("New number").setMinValue(1).setMaxValue(FRIDGE_MAX_NUMBER)),

        new SlashCommandBuilder()
            .setName("delfridges")
            .setDescription("Delete a fridge color, or just one number of that color")
            .addStringOption(o => o.setName("color").setDescription("Fridge color").setRequired(true).setAutocomplete(true))
            .addStringOption(o => o.setName("number").setDescription("Number to delete (leave empty to delete the whole color)").setAutocomplete(true)),

        new SlashCommandBuilder()
            .setName("fridges")
            .setDescription("Show which fridges are free or occupied"),

        new SlashCommandBuilder()
            .setName("shoppanels")
            .setDescription("Refresh the Shop Catalog and How to Buy messages"),

        new SlashCommandBuilder()
            .setName("deldino")
            .setDescription("Delete a species (removes its sheet and prices)")
            .addStringOption(o =>
                o.setName("species")
                    .setDescription("Species to delete")
                    .setRequired(true)
                    .setAutocomplete(true)),

        new SlashCommandBuilder()
            .setName("delcolor")
            .setDescription("Delete a color from a species")
            .addStringOption(o =>
                o.setName("species")
                    .setDescription("Species")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o =>
                o.setName("color")
                    .setDescription("Color to delete")
                    .setRequired(true)
                    .setAutocomplete(true)),

        new SlashCommandBuilder()
            .setName("editdino")
            .setDescription("Edit a species (name and/or prices)")
            .addStringOption(o =>
                o.setName("species")
                    .setDescription("Species to edit")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o => o.setName("new_name").setDescription("New species name"))
            .addNumberOption(o => o.setName("pair").setDescription("New Pair price in Tek").setMinValue(0))
            .addNumberOption(o => o.setName("bundle").setDescription("New Bundle price in Tek").setMinValue(0))
            .addNumberOption(o => o.setName("extra").setDescription("New extra dino price in Tek").setMinValue(0))
            .addNumberOption(o => o.setName("pair_paypal").setDescription("New Pair price in PayPal (€)").setMinValue(0))
            .addNumberOption(o => o.setName("bundle_paypal").setDescription("New Bundle price in PayPal (€)").setMinValue(0))
            .addNumberOption(o => o.setName("extra_paypal").setDescription("New extra dino price in PayPal (€)").setMinValue(0)),

        new SlashCommandBuilder()
            .setName("editcolor")
            .setDescription("Rename a color of a species; stock is adjusted via the dashboard")
            .addStringOption(o =>
                o.setName("species")
                    .setDescription("Species")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o =>
                o.setName("color")
                    .setDescription("Color to edit")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o => o.setName("new_name").setDescription("New color name"))
    ].map(cmd => cmd.toJSON());

    const rest = new REST({ version: "10" }).setToken(TOKEN);

    await rest.put(
        Routes.applicationGuildCommands(client.user.id, GUILD_ID),
        { body: commands }
    );

    console.log("Slash commands registered!");
}

/* -----------------------------------------------------------
   INTERACTIONS (tudo num só handler)
------------------------------------------------------------*/
client.on("interactionCreate", async (interaction) => {
    try {
        // Autocomplete das espécies no /addcolor
        if (interaction.isAutocomplete()) {
            const focusedOption = interaction.options.getFocused(true);
            const typed = focusedOption.value.toLowerCase();

            let list = [];

            if (["addfridges", "editfridges", "delfridges"].includes(interaction.commandName)) {
                // [PARTE 3] Cores e números dos fridges
                const fridges = await getFridges();
                if (focusedOption.name === "color") {
                    list = fridges.map(f => f.color);
                } else if (focusedOption.name === "number") {
                    const chosen = (interaction.options.getString("color") || "").trim().toLowerCase();
                    const f = fridges.find(x => x.color.toLowerCase() === chosen);
                    list = f ? f.numbers.map(padNum) : [];
                }
            } else if (focusedOption.name === "species") {
                list = await getSpeciesSheets();
            } else if (focusedOption.name === "color" && ["delcolor", "editcolor"].includes(interaction.commandName)) {
                // Só mostra cores se a espécie já tiver sido escolhida
                const especie = interaction.options.getString("species");
                if (especie) {
                    const species = await getSpeciesSheets();
                    if (species.includes(especie)) list = await getColorsForSpecies(especie);
                }
            }

            const choices = list
                .filter(s => s.toLowerCase().includes(typed))
                .slice(0, 25) // limite do Discord
                .map(s => ({ name: s, value: s }));

            return interaction.respond(choices);
        }

        // [PARTE 3] Modais (PIN do fridge)
        if (interaction.isModalSubmit()) return handleModalSubmit(interaction);

        // Componentes da dashboard e do sistema de tickets
        if (interaction.isButton() || interaction.isStringSelectMenu()) {
            if (interaction.customId?.startsWith("dash_")) return handleDashboardComponent(interaction);
            // [PARTE 3] Botões/menus das orders (canal de logs): claim, fridge, PIN, ready
            if (interaction.customId?.startsWith("o_")) return handleOrderComponent(interaction);
            return handleTicketComponent(interaction);
        }

        if (!interaction.isChatInputCommand()) return;

        /* ---------- /stock (aberto a todos) ---------- */
        if (interaction.commandName === "stock") {
            await interaction.deferReply();

            const data = await getStockData();
            if (data.length === 0) {
                return interaction.editReply("No species registered yet.");
            }

            const lines = data.map(d => `${stockBall(d.bundles)}  **${d.name}**`);
            let description = lines.join("\n");
            if (description.length > 3800) description = description.slice(0, 3800) + "\n…";

            const embed = new EmbedBuilder()
                .setTitle("📦 Stock")
                .setDescription(description)
                .setFooter({ text: "🟢 Available  🟡 2 bundles  🟠 Last bundle  🔴 Out of stock" });

            return interaction.editReply({ embeds: [embed] });
        }

        if (!isStaff(interaction)) {
            return interaction.reply({
                content: "Only staff can use this command.",
                flags: MessageFlags.Ephemeral
            });
        }

        /* ---------- /adddino ---------- */
        if (interaction.commandName === "adddino") {
            await interaction.deferReply();

            const nome = interaction.options.getString("name").trim();
            const pair = interaction.options.getNumber("pair");
            const bundle = interaction.options.getNumber("bundle");
            const extra = interaction.options.getNumber("extra");
            const pairPP = interaction.options.getNumber("pair_paypal");
            const bundlePP = interaction.options.getNumber("bundle_paypal");
            const extraPP = interaction.options.getNumber("extra_paypal");

            if (!nome || nome.length > 100 || /[\[\]\*\?:\/\\]/.test(nome)) {
                return interaction.editReply("🔴 Invalid name. Don't use `[ ] * ? : / \\` and max 100 characters.");
            }

            const existing = await getSpeciesSheets();
            if (existing.some(s => s.toLowerCase() === nome.toLowerCase()) ||
                IGNORED_SHEETS.some(s => s.toLowerCase() === nome.toLowerCase())) {
                return interaction.editReply(`🔴 The species **${nome}** already exists.`);
            }

            // Criar folha
            await sheetsClient.spreadsheets.batchUpdate({
                spreadsheetId: SHEET_ID,
                resource: { requests: [{ addSheet: { properties: { title: nome } } }] }
            });

            // Cabeçalho
            await sheetsClient.spreadsheets.values.update({
                spreadsheetId: SHEET_ID,
                range: `${q(nome)}!A1:C1`,
                valueInputOption: "USER_ENTERED",
                resource: { values: [["Color", "Male", "Female"]] }
            });

            // Preços
            await sheetsClient.spreadsheets.values.append({
                spreadsheetId: SHEET_ID,
                range: "Precos!A:G",
                valueInputOption: "USER_ENTERED",
                resource: { values: [[nome, pair, bundle, extra, pairPP, bundlePP, extraPP]] }
            });

            await updateStockSheet();

            return interaction.editReply(
                `🟢 Species **${nome}** created successfully!\n` +
                `Tek — Pair: **${pair}** · Bundle: **${bundle}** · Extra: **${extra}**\n` +
                `PayPal — Pair: **${pairPP}€** · Bundle: **${bundlePP}€** · Extra: **${extraPP}€**`
            );
        }

        /* ---------- /addcolor ---------- */
        if (interaction.commandName === "addcolor") {
            await interaction.deferReply();

            const especie = interaction.options.getString("species");
            const color = interaction.options.getString("color").trim();

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 The species **${especie}** doesn't exist. Pick one from the list.`);
            }
            if (!color || color.length > 100) return interaction.editReply("🔴 Invalid color name (max 100 characters).");

            const existingColors = await getColorsForSpecies(especie);
            if (existingColors.some(c => c.toLowerCase() === color.toLowerCase())) {
                return interaction.editReply(`🔴 The color **${color}** already exists in **${especie}**.`);
            }

            // A cor é criada a zero; a staff ajusta o stock pelos botões da dashboard.
            await sheetsClient.spreadsheets.values.append({
                spreadsheetId: SHEET_ID,
                range: `${q(especie)}!A:C`,
                valueInputOption: "USER_ENTERED",
                resource: { values: [[color, 0, 0]] }
            });

            dashboardSpecies = especie;
            dashboardColor = color;
            await updateStockSheet();
            await refreshDashboard();

            return interaction.editReply(`🟢 Color **${color}** added to **${especie}** with zero stock. Use the dashboard buttons to add quantity.`);
        }

        /* ---------- /deldino ---------- */
        if (interaction.commandName === "deldino") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });

            const especie = interaction.options.getString("species");

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 The species **${especie}** doesn't exist.`);
            }

            // Pedido de confirmação (apaga a folha inteira!)
            const buttons = new ActionRowBuilder().addComponents(
                new ButtonBuilder()
                    .setCustomId("confirm_deldino")
                    .setLabel("Delete")
                    .setStyle(ButtonStyle.Danger),
                new ButtonBuilder()
                    .setCustomId("cancel_deldino")
                    .setLabel("Cancel")
                    .setStyle(ButtonStyle.Secondary)
            );

            const msg = await interaction.editReply({
                content: `⚠️ Are you sure you want to delete **${especie}**?\nThis deletes the sheet with all colors and removes the prices. **This cannot be undone.**`,
                components: [buttons]
            });

            let click;
            try {
                click = await msg.awaitMessageComponent({
                    filter: i => i.user.id === interaction.user.id,
                    time: 30_000
                });
            } catch {
                return interaction.editReply({ content: "⏱️ Timed out. Nothing was deleted.", components: [] });
            }

            if (click.customId !== "confirm_deldino") {
                return click.update({ content: "Cancelled. Nothing was deleted.", components: [] });
            }

            await click.update({ content: "Deleting...", components: [] });

            // 1) Apagar a folha da espécie
            const sheetId = await getSheetId(especie);
            if (sheetId === null) throw new Error(`Sheet "${especie}" not found.`);

            await sheetsClient.spreadsheets.batchUpdate({
                spreadsheetId: SHEET_ID,
                resource: { requests: [{ deleteSheet: { sheetId } }] }
            });

            // 2) Remover a linha na folha Precos
            const precos = await sheetsClient.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: "Precos!A:A"
            });
            const precoRows = precos.data.values || [];
            const precoIndex = precoRows.findIndex(r => (r[0] || "").trim() === especie);
            if (precoIndex !== -1) await deleteRow("Precos", precoIndex);

            // 3) Atualizar Stock
            await updateStockSheet();

            return interaction.editReply(`🟢 Species **${especie}** deleted successfully!`);
        }

        /* ---------- /delcolor ---------- */
        if (interaction.commandName === "delcolor") {
            await interaction.deferReply();

            const especie = interaction.options.getString("species");
            const color = interaction.options.getString("color").trim();

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 The species **${especie}** doesn't exist. Pick one from the list.`);
            }

            const res = await sheetsClient.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: `${q(especie)}!A:A`
            });
            const rows = res.data.values || [];

            // Começa em 1 para nunca apagar o cabeçalho (linha 0)
            let rowIndex = -1;
            for (let i = 1; i < rows.length; i++) {
                if ((rows[i][0] || "").trim().toLowerCase() === color.toLowerCase()) {
                    rowIndex = i;
                    break;
                }
            }

            if (rowIndex === -1) {
                return interaction.editReply(`🔴 The color **${color}** doesn't exist in **${especie}**.`);
            }

            await deleteRow(especie, rowIndex);
            await updateStockSheet();

            return interaction.editReply(`🟢 Color **${color}** deleted from **${especie}**!`);
        }

        /* ---------- /editdino ---------- */
        if (interaction.commandName === "editdino") {
            await interaction.deferReply();

            const especie = interaction.options.getString("species");
            const novoNomeRaw = interaction.options.getString("new_name");
            const novoNome = novoNomeRaw ? novoNomeRaw.trim() : null;
            const pair = interaction.options.getNumber("pair");
            const bundle = interaction.options.getNumber("bundle");
            const extra = interaction.options.getNumber("extra");

            const pairPP = interaction.options.getNumber("pair_paypal");
            const bundlePP = interaction.options.getNumber("bundle_paypal");
            const extraPP = interaction.options.getNumber("extra_paypal");

            if (!novoNome && pair === null && bundle === null && extra === null && pairPP === null && bundlePP === null && extraPP === null) {
                return interaction.editReply("🔴 Provide at least one thing to change: `new_name`, `pair`, `bundle`, `extra`, `pair_paypal`, `bundle_paypal` or `extra_paypal`.");
            }

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 The species **${especie}** doesn't exist. Pick one from the list.`);
            }

            const renaming = novoNome && novoNome !== especie;

            if (renaming) {
                if (novoNome.length > 100 || /[\[\]\*\?:\/\\]/.test(novoNome)) {
                    return interaction.editReply("🔴 Invalid name. Don't use `[ ] * ? : / \\` and max 100 characters.");
                }
                const lower = novoNome.toLowerCase();
                if (species.some(s => s.toLowerCase() === lower && s !== especie) ||
                    IGNORED_SHEETS.some(s => s.toLowerCase() === lower)) {
                    return interaction.editReply(`🔴 A species named **${novoNome}** already exists.`);
                }
            }

            // Linha na folha Precos
            const precos = await sheetsClient.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: "Precos!A:G"
            });
            const precoRows = precos.data.values || [];
            const precoIndex = precoRows.findIndex(r => (r[0] || "").trim() === especie);
            const current = precoIndex !== -1 ? precoRows[precoIndex] : [especie, "", "", "", "", "", ""];

            // 1) Renomear a folha
            if (renaming) {
                const sheetId = await getSheetId(especie);
                if (sheetId === null) throw new Error(`Sheet "${especie}" not found.`);

                await sheetsClient.spreadsheets.batchUpdate({
                    spreadsheetId: SHEET_ID,
                    resource: {
                        requests: [{
                            updateSheetProperties: {
                                properties: { sheetId, title: novoNome },
                                fields: "title"
                            }
                        }]
                    }
                });
            }

            // 2) Atualizar Precos (nome + preços)
            const finalName = renaming ? novoNome : especie;
            const newRow = [
                finalName,
                pair !== null ? pair : (current[1] ?? ""),
                bundle !== null ? bundle : (current[2] ?? ""),
                extra !== null ? extra : (current[3] ?? ""),
                pairPP !== null ? pairPP : (current[4] ?? ""),
                bundlePP !== null ? bundlePP : (current[5] ?? ""),
                extraPP !== null ? extraPP : (current[6] ?? "")
            ];

            if (precoIndex !== -1) {
                await sheetsClient.spreadsheets.values.update({
                    spreadsheetId: SHEET_ID,
                    range: `Precos!A${precoIndex + 1}:G${precoIndex + 1}`,
                    valueInputOption: "USER_ENTERED",
                    resource: { values: [newRow] }
                });
            } else {
                await sheetsClient.spreadsheets.values.append({
                    spreadsheetId: SHEET_ID,
                    range: "Precos!A:G",
                    valueInputOption: "USER_ENTERED",
                    resource: { values: [newRow] }
                });
            }

            // 3) Atualizar Stock (o nome pode ter mudado)
            if (renaming && dashboardSpecies === especie) dashboardSpecies = novoNome;
            await updateStockSheet();

            return interaction.editReply(
                `🟢 Species **${finalName}** updated!` +
                (renaming ? `\nName: ${especie} → **${finalName}**` : "") +
                (pair !== null ? `\nPair Tek: **${pair}**` : "") +
                (bundle !== null ? `\nBundle Tek: **${bundle}**` : "") +
                (extra !== null ? `\nExtra Tek: **${extra}**` : "") +
                (pairPP !== null ? `\nPair PayPal: **${pairPP}€**` : "") +
                (bundlePP !== null ? `\nBundle PayPal: **${bundlePP}€**` : "") +
                (extraPP !== null ? `\nExtra PayPal: **${extraPP}€**` : "")
            );
        }

        /* ---------- /editcolor ---------- */
        if (interaction.commandName === "editcolor") {
            await interaction.deferReply();

            const especie = interaction.options.getString("species");
            const color = interaction.options.getString("color").trim();
            const novoNomeRaw = interaction.options.getString("new_name");
            const novoNome = novoNomeRaw ? novoNomeRaw.trim() : null;
            if (!novoNome) {
                return interaction.editReply("🔴 Provide the color's `new_name`. To change quantities, use the dashboard buttons.");
            }

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 The species **${especie}** doesn't exist. Pick one from the list.`);
            }

            const res = await sheetsClient.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: `${q(especie)}!A:C`
            });
            const rows = res.data.values || [];

            // Começa em 1 para nunca editar o cabeçalho
            let rowIndex = -1;
            for (let i = 1; i < rows.length; i++) {
                if ((rows[i][0] || "").trim().toLowerCase() === color.toLowerCase()) {
                    rowIndex = i;
                    break;
                }
            }

            if (rowIndex === -1) {
                return interaction.editReply(`🔴 The color **${color}** doesn't exist in **${especie}**.`);
            }

            const currentRow = rows[rowIndex];
            const currentName = (currentRow[0] || "").trim();

            if (novoNome && novoNome.toLowerCase() !== currentName.toLowerCase()) {
                const clash = rows.some((r, i) => i > 0 && i !== rowIndex &&
                    (r[0] || "").trim().toLowerCase() === novoNome.toLowerCase());
                if (clash) {
                    return interaction.editReply(`🔴 A color **${novoNome}** already exists in **${especie}**.`);
                }
            }

            const finalName = novoNome;

            await sheetsClient.spreadsheets.values.update({
                spreadsheetId: SHEET_ID,
                range: `${q(especie)}!A${rowIndex + 1}:C${rowIndex + 1}`,
                valueInputOption: "USER_ENTERED",
                resource: { values: [[finalName, currentRow[1] ?? 0, currentRow[2] ?? 0]] }
            });

            if (dashboardSpecies === especie && dashboardColor === currentName) dashboardColor = finalName;
            await updateStockSheet();

            return interaction.editReply(`🟢 Color **${currentName}** of **${especie}** renamed to **${finalName}**. Quantities are changed on the dashboard.`);
        }

        /* ---------- /addfridges ---------- */
        if (interaction.commandName === "addfridges") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });

            const color = interaction.options.getString("color").trim();
            const numbersRaw = interaction.options.getString("numbers");

            if (!validFridgeColor(color)) {
                return interaction.editReply("🔴 Invalid color name (max 40 characters, no `|`).");
            }
            // Sem "numbers" → #01 … #10
            const nums = numbersRaw
                ? parseNumbers(numbersRaw)
                : Array.from({ length: FRIDGE_DEFAULT_COUNT }, (_, n) => n + 1);
            if (!nums) {
                return interaction.editReply(`🔴 Invalid numbers. Use for example \`1-10\`, \`11\` or \`1,2,5\` (values from 1 to ${FRIDGE_MAX_NUMBER}).`);
            }

            const fridges = await getFridges();
            const existing = fridges.find(f => f.color.toLowerCase() === color.toLowerCase());

            if (existing) {
                // Cor já existe → acrescenta só os números novos
                const toAdd = nums.filter(n => !existing.numbers.includes(n));
                if (!toAdd.length) {
                    return interaction.editReply(`🔴 **${existing.color}** already has those numbers.`);
                }
                if (existing.numbers.length + toAdd.length > FRIDGE_MAX_PER_COLOR) {
                    return interaction.editReply(`🔴 A color can have at most ${FRIDGE_MAX_PER_COLOR} numbers.`);
                }
                existing.numbers = [...existing.numbers, ...toAdd].sort((a, b) => a - b);
                await saveFridges(fridges);
                return interaction.editReply(`🟢 Added ${formatNumbers(toAdd)} to **${existing.color}**.\nNow: ${formatNumbers(existing.numbers)}`);
            }

            // Cor nova
            if (fridges.length >= FRIDGE_MAX_COLORS) {
                return interaction.editReply(`🔴 You can have at most ${FRIDGE_MAX_COLORS} fridge colors.`);
            }
            if (nums.length > FRIDGE_MAX_PER_COLOR) {
                return interaction.editReply(`🔴 A color can have at most ${FRIDGE_MAX_PER_COLOR} numbers.`);
            }
            fridges.push({ color, numbers: nums, occupied: {} });
            await saveFridges(fridges);
            return interaction.editReply(`🟢 Fridge color **${color}** created with ${formatNumbers(nums)}.`);
        }

        /* ---------- /editfridges ---------- */
        if (interaction.commandName === "editfridges") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });

            const color = interaction.options.getString("color").trim();
            const newName = (interaction.options.getString("new_name") || "").trim() || null;
            const numberRaw = interaction.options.getString("number");
            const newNumber = interaction.options.getInteger("new_number");

            if (!newName && numberRaw === null && newNumber === null) {
                return interaction.editReply("🔴 Provide `new_name`, or `number` + `new_number`.");
            }
            if ((numberRaw === null) !== (newNumber === null)) {
                return interaction.editReply("🔴 To change a number, provide both `number` and `new_number`.");
            }

            const fridges = await getFridges();
            const fridge = fridges.find(f => f.color.toLowerCase() === color.toLowerCase());
            if (!fridge) return interaction.editReply(`🔴 The fridge color **${color}** doesn't exist.`);

            const changes = [];

            // Renomear a cor
            if (newName && newName !== fridge.color) {
                if (!validFridgeColor(newName)) {
                    return interaction.editReply("🔴 Invalid color name (max 40 characters, no `|`).");
                }
                if (fridges.some(f => f !== fridge && f.color.toLowerCase() === newName.toLowerCase())) {
                    return interaction.editReply(`🔴 A fridge color named **${newName}** already exists.`);
                }
                changes.push(`Color: ${fridge.color} → **${newName}**`);
                fridge.color = newName;
            }

            // Trocar um número
            if (numberRaw !== null) {
                const oldN = parseInt(numberRaw.replace("#", ""), 10);
                if (isNaN(oldN) || !fridge.numbers.includes(oldN)) {
                    return interaction.editReply(`🔴 **${fridge.color}** has no fridge ${numberRaw}.`);
                }
                if (newNumber !== oldN && fridge.numbers.includes(newNumber)) {
                    return interaction.editReply(`🔴 **${fridge.color}** already has #${padNum(newNumber)}.`);
                }
                if (fridge.occupied[oldN]) {
                    return interaction.editReply(`🔴 **${fridge.color} #${padNum(oldN)}** is occupied by an order. Mark that order as collected first.`);
                }
                fridge.numbers = fridge.numbers.map(n => n === oldN ? newNumber : n).sort((a, b) => a - b);
                changes.push(`Number: #${padNum(oldN)} → **#${padNum(newNumber)}**`);
            }

            if (!changes.length) return interaction.editReply("Nothing to change.");
            await saveFridges(fridges);
            return interaction.editReply(`🟢 **${fridge.color}** updated!\n${changes.join("\n")}`);
        }

        /* ---------- /delfridges ---------- */
        if (interaction.commandName === "delfridges") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });

            const color = interaction.options.getString("color").trim();
            const numberRaw = interaction.options.getString("number");

            const fridges = await getFridges();
            const fridge = fridges.find(f => f.color.toLowerCase() === color.toLowerCase());
            if (!fridge) return interaction.editReply(`🔴 The fridge color **${color}** doesn't exist.`);

            // Apagar apenas UM número da cor
            if (numberRaw !== null) {
                const n = parseInt(numberRaw.replace("#", ""), 10);
                if (isNaN(n) || !fridge.numbers.includes(n)) {
                    return interaction.editReply(`🔴 **${fridge.color}** has no fridge ${numberRaw}.`);
                }
                if (fridge.occupied[n]) {
                    return interaction.editReply(`🔴 **${fridge.color} #${padNum(n)}** is occupied by an order. Mark that order as collected first.`);
                }
                fridge.numbers = fridge.numbers.filter(x => x !== n);
                // Se já não sobrar nenhum número, a cor deixa de fazer sentido
                const remaining = fridge.numbers.length ? fridges : fridges.filter(f => f !== fridge);
                await saveFridges(remaining);
                return interaction.editReply(
                    `🟢 Fridge **${fridge.color} #${padNum(n)}** deleted.` +
                    (fridge.numbers.length ? `\nNow: ${formatNumbers(fridge.numbers)}` : `\nThat was the last number, so the color **${fridge.color}** was removed.`)
                );
            }

            // Não deixa apagar uma cor com fridges ocupados
            if (Object.keys(fridge.occupied).length) {
                return interaction.editReply(`🔴 **${fridge.color}** has occupied fridges (${formatNumbers(Object.keys(fridge.occupied).map(Number).sort((a, b) => a - b))}). Mark those orders as collected first.`);
            }

            // Apagar a cor inteira (pede confirmação)
            const buttons = new ActionRowBuilder().addComponents(
                new ButtonBuilder().setCustomId("confirm_delfridge").setLabel("Delete").setStyle(ButtonStyle.Danger),
                new ButtonBuilder().setCustomId("cancel_delfridge").setLabel("Cancel").setStyle(ButtonStyle.Secondary)
            );
            const msg = await interaction.editReply({
                content: `⚠️ Delete the fridge color **${fridge.color}** with all its numbers (${formatNumbers(fridge.numbers)})?`,
                components: [buttons]
            });

            let click;
            try {
                click = await msg.awaitMessageComponent({ filter: i => i.user.id === interaction.user.id, time: 30_000 });
            } catch {
                return interaction.editReply({ content: "⏱️ Timed out. Nothing was deleted.", components: [] });
            }
            if (click.customId !== "confirm_delfridge") {
                return click.update({ content: "Cancelled. Nothing was deleted.", components: [] });
            }

            await click.update({ content: "Deleting...", components: [] });
            await saveFridges(fridges.filter(f => f !== fridge));
            return interaction.editReply(`🟢 Fridge color **${fridge.color}** deleted.`);
        }

        /* ---------- /fridges (lista livres/ocupados) ---------- */
        if (interaction.commandName === "fridges") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });

            const fridges = await getFridges();
            if (!fridges.length) {
                return interaction.editReply("No fridges configured yet. Use `/addfridges` first.");
            }

            const fields = fridges.slice(0, 25).map(f => {
                const grid = f.numbers.map(n => `${f.occupied[n] ? "🔴" : "🟢"}${padNum(n)}`).join(" ");
                const who = f.numbers.filter(n => f.occupied[n])
                    .map(n => `#${padNum(n)} → ${f.occupied[n].customer || "order"}`).join("\n");
                return { name: f.color, value: (grid + (who ? `\n${who}` : "")).slice(0, 1024) || "—" };
            });
            const embed = new EmbedBuilder()
                .setTitle("🧊 Fridges")
                .setDescription("🟢 free · 🔴 occupied")
                .addFields(fields)
                .setFooter({ text: "A fridge is freed when its order is marked as collected." });
            return interaction.editReply({ embeds: [embed] });
        }

        /* ---------- /shoppanels (força a atualização das mensagens informativas) ---------- */
        if (interaction.commandName === "shoppanels") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });
            await refreshShopInfo(null, true);
            return interaction.editReply("🟢 Shop Catalog and How to Buy messages refreshed.");
        }

        /* ---------- /ticketpanel ---------- */
        if (interaction.commandName === "ticketpanel") {
            await publishTicketPanel(interaction.channel);
            return interaction.reply({ content: "🟢 Panel posted.", flags: MessageFlags.Ephemeral });
        }
    } catch (err) {
        console.error("Interaction error:", err);

        // Autocomplete não pode receber mensagem de texto
        if (interaction.isAutocomplete()) return;

        const msg = `🔴 Error: ${err.message}`;
        try {
            // Em botões/menus não editamos a mensagem (perdia o carrinho)
            if (interaction.isMessageComponent() || interaction.isModalSubmit()) {
                const opts = { content: msg, flags: MessageFlags.Ephemeral };
                if (interaction.deferred || interaction.replied) await interaction.followUp(opts);
                else await interaction.reply(opts);
                return;
            }

            if (interaction.deferred || interaction.replied) {
                await interaction.editReply(msg);
            } else {
                await interaction.reply({ content: msg, flags: MessageFlags.Ephemeral });
            }
        } catch (e) {
            console.error("Could not reply with the error:", e);
        }
    }
});

/* -----------------------------------------------------------
   SISTEMA DE TICKETS
------------------------------------------------------------*/
let TICKET_CATEGORY_ID = null;  // definido automaticamente pela categoria do canal open-a-ticket
const STAFF_CATEGORY_ID = null;   // (opcional) ID da categoria dos pedidos para a staff
const PAYPAL_SYMBOL = "€";

// Estado de cada ticket aberto: channelId -> { userId, payment, mode, species, color, cart }
// (perde-se se o bot reiniciar)
const tickets = new Map();

const PAY_LABEL = { tek: "Tek", paypal: "PayPal" };

function money(method, value) {
    const v = Math.round(value * 100) / 100;
    return method === "tek" ? `${v} Tek` : `${v}${PAYPAL_SYMBOL}`;
}

function safeName(s) {
    return s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "cliente";
}

// Precos: A=Nome | B=Tek Pair | C=Tek Bundle | D=Tek Extra | E=PayPal Pair | F=PayPal Bundle | G=PayPal Extra
async function getPrices(species) {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: "Precos!A:G"
    });
    const row = (res.data.values || []).find(r => (r[0] || "").trim() === species);
    if (!row) return null;

    const n = v => {
        const x = parseFloat(String(v ?? "").replace(",", "."));
        return isNaN(x) ? null : x;
    };

    return {
        tek: { pair: n(row[1]), bundle: n(row[2]), dino: n(row[3]) },
        paypal: { pair: n(row[4]), bundle: n(row[5]), dino: n(row[6]) }
    };
}

function priceText(prices, method, kind) {
    const v = prices?.[method]?.[kind];
    return v === null || v === undefined ? "⚠️ not set" : money(method, v);
}

async function getColorStock(species) {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${q(species)}!A2:C`
    });
    return (res.data.values || [])
        .filter(r => (r[0] || "").trim() !== "")
        .map(r => ({
            color: r[0].trim(),
            male: parseInt(r[1] || "0", 10) || 0,
            female: parseInt(r[2] || "0", 10) || 0
        }));
}

function cartQty(st, species, color, gender) {
    const it = st.cart.find(c =>
        c.type === "dino" && c.species === species && c.color === color && c.gender === gender);
    return it ? it.qty : 0;
}

// Calcula linhas e total do carrinho com o método de pagamento escolhido
async function calcCart(st) {
    const cache = new Map();
    const lines = [];
    const errors = [];
    let total = 0;

    for (const it of st.cart) {
        if (!cache.has(it.species)) cache.set(it.species, await getPrices(it.species));
        const p = cache.get(it.species);
        const priceKind = it.type; // "bundle" | "pair" | "dino" — cada um com o seu preço
        const unit = p ? p[st.payment][priceKind] : null;

        const label = it.type === "bundle"
            ? `📦 **${it.species}** — Bundle ×${it.qty}`
            : it.type === "pair"
                ? `🦖 **${it.species}** — Pair ${it.color} (♂ + ♀) ×${it.qty}`
                : `🦖 **${it.species}** — ${it.color} ${it.gender === "m" ? "♂" : "♀"} ×${it.qty}`;

        if (unit === null || unit === undefined) {
            errors.push(it.species);
            lines.push(`${label} — ⚠️ ${PAY_LABEL[st.payment]} price not set`);
            continue;
        }

        const sub = unit * it.qty;
        total += sub;
        lines.push(`${label} — ${money(st.payment, sub)}`);
    }

    return { lines, total, errors };
}

/* ---------- Ecrãs (mensagens com botões/menus) ---------- */
const cancelButton = () =>
    new ButtonBuilder().setCustomId("t_cancel").setLabel("Cancel").setStyle(ButtonStyle.Secondary);

const backButton = () =>
    new ButtonBuilder().setCustomId("t_more").setLabel("⬅️ Back").setStyle(ButtonStyle.Secondary);

const closeRow = () =>
    new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("t_close").setLabel("Close").setEmoji("🔒").setStyle(ButtonStyle.Danger)
    );

function paymentScreen() {
    return {
        content: "",
        embeds: [new EmbedBuilder().setTitle("💳 Payment method").setDescription("How would you like to pay?")],
        components: [new ActionRowBuilder().addComponents(
            new ButtonBuilder().setCustomId("t_pay_tek").setLabel("Tek").setStyle(ButtonStyle.Primary),
            new ButtonBuilder().setCustomId("t_pay_paypal").setLabel("PayPal").setStyle(ButtonStyle.Primary),
            cancelButton()
        )]
    };
}

function typeScreen(st) {
    return {
        content: "",
        embeds: [new EmbedBuilder()
            .setTitle("🛒 What would you like to buy?")
            .setDescription(`Payment: **${PAY_LABEL[st.payment]}**`)],
        components: [new ActionRowBuilder().addComponents(
            new ButtonBuilder().setCustomId("t_type_bundle").setLabel("Bundle").setEmoji("📦").setStyle(ButtonStyle.Primary),
            new ButtonBuilder().setCustomId("t_type_dino").setLabel("Dino").setEmoji("🦖").setStyle(ButtonStyle.Primary),
            cancelButton()
        )]
    };
}

async function speciesScreen(st, warning) {
    const species = (await getSpeciesSheets()).slice(0, 25); // limite do Discord

    const embed = new EmbedBuilder()
        .setTitle(st.mode === "bundle" ? "📦 Bundle" : "🦖 Dino")
        .setDescription(species.length ? "Choose the species:" : "No species available yet.");
    if (warning) embed.setFooter({ text: warning });

    const components = [];
    if (species.length) {
        components.push(new ActionRowBuilder().addComponents(
            new StringSelectMenuBuilder()
                .setCustomId("t_species")
                .setPlaceholder("Choose the species")
                .addOptions(species.map(s => ({ label: s, value: s })))
        ));
    }
    components.push(new ActionRowBuilder().addComponents(cancelButton()));

    return { content: "", embeds: [embed], components };
}

async function bundleScreen(st) {
    const bundles = await getBundlesForSpecies(st.species);
    if (bundles <= 0) return speciesScreen(st, `⚠️ ${st.species} has no bundles available.`);

    const prices = await getPrices(st.species);
    const max = Math.min(bundles, 25);

    const menu = new StringSelectMenuBuilder()
        .setCustomId("t_qty")
        .setPlaceholder("How many bundles?")
        .addOptions(Array.from({ length: max }, (_, n) => ({
            label: `${n + 1} bundle${n ? "s" : ""}`,
            value: String(n + 1)
        })));

    return {
        content: "",
        embeds: [new EmbedBuilder()
            .setTitle(`📦 Bundle — ${st.species}`)
            .setDescription(
                `Available: **${bundles}** ${stockBall(bundles)}\n` +
                `Price per bundle: **${priceText(prices, st.payment, "bundle")}**`
            )],
        components: [
            new ActionRowBuilder().addComponents(menu),
            new ActionRowBuilder().addComponents(backButton(), cancelButton())
        ]
    };
}

function bundleContinueScreen(st) {
    return {
        content: "",
        embeds: [new EmbedBuilder()
            .setTitle("📦 Bundle added")
            .setDescription("Would you like to make another purchase?")],
        components: [new ActionRowBuilder().addComponents(
            new ButtonBuilder().setCustomId("t_buy_more").setLabel("Yes, buy more").setStyle(ButtonStyle.Primary),
            new ButtonBuilder().setCustomId("t_no_more").setLabel("No, continue").setStyle(ButtonStyle.Success),
            cancelButton()
        )]
    };
}

async function colorScreen(st) {
    const stock = await getColorStock(st.species);
    const prices = await getPrices(st.species);

    const lines = stock.map(c => `**${c.color}** — ♂ ${c.male} · ♀ ${c.female}`);
    const available = stock.filter(c => c.male > 0 || c.female > 0).slice(0, 25);

    const embed = new EmbedBuilder()
        .setTitle(`🦖 ${st.species} — Stock by color`)
        .setDescription(lines.length ? lines.join("\n") : "No colors registered.")
        .setFooter({ text: `Pair: ${priceText(prices, st.payment, "pair")} • Dino: ${priceText(prices, st.payment, "dino")}` });

    const components = [];
    if (available.length) {
        components.push(new ActionRowBuilder().addComponents(
            new StringSelectMenuBuilder()
                .setCustomId("t_color")
                .setPlaceholder("Choose the color")
                .addOptions(available.map(c => ({
                    label: c.color.slice(0, 100),
                    value: c.color.slice(0, 100),
                    description: `♂ ${c.male} · ♀ ${c.female}`
                })))
        ));
    } else {
        embed.setDescription((lines.join("\n") || "No colors registered.") + "\n\n⚠️ No stock available.");
    }
    components.push(new ActionRowBuilder().addComponents(backButton(), cancelButton()));

    return { content: "", embeds: [embed], components };
}

async function genderScreen(st) {
    const stock = await getColorStock(st.species);
    const prices = await getPrices(st.species);
    const c = stock.find(x => x.color === st.color);
    const pairsInCart = st.cart
        .filter(x => x.type === "pair" && x.species === st.species && x.color === st.color)
        .reduce((sum, x) => sum + x.qty, 0);
    const m = (c?.male ?? 0) - cartQty(st, st.species, st.color, "m") - pairsInCart;
    const f = (c?.female ?? 0) - cartQty(st, st.species, st.color, "f") - pairsInCart;

    return {
        content: "",
        embeds: [new EmbedBuilder()
            .setTitle(`🦖 ${st.species} — ${st.color}`)
            .setDescription(
                "Choose the gender:\n" +
                `Pair (♂ + ♀): **${priceText(prices, st.payment, "pair")}**\n` +
                `Single dino: **${priceText(prices, st.payment, "dino")}**`
            )],
        components: [new ActionRowBuilder().addComponents(
            new ButtonBuilder().setCustomId("t_gender_m").setLabel(`♂ Male (${Math.max(m, 0)})`)
                .setStyle(ButtonStyle.Primary).setDisabled(m <= 0),
            new ButtonBuilder().setCustomId("t_gender_f").setLabel(`♀ Female (${Math.max(f, 0)})`)
                .setStyle(ButtonStyle.Primary).setDisabled(f <= 0),
            new ButtonBuilder().setCustomId("t_gender_pair")
                .setLabel(`Pair (♂ + ♀)`)
                .setStyle(ButtonStyle.Success).setDisabled(m <= 0 || f <= 0),
            backButton(),
            cancelButton()
        )]
    };
}

// Texto simples (sem markdown) de uma linha do carrinho — usado nos menus
function cartItemLabel(it) {
    let label;
    if (it.type === "bundle") label = `${it.species} — Bundle ×${it.qty}`;
    else if (it.type === "pair") label = `${it.species} — ${it.color} Pair (♂ + ♀) ×${it.qty}`;
    else label = `${it.species} — ${it.color} ${it.gender === "m" ? "♂ Male" : "♀ Female"} ×${it.qty}`;
    return label.slice(0, 100);
}

// Resumo do carrinho (agora com o botão "Remove item")
async function summaryScreen(st) {
    const { lines, total, errors } = await calcCart(st);

    const embed = new EmbedBuilder()
        .setTitle("🧾 Your order")
        .setDescription(`${lines.join("\n")}\n\n**Total: ${money(st.payment, total)}**`)
        .setFooter({ text: `Payment: ${PAY_LABEL[st.payment]}` });

    const row1 = new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("t_buy_more").setLabel("Add more purchases").setEmoji("➕").setStyle(ButtonStyle.Primary),
        // [PARTE 2] Remover apenas 1 item do carrinho
        new ButtonBuilder().setCustomId("t_remove").setLabel("Remove item").setEmoji("➖").setStyle(ButtonStyle.Secondary).setDisabled(st.cart.length === 0),
        new ButtonBuilder().setCustomId("t_clear").setLabel("Clear all").setEmoji("🗑️").setStyle(ButtonStyle.Secondary)
    );
    const row2 = new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("t_finish").setLabel("Complete order").setEmoji("✅").setStyle(ButtonStyle.Success).setDisabled(errors.length > 0),
        cancelButton()
    );

    return { content: "", embeds: [embed], components: [row1, row2] };
}

// [PARTE 2] Ecrã para escolher qual item remover do carrinho
function removeScreen(st) {
    const items = st.cart.slice(0, 25); // limite do Discord

    const menu = new StringSelectMenuBuilder()
        .setCustomId("t_remove_pick")
        .setPlaceholder("Choose the item to remove")
        .addOptions(items.map((it, idx) => ({
            label: cartItemLabel(it),
            value: String(idx),
            description: "Removes only this entry from your cart"
        })));

    return {
        content: "",
        embeds: [new EmbedBuilder()
            .setTitle("➖ Remove an item")
            .setDescription("Select the dino (and color) you want to remove.\nOnly that entry is removed — the rest of your cart stays untouched.")],
        components: [
            new ActionRowBuilder().addComponents(menu),
            new ActionRowBuilder().addComponents(
                new ButtonBuilder().setCustomId("t_summary").setLabel("⬅️ Back").setStyle(ButtonStyle.Secondary),
                cancelButton()
            )
        ]
    };
}

// [PARTE 3] Estado da order lido do embed (funciona mesmo depois de reiniciar o bot)
function orderState(embed) {
    const claim = getEmbedField(embed, F_CLAIM) || "";
    const color = getEmbedField(embed, F_COLOR);
    const pin = getEmbedField(embed, F_PIN);
    const status = getEmbedField(embed, F_STATUS) || "";
    return {
        claimerId: claim.match(/<@!?(\d+)>/)?.[1] || null,
        fridgeSet: !!color && color !== PIN_PLACEHOLDER,
        pinSet: !!pin && pin !== PIN_PLACEHOLDER,
        ready: status.includes("Ready") || status.includes("Completed"), // pronta (ou já recolhida)
        collected: status.includes("Completed")                         // cliente já recolheu → fridge livre
    };
}

// [PARTE 3] Botões da mensagem da order (canal orders-logs)
function staffOrderRow({ claimerId = null, fridgeSet = false, pinSet = false, ready = false, collected = false } = {}) {
    const claimed = !!claimerId;
    return new ActionRowBuilder().addComponents(
        // Claim / Release (alterna)
        new ButtonBuilder().setCustomId("o_claim").setLabel(claimed ? "Release" : "Claim")
            .setEmoji(claimed ? "↩️" : "🙋").setStyle(claimed ? ButtonStyle.Secondary : ButtonStyle.Primary).setDisabled(ready),
        // Escolher cor + número do fridge (só depois de claim)
        new ButtonBuilder().setCustomId("o_fridge").setLabel(fridgeSet ? "Change fridge" : "Set fridge")
            .setEmoji("🧊").setStyle(ButtonStyle.Primary).setDisabled(!claimed || ready),
        // Definir o PIN
        new ButtonBuilder().setCustomId("o_pin").setLabel(pinSet ? "Change PIN" : "Set PIN")
            .setEmoji("🔑").setStyle(ButtonStyle.Primary).setDisabled(!claimed || ready),
        // Terminar (só com claim + fridge + PIN)
        new ButtonBuilder().setCustomId("o_ready").setLabel("Order ready")
            .setEmoji("✅").setStyle(ButtonStyle.Success).setDisabled(!(claimed && fridgeSet && pinSet) || ready),
        // Collected: o cliente já recolheu → liberta o fridge (só depois de "ready")
        new ButtonBuilder().setCustomId("o_collected").setLabel("Collected")
            .setEmoji("📤").setStyle(ButtonStyle.Secondary).setDisabled(!ready || collected)
    );
}

// Helpers para ler/escrever campos do embed da order
function getEmbedField(embed, name) {
    return embed?.fields?.find(f => f.name === name)?.value ?? null;
}
function setEmbedField(builder, name, value) {
    const idx = (builder.data.fields || []).findIndex(f => f.name === name);
    const field = { name, value, inline: true };
    if (idx === -1) builder.addFields(field);
    else builder.spliceFields(idx, 1, field);
    return builder;
}

/* ---------- Abrir ticket ---------- */
async function openTicket(i) {
    await i.deferReply({ flags: MessageFlags.Ephemeral });

    const existing = i.guild.channels.cache.find(c => c.topic === `ticket:${i.user.id}`);
    if (existing) return i.editReply(`🔴 You already have an open ticket: ${existing}`);

    const allow = [
        PermissionFlagsBits.ViewChannel,
        PermissionFlagsBits.SendMessages,
        PermissionFlagsBits.ReadMessageHistory
    ];

    const parentId = TICKET_CATEGORY_ID || i.channel.parentId || undefined;
    const channel = await i.guild.channels.create({
        name: `ticket-${safeName(i.user.username)}`,
        topic: `ticket:${i.user.id}`,
        parent: parentId,
        permissionOverwrites: [
            { id: i.guild.id, deny: [PermissionFlagsBits.ViewChannel] },
            { id: i.user.id, allow },
            { id: STAFF_ROLE_ID, allow },
            { id: client.user.id, allow: [...allow, PermissionFlagsBits.ManageChannels, PermissionFlagsBits.EmbedLinks] }
        ]
    });

    const st = { userId: i.user.id, payment: null, mode: null, species: null, color: null, cart: [] };
    tickets.set(channel.id, st);

    const screen = paymentScreen();
    screen.content = `${i.user}, welcome to your ticket!`;
    await channel.send(screen);

    return i.editReply(`🟢 Ticket created: ${channel}`);
}

/* ---------- Descontar stock após validar todo o carrinho ---------- */
async function deductStockForCart(st) {
    // Agrupa o consumo por espécie/cor para validar e atualizar cada linha apenas uma vez.
    const deltas = new Map();
    const sheetCache = new Map();

    async function loadSpecies(species) {
        if (!sheetCache.has(species)) {
            const res = await sheetsClient.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: `${q(species)}!A:C`
            });
            const rows = res.data.values || [];
            sheetCache.set(species, rows);
        }
        return sheetCache.get(species);
    }

    function addDelta(species, rowIndex, gender, qty) {
        const key = `${species}::${rowIndex}`;
        if (!deltas.has(key)) deltas.set(key, { species, rowIndex, male: 0, female: 0 });
        if (gender === "m") deltas.get(key).male += qty;
        else deltas.get(key).female += qty;
    }

    for (const item of st.cart) {
        const rows = await loadSpecies(item.species);
        const colorRows = rows.slice(1);
        if (item.type === "bundle") {
            for (let idx = 0; idx < colorRows.length; idx++) {
                const row = colorRows[idx];
                if (!(row?.[0] || "").trim()) continue;
                addDelta(item.species, idx + 1, "m", item.qty);
                addDelta(item.species, idx + 1, "f", item.qty);
            }
        } else {
            const idx = colorRows.findIndex(r => (r?.[0] || "").trim() === item.color);
            if (idx < 0) throw new Error(`The color "${item.color}" no longer exists in ${item.species}.`);
            if (item.type === "pair") {
                addDelta(item.species, idx + 1, "m", item.qty);
                addDelta(item.species, idx + 1, "f", item.qty);
            } else {
                addDelta(item.species, idx + 1, item.gender, item.qty);
            }
        }
    }

    const updates = [];
    for (const delta of deltas.values()) {
        const rows = await loadSpecies(delta.species);
        const row = rows[delta.rowIndex];
        const color = (row?.[0] || "").trim();
        const male = parseInt(row?.[1] || "0", 10) || 0;
        const female = parseInt(row?.[2] || "0", 10) || 0;
        const newMale = male - delta.male;
        const newFemale = female - delta.female;
        if (newMale < 0 || newFemale < 0) {
            throw new Error(`Not enough stock for ${delta.species} — ${color}. Available: ♂ ${male}, ♀ ${female}.`);
        }
        updates.push({
            range: `${q(delta.species)}!A${delta.rowIndex + 1}:C${delta.rowIndex + 1}`,
            values: [[color, newMale, newFemale]]
        });
    }

    if (updates.length) {
        await sheetsClient.spreadsheets.values.batchUpdate({
            spreadsheetId: SHEET_ID,
            resource: { valueInputOption: "USER_ENTERED", data: updates }
        });
    }
    await updateStockSheet();
}

/* ---------- Pedido para a staff (tudo no canal orders-logs) ---------- */
async function createStaffOrder(i, st) {
    const { lines, total, errors } = await calcCart(st);

    if (errors.length || st.cart.length === 0) {
        return i.followUp({
            content: "🔴 Could not complete the order (empty cart or items without a price). Please contact staff.",
            flags: MessageFlags.Ephemeral
        });
    }

    // [PARTE 3] Canal onde ficam TODAS as orders (verificado ANTES de descontar stock)
    await i.guild.channels.fetch();
    const logsChannel = findTextChannel(i.guild, ORDERS_LOG_CHANNEL_NAME);
    const canSend = logsChannel?.permissionsFor(client.user)?.has([
        PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages, PermissionFlagsBits.EmbedLinks
    ]);
    if (!logsChannel || !canSend) {
        console.error('Orders log channel "⛔│「orders-logs」" not found or the bot cannot write in it.');
        return i.followUp({
            content: "🔴 Could not complete the order right now. Please contact staff.",
            flags: MessageFlags.Ephemeral
        });
    }

    // Desconta o stock (se falhar, o erro segue para o handler geral e nada é enviado)
    await deductStockForCart(st);

    const embed = new EmbedBuilder()
        .setTitle("🧾 New order")
        .setDescription(lines.join("\n"))
        .addFields(
            { name: F_CUSTOMER, value: `${i.user} (${i.user.username})`, inline: true },
            { name: "Type", value: st.mode === "bundle" ? "Bundle" : "Dino", inline: true },
            { name: "Payment", value: PAY_LABEL[st.payment], inline: true },
            { name: "Total", value: `**${money(st.payment, total)}**`, inline: true },
            { name: F_TICKET, value: `${i.channel}`, inline: true },
            // Campos INTERNOS (o cliente nunca vê este canal)
            { name: F_CLAIM, value: UNCLAIMED, inline: true },
            { name: F_COLOR, value: PIN_PLACEHOLDER, inline: true },
            { name: F_NUMBER, value: PIN_PLACEHOLDER, inline: true },
            { name: F_PIN, value: PIN_PLACEHOLDER, inline: true },
            { name: F_STATUS, value: "🟡 Pending", inline: true }
        )
        .setTimestamp();

    const orderMsg = await logsChannel.send({
        content: `<@&${STAFF_ROLE_ID}>`,
        embeds: [embed],
        components: [staffOrderRow()]
    });

    // [PARTE 3] Regista a order na folha TradingPost
    try {
        await saveTradingPostOrder({ messageId: orderMsg.id, userId: i.user.id, username: i.user.username });
    } catch (err) {
        console.error("Could not save the order in the TradingPost sheet:", err);
        await logOrderEvent(orderMsg, "⚠️ Could not save this order in the **TradingPost** sheet. Please add it manually.");
    }

    tickets.delete(i.channelId);

    await i.editReply({
        content: "✅ Order sent! Staff will talk to you here in this ticket.\nYou'll also receive a DM when your order is ready — please keep your DMs open.",
        embeds: [new EmbedBuilder()
            .setTitle("🧾 Your order")
            .setDescription(`${lines.join("\n")}\n\n**Total: ${money(st.payment, total)}**`)
            .setFooter({ text: `Payment: ${PAY_LABEL[st.payment]}` })],
        components: [closeRow()]
    });
}

/* ---------- Handler de botões e menus ---------- */
async function handleTicketComponent(i) {
    const id = i.customId;
    if (!id.startsWith("t_")) return;

    if (id === "t_open") return openTicket(i);

    // Fechar: só staff
    if (id === "t_close") {
        if (!isStaff(i)) {
            return i.reply({ content: "Only staff can close this.", flags: MessageFlags.Ephemeral });
        }
        await i.reply("🔒 Closing this channel in 5 seconds...");
        tickets.delete(i.channelId);
        setTimeout(() => i.channel.delete().catch(console.error), 5000);
        return;
    }

    const st = tickets.get(i.channelId);
    if (!st) {
        return i.reply({
            content: "This ticket is no longer active (the bot may have restarted). Ask a staff member to close it and open a new one.",
            flags: MessageFlags.Ephemeral
        });
    }

    if (i.user.id !== st.userId) {
        return i.reply({ content: "Only the person who opened the ticket can use this.", flags: MessageFlags.Ephemeral });
    }

    if (id === "t_cancel") {
        await i.reply("❌ Ticket cancelled. Closing in 5 seconds...");
        tickets.delete(i.channelId);
        setTimeout(() => i.channel.delete().catch(console.error), 5000);
        return;
    }

    // A partir daqui editamos a mensagem original (evita o limite de 3s)
    await i.deferUpdate();

    // 1) Pagamento
    if (id === "t_pay_tek" || id === "t_pay_paypal") {
        st.payment = id === "t_pay_tek" ? "tek" : "paypal";
        return i.editReply(typeScreen(st));
    }

    // 2) Bundle ou Dino
    if (id === "t_type_bundle" || id === "t_type_dino") {
        st.mode = id === "t_type_bundle" ? "bundle" : "dino";
        return i.editReply(await speciesScreen(st));
    }

    // Adicionar mais compras / voltar ao ecrã de espécies
    if (id === "t_more" || id === "t_buy_more") {
        return i.editReply(typeScreen(st));
    }
    if (id === "t_no_more") {
        return i.editReply(await summaryScreen(st));
    }

    // 3) Espécie escolhida
    if (id === "t_species") {
        st.species = i.values[0];
        return i.editReply(st.mode === "bundle" ? await bundleScreen(st) : await colorScreen(st));
    }

    // 4a) Quantidade de bundles → vai direto para o resumo
    if (id === "t_qty") {
        const qty = parseInt(i.values[0], 10);
        const bundles = await getBundlesForSpecies(st.species);
        if (qty > bundles) return i.editReply(await bundleScreen(st));

        const existingBundle = st.cart.find(x => x.type === "bundle" && x.species === st.species);
        if (existingBundle) existingBundle.qty += qty;
        else st.cart.push({ type: "bundle", species: st.species, qty });
        return i.editReply(bundleContinueScreen(st));
    }

    // 4b) Cor escolhida → sexo
    if (id === "t_color") {
        st.color = i.values[0];
        return i.editReply(await genderScreen(st));
    }

    // 5) Sexo escolhido → adiciona ao carrinho
    if (id === "t_gender_pair") {
        const stock = await getColorStock(st.species);
        const c = stock.find(x => x.color === st.color);
        const alreadyPairs = st.cart
            .filter(x => x.type === "pair" && x.species === st.species && x.color === st.color)
            .reduce((sum, x) => sum + x.qty, 0);
        const maleInCart = cartQty(st, st.species, st.color, "m");
        const femaleInCart = cartQty(st, st.species, st.color, "f");
        if (!c || c.male - maleInCart - alreadyPairs < 1 || c.female - femaleInCart - alreadyPairs < 1) {
            return i.editReply(await genderScreen(st));
        }
        const existingPair = st.cart.find(x => x.type === "pair" && x.species === st.species && x.color === st.color);
        if (existingPair) existingPair.qty += 1;
        else st.cart.push({ type: "pair", species: st.species, color: st.color, qty: 1 });
        return i.editReply(await summaryScreen(st));
    }

    if (id === "t_gender_m" || id === "t_gender_f") {
        const g = id === "t_gender_m" ? "m" : "f";
        const stock = await getColorStock(st.species);
        const c = stock.find(x => x.color === st.color);
        const have = (g === "m" ? c?.male : c?.female) ?? 0;
        const pairQty = st.cart
            .filter(x => x.type === "pair" && x.species === st.species && x.color === st.color)
            .reduce((sum, x) => sum + x.qty, 0);

        if (cartQty(st, st.species, st.color, g) + pairQty + 1 > have) {
            return i.editReply(await genderScreen(st));
        }

        const existing = st.cart.find(x =>
            x.type === "dino" && x.species === st.species && x.color === st.color && x.gender === g);
        if (existing) existing.qty += 1;
        else st.cart.push({ type: "dino", species: st.species, color: st.color, gender: g, qty: 1 });

        return i.editReply(await summaryScreen(st));
    }

    // Voltar ao resumo do carrinho
    if (id === "t_summary") {
        return i.editReply(st.cart.length ? await summaryScreen(st) : typeScreen(st));
    }

    // [PARTE 2] Abrir o ecrã de remoção de um item
    if (id === "t_remove") {
        if (!st.cart.length) return i.editReply(typeScreen(st));
        return i.editReply(removeScreen(st));
    }

    // [PARTE 2] Item escolhido → remove apenas essa linha do carrinho
    if (id === "t_remove_pick") {
        const idx = parseInt(i.values[0], 10);
        const item = st.cart[idx];
        if (!item) return i.editReply(await summaryScreen(st)); // carrinho mudou entretanto

        const label = cartItemLabel(item);
        st.cart.splice(idx, 1);

        // Carrinho vazio → volta ao início da compra
        if (st.cart.length === 0) {
            st.mode = null;
            const screen = typeScreen(st);
            screen.content = `🗑️ Removed: **${label}**\nYour cart is now empty.`;
            return i.editReply(screen);
        }

        const screen = await summaryScreen(st);
        screen.content = `🗑️ Removed: **${label}**`;
        return i.editReply(screen);
    }

    // Limpar carrinho
    if (id === "t_clear") {
        st.cart = [];
        st.mode = null;
        return i.editReply(typeScreen(st));
    }

    // Finalizar → envia a order para a staff
    if (id === "t_finish") {
        if (!st.cart.length) return i.editReply(typeScreen(st));
        return createStaffOrder(i, st);
    }
}

/* ===========================================================
   [PARTE 3] ORDERS (canal orders-logs): claim, fridge, PIN, ready
   Nada daqui é visível ao cliente.
=========================================================== */

// Escreve uma linha de log por baixo da order (resposta sem pingar ninguém)
function logOrderEvent(message, text) {
    return message.reply({ content: text, allowedMentions: { parse: [] } })
        .catch(err => console.warn("Could not write the order log:", err.message));
}

// Quem pode mexer na order: quem fez claim (ou alguém com Manage Server)
function canManageOrder(i, claimerId) {
    return claimerId === i.user.id || !!i.memberPermissions?.has(PermissionFlagsBits.ManageGuild);
}

// Valida claim/estado antes de uma ação; devolve true se a ação pode continuar
async function guardOrderAction(i, state) {
    if (state.ready) {
        await i.reply({ content: "This order is already completed.", flags: MessageFlags.Ephemeral });
        return false;
    }
    if (!state.claimerId) {
        await i.reply({ content: "Claim this order first.", flags: MessageFlags.Ephemeral });
        return false;
    }
    if (!canManageOrder(i, state.claimerId)) {
        await i.reply({ content: `This order was claimed by <@${state.claimerId}>. Only they can do this.`, flags: MessageFlags.Ephemeral });
        return false;
    }
    return true;
}

async function handleOrderComponent(i) {
    const id = i.customId;

    if (!isStaff(i)) {
        return i.reply({ content: "Only staff can manage orders.", flags: MessageFlags.Ephemeral });
    }

    /* ---------- Menus efémeros: cor → número do fridge (só mostra os LIVRES) ---------- */
    // customId: o_fcolor|<idMensagemOrder>
    if (id.startsWith("o_fcolor|")) {
        const msgId = id.split("|")[1];
        const color = i.values[0];
        await i.deferUpdate();

        const fridge = (await getFridges()).find(f => f.color === color);
        const free = fridge ? fridge.numbers.filter(n => isFridgeFree(fridge, n, msgId)) : [];
        if (!free.length) {
            return i.editReply({ content: "🔴 There are no free fridges of that color anymore.", components: [] });
        }
        const menu = new StringSelectMenuBuilder()
            .setCustomId(`o_fnum|${msgId}|${color}`)
            .setPlaceholder("Choose the fridge number")
            .addOptions(free.slice(0, 25).map(n => ({ label: `#${padNum(n)}`, value: String(n) })));
        return i.editReply({
            content: `🧊 **${color}** — choose a free number:`,
            components: [new ActionRowBuilder().addComponents(menu)]
        });
    }

    // customId: o_fnum|<idMensagemOrder>|<cor>
    if (id.startsWith("o_fnum|")) {
        const [, msgId, color] = id.split("|");
        const num = parseInt(i.values[0], 10);
        const number = `#${padNum(num)}`;
        await i.deferUpdate();

        const orderMsg = await i.channel.messages.fetch(msgId).catch(() => null);
        if (!orderMsg) return i.editReply({ content: "🔴 Could not find the order message.", components: [] });

        const state = orderState(orderMsg.embeds[0]);
        if (state.ready || !state.claimerId || !canManageOrder(i, state.claimerId)) {
            return i.editReply({ content: "🔴 You can't change this order (not claimed by you, or already completed).", components: [] });
        }

        // Marca o fridge como OCUPADO (e liberta o anterior desta order, se houver)
        const customer = (getEmbedField(orderMsg.embeds[0], F_CUSTOMER) || "").match(/\(([^)]*)\)\s*$/)?.[1] || "";
        try {
            await assignFridge(orderMsg.id, customer, color, num);
        } catch (err) {
            return i.editReply({ content: `🔴 ${err.message}`, components: [] });
        }

        const updated = EmbedBuilder.from(orderMsg.embeds[0]);
        setEmbedField(updated, F_COLOR, color);
        setEmbedField(updated, F_NUMBER, number);
        await orderMsg.edit({ embeds: [updated], components: [staffOrderRow({ ...state, fridgeSet: true })] });

        await updateTradingPostRow(orderMsg.id, { color, number }).catch(err => console.error("TradingPost update error:", err));
        await logOrderEvent(orderMsg, `🧊 ${i.user} set the fridge to **${color} ${number}** (now occupied).`);
        return i.editReply({ content: `🟢 Fridge set to **${color} ${number}**.`, components: [] });
    }

    /* ---------- Botões da mensagem da order ---------- */
    const embed = i.message.embeds[0];
    const state = orderState(embed);

    // Claim / Release
    if (id === "o_claim") {
        if (state.ready) {
            return i.reply({ content: "This order is already completed.", flags: MessageFlags.Ephemeral });
        }

        // Já tem dono: só o dono (ou Manage Server) pode libertar
        if (state.claimerId) {
            if (!canManageOrder(i, state.claimerId)) {
                return i.reply({ content: `This order was already claimed by <@${state.claimerId}>.`, flags: MessageFlags.Ephemeral });
            }
            await i.deferUpdate();
            const updated = EmbedBuilder.from(embed);
            setEmbedField(updated, F_CLAIM, UNCLAIMED);
            setEmbedField(updated, F_STATUS, "🟡 Pending");
            await i.editReply({ embeds: [updated], components: [staffOrderRow({ ...state, claimerId: null })] });
            await updateTradingPostRow(i.message.id, { status: "Pending", staff: "" }).catch(err => console.error("TradingPost update error:", err));
            return logOrderEvent(i.message, `↩️ ${i.user} released this order.`);
        }

        // Sem dono: quem clicou fica com a order
        await i.deferUpdate();
        const updated = EmbedBuilder.from(embed);
        setEmbedField(updated, F_CLAIM, `${i.user}`);
        setEmbedField(updated, F_STATUS, "🔵 In progress");
        await i.editReply({ embeds: [updated], components: [staffOrderRow({ ...state, claimerId: i.user.id })] });
        await updateTradingPostRow(i.message.id, { status: "In progress", staff: i.user.username }).catch(err => console.error("TradingPost update error:", err));
        return logOrderEvent(i.message, `👤 ${i.user} claimed this order.`);
    }

    // Set fridge → dropdown com as cores definidas pela staff
    if (id === "o_fridge") {
        if (!(await guardOrderAction(i, state))) return;
        await i.deferReply({ flags: MessageFlags.Ephemeral });

        const allFridges = await getFridges();
        if (!allFridges.length) {
            return i.editReply("🔴 No fridges configured yet. Use `/addfridges` first.");
        }
        // Só aparecem cores que ainda tenham pelo menos 1 fridge livre
        const fridges = allFridges.filter(f => f.numbers.some(n => isFridgeFree(f, n, i.message.id)));
        if (!fridges.length) {
            return i.editReply("🔴 All fridges are occupied. Add more with `/addfridges`, or mark a finished order as **Collected** to free one.");
        }
        const menu = new StringSelectMenuBuilder()
            .setCustomId(`o_fcolor|${i.message.id}`)
            .setPlaceholder("Choose the fridge color")
            .addOptions(fridges.slice(0, 25).map(f => ({
                label: f.color.slice(0, 100),
                value: f.color,
                description: `${f.numbers.filter(n => isFridgeFree(f, n, i.message.id)).length} free of ${f.numbers.length}`
            })));
        return i.editReply({
            content: "🧊 Choose the fridge color:",
            components: [new ActionRowBuilder().addComponents(menu)]
        });
    }

    // Set PIN → modal
    if (id === "o_pin") {
        if (!(await guardOrderAction(i, state))) return;
        const modal = new ModalBuilder()
            .setCustomId("o_pin_modal")
            .setTitle("Fridge PIN")
            .addComponents(new ActionRowBuilder().addComponents(
                new TextInputBuilder()
                    .setCustomId("pin_value")
                    .setLabel("PIN of the fridge")
                    .setStyle(TextInputStyle.Short)
                    .setMinLength(3)
                    .setMaxLength(12)
                    .setRequired(true)
            ));
        return i.showModal(modal);
    }

    // Collected → o cliente já recolheu a encomenda: liberta o fridge
    if (id === "o_collected") {
        if (!state.ready || state.collected) {
            return i.reply({ content: "This order can't be marked as collected.", flags: MessageFlags.Ephemeral });
        }
        if (!canManageOrder(i, state.claimerId)) {
            return i.reply({ content: `This order was claimed by <@${state.claimerId}>. Only they can do this.`, flags: MessageFlags.Ephemeral });
        }
        await i.deferUpdate();

        try {
            await releaseFridge(i.message.id);
        } catch (err) {
            console.error("Could not free the fridge:", err);
            return i.followUp({ content: "🔴 Could not free the fridge in the sheet. Try again.", flags: MessageFlags.Ephemeral });
        }

        const color = getEmbedField(embed, F_COLOR);
        const number = getEmbedField(embed, F_NUMBER);
        const updated = setEmbedField(EmbedBuilder.from(embed), F_STATUS, "⚪ Completed — collected");
        await i.editReply({ embeds: [updated], components: [staffOrderRow({ ...state, ready: true, collected: true })] });
        await updateTradingPostRow(i.message.id, { status: "Completed" }).catch(err => console.error("TradingPost update error:", err));
        return logOrderEvent(i.message, `📤 ${i.user} marked this order as collected — fridge **${color} ${number}** is free again.`);
    }

    // Order ready → DM ao cliente
    if (id === "o_ready") {
        if (!(await guardOrderAction(i, state))) return;
        if (!state.fridgeSet || !state.pinSet) {
            return i.reply({ content: "🔴 Set the fridge and the PIN first.", flags: MessageFlags.Ephemeral });
        }
        return markOrderReady(i);
    }
}

/* ---------- Modal do PIN ---------- */
async function handleModalSubmit(i) {
    if (i.customId !== "o_pin_modal") return;

    if (!isStaff(i)) {
        return i.reply({ content: "Only staff can set the PIN.", flags: MessageFlags.Ephemeral });
    }

    const embed = i.message.embeds[0];
    const state = orderState(embed);
    if (state.ready || !state.claimerId || !canManageOrder(i, state.claimerId)) {
        return i.reply({ content: "🔴 You can't change this order (not claimed by you, or already completed).", flags: MessageFlags.Ephemeral });
    }

    const pin = i.fields.getTextInputValue("pin_value").trim();
    if (!/^\S{3,12}$/.test(pin)) {
        return i.reply({ content: "🔴 The PIN must have 3–12 characters with no spaces.", flags: MessageFlags.Ephemeral });
    }

    await i.deferUpdate();
    const updated = setEmbedField(EmbedBuilder.from(embed), F_PIN, pin);
    await i.editReply({ embeds: [updated], components: [staffOrderRow({ ...state, pinSet: true })] });
    return logOrderEvent(i.message, `🔑 ${i.user} set the PIN.`); // o PIN não é escrito no log
}

/* ---------- Mensagem "order ready" (imagem do servidor + texto por baixo) ---------- */
// Devolve { embeds, files }. Em cada envio cria-se um novo payload (um anexo não se reutiliza).
function buildReadyPayload({ description, color, number, pin }) {
    const embeds = [];
    const files = [];

    // 1) Imagem do servidor (se o ficheiro existir)
    if (fs.existsSync(DM_BANNER_FILE)) {
        files.push(new AttachmentBuilder(DM_BANNER_FILE, { name: DM_BANNER_NAME }));
        embeds.push(new EmbedBuilder().setColor(SHOP_COLOR).setImage(`attachment://${DM_BANNER_NAME}`));
    } else {
        console.warn(`Banner image not found at ${DM_BANNER_FILE} — sending the DM without it.`);
    }

    // 2) Texto por baixo da imagem: cor, número e PIN do fridge
    embeds.push(new EmbedBuilder()
        .setColor(SHOP_COLOR)
        .setTitle("✅ Your order is ready!")
        .setDescription(`Good news — your order from **TUGAS SHOP** is ready to be collected.\n\n${description || ""}`)
        .addFields(
            { name: "🧊 Fridge Color", value: color, inline: true },
            { name: "🧊 Fridge Number", value: number, inline: true },
            { name: "🔑 PIN", value: pin, inline: true }
        )
        .setFooter({ text: "Thank you for shopping with us! 💚" })
        .setTimestamp());

    return { embeds, files };
}

/* ---------- Order terminada: DM ao cliente ---------- */
async function markOrderReady(i) {
    await i.deferUpdate();

    const base = i.message.embeds[0];
    const state = orderState(base);
    const pin = getEmbedField(base, F_PIN);
    const color = getEmbedField(base, F_COLOR);
    const number = getEmbedField(base, F_NUMBER);
    const ticketId = (getEmbedField(base, F_TICKET) || "").match(/\d+/)?.[0];
    const userId = (getEmbedField(base, F_CUSTOMER) || "").match(/<@!?(\d+)>/)?.[1];

    if (!userId) {
        return i.followUp({ content: "🔴 Could not find the customer for this order.", flags: MessageFlags.Ephemeral });
    }

    // Dados da mensagem enviada ao cliente (em inglês, sem qualquer referência à staff)
    const readyData = { description: base.description, color, number, pin };

    // 1) Tenta enviar DM (imagem + texto)
    let dmOk = false;
    try {
        const user = await client.users.fetch(userId);
        await user.send(buildReadyPayload(readyData));
        dmOk = true;
    } catch (err) {
        console.warn(`Could not DM user ${userId}:`, err.message);
    }

    // 2) Se a DM falhar (DMs fechadas), avisa no ticket do cliente
    let ticketOk = false;
    if (!dmOk && ticketId) {
        const ticketChannel = await i.guild.channels.fetch(ticketId).catch(() => null);
        if (ticketChannel) {
            await ticketChannel.send({
                content: `<@${userId}> your order is ready! (We couldn't send you a DM — please check your privacy settings.)`,
                ...buildReadyPayload(readyData)
            }).then(() => { ticketOk = true; }).catch(() => {});
        }
    }

    // 3) Atualiza a order, a folha TradingPost e o log
    const statusText = dmOk
        ? "🟢 Ready — customer notified by DM"
        : ticketOk ? "🟠 Ready — DM failed, notified in ticket" : "🔴 Ready — customer NOT notified";
    const updated = setEmbedField(EmbedBuilder.from(base), F_STATUS, statusText);
    await i.editReply({ embeds: [updated], components: [staffOrderRow({ ...state, ready: true })] });
    await updateTradingPostRow(i.message.id, { status: "Ready" }).catch(err => console.error("TradingPost update error:", err));
    await logOrderEvent(i.message, `✅ ${i.user} marked this order as ready (${statusText.replace(/^\S+ /, "")}).`);

    // 4) Feedback para a staff
    const feedback = dmOk
        ? "🟢 The customer was notified by DM."
        : ticketOk
            ? "🟠 Could not DM the customer (DMs closed). A message was posted in their ticket instead."
            : "🔴 Could not notify the customer. Please contact them manually.";
    return i.followUp({ content: feedback, flags: MessageFlags.Ephemeral });
}

async function publishTicketPanel(channel) {
    const embed = new EmbedBuilder()
        .setTitle("🎫 Shop")
        .setDescription("Click the button below to open a ticket and place your order.");
    const row = new ActionRowBuilder().addComponents(
        new ButtonBuilder()
            .setCustomId("t_open")
            .setLabel("Open ticket")
            .setEmoji("🎫")
            .setStyle(ButtonStyle.Success)
    );
    await channel.send({ embeds: [embed], components: [row] });
}

async function ensureTicketPanel() {
    const guild = await client.guilds.fetch(GUILD_ID);
    await guild.channels.fetch();
    const channel = guild.channels.cache.find(c =>
        c.isTextBased() && c.name.includes("open-a-ticket"));
    if (!channel) {
        console.warn('Channel "📝│「open-a-ticket」" not found. Create it or check the name/permissions.');
        return;
    }
    TICKET_CATEGORY_ID = channel.parentId || null;
    const messages = await channel.messages.fetch({ limit: 50 });
    const exists = messages.some(m =>
        m.author.id === client.user.id &&
        m.components.some(row => row.components.some(component => component.customId === "t_open")));
    if (!exists) {
        await publishTicketPanel(channel);
        console.log(`Ticket panel posted automatically in #${channel.name}`);
    } else {
        console.log(`Ticket panel already exists in #${channel.name}`);
    }
}

/* -----------------------------------------------------------
   BOT READY
------------------------------------------------------------*/
client.once("clientReady", async () => {
    console.log(`Bot logged in as ${client.user.tag}`);
    await initSheets();
    await registerCommands();
    await ensureTicketPanel();
    await refreshShopInfo(null, true); // [PARTE 1] publica Shop Catalog + How to Buy
    await refreshDashboard();
    setInterval(() => refreshDashboard().catch(err => console.error("Dashboard refresh error:", err)), 30 * 60 * 1000);
});

// Evita que o bot crashe com erros inesperados
client.on("error", console.error);
process.on("unhandledRejection", console.error);

client.login(TOKEN);