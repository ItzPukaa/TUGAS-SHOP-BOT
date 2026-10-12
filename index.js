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
    ChannelType
} from "discord.js";
import { google } from "googleapis";

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
    for (const title of ["Precos", "Stock"]) {
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

const IGNORED_SHEETS = ["Stock", "Precos"];

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

    if (client.isReady()) await refreshDashboard();
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

            if (focusedOption.name === "species") {
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

        // Componentes da dashboard e do sistema de tickets
        if (interaction.isButton() || interaction.isStringSelectMenu()) {
            if (interaction.customId?.startsWith("dash_")) return handleDashboardComponent(interaction);
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
            if (interaction.isMessageComponent()) {
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

async function summaryScreen(st) {
    const { lines, total, errors } = await calcCart(st);

    const embed = new EmbedBuilder()
        .setTitle("🧾 Your order")
        .setDescription(`${lines.join("\n")}\n\n**Total: ${money(st.payment, total)}**`)
        .setFooter({ text: `Payment: ${PAY_LABEL[st.payment]}` });

    const row = new ActionRowBuilder();
    row.addComponents(new ButtonBuilder().setCustomId("t_buy_more").setLabel("Add more purchases")
        .setEmoji("➕").setStyle(ButtonStyle.Primary));
    row.addComponents(
        new ButtonBuilder().setCustomId("t_finish").setLabel("Complete order")
            .setEmoji("✅").setStyle(ButtonStyle.Success).setDisabled(errors.length > 0),
        new ButtonBuilder().setCustomId("t_clear").setLabel("Clear").setEmoji("🗑️").setStyle(ButtonStyle.Secondary),
        cancelButton()
    );

    return { content: "", embeds: [embed], components: [row] };
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

/* ---------- Pedido para a staff ---------- */
async function createStaffOrder(i, st) {
    const { lines, total, errors } = await calcCart(st);

    if (errors.length || st.cart.length === 0) {
        return i.followUp({
            content: "🔴 Could not complete the order (empty cart or items without a price). Please contact staff.",
            flags: MessageFlags.Ephemeral
        });
    }

    const allow = [
        PermissionFlagsBits.ViewChannel,
        PermissionFlagsBits.SendMessages,
        PermissionFlagsBits.ReadMessageHistory
    ];

    await i.guild.channels.fetch();
    const staffCategory = i.guild.channels.cache.find(c =>
        c.type === ChannelType.GuildCategory && c.name === "⛔「STAFF」⛔");

    const staffChannel = await i.guild.channels.create({
        name: `order-${safeName(i.user.username)}`,
        topic: `order:${i.user.id}`,
        parent: staffCategory?.id || STAFF_CATEGORY_ID || undefined,
        permissionOverwrites: [
            { id: i.guild.id, deny: [PermissionFlagsBits.ViewChannel] },
            { id: STAFF_ROLE_ID, allow },
            { id: client.user.id, allow: [...allow, PermissionFlagsBits.ManageChannels, PermissionFlagsBits.EmbedLinks] }
        ]
    });

    if (staffCategory) {
        const siblings = staffCategory.children.cache.filter(c => c.id !== staffChannel.id);
        const lastPosition = siblings.size ? Math.max(...siblings.map(c => c.position)) : staffCategory.position;
        await staffChannel.setPosition(lastPosition + 1).catch(err => console.warn("Could not move the order channel to the end of the STAFF category:", err.message));
    }

    try {
        await deductStockForCart(st);
    } catch (err) {
        await staffChannel.delete().catch(() => {});
        throw err;
    }

    const embed = new EmbedBuilder()
        .setTitle("🧾 New order")
        .setDescription(lines.join("\n"))
        .addFields(
            { name: "Customer", value: `${i.user} (${i.user.username})`, inline: true },
            { name: "Type", value: st.mode === "bundle" ? "Bundle" : "Dino", inline: true },
            { name: "Payment", value: PAY_LABEL[st.payment], inline: true },
            { name: "Total", value: `**${money(st.payment, total)}**`, inline: true },
            { name: "Customer ticket", value: `${i.channel}`, inline: true }
        )
        .setTimestamp();

    await staffChannel.send({
        content: `<@&${STAFF_ROLE_ID}>`,
        embeds: [embed],
        components: [closeRow()]
    });

    tickets.delete(i.channelId);

    await i.editReply({
        content: "✅ Order sent! Staff will talk to you here in this ticket.",
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

    // Limpar carrinho
    if (id === "t_clear") {
        st.cart = [];
        st.mode = null;
        return i.editReply(typeScreen(st));
    }

    // Finalizar
    if (id === "t_finish") {
        return createStaffOrder(i, st);
    }
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
    await refreshDashboard();
    setInterval(() => refreshDashboard().catch(err => console.error("Dashboard refresh error:", err)), 30 * 60 * 1000);
});

// Evita que o bot crashe com erros inesperados
client.on("error", console.error);
process.on("unhandledRejection", console.error);

client.login(TOKEN);