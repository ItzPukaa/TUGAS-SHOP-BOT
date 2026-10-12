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
    console.error("Falta o DISCORD_TOKEN no ficheiro .env");
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
        console.error("GOOGLE_SERVICE_ACCOUNT_JSON não contém JSON válido.", error.message);
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
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: "Precos!A1:E1", valueInputOption: "USER_ENTERED",
        resource: { values: [["Nome", "Tek por Pair/Bundle", "Tek por Extra/Dino", "PayPal por Pair/Bundle", "PayPal por Extra/Dino"]] }
    });
    await sheetsClient.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: "Stock!A1:C1", valueInputOption: "USER_ENTERED",
        resource: { values: [["Nome", "Bundles", "Estado"]] }
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
    const values = [["Nome", "Bundles", "Estado"]];

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
    if (sheetId === null) throw new Error(`Folha "${sheetName}" não encontrada.`);

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
            .setDescription("Ainda não há dinos registados. Usa `/adddino` para criar a primeira espécie.")
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
    let description = lines.length ? lines.join("\n") : "Ainda não há cores. Usa `/addcolor` para adicionar uma cor com stock inicial a zero.";
    if (description.length > 3800) description = description.slice(0, 3750) + "\n… (lista encurtada; usa o menu para selecionar uma cor)";

    const selectedColor = colors.find(c => c.color === dashboardColor);
    const selectedQty = selectedColor ? (dashboardGender === "m" ? selectedColor.male : selectedColor.female) : 0;
    const embed = new EmbedBuilder()
        .setTitle(`📦 TUGAS SHOP — ${dashboardSpecies}`)
        .setDescription(description)
        .addFields(
            { name: "Pair / Bundle — Tek", value: prices?.tek?.bundle == null ? "Não definido" : `${prices.tek.bundle} Tek`, inline: true },
            { name: "Extra — Tek", value: prices?.tek?.dino == null ? "Não definido" : `${prices.tek.dino} Tek`, inline: true },
            { name: "Pair / Bundle — PayPal", value: prices?.paypal?.bundle == null ? "Não definido" : `${prices.paypal.bundle}€`, inline: true },
            { name: "Extra — PayPal", value: prices?.paypal?.dino == null ? "Não definido" : `${prices.paypal.dino}€`, inline: true },
            { name: "Ajustar stock", value: selectedColor ? `Cor: **${selectedColor.color}**\nGénero: **${dashboardGender === "m" ? "♂ Macho" : "♀ Fêmea"}**\nQuantidade atual: **${selectedQty}**` : "Seleciona uma cor para ajustar o stock." }
        )
        .setFooter({ text: `${DASHBOARD_MARKER} • Atualização automática a cada 30 minutos e após alterações de stock` });

    const rows = [];
    rows.push(new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
            .setCustomId("dash_species")
            .setPlaceholder("Seleciona um dino")
            .addOptions(species.slice(0, 25).map(name => ({ label: name.slice(0, 100), value: name, default: name === dashboardSpecies })))
    ));

    if (colors.length) {
        rows.push(new ActionRowBuilder().addComponents(
            new StringSelectMenuBuilder()
                .setCustomId("dash_color")
                .setPlaceholder("Seleciona uma cor")
                .addOptions(colors.slice(0, 25).map(c => ({
                    label: c.color.slice(0, 100),
                    value: c.color,
                    description: `♂ ${c.male} · ♀ ${c.female}`.slice(0, 100),
                    default: c.color === dashboardColor
                })))
        ));
    }

    rows.push(new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("dash_gender_m").setLabel("♂ Macho").setStyle(dashboardGender === "m" ? ButtonStyle.Success : ButtonStyle.Secondary).setDisabled(!selectedColor),
        new ButtonBuilder().setCustomId("dash_gender_f").setLabel("♀ Fêmea").setStyle(dashboardGender === "f" ? ButtonStyle.Success : ButtonStyle.Secondary).setDisabled(!selectedColor)
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
            console.warn("Canal da dashboard não encontrado. Cria um canal com nome como ⛔│「dashboard」.");
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
        console.error("Erro ao atualizar a dashboard de stock:", err);
    } finally {
        dashboardRefreshRunning = false;
    }
}

async function handleDashboardComponent(i) {
    if (!i.customId?.startsWith("dash_")) return false;
    if (!isStaff(i)) {
        await i.reply({ content: "Apenas staff pode alterar ou navegar na dashboard de stock.", flags: MessageFlags.Ephemeral });
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
            await i.reply({ content: "Ação da dashboard desconhecida.", flags: MessageFlags.Ephemeral });
            return true;
        }
        if (!dashboardSpecies || !dashboardColor) {
            await i.reply({ content: "Seleciona primeiro um dino e uma cor.", flags: MessageFlags.Ephemeral });
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
        if (rowIndex < 1) throw new Error(`Não encontrei a cor ${dashboardColor} em ${dashboardSpecies}.`);

        const row = rows[rowIndex];
        const male = parseInt(row[1] || "0", 10) || 0;
        const female = parseInt(row[2] || "0", 10) || 0;
        const current = dashboardGender === "m" ? male : female;
        const next = current + delta;
        if (next < 0) throw new Error(`Stock insuficiente: só existem ${current} unidades para retirar.`);

        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: SHEET_ID,
            range: `${q(dashboardSpecies)}!A${rowIndex + 1}:C${rowIndex + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: { values: [[dashboardColor, dashboardGender === "m" ? next : male, dashboardGender === "f" ? next : female]] }
        });
        await updateStockSheet();
        await refreshDashboard();
    } catch (err) {
        console.error("Erro no controlo de stock da dashboard:", err);
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
            .setDescription("Adicionar nova espécie")
            .addStringOption(o => o.setName("nome").setDescription("Nome da espécie").setRequired(true))
            .addNumberOption(o => o.setName("pair").setDescription("Preço do bundle/pair em Tek").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("extra").setDescription("Preço por dino extra em Tek").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("pair_paypal").setDescription("Preço do bundle/pair em PayPal (€)").setRequired(true).setMinValue(0))
            .addNumberOption(o => o.setName("extra_paypal").setDescription("Preço por dino extra em PayPal (€)").setRequired(true).setMinValue(0)),

        new SlashCommandBuilder()
            .setName("addcolor")
            .setDescription("Adicionar cor a uma espécie")
            .addStringOption(o =>
                o.setName("especie")
                    .setDescription("Espécie")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o => o.setName("color").setDescription("Nome da cor").setRequired(true)),

        new SlashCommandBuilder()
            .setName("stock")
            .setDescription("Ver o stock de todas as espécies"),

        new SlashCommandBuilder()
            .setName("ticketpanel")
            .setDescription("Publicar neste canal o painel para abrir tickets"),

        new SlashCommandBuilder()
            .setName("deldino")
            .setDescription("Eliminar uma espécie (apaga a folha e o preço)")
            .addStringOption(o =>
                o.setName("especie")
                    .setDescription("Espécie a eliminar")
                    .setRequired(true)
                    .setAutocomplete(true)),

        new SlashCommandBuilder()
            .setName("delcolor")
            .setDescription("Eliminar uma cor de uma espécie")
            .addStringOption(o =>
                o.setName("especie")
                    .setDescription("Espécie")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o =>
                o.setName("color")
                    .setDescription("Cor a eliminar")
                    .setRequired(true)
                    .setAutocomplete(true)),

        new SlashCommandBuilder()
            .setName("editdino")
            .setDescription("Editar uma espécie (nome e/ou preços)")
            .addStringOption(o =>
                o.setName("especie")
                    .setDescription("Espécie a editar")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o => o.setName("novo_nome").setDescription("Novo nome da espécie"))
            .addNumberOption(o => o.setName("pair").setDescription("Novo preço Pair / Bundle em Tek").setMinValue(0))
            .addNumberOption(o => o.setName("extra").setDescription("Novo preço por dino extra em Tek").setMinValue(0))
            .addNumberOption(o => o.setName("pair_paypal").setDescription("Novo preço Pair / Bundle (PayPal)").setMinValue(0))
            .addNumberOption(o => o.setName("extra_paypal").setDescription("Novo Extra Female / Dino (PayPal)").setMinValue(0)),

        new SlashCommandBuilder()
            .setName("editcolor")
            .setDescription("Renomear uma cor de uma espécie; stock ajustado pela dashboard")
            .addStringOption(o =>
                o.setName("especie")
                    .setDescription("Espécie")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o =>
                o.setName("color")
                    .setDescription("Cor a editar")
                    .setRequired(true)
                    .setAutocomplete(true))
            .addStringOption(o => o.setName("novo_nome").setDescription("Novo nome da cor"))
    ].map(cmd => cmd.toJSON());

    const rest = new REST({ version: "10" }).setToken(TOKEN);

    await rest.put(
        Routes.applicationGuildCommands(client.user.id, GUILD_ID),
        { body: commands }
    );

    console.log("Slash commands registados!");
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

            if (focusedOption.name === "especie") {
                list = await getSpeciesSheets();
            } else if (focusedOption.name === "color" && ["delcolor", "editcolor"].includes(interaction.commandName)) {
                // Só mostra cores se a espécie já tiver sido escolhida
                const especie = interaction.options.getString("especie");
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
                return interaction.editReply("Ainda não há espécies registadas.");
            }

            const lines = data.map(d => `${stockBall(d.bundles)}  **${d.name}**`);
            let description = lines.join("\n");
            if (description.length > 3800) description = description.slice(0, 3800) + "\n…";

            const embed = new EmbedBuilder()
                .setTitle("📦 Stock")
                .setDescription(description)
                .setFooter({ text: "🟢 Disponível  🟡 2 bundles  🟠 Último bundle  🔴 Sem stock" });

            return interaction.editReply({ embeds: [embed] });
        }

        if (!isStaff(interaction)) {
            return interaction.reply({
                content: "Apenas staff pode usar este comando.",
                flags: MessageFlags.Ephemeral
            });
        }

        /* ---------- /adddino ---------- */
        if (interaction.commandName === "adddino") {
            await interaction.deferReply();

            const nome = interaction.options.getString("nome").trim();
            const pair = interaction.options.getNumber("pair");
            const extra = interaction.options.getNumber("extra");
            const pairPP = interaction.options.getNumber("pair_paypal");
            const extraPP = interaction.options.getNumber("extra_paypal");

            if (!nome || nome.length > 100 || /[\[\]\*\?:\/\\]/.test(nome)) {
                return interaction.editReply("🔴 Nome inválido. Não uses `[ ] * ? : / \\` e máximo 100 caracteres.");
            }

            const existing = await getSpeciesSheets();
            if (existing.some(s => s.toLowerCase() === nome.toLowerCase()) ||
                IGNORED_SHEETS.some(s => s.toLowerCase() === nome.toLowerCase())) {
                return interaction.editReply(`🔴 A espécie **${nome}** já existe.`);
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
                range: "Precos!A:E",
                valueInputOption: "USER_ENTERED",
                resource: { values: [[nome, pair, extra, pairPP, extraPP]] }
            });

            await updateStockSheet();

            return interaction.editReply(`🟢 Espécie **${nome}** criada com sucesso!`);
        }

        /* ---------- /addcolor ---------- */
        if (interaction.commandName === "addcolor") {
            await interaction.deferReply();

            const especie = interaction.options.getString("especie");
            const color = interaction.options.getString("color").trim();

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 A espécie **${especie}** não existe. Escolhe uma da lista.`);
            }
            if (!color || color.length > 100) return interaction.editReply("🔴 Nome de cor inválido (máximo 100 caracteres).");

            const existingColors = await getColorsForSpecies(especie);
            if (existingColors.some(c => c.toLowerCase() === color.toLowerCase())) {
                return interaction.editReply(`🔴 A cor **${color}** já existe em **${especie}**.`);
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

            return interaction.editReply(`🟢 Cor **${color}** adicionada a **${especie}** com stock a zero. Usa a dashboard para adicionar quantidade com os botões.`);
        }

        /* ---------- /deldino ---------- */
        if (interaction.commandName === "deldino") {
            await interaction.deferReply({ flags: MessageFlags.Ephemeral });

            const especie = interaction.options.getString("especie");

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 A espécie **${especie}** não existe.`);
            }

            // Pedido de confirmação (apaga a folha inteira!)
            const buttons = new ActionRowBuilder().addComponents(
                new ButtonBuilder()
                    .setCustomId("confirm_deldino")
                    .setLabel("Eliminar")
                    .setStyle(ButtonStyle.Danger),
                new ButtonBuilder()
                    .setCustomId("cancel_deldino")
                    .setLabel("Cancelar")
                    .setStyle(ButtonStyle.Secondary)
            );

            const msg = await interaction.editReply({
                content: `⚠️ Tens a certeza que queres eliminar **${especie}**?\nIsto apaga a folha com todas as cores e remove o preço. **Não dá para desfazer.**`,
                components: [buttons]
            });

            let click;
            try {
                click = await msg.awaitMessageComponent({
                    filter: i => i.user.id === interaction.user.id,
                    time: 30_000
                });
            } catch {
                return interaction.editReply({ content: "⏱️ Tempo esgotado. Nada foi eliminado.", components: [] });
            }

            if (click.customId !== "confirm_deldino") {
                return click.update({ content: "Cancelado. Nada foi eliminado.", components: [] });
            }

            await click.update({ content: "A eliminar...", components: [] });

            // 1) Apagar a folha da espécie
            const sheetId = await getSheetId(especie);
            if (sheetId === null) throw new Error(`Folha "${especie}" não encontrada.`);

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

            return interaction.editReply(`🟢 Espécie **${especie}** eliminada com sucesso!`);
        }

        /* ---------- /delcolor ---------- */
        if (interaction.commandName === "delcolor") {
            await interaction.deferReply();

            const especie = interaction.options.getString("especie");
            const color = interaction.options.getString("color").trim();

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 A espécie **${especie}** não existe. Escolhe uma da lista.`);
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
                return interaction.editReply(`🔴 A cor **${color}** não existe em **${especie}**.`);
            }

            await deleteRow(especie, rowIndex);
            await updateStockSheet();

            return interaction.editReply(`🟢 Cor **${color}** eliminada de **${especie}**!`);
        }

        /* ---------- /editdino ---------- */
        if (interaction.commandName === "editdino") {
            await interaction.deferReply();

            const especie = interaction.options.getString("especie");
            const novoNomeRaw = interaction.options.getString("novo_nome");
            const novoNome = novoNomeRaw ? novoNomeRaw.trim() : null;
            const pair = interaction.options.getNumber("pair");
            const extra = interaction.options.getNumber("extra");

            const pairPP = interaction.options.getNumber("pair_paypal");
            const extraPP = interaction.options.getNumber("extra_paypal");

            if (!novoNome && pair === null && extra === null && pairPP === null && extraPP === null) {
                return interaction.editReply("🔴 Indica pelo menos uma coisa para alterar: `novo_nome`, `pair`, `extra`, `pair_paypal` ou `extra_paypal`.");
            }

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 A espécie **${especie}** não existe. Escolhe uma da lista.`);
            }

            const renaming = novoNome && novoNome !== especie;

            if (renaming) {
                if (novoNome.length > 100 || /[\[\]\*\?:\/\\]/.test(novoNome)) {
                    return interaction.editReply("🔴 Nome inválido. Não uses `[ ] * ? : / \\` e máximo 100 caracteres.");
                }
                const lower = novoNome.toLowerCase();
                if (species.some(s => s.toLowerCase() === lower && s !== especie) ||
                    IGNORED_SHEETS.some(s => s.toLowerCase() === lower)) {
                    return interaction.editReply(`🔴 Já existe uma espécie com o nome **${novoNome}**.`);
                }
            }

            // Linha na folha Precos
            const precos = await sheetsClient.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: "Precos!A:E"
            });
            const precoRows = precos.data.values || [];
            const precoIndex = precoRows.findIndex(r => (r[0] || "").trim() === especie);
            const current = precoIndex !== -1 ? precoRows[precoIndex] : [especie, 0, 0, "", ""];

            // 1) Renomear a folha
            if (renaming) {
                const sheetId = await getSheetId(especie);
                if (sheetId === null) throw new Error(`Folha "${especie}" não encontrada.`);

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
                pair !== null ? pair : (current[1] ?? 0),
                extra !== null ? extra : (current[2] ?? 0),
                pairPP !== null ? pairPP : (current[3] ?? ""),
                extraPP !== null ? extraPP : (current[4] ?? "")
            ];

            if (precoIndex !== -1) {
                await sheetsClient.spreadsheets.values.update({
                    spreadsheetId: SHEET_ID,
                    range: `Precos!A${precoIndex + 1}:E${precoIndex + 1}`,
                    valueInputOption: "USER_ENTERED",
                    resource: { values: [newRow] }
                });
            } else {
                await sheetsClient.spreadsheets.values.append({
                    spreadsheetId: SHEET_ID,
                    range: "Precos!A:E",
                    valueInputOption: "USER_ENTERED",
                    resource: { values: [newRow] }
                });
            }

            // 3) Atualizar Stock (o nome pode ter mudado)
            if (renaming && dashboardSpecies === especie) dashboardSpecies = novoNome;
            await updateStockSheet();

            return interaction.editReply(
                `🟢 Espécie **${finalName}** atualizada!` +
                (renaming ? `\nNome: ${especie} → **${finalName}**` : "") +
                (pair !== null ? `\nPair: **${pair}**` : "") +
                (extra !== null ? `\nExtra Female: **${extra}**` : "") +
                (pairPP !== null ? `\nPair PayPal: **${pairPP}**` : "") +
                (extraPP !== null ? `\nExtra PayPal: **${extraPP}**` : "")
            );
        }

        /* ---------- /editcolor ---------- */
        if (interaction.commandName === "editcolor") {
            await interaction.deferReply();

            const especie = interaction.options.getString("especie");
            const color = interaction.options.getString("color").trim();
            const novoNomeRaw = interaction.options.getString("novo_nome");
            const novoNome = novoNomeRaw ? novoNomeRaw.trim() : null;
            if (!novoNome) {
                return interaction.editReply("🔴 Indica o `novo_nome` da cor. Para alterar quantidades, usa os botões da dashboard.");
            }

            const species = await getSpeciesSheets();
            if (!species.includes(especie)) {
                return interaction.editReply(`🔴 A espécie **${especie}** não existe. Escolhe uma da lista.`);
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
                return interaction.editReply(`🔴 A cor **${color}** não existe em **${especie}**.`);
            }

            const currentRow = rows[rowIndex];
            const currentName = (currentRow[0] || "").trim();

            if (novoNome && novoNome.toLowerCase() !== currentName.toLowerCase()) {
                const clash = rows.some((r, i) => i > 0 && i !== rowIndex &&
                    (r[0] || "").trim().toLowerCase() === novoNome.toLowerCase());
                if (clash) {
                    return interaction.editReply(`🔴 Já existe uma cor **${novoNome}** em **${especie}**.`);
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

            return interaction.editReply(`🟢 Cor **${currentName}** de **${especie}** renomeada para **${finalName}**. As quantidades são alteradas na dashboard.`);
        }

        /* ---------- /ticketpanel ---------- */
        if (interaction.commandName === "ticketpanel") {
            await publishTicketPanel(interaction.channel);
            return interaction.reply({ content: "🟢 Painel publicado.", flags: MessageFlags.Ephemeral });
        }
    } catch (err) {
        console.error("Erro na interação:", err);

        // Autocomplete não pode receber mensagem de texto
        if (interaction.isAutocomplete()) return;

        const msg = `🔴 Erro: ${err.message}`;
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
            console.error("Não foi possível responder ao erro:", e);
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

// Precos: A=Nome | B=Pair/Bundle Tek | C=Extra/Dino Tek | D=Pair/Bundle PayPal | E=Extra/Dino PayPal
async function getPrices(species) {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: "Precos!A:E"
    });
    const row = (res.data.values || []).find(r => (r[0] || "").trim() === species);
    if (!row) return null;

    const n = v => {
        const x = parseFloat(String(v ?? "").replace(",", "."));
        return isNaN(x) ? null : x;
    };

    return {
        tek: { bundle: n(row[1]), dino: n(row[2]) },
        paypal: { bundle: n(row[3]), dino: n(row[4]) }
    };
}

function priceText(prices, method, kind) {
    const v = prices?.[method]?.[kind];
    return v === null || v === undefined ? "⚠️ não definido" : money(method, v);
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
        const priceKind = it.type === "pair" ? "bundle" : it.type;
        const unit = p ? p[st.payment][priceKind] : null;

        const label = it.type === "bundle"
            ? `📦 **${it.species}** — Bundle ×${it.qty}`
            : it.type === "pair"
                ? `🦖 **${it.species}** — Pair ${it.color} (♂ + ♀) ×${it.qty}`
                : `🦖 **${it.species}** — ${it.color} ${it.gender === "m" ? "♂" : "♀"} ×${it.qty}`;

        if (unit === null || unit === undefined) {
            errors.push(it.species);
            lines.push(`${label} — ⚠️ preço ${PAY_LABEL[st.payment]} não definido`);
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
    new ButtonBuilder().setCustomId("t_cancel").setLabel("Cancelar").setStyle(ButtonStyle.Secondary);

const backButton = () =>
    new ButtonBuilder().setCustomId("t_more").setLabel("⬅️ Voltar").setStyle(ButtonStyle.Secondary);

const closeRow = () =>
    new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("t_close").setLabel("Fechar").setEmoji("🔒").setStyle(ButtonStyle.Danger)
    );

function paymentScreen() {
    return {
        content: "",
        embeds: [new EmbedBuilder().setTitle("💳 Método de pagamento").setDescription("Como queres pagar?")],
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
            .setTitle("🛒 O que queres comprar?")
            .setDescription(`Pagamento: **${PAY_LABEL[st.payment]}**`)],
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
        .setDescription(species.length ? "Escolhe a espécie:" : "Ainda não há espécies disponíveis.");
    if (warning) embed.setFooter({ text: warning });

    const components = [];
    if (species.length) {
        components.push(new ActionRowBuilder().addComponents(
            new StringSelectMenuBuilder()
                .setCustomId("t_species")
                .setPlaceholder("Escolhe a espécie")
                .addOptions(species.map(s => ({ label: s, value: s })))
        ));
    }
    components.push(new ActionRowBuilder().addComponents(cancelButton()));

    return { content: "", embeds: [embed], components };
}

async function bundleScreen(st) {
    const bundles = await getBundlesForSpecies(st.species);
    if (bundles <= 0) return speciesScreen(st, `⚠️ ${st.species} não tem bundles disponíveis.`);

    const prices = await getPrices(st.species);
    const max = Math.min(bundles, 25);

    const menu = new StringSelectMenuBuilder()
        .setCustomId("t_qty")
        .setPlaceholder("Quantos bundles?")
        .addOptions(Array.from({ length: max }, (_, n) => ({
            label: `${n + 1} bundle${n ? "s" : ""}`,
            value: String(n + 1)
        })));

    return {
        content: "",
        embeds: [new EmbedBuilder()
            .setTitle(`📦 Bundle — ${st.species}`)
            .setDescription(
                `Disponíveis: **${bundles}** ${stockBall(bundles)}\n` +
                `Preço por bundle: **${priceText(prices, st.payment, "bundle")}**`
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
            .setTitle("📦 Bundle adicionado")
            .setDescription("Pretendes realizar mais alguma compra?")],
        components: [new ActionRowBuilder().addComponents(
            new ButtonBuilder().setCustomId("t_buy_more").setLabel("Sim, comprar mais").setStyle(ButtonStyle.Primary),
            new ButtonBuilder().setCustomId("t_no_more").setLabel("Não, continuar").setStyle(ButtonStyle.Success),
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
        .setTitle(`🦖 ${st.species} — Stock por cor`)
        .setDescription(lines.length ? lines.join("\n") : "Sem cores registadas.")
        .setFooter({ text: `Preço por dino: ${priceText(prices, st.payment, "dino")}` });

    const components = [];
    if (available.length) {
        components.push(new ActionRowBuilder().addComponents(
            new StringSelectMenuBuilder()
                .setCustomId("t_color")
                .setPlaceholder("Escolhe a cor")
                .addOptions(available.map(c => ({
                    label: c.color.slice(0, 100),
                    value: c.color.slice(0, 100),
                    description: `♂ ${c.male} · ♀ ${c.female}`
                })))
        ));
    } else {
        embed.setDescription((lines.join("\n") || "Sem cores registadas.") + "\n\n⚠️ Sem stock disponível.");
    }
    components.push(new ActionRowBuilder().addComponents(backButton(), cancelButton()));

    return { content: "", embeds: [embed], components };
}

async function genderScreen(st) {
    const stock = await getColorStock(st.species);
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
            .setDescription("Escolhe o sexo:")],
        components: [new ActionRowBuilder().addComponents(
            new ButtonBuilder().setCustomId("t_gender_m").setLabel(`♂ Macho (${Math.max(m, 0)})`)
                .setStyle(ButtonStyle.Primary).setDisabled(m <= 0),
            new ButtonBuilder().setCustomId("t_gender_f").setLabel(`♀ Fêmea (${Math.max(f, 0)})`)
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
        .setTitle("🧾 O teu pedido")
        .setDescription(`${lines.join("\n")}\n\n**Total: ${money(st.payment, total)}**`)
        .setFooter({ text: `Pagamento: ${PAY_LABEL[st.payment]}` });

    const row = new ActionRowBuilder();
    row.addComponents(new ButtonBuilder().setCustomId("t_buy_more").setLabel("Adicionar mais compras")
        .setEmoji("➕").setStyle(ButtonStyle.Primary));
    row.addComponents(
        new ButtonBuilder().setCustomId("t_finish").setLabel("Finalizar pedido")
            .setEmoji("✅").setStyle(ButtonStyle.Success).setDisabled(errors.length > 0),
        new ButtonBuilder().setCustomId("t_clear").setLabel("Limpar").setEmoji("🗑️").setStyle(ButtonStyle.Secondary),
        cancelButton()
    );

    return { content: "", embeds: [embed], components: [row] };
}

/* ---------- Abrir ticket ---------- */
async function openTicket(i) {
    await i.deferReply({ flags: MessageFlags.Ephemeral });

    const existing = i.guild.channels.cache.find(c => c.topic === `ticket:${i.user.id}`);
    if (existing) return i.editReply(`🔴 Já tens um ticket aberto: ${existing}`);

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
    screen.content = `${i.user}, bem-vindo ao teu ticket!`;
    await channel.send(screen);

    return i.editReply(`🟢 Ticket criado: ${channel}`);
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
            if (idx < 0) throw new Error(`A cor "${item.color}" já não existe em ${item.species}.`);
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
            throw new Error(`Stock insuficiente para ${delta.species} — ${color}. Disponível: ♂ ${male}, ♀ ${female}.`);
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
            content: "🔴 Não foi possível finalizar (carrinho vazio ou itens sem preço). Fala com a staff.",
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
        name: `pedido-${safeName(i.user.username)}`,
        topic: `pedido:${i.user.id}`,
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
        await staffChannel.setPosition(lastPosition + 1).catch(err => console.warn("Não foi possível posicionar o canal de pedido no fim da categoria STAFF:", err.message));
    }

    try {
        await deductStockForCart(st);
    } catch (err) {
        await staffChannel.delete().catch(() => {});
        throw err;
    }

    const embed = new EmbedBuilder()
        .setTitle("🧾 Novo pedido")
        .setDescription(lines.join("\n"))
        .addFields(
            { name: "Cliente", value: `${i.user} (${i.user.username})`, inline: true },
            { name: "Tipo", value: st.mode === "bundle" ? "Bundle" : "Dino", inline: true },
            { name: "Pagamento", value: PAY_LABEL[st.payment], inline: true },
            { name: "Total", value: `**${money(st.payment, total)}**`, inline: true },
            { name: "Ticket do cliente", value: `${i.channel}`, inline: true }
        )
        .setTimestamp();

    await staffChannel.send({
        content: `<@&${STAFF_ROLE_ID}>`,
        embeds: [embed],
        components: [closeRow()]
    });

    tickets.delete(i.channelId);

    await i.editReply({
        content: "✅ Pedido enviado! A staff vai falar contigo aqui neste ticket.",
        embeds: [new EmbedBuilder()
            .setTitle("🧾 O teu pedido")
            .setDescription(`${lines.join("\n")}\n\n**Total: ${money(st.payment, total)}**`)
            .setFooter({ text: `Pagamento: ${PAY_LABEL[st.payment]}` })],
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
            return i.reply({ content: "Apenas staff pode fechar.", flags: MessageFlags.Ephemeral });
        }
        await i.reply("🔒 A fechar este canal em 5 segundos...");
        tickets.delete(i.channelId);
        setTimeout(() => i.channel.delete().catch(console.error), 5000);
        return;
    }

    const st = tickets.get(i.channelId);
    if (!st) {
        return i.reply({
            content: "Este ticket já não está ativo (o bot pode ter reiniciado). Pede a um staff para o fechar e abre outro.",
            flags: MessageFlags.Ephemeral
        });
    }

    if (i.user.id !== st.userId) {
        return i.reply({ content: "Só quem abriu o ticket pode usar isto.", flags: MessageFlags.Ephemeral });
    }

    if (id === "t_cancel") {
        await i.reply("❌ Ticket cancelado. A fechar em 5 segundos...");
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
        .setTitle("🎫 Loja")
        .setDescription("Clica no botão abaixo para abrir um ticket e fazer o teu pedido.");
    const row = new ActionRowBuilder().addComponents(
        new ButtonBuilder()
            .setCustomId("t_open")
            .setLabel("Abrir ticket")
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
        console.warn('Canal "📝│「open-a-ticket」" não encontrado. Cria-o ou verifica o nome/permissões.');
        return;
    }
    TICKET_CATEGORY_ID = channel.parentId || null;
    const messages = await channel.messages.fetch({ limit: 50 });
    const exists = messages.some(m =>
        m.author.id === client.user.id &&
        m.components.some(row => row.components.some(component => component.customId === "t_open")));
    if (!exists) {
        await publishTicketPanel(channel);
        console.log(`Painel de tickets publicado automaticamente em #${channel.name}`);
    } else {
        console.log(`Painel de tickets já existe em #${channel.name}`);
    }
}

/* -----------------------------------------------------------
   BOT READY
------------------------------------------------------------*/
client.once("clientReady", async () => {
    console.log(`Bot ligado como ${client.user.tag}`);
    await initSheets();
    await registerCommands();
    await ensureTicketPanel();
    await refreshDashboard();
    setInterval(() => refreshDashboard().catch(err => console.error("Erro no refresh da dashboard:", err)), 30 * 60 * 1000);
});

// Evita que o bot crashe com erros inesperados
client.on("error", console.error);
process.on("unhandledRejection", console.error);

client.login(TOKEN);
