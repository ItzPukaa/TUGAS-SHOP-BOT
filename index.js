import { Client, GatewayIntentBits, Partials } from "discord.js";
import { google } from "googleapis";
import fs from "fs";
import { REST, Routes } from "discord.js";

const STAFF_ROLE_ID = "1389265558171684904";
const credentials = JSON.parse(fs.readFileSync(process.env.GOOGLE_APPLICATION_CREDENTIALS, "utf8"));
const config = {
    token: process.env.TOKEN,
    sheetId: process.env.SHEET_ID,
    staffRoleId: process.env.STAFF_ROLE_ID
};

const client = new Client({
    intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages
    ],
    partials: [Partials.Channel]
});

let sheetsClient;

// Inicializar Google Sheets
async function initSheets() {
    const auth = new google.auth.GoogleAuth({
        keyFile: "credentials.json",
        scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    });

    const authClient = await auth.getClient();
    sheetsClient = google.sheets({ version: "v4", auth: authClient });
}

// Ler stock da Google Sheet
async function getStockData() {
    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: config.sheetId,
        range: "Stock!A2:B"
    });

    const rows = res.data.values || [];

    return rows.map(row => {
    const raw = row[1] || "FALSE";

    const stock = (raw === "TRUE") ? 1 : 0;

    return {
        name: row[0],
        stock
    };
});
}

// Quando o Ticket Tool cria um ticket
client.on("channelCreate", async (channel) => {
    try {
        if (!channel.name.startsWith("ticket-")) return;

        const stockData = await getStockData();

        const lines = stockData.map(item => {
            const emoji = item.stock > 0 ? "🟢" : "🔴";
            return `${emoji} ${item.name}`;
        });

        // Embed de boas-vindas
        await channel.send({
            embeds: [
                {
                    title: "👋 Welcome to the ticket!",
                    description: "Thanks for opening a ticket. The team will get back to you as soon as possible.\n\n Here is the updated stock:",
                    color: 0x2ecc71 // verde
                }
            ]
        });

        // Embed do stock
        await channel.send({
            embeds: [
                {
                    title: "🧾 Current Stock",
                    description: lines.join("\n"),
                    color: 0x3498db // azul
                }
            ]
        });


    } catch (err) {
        console.error("Error sending stock:", err);
    }
});

//editar tabela do sheets
client.on("messageCreate", async (message) => {
    if (message.author.bot) return;
    if (!message.member.roles.cache.has(STAFF_ROLE_ID)) return; // só staff

    const args = message.content.split(" ");
    const cmd = args.shift().toLowerCase();

    // Comando: adicionar item
    if (cmd === "!additem") {
        const itemName = args.join(" ");
        if (!itemName) return message.reply("❌ Escreve o nome do item.");

        await sheetsClient.spreadsheets.values.append({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B",
            valueInputOption: "USER_ENTERED",
            resource: {
                values: [[itemName, "FALSE"]]
            }
        });

        return message.reply(`🟢 Item **${itemName}** adicionado com sucesso!`);
    }

    // Comando: remover item
    if (cmd === "!delitem") {
        const itemName = args.join(" ");
        if (!itemName) return message.reply("❌ Escreve o nome do item.");

        const res = await sheetsClient.spreadsheets.values.get({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B"
        });

        const rows = res.data.values || [];
        const index = rows.findIndex(r => r[0] === itemName);

        if (index === -1) return message.reply("❌ Item não encontrado.");

        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: config.sheetId,
            range: `Stock!A${index + 1}:B${index + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: {
                values: [["", ""]]
            }
        });

        return message.reply(`🟢 Item **${itemName}** removido.`);
    }

    // Comando: alterar checkbox
    if (cmd === "!stock") {
        const itemName = args.shift();
        const newValue = args.shift();

        if (!itemName || !newValue) {
            return message.reply("❌ Usa: `!stock <item> <on/off>`");
        }

        const value = newValue.toLowerCase() === "on" ? "TRUE" : "FALSE";

        const res = await sheetsClient.spreadsheets.values.get({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B"
        });

        const rows = res.data.values || [];
        const index = rows.findIndex(r => r[0] === itemName);

        if (index === -1) return message.reply("❌ Item não encontrado.");

        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: config.sheetId,
            range: `Stock!B${index + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: {
                values: [[value]]
            }
        });

        return message.reply(`🟢 Stock de **${itemName}** atualizado para **${value}**.`);
    }
});

client.once("ready", async () => {
    console.log(`Bot ligado como ${client.user.tag}`);
    await initSheets();
    await registerCommands(); // Slash comands
});

async function registerCommands() {
    const commands = [
        {
            name: "stock",
            description: "Alterar o stock de um item",
            options: [
                {
                    name: "item",
                    description: "Escolhe o item",
                    type: 3,
                    required: true,
                    autocomplete: true
                },
                {
                    name: "estado",
                    description: "on = TRUE | off = FALSE",
                    type: 3,
                    required: true,
                    choices: [
                        { name: "on", value: "on" },
                        { name: "off", value: "off" }
                    ]
                }
            ]
        },
        {
            name: "delitem",
            description: "Apagar um item da tabela",
            options: [
                {
                    name: "item",
                    description: "Escolhe o item",
                    type: 3,
                    required: true,
                    autocomplete: true
                }
            ]
        },
        {
            name: "additem",
            description: "Adicionar um item à tabela",
            options: [
                {
                    name: "nome",
                    description: "Nome do item",
                    type: 3,
                    required: true
                }
            ]
        },
        {
            name: "showstock",
            description: "Mostra o stock completo da tabela"
        }
    ];

    const rest = new REST({ version: "10" }).setToken(config.token);

    await rest.put(
    Routes.applicationGuildCommands("1558956018476515331", "1383063106594406510"),
    { body: commands }
    );

    console.log("Slash commands registados!");
}

client.on("interactionCreate", async (interaction) => {
    if (!interaction.isAutocomplete()) return;

    const focused = interaction.options.getFocused();

    const res = await sheetsClient.spreadsheets.values.get({
        spreadsheetId: config.sheetId,
        range: "Stock!A:A"
    });

    const items = (res.data.values || []).map(r => r[0]);

    const filtered = items.filter(i =>
        i.toLowerCase().includes(focused.toLowerCase())
    );

    await interaction.respond(
        filtered.slice(0, 25).map(i => ({ name: i, value: i }))
    );
});

client.on("interactionCreate", async (interaction) => {
    if (!interaction.isChatInputCommand()) return;

    const { commandName } = interaction;

    // /verstock
    if (commandName === "showstock") {
        const res = await sheetsClient.spreadsheets.values.get({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B"
        });

        const rows = res.data.values || [];

        const lines = rows.map(r => {
            const name = r[0];
            const raw = r[1] || "FALSE";
            const emoji = raw === "TRUE" ? "🟢" : "🔴";
            return `${emoji} ${name}`;
        });

        return interaction.reply({
            embeds: [
                {
                    title: "📦 Stock Atual",
                    description: lines.join("\n"),
                    color: 0x3498db
                }
            ]
        });
    }

    // /additem
    if (commandName === "additem") {
        const nome = interaction.options.getString("nome");

        await sheetsClient.spreadsheets.values.append({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B",
            valueInputOption: "USER_ENTERED",
            resource: { values: [[nome, "FALSE"]] }
        });

        return interaction.reply(`🟢 Item **${nome}** adicionado!`);
    }

    // /delitem
    if (commandName === "delitem") {
        const item = interaction.options.getString("item");

        const res = await sheetsClient.spreadsheets.values.get({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B"
        });

        const rows = res.data.values || [];
        const index = rows.findIndex(r => r[0] === item);

        if (index === -1)
            return interaction.reply("❌ Item não encontrado.");

        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: config.sheetId,
            range: `Stock!A${index + 1}:B${index + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: { values: [["", ""]] }
        });

        return interaction.reply(`🟢 Item **${item}** removido.`);
    }

    // /stock
    if (commandName === "stock") {
        const item = interaction.options.getString("item");
        const estado = interaction.options.getString("estado");
        const value = estado === "on" ? "TRUE" : "FALSE";

        const res = await sheetsClient.spreadsheets.values.get({
            spreadsheetId: config.sheetId,
            range: "Stock!A:B"
        });

        const rows = res.data.values || [];
        const index = rows.findIndex(r => r[0] === item);

        if (index === -1)
            return interaction.reply("❌ Item não encontrado.");

        await sheetsClient.spreadsheets.values.update({
            spreadsheetId: config.sheetId,
            range: `Stock!B${index + 1}`,
            valueInputOption: "USER_ENTERED",
            resource: { values: [[value]] }
        });

        return interaction.reply(`🟢 Stock de **${item}** atualizado para **${value}**.`);
    }
});


client.login(config.token);
