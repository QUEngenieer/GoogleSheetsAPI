
const { google } = require('googleapis');

async function authenticate() {
    try {
        const auth = await google.auth.getClient({
        keyFile: "credentials.json",
        scopes: "https://www.googleapis.com/auth/spreadsheets",
        });

        return auth;
    } catch (error) {
        console.error("Error during authentication:", error.message);
        throw error;
    }
    }

    async function readSpreadsheet(auth) {
    try {
        const sheets = google.sheets({ version: "v4", auth });
        const spreadsheetId = "1kn7v8PNAGV2cPsgJ-t9uUMesPGAiXMALeSQOPsBeamA";
        const range = "engenharia_de_software"; 

        const response = await sheets.spreadsheets.get({
        spreadsheetId,
        });

        const sheetNames = response.data.sheets.map(sheet => sheet.properties.title);
        console.log("Sheets found:", sheetNames);

        const sheet = response.data.sheets.find(sheet => sheet.properties.title === range);

        if (!sheet) {
        throw new Error(`Sheet "${range}" not found.`);
        }

        const values = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${range}!A1:H${sheet.properties.gridProperties.rowCount}`,
        });

        return values.data.values;
    } catch (error) {
        console.error("Error while reading spreadsheet:", error.message);
        throw error;
    }
    }

    async function calculateAndUpdateSheet(auth, data) {
    try {
        const sheets = google.sheets({ version: "v4", auth });
        const spreadsheetId = "1kn7v8PNAGV2cPsgJ-t9uUMesPGAiXMALeSQOPsBeamA";
        const rangeToUpdate = "engenharia_de_software!G4:K100"; 

        const valuesToWrite = []

        // Iterate over the data and calculate the situation, the final approval grade, and the average
        for (const row of data.slice(3)) { // Start from row 4
        const [matricula, Aluno, absences, p1, p2, p3] = row.map(Number);
        let average = ((p1 + p2 + p3) / 30).toFixed(2);
        let situacao = "";
        let naf = 0; 

        if (absences > 0.25 * 60) {
            situacao = "Reprovado por Falta";
        } else if (average < 5) {
            situacao = "Reprovado por Nota";
        } else if (average >= 5 && average < 7) {
            situacao = "Exame Final";

            // Calculate the Final Approval Grade 
            naf = Math.max(0, Math.ceil(10 - average));
        } else if (average >= 7) {
            situacao = "Aprovado";
        }

        valuesToWrite.push([situacao, naf]);
        }

        // Results back to the spreadsheet
        await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: rangeToUpdate,
        valueInputOption: "USER_ENTERED",
        resource: { values: valuesToWrite },
        });
    } catch (error) {
        console.error("Error during calculation and update:", error.message);
        throw error;
    }
    }

    async function main() {
    try {
        const auth = await authenticate();
        const spreadsheetData = await readSpreadsheet(auth);

        await calculateAndUpdateSheet(auth, spreadsheetData);

        console.log("Processing completed.");
    } catch (error) {
        console.error("Error:", error.message);
    }
    }

    main();