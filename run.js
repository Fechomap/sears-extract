const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');

// Configuración de logging
const logsDir = path.join(__dirname, 'logs');
// Asegurar que la carpeta logs existe
if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
}
const logFile = path.join(logsDir, 'run.log'); // Cambiado a logs/run.log
const logStream = fs.createWriteStream(logFile, { flags: 'a' });

function log(message) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] ${message}\n`;
    console.log(logMessage.trim());
    logStream.write(logMessage);
}

function runScript(scriptName) {
    return new Promise((resolve, reject) => {
        log(`Ejecutando ${scriptName}...`);
        exec(`python ${scriptName}`, (error, stdout, stderr) => {
            if (error) {
                log(`Error al ejecutar ${scriptName}: ${stderr || error.message}`);
                reject(error);
            } else {
                log(`${scriptName} ejecutado correctamente.`);
                if (stdout) log(`Salida de ${scriptName}: ${stdout}`);
                resolve();
            }
        });
    });
}

async function main() {
    try {
        // Lista de scripts a ejecutar en orden
        const scripts = [
            "scripts/extract.py",
            "scripts/merge_data.py",
            "scripts/merge_csv_data.py"
        ];

        for (const script of scripts) {
            await runScript(script);
        }

        log("Todos los scripts se han ejecutado correctamente.");
    } catch (error) {
        log(`El proceso falló: ${error.message}`);
    } finally {
        logStream.end();
    }
}

main();
