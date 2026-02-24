// filepath: c:\Users\HKAN739291\OneDrive - WSP O365\powershell\PowerShell(using)\server.js
const express = require('express');
const { exec } = require('child_process');
const path = require('path');

const app = express();
const port = 3000;

app.use(express.json());
app.use(express.static(__dirname)); // Serve static files (e.g., HTML)

app.post('/run-script', (req, res) => {
    const scriptName = req.body.scriptName;

    // Validate the script name to prevent arbitrary command execution
    const allowedScripts = ['GetACLReport_byAlbert(2.0).ps1', 'ADDirector.ps1', 'ADReaper.ps1', 'run_collectinfo.bat'];
    if (!allowedScripts.includes(scriptName)) {
        return res.status(400).json({ message: 'Invalid script name.' });
    }

    // Properly quote the file path
    const scriptPath = `"C:\\temp\\PowerShell(using)\\${scriptName}"`;

    // Construct the command
    const command = scriptName.endsWith('.ps1')
        ? `powershell.exe -NoExit -File ${scriptPath}`
        : `cmd.exe /c ${scriptPath}`;

    exec(command, (error, stdout, stderr) => {
        if (error) {
            console.error(`Error: ${error.message}`);
            return res.status(500).json({ message: 'Failed to run script.' });
        }
        if (stderr) {
            console.error(`Stderr: ${stderr}`);
            return res.status(500).json({ message: 'Script error occurred.' });
        }
        console.log(`Stdout: ${stdout}`);
        res.status(204).end(); // Send a "No Content" response without any JSON data
    });
});

app.post('/execute-batch', (req, res) => {
    const scriptPath = 'C:\\temp\\PowerShell(using)\\CollectInfo.ps1';
    exec(`start cmd.exe /k powershell.exe -NoExit -File "${scriptPath}"`, () => {
        res.status(204).end(); // Send a "No Content" response without any JSON data
    });
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});