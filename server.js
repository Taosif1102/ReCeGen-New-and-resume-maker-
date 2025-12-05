const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, '.')));

// Ensure output directory exists
const outputDir = path.join(__dirname, 'docx');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}

// Helper function to run Python script
const runPythonScript = (scriptPath, data, res, outputFilename) => {
    const tempJsonPath = path.join(__dirname, 'temp_data.json');
    fs.writeFileSync(tempJsonPath, JSON.stringify(data));

    console.log(`Executing python script: ${scriptPath}`);
    const pythonProcess = spawn('python', [scriptPath, tempJsonPath]);

    let errorOutput = '';

    pythonProcess.stdout.on('data', (data) => {
        console.log(`stdout: ${data}`);
    });

    pythonProcess.stderr.on('data', (data) => {
        console.error(`stderr: ${data}`);
        errorOutput += data.toString();
    });

    pythonProcess.on('error', (err) => {
        console.error('Failed to start subprocess:', err);
        res.status(500).send(`Failed to execute Python script: ${err.message}`);
    });

    pythonProcess.on('close', (code) => {
        if (fs.existsSync(tempJsonPath)) {
            fs.unlinkSync(tempJsonPath); // Clean up temp file
        }

        if (code === 0) {
            const filePath = path.join(outputDir, outputFilename);
            if (fs.existsSync(filePath)) {
                res.download(filePath, outputFilename, (err) => {
                    if (err) {
                        console.error('Error downloading file:', err);
                        if (!res.headersSent) {
                            res.status(500).send('Error downloading file');
                        }
                    }
                });
            } else {
                console.error('Generated file not found');
                res.status(500).send('Generated file not found');
            }
        } else {
            console.error(`Python script exited with code ${code}`);
            res.status(500).send(`Error generating document (Exit Code ${code}): ${errorOutput}`);
        }
    });
};

app.post('/api/generate-resume', (req, res) => {
    const data = req.body;
    const scriptPath = path.join(__dirname, 'python', 'resume_builder.py');
    runPythonScript(scriptPath, data, res, 'resume_output.docx');
});

app.post('/api/generate-cv', (req, res) => {
    const data = req.body;
    const scriptPath = path.join(__dirname, 'python', 'generate_cv.py');
    runPythonScript(scriptPath, data, res, 'cv_output.docx');
});

app.post('/api/generate-ats-cv', (req, res) => {
    const data = req.body;
    const scriptPath = path.join(__dirname, 'python', 'ats_cv_builder.py');
    runPythonScript(scriptPath, data, res, 'ats_cv_output.docx');
});

app.post('/api/generate-ats-resume', (req, res) => {
    const data = req.body;
    const scriptPath = path.join(__dirname, 'python', 'ats_resume_builder.py');
    runPythonScript(scriptPath, data, res, 'ats_resume_output.docx');
});

app.post('/api/generate-modern-resume', (req, res) => {
    const data = req.body;
    const scriptPath = path.join(__dirname, 'python', 'modern_resume_builder.py');
    runPythonScript(scriptPath, data, res, 'modern_resume.docx');
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
