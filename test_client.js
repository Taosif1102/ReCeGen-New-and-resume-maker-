const http = require('http');
const fs = require('fs');

const data = JSON.stringify({
    name: "Test User",
    title: "Developer",
    email: "test@example.com",
    phone: "1234567890",
    address: "Test Address",
    objective: "Test Objective",
    skills: "Python, JavaScript",
    experience: [],
    education: []
});

const options = {
    hostname: 'localhost',
    port: 3000,
    path: '/api/generate-resume',
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
        'Content-Length': data.length
    }
};

const req = http.request(options, (res) => {
    console.log(`StatusCode: ${res.statusCode}`);
    
    if (res.statusCode === 200) {
        const file = fs.createWriteStream("test_resume.docx");
        res.pipe(file);
        file.on('finish', () => {
            file.close();
            console.log("Download completed: test_resume.docx");
        });
    } else {
        res.on('data', (d) => {
            process.stdout.write(d);
        });
    }
});

req.on('error', (error) => {
    console.error(error);
});

req.write(data);
req.end();
