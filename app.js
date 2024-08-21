const express = require('express');
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const path = require('path');
const app = express();
const port = 8000;

app.use(fileUpload());
const publicPath = path.join(__dirname,'public');
app.get('/',(req,res)=>{
    res.sendFile(`${publicPath}/index.html`);
});
app.post('/upload', (req, res) => {
    if (!req.files || !req.files.excelFile) {
        return res.status(400).send({ message: 'No file uploaded.' });
    }

    const file = req.files.excelFile;
    const filePath = path.join(__dirname, 'uploads', file.name);

    file.mv(filePath, async (err) => {
        if (err) {
            console.error('File move error:', err);
            return res.status(500).send({ message: 'Error uploading file.' });
        }
    for(let i = 1; i<=5; i++){
        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            const worksheet = workbook.getWorksheet(1);

            let transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: {
                    user: 'puneetkhoiyabnl8@gmail.com',
                    pass: 'xadgnnjutdlvumjb'
                }
            });

            worksheet.eachRow({ includeEmpty: false }, (row) => {
                const name = row.getCell(1).value;
                let email = row.getCell(2).value;
                const company = row.getCell(3).value;

                if (typeof email === 'object' && email !== null) {
                    email = email.text || email.hyperlink || null;
                }

                if (typeof email === 'string' && email.includes('@')) {
                    const mailOptions = {
                        from: 'puneetkhoiyabnl8@gmail.com',
                        to: email,
                        subject: `Campus Hiring Opportunities at ${company}`,
                        text: `Hi ${name},\n\nI’m Puneet Khoiya, a Computer Science student at NIT Jalandhar and currently serving as the Internship Representative for our batch. I’m reaching out to discuss potential opportunities for campus hiring at ${company}. Our college has a strong pool of talented students eager to contribute and grow in the industry.\n\nI would love to connect and explore how ${company} can collaborate with NIT Jalandhar for internship and placement opportunities.\n\nLooking forward to your response!\n\nBest Regards,\nPuneet Khoiya`
                    };

                    transporter.sendMail(mailOptions, (error, info) => {
                        if (error) {
                            console.log('Error sending email to:', email, error);
                        } else {
                            console.log('Email sent to:', email, info.response);
                        }
                    });
                } else {
                    console.log('Invalid email address:', email);
                }
            });

            res.send({ message: 'Emails sent successfully.' });

        } catch (error) {
            console.error('Error during file processing:', error);
            res.status(500).send({ message: 'Error processing file.' });
        }
    }    
    });
});

app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});