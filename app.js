const express = require("express");
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
var app = express();
var bodyParser = require("body-parser");
const PORT = 7000;

app.use(bodyParser.json());app.use(bodyParser.urlencoded({ extended: true }));

app.use(express.static("./Public"));
app.set("view engine", "ejs");

app.get("/", (req, res) => {
    res.render("page");
});
app.get("/form", (req, res) => {
    res.render("form");
});


app.post("/doc", (req, res) => {
    console.log(req.body);
    const selectedValue = req.body.course;

    const options = {
        "CO": "Computer Engineering",
        "IF": "Information Technology",
        "EE": "Electrical Engineering",
        "ME": "Mechanical Engineering",
        "EJ": "Electronics & Telecommunication Engineering"
    };

    const selectedText = options[selectedValue];

    const selectedValueYear = req.body.sem; // This will hold the selected value, e.g., "CO"

    const optionsYear = {
        "1": "First",
        "2": "First",
        "3": "Second",
        "4": "Second",
        "5": "Third",
        "6": "Third"
    };


    const selectedTextYear = optionsYear[selectedValueYear];
    try {
        const templateFile = fs.readFileSync(path.resolve(__dirname, './Public/templateDocx/Sample_template_certificate.docx'), 'binary');
        const zip = new PizZip(templateFile);
        let outputDocument = new Docxtemplater(zip);
        console.log( req.body.name2);

        const dataToAdd_certificate = {
            
                STD_TITLE1: req.body.studentTitle1 || "Mr./Ms.",
                STD_TITLE2: req.body.studentTitle2 !== undefined ? ", " + req.body.studentTitle2 : "",
                STD_TITLE3: req.body.studentTitle3 !== undefined ? ", " + req.body.studentTitle3 : "",
                STD_TITLE4: req.body.studentTitle4 !== undefined ? ", " + req.body.studentTitle4 : "",
                STD_TITLE5: req.body.studentTitle5 !== undefined ? ", " + req.body.studentTitle5 : "",
                // ... other fields
            

            STUDENT_NAME1: req.body.name1 || "Omanand Prashant Swami",
STUDENT_NAME2: req.body.name2  || "",
STUDENT_NAME3: req.body.name3  || "",
STUDENT_NAME4: req.body.name4  || "",
STUDENT_NAME5: req.body.name5  || "",


            STUDENT_ENR1: req.body.studentEnrollment1 ||"2110950050",
            STUDENT_ENR2: req.body.studentEnrollment2 ||"",
            STUDENT_ENR3: req.body.studentEnrollment3 ||"",
            STUDENT_ENR4: req.body.studentEnrollment4 ||"",
            STUDENT_ENR5: req.body.studentEnrollment5 ||"",
            COI: selectedValue||"CO",
            SCHEME: req.body.scheme || "I",
            MICROPROJECT_TITLE: req.body.mpTitle || "Micro Project Title",
            COSUBJECT: req.body.subject || "Subject Name",
            SUBCODE: req.body.subCode || "Subject Code",
            TEACHERTITLE: req.body.teacherTitle || "Mr./Mr.",
            TEACHER_NAME: req.body.teacherName || "Teacher Name",
            ACAYEAR: req.body.acaYear || "2023-24",
            SEM: req.body.sem || "5",
            COURSE: selectedText || "Computer Engineering",
            YEAR: selectedTextYear || "Third",
            PRO: req.body.pro || "has",
        };

        outputDocument.setData(dataToAdd_certificate);
    
        
    
        try {
            outputDocument.render()
            let outputDocumentBuffer = outputDocument.getZip().generate({ type: 'nodebuffer' });

            // Set appropriate headers for the response
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', `attachment; filename=${req.body.name1||"Omanand"}-${req.body.subject||"Subject"}-certificate.docx`);

            // Send the generated DOCX file to the client for download
            res.send(outputDocumentBuffer);
        }
        catch (error) {
            console.error('ERROR Filling out Template:');
            console.error(error);
            res.status(500).send('Internal Server Error');
        }
    } catch (error) {
        console.error('ERROR Loading Template:');
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
});


app.post("/civilcertificate", (req, res) => {
    console.log(req.body);
    try {
        const templateFile = fs.readFileSync(path.resolve(__dirname, './Public/templateDocx/civil_certificate.docx'), 'binary');
        const zip = new PizZip(templateFile);
        let outputDocument = new Docxtemplater(zip);

        const dataToAdd = {
            MicroprojectTitle: req.body.MicroprojectTitle,
            yr: req.body.yr,
            STUD_NAME: req.body.stud_name,
            stdName2: req.body.stdName2,
            stdName3: req.body.stdName3,
            stdName4: req.body.stdName4,
            STUD_ENR: req.body.stud_enr,
            stdEnr2: req.body.stdEnr2,
            stdEnr3: req.body.stdEnr3,
            stdEnr4: req.body.stdEnr4,
            profSirName: req.body.profSirName,
            Subject: req.body.Subject,
            HOD: req.body.hod,
            Principal: req.body.Principal,
            sem: req.body.sem,
        };
        outputDocument.setData(dataToAdd);

        try {
            outputDocument.render()
            let outputDocumentBuffer = outputDocument.getZip().generate({ type: 'nodebuffer' });
            // Set appropriate headers for the response
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', `attachment; filename=${req.body.stud_name}-computer-certificate.docx`);

            // Send the generated DOCX file to the client for download
            res.send(outputDocumentBuffer);
        }
        catch (error) {
            console.error(`ERROR Filling out Template:`);
            console.error(error);
            res.status(500).send('Internal Server Error');
        }

    } catch (error) {
        console.log("Error Loading Template:");
        console.log(error);
        res.status(500).send('Internal Server Error');
    }
});


app.listen(PORT, ()=>{
    console.log("listing on port" + PORT);
})