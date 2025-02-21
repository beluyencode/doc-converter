const express = require('express');
const cors = require('cors');
const { PDFNet } = require('@pdftron/pdfnet-node');
const pdfParse = require('pdf-parse');
const fs = require('fs');
const path = require('path');
const app = express();
const port = 3000;
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("xmldom");
const jsDiff = require('diff');

// const mammoth = require("mammoth");

const allowCrossDomain = (req, res, next) => {
    res.header(`Access-Control-Allow-Origin`, `*`);
    res.header(`Access-Control-Allow-Methods`, `GET,PUT,POST,DELETE`);
    res.header(`Access-Control-Allow-Headers`, `Content-Type`);
    next();
};

app.use(allowCrossDomain);

app.use(express.static('public'));
app.use(cors({
    origin: '*',
}));

const upload = require('multer')({
    dest: 'uploads/'
});

// ----------------- Helper -----------------

async function modifyDocxDirectly(newPath, segments) {
    try {
        const data = fs.readFileSync(newPath);
        const zip = await JSZip.loadAsync(data);

        const docXmlPath = "word/document.xml";
        const docXmlContent = await zip.file(docXmlPath).async("text");

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(docXmlContent, "text/xml");

        const textNodes = xmlDoc.getElementsByTagName("w:t");

        let fullText = "";
        let textElements = [];

        // Duyệt qua tất cả <w:t> để xây dựng văn bản đầy đủ
        for (let i = 0; i < textNodes.length; i++) {
            textElements.push({
                node: textNodes[i],
                text: textNodes[i].textContent,
            });
            fullText += textNodes[i].textContent;
        }

        let searchStartIndex = 0;

        segments.forEach(({ original, translated }) => {
            let startIndex = fullText.indexOf(original.trim(), searchStartIndex);
            if (startIndex !== -1) {
                console.log(`🔍 Tìm thấy "${original}" tại vị trí ${startIndex}`);

                let endIndex = startIndex + original.length;
                let currentIndex = 0;
                let lastNode = null; // Node cuối cùng chứa phần segment

                for (let i = 0; i < textElements.length; i++) {
                    let { node, text } = textElements[i];

                    if (currentIndex + text.length > startIndex) {
                        if (currentIndex < endIndex) {
                            lastNode = node;
                        }
                    }

                    currentIndex += text.length;
                }

                if (lastNode) {
                    let parentRun = lastNode.parentNode;

                    // Thêm văn bản marianData[index] sau đoạn gốc
                    let newTextNode = xmlDoc.createElement("w:t");
                    newTextNode.textContent = " / " + translated;

                    let newRunNode = xmlDoc.createElement("w:r");
                    let newRunProperties = xmlDoc.createElement("w:rPr");
                    let colorNode = xmlDoc.createElement("w:color");
                    colorNode.setAttribute("w:val", "FF0000"); // Màu đỏ

                    newRunProperties.appendChild(colorNode);
                    newRunNode.appendChild(newRunProperties);
                    newRunNode.appendChild(newTextNode);

                    parentRun.parentNode.insertBefore(newRunNode, parentRun.nextSibling);
                }

                searchStartIndex = endIndex;
            } else {
                console.log(`❌ Không tìm thấy "${segment}" trong tài liệu.`);
            }
        });

        const serializer = new XMLSerializer();
        const newXml = serializer.serializeToString(xmlDoc);

        zip.file(docXmlPath, newXml);

        zip.generateAsync({ type: "nodebuffer" }).then((buffer) => {
            fs.writeFileSync(newPath, buffer);
            console.log("✅ File DOCX đã được cập nhật thành công!");
        });
    } catch (error) {
        console.error("❌ Lỗi khi chỉnh sửa file DOCX:", error);
    }
}


function guidGenerator() {
    var S4 = function () {
        return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
    };
    return (S4() + S4() + "-" + S4() + "-" + S4() + "-" + S4() + "-" + S4() + S4() + S4());
}

async function extractTextFromDocx(docxPath) {
    try {
        const data = fs.readFileSync(docxPath);
        const zip = await JSZip.loadAsync(data);
        const contentXml = await zip.file("word/document.xml").async("text");
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(contentXml, "text/xml");

        let text = "";
        const paragraphs = xmlDoc.getElementsByTagName("w:p");

        for (let i = 0; i < paragraphs.length; i++) {
            let paragraphText = "";
            const textNodes = paragraphs[i].getElementsByTagName("w:t");

            for (let j = 0; j < textNodes.length; j++) {
                paragraphText += textNodes[j].textContent;
            }

            if (paragraphText.trim().length > 0) {
                text += paragraphText + "\n"; // Xuống dòng giữa các đoạn
            }
        }

        return text;
    } catch (error) {
        console.error("Lỗi khi đọc DOCX:", error);
        return "";
    }
}

function removeSegmentNotTranslated(segments, marianData) {
    const data = [];
    for (let i = 0; i < segments.length; i++) {
        if (marianData[i] !== segments[i]) {
            data.push({
                original: segments[i],
                translated: marianData[i],
            });
        }
    }
    return data;
}

function findAddedAndReplacedText(originalText, newText) {
    const diffResult = jsDiff.diffWords(originalText, newText);

    const data = [];

    let tempDeleted = null;

    for (const part of diffResult) {
        if (part.added) {
            if (tempDeleted) {
                // Nếu có phần bị xóa trước đó, coi như bị thay thế
                data.push({ old: tempDeleted, new: part.value.trim() });
                tempDeleted = null;
            } else {
                data.push({
                    old: null, // Không có trong văn bản gốc
                    new: part.value.trim()
                });
            }
        } else if (part.removed) {
            tempDeleted = part.value.trim();
        } else {
            tempDeleted = null; // Nếu có phần chung, reset tempDeleted
        }
    }

    return data.reduce((prev, next) => {
        if (next.new.includes("\n")) {
            const split = next.new.split("\n");
            return [...prev, ...split];
        }
        return [...prev, next.new];
    }, []);;
}

async function translateTexts(listText) {
    const data = await fetch('http://localhost:8000/translates', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            source_text: listText,
        }),
    }).then(response => response.json())
    return data;
}

// ---------------------------------------------------

// API to handle PDF upload
app.post('/convert', upload.fields([{ name: 'fileOrigin', maxCount: 1 }, { name: 'fileTranslation', maxCount: 1 }]), async (req, res) => {
    if (!req.files) {
        return res.status(400).send('No file uploaded.');
    }
    const pdfDir = path.join(__dirname, 'pdf');
    if (!fs.existsSync(pdfDir)) {
        fs.mkdirSync(pdfDir);
    }

    const fileTranslationNewPath = path.join(pdfDir, req.files.fileTranslation[0].filename) + ".pdf";
    fs.renameSync(req.files.fileTranslation[0].path, fileTranslationNewPath);
    req.files.fileTranslation[0].path = fileTranslationNewPath;

    const fileOriginPath = path.join(pdfDir, req.files.fileOrigin[0].filename) + ".pdf";
    fs.renameSync(req.files.fileOrigin[0].path, fileOriginPath);
    req.files.fileOrigin[0].path = fileOriginPath;

    const name = guidGenerator();
    const pathOrigin = path.join(__dirname, 'public') + '/' + name + '_origin.docx';
    const pathTranslation = path.join(__dirname, 'public') + '/' + name + '_translation.docx';
    await PDFNet.runWithCleanup(async () => {
        await PDFNet.addResourceSearchPath('./Lib/');
        // check if the module is available
        if (!(await PDFNet.StructuredOutputModule.isModuleAvailable())) {
            return;
        }

        await PDFNet.Convert.fileToWord(fileTranslationNewPath, pathTranslation);
        await PDFNet.Convert.fileToWord(fileOriginPath, pathOrigin);
    }, 'demo:1739949060645:617adfb903000000000935bcbf740717e9c6b2c11e6fd7ac9496321dc6')
        .catch(err => {
            console.error(err);
        })
        .then(async () => {
            PDFNet.shutdown();

            const originText = await extractTextFromDocx(pathOrigin);
            const translationText = await extractTextFromDocx(pathTranslation);
            const segments = findAddedAndReplacedText(originText, translationText);
            const marianData = await translateTexts(segments);
            const segmentsTranslate = removeSegmentNotTranslated(segments, marianData.translation);
            console.log("segments:", segments);

            console.log("marianData:", marianData);

            console.log("segmentsTranslate:", segmentsTranslate);
            await modifyDocxDirectly(pathTranslation, segmentsTranslate);
            res.json({ message: 'File converted successfully', path: `${name}_translation.docx`, changed: segments });
        });
});

// API to convert DOCX to HTML
// app.post('/convert-docx-to-html', upload.single('file'), async (req, res) => {
//     if (!req.file) {
//         return res.status(400).send('No file uploaded.');
//     }

//     const docxFilePath = req.file.path;

//     try {
//         const result = await mammoth.convertToHtml({ path: docxFilePath });
//         res.send(result.value); // The generated HTML
//     } catch (error) {
//         console.error("Error converting DOCX to HTML:", error);
//         res.status(500).send('Error converting DOCX to HTML.');
//     } finally {
//         // Clean up the uploaded file
//         fs.unlinkSync(docxFilePath);
//     }
// });

app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});