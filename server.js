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
async function modifyDocxDirectly(filePath, segments) {
    console.log(segments);
    
    try {
        // Đọc file DOCX
        const data = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(data);

        // Đọc nội dung XML chính
        const docXmlPath = "word/document.xml";
        const docXmlContent = await zip.file(docXmlPath).async("text");

        // Chuyển XML thành DOM
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(docXmlContent, "text/xml");

        // Tìm tất cả các thẻ <w:t> (chứa nội dung văn bản)
        const textNodes = xmlDoc.getElementsByTagName("w:t");

        // Gộp toàn bộ văn bản thành một chuỗi để tìm kiếm
        let fullText = "";
        let textElements = [];
        
        for (let i = 0; i < textNodes.length; i++) {
            textElements.push({
                node: textNodes[i],
                text: textNodes[i].textContent
            });
            fullText += textNodes[i].textContent;
        }

        // Duyệt qua từng đoạn cần bôi vàng
        segments.forEach(({ segment }) => {
            let startIndex = fullText.indexOf(segment);
            if (startIndex !== -1) {
                console.log(`🔍 Tìm thấy "${segment}" tại vị trí ${startIndex}`);

                // Tạo thẻ highlight
                const highlightNode = xmlDoc.createElement("w:highlight");
                highlightNode.setAttribute("w:val", "yellow");

                // Chèn highlight vào các thẻ <w:rPr> tương ứng
                let currentIndex = 0;
                for (let i = 0; i < textElements.length; i++) {
                    let { node, text } = textElements[i];

                    // Nếu đoạn văn bản nằm trong khoảng cần bôi vàng
                    if (
                        currentIndex >= startIndex &&
                        currentIndex < startIndex + segment.length
                    ) {
                        let parentRun = node.parentNode;
                        let rPrNode = xmlDoc.createElement("w:rPr");
                        rPrNode.appendChild(highlightNode.cloneNode(true));
                        parentRun.insertBefore(rPrNode, node);
                    }

                    currentIndex += text.length;
                }
            }
        });

        // Chuyển DOM về chuỗi XML
        const serializer = new XMLSerializer();
        const newXml = serializer.serializeToString(xmlDoc);

        // Ghi lại nội dung mới vào ZIP
        zip.file(docXmlPath, newXml);

        // Ghi đè lại chính file DOCX gốc
        zip.generateAsync({ type: "nodebuffer" }).then((buffer) => {
            fs.writeFileSync(filePath, buffer);
            console.log("✅ File DOCX đã được cập nhật thành công!");
        });
    } catch (error) {
        console.error("❌ Lỗi khi chỉnh sửa file DOCX:", error);
    }
}

function guidGenerator() {
    var S4 = function() {
       return (((1+Math.random())*0x10000)|0).toString(16).substring(1);
    };
    return (S4()+S4()+"-"+S4()+"-"+S4()+"-"+S4()+"-"+S4()+S4()+S4());
}

async function extractTextFromPDF(pdfPath) {
    try {
        const dataBuffer = fs.readFileSync(pdfPath);
        const data = await pdfParse(dataBuffer);
        return data.text;
    } catch (error) {
        console.error("Lỗi khi đọc PDF:", error);
    }
}

function findAddedText(originalText, newText) {
    const originalWords = originalText.split(/\s+/); // Tách từ dựa trên khoảng trắng
    const newWords = newText.split(/\s+/);
    console.log(originalWords);
    console.log(newWords);
    const dp = Array.from({ length: originalWords.length + 1 }, () => Array(newWords.length + 1).fill(0));

    // Xây dựng bảng DP cho LCS
    for (let i = 1; i <= originalWords.length; i++) {
        for (let j = 1; j <= newWords.length; j++) {
            if (originalWords[i - 1] === newWords[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1] + 1;
            } else {
                dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
            }
        }
    }

    // Truy vết ngược để tìm các đoạn được thêm và vị trí của chúng
    let i = originalWords.length;
    let j = newWords.length;
    const addedSegments = [];

    let tempSegment = [];
    let tempOriginalIndexes = [];  // Mảng chứa chỉ mục trong `originalWords` cho từ được thêm
    let tempNewIndexes = []; // Mảng chứa chỉ mục trong `newWords` cho từ được thêm

    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && originalWords[i - 1] === newWords[j - 1]) {
            if (tempSegment.length > 0) {
                addedSegments.unshift({
                    segment: tempSegment.join(" "),
                    originalIndexes: tempOriginalIndexes,
                    newIndexes: tempNewIndexes
                });
                tempSegment = [];
                tempOriginalIndexes = [];
                tempNewIndexes = [];
            }
            i--;
            j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
            tempSegment.unshift(newWords[j - 1]); // Thêm từ vào đoạn tạm thời
            tempNewIndexes.unshift(j - 1); // Lưu chỉ mục của từ trong `newWords`
            j--;
        } else {
            i--;
        }
    }

    // Nếu còn đoạn văn bản mới chưa được thêm
    if (tempSegment.length > 0) {
        addedSegments.unshift({
            segment: tempSegment.join(" "),
            originalIndexes: [], // Không có chỉ mục trong `originalWords` vì đoạn này chỉ có trong văn bản mới
            newIndexes: tempNewIndexes
        });
    }

    return addedSegments;
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
    
    const fileOriginPath = req.files.fileOrigin[0].path;

    const originText = await extractTextFromPDF(fileOriginPath);
    const translationText = await extractTextFromPDF(fileTranslationNewPath);

    const name = guidGenerator();

    await PDFNet.runWithCleanup(async () => {
        
        await PDFNet.addResourceSearchPath('./Lib/');
            // check if the module is available
            if (!(await PDFNet.StructuredOutputModule.isModuleAvailable())) {
                return;
            }
            await PDFNet.Convert.fileToWord(fileTranslationNewPath, path.join(__dirname, 'public') + '/' + name + '.docx');
        }, 'demo:1739949060645:617adfb903000000000935bcbf740717e9c6b2c11e6fd7ac9496321dc6')
            .catch(err => {
                console.error(err);
            })
            .then(async function () {
                PDFNet.shutdown();
                const segments = findAddedText(originText, translationText);
                
                await modifyDocxDirectly(path.join(__dirname, 'public') + '/' + name + '.docx', segments);
                res.json({ message: 'File converted successfully', path: `${name}.docx`, changed: segments });
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