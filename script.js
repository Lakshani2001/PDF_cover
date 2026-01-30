document.addEventListener('DOMContentLoaded', () => {
    // --- Elements ---
    const form = document.getElementById('coverPageForm');
    const universityInput = document.getElementById('university');
    const facultyInput = document.getElementById('faculty');
    const moduleInput = document.getElementById('module');
    const assignmentTitleInput = document.getElementById('assignmentTitle');
    const studentNameInput = document.getElementById('studentName');
    const studentIdInput = document.getElementById('studentId');
    const submissionDateInput = document.getElementById('submissionDate');
    const logoUploadInput = document.getElementById('logoUpload');
    const fileNameDisplay = document.getElementById('fileName');

    // Customization Elements
    const themeColorInput = document.getElementById('themeColor');
    const pageColorInput = document.getElementById('pageColor');

    // Typography Elements
    const fontFamilyInput = document.getElementById('fontFamily');
    const titleCaseInput = document.getElementById('titleCase');
    const titleSizeInput = document.getElementById('titleSize');

    // Preview Elements
    const paperPreview = document.getElementById('paperPreview');
    const prevUniversity = document.getElementById('prevUniversity');
    const prevFaculty = document.getElementById('prevFaculty');
    const prevModule = document.getElementById('prevModule');
    const prevAssignment = document.getElementById('prevAssignment');
    const prevName = document.getElementById('prevName');
    const prevId = document.getElementById('prevId');
    const prevDate = document.getElementById('prevDate');
    const previewLogoContainer = document.getElementById('previewLogoContainer');

    let logoBase64 = null;

    // --- Tab Switching Logic ---
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            // Remove active class from all
            tabBtns.forEach(b => b.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));

            // Add active to clicked
            btn.classList.add('active');
            const tabId = btn.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');
        });
    });

    // --- Real-time Preview Updates ---
    function updatePreview() {
        if (universityInput.value) prevUniversity.textContent = universityInput.value;
        if (facultyInput.value) prevFaculty.textContent = facultyInput.value;
        if (moduleInput.value) prevModule.textContent = moduleInput.value;

        // Assignment Title Logic (Case)
        let titleText = assignmentTitleInput.value;
        const caseStyle = titleCaseInput.value;
        if (caseStyle === 'uppercase') titleText = titleText.toUpperCase();
        else if (caseStyle === 'lowercase') titleText = titleText.toLowerCase();
        // Capitalize is handled via CSS in preview generally, but let's be explicit if we can. 
        // For preview CSS is easier:
        prevAssignment.style.textTransform = caseStyle;
        prevAssignment.textContent = assignmentTitleInput.value; // Let CSS handle transform

        if (studentNameInput.value) prevName.textContent = studentNameInput.value;
        if (studentIdInput.value) prevId.textContent = studentIdInput.value;
        if (submissionDateInput.value) prevDate.textContent = submissionDateInput.value;

        // Apply Styles
        const color = themeColorInput.value;
        const pageColor = pageColorInput.value;
        const borderStyle = 'solid';

        // Typography
        const fontFamily = fontFamilyInput.value;
        const titleSize = titleSizeInput.value;

        // Paper Background
        paperPreview.style.background = pageColor;

        // Font Family (Map simple names to CSS)
        // times -> 'Times New Roman'
        // helvetica -> 'Arial', sans-serif
        // courier -> 'Courier New', monospace
        let cssFont = 'Times New Roman';
        if (fontFamily === 'helvetica') cssFont = 'Arial, sans-serif';
        if (fontFamily === 'courier') cssFont = '"Courier New", monospace';
        if (fontFamily === 'georgia') cssFont = 'Georgia, serif';
        if (fontFamily === 'verdana') cssFont = 'Verdana, sans-serif';
        if (fontFamily === 'sinhala') cssFont = '"Iskoola Pota", "Nirmala UI", sans-serif';

        paperPreview.style.fontFamily = cssFont;

        // Text Colors
        prevUniversity.style.color = color;
        prevAssignment.style.color = color;
        prevAssignment.style.fontSize = `${titleSize}px`; // dynamic size

        // Faculty is dark grey
        prevFaculty.style.color = '#333';

        // Border Logic for Preview
        const contentBox = document.querySelector('.preview-content');
        if (borderStyle === 'solid') {
            contentBox.style.border = `2px solid ${color}`;
        }
    }

    // Add listeners to regular inputs
    [universityInput, facultyInput, moduleInput, assignmentTitleInput, studentNameInput, studentIdInput, submissionDateInput, themeColorInput, pageColorInput, fontFamilyInput, titleCaseInput, titleSizeInput].forEach(input => {
        input.addEventListener('input', updatePreview);
    });

    // --- Logo Upload Handling ---
    logoUploadInput.addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (file) {
            fileNameDisplay.textContent = file.name;

            const reader = new FileReader();
            reader.onload = function (event) {
                logoBase64 = event.target.result;
                // Update Preview Image
                previewLogoContainer.innerHTML = `<img src="${logoBase64}" alt="University Logo" style="max-width: 100%; max-height: 100%;">`;
            };
            reader.readAsDataURL(file);
        } else {
            fileNameDisplay.textContent = "No file chosen";
            logoBase64 = null;
            previewLogoContainer.innerHTML = `<i class="fa-solid fa-university"></i>`;
        }
    });

    // --- PDF Generation ---
    document.getElementById('btnPdf').addEventListener('click', async () => {
        // Warning for Sinhala
        if (fontFamilyInput.value === 'sinhala') {
            alert("Note: PDF generation for Sinhala characters may not render correctly due to browser limitations. For best results with Sinhala text, please use the 'Download Word' option.");
        }

        // Import jsPDF from window
        const { jsPDF } = window.jspdf;

        // Create new PDF (A4 Portrait)
        const doc = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 20;
        const centerX = pageWidth / 2;

        // Background
        const pageColor = pageColorInput.value;
        if (pageColor !== '#ffffff') {
            doc.setFillColor(pageColor);
            doc.rect(0, 0, pageWidth, pageHeight, 'F');
        }

        // Font Family Map
        // jsPDF standard fonts: times, helvetica, courier
        let pdfFont = 'times';
        if (fontFamilyInput.value === 'helvetica' || fontFamilyInput.value === 'verdana') pdfFont = 'helvetica';
        if (fontFamilyInput.value === 'courier') pdfFont = 'courier';

        doc.setFont(pdfFont, "bold"); // Standard academic font

        // Colors
        const hexColor = themeColorInput.value;
        doc.setTextColor(hexColor);
        doc.setDrawColor(hexColor);

        // 1. Border (Always Solid now)
        doc.setLineWidth(1.5); // Slightly thicker for professional look
        doc.rect(margin, margin, pageWidth - (margin * 2), pageHeight - (margin * 2));

        let currentY = 40;

        // 2. Logo
        if (logoBase64) {
            const logoWidth = 30;
            const logoHeight = 30; // Aspect ratio adjustment could be better but keeping simple
            doc.addImage(logoBase64, 'JPEG', centerX - (logoWidth / 2), currentY, logoWidth, logoHeight);
            currentY += 40;
        } else {
            currentY += 10; // Extra spacing if no logo
        }

        // 3. University Name
        doc.setFontSize(22);
        // Text Color set above
        doc.text(universityInput.value.toUpperCase(), centerX, currentY, { align: 'center' });

        currentY += 15;

        // 4. Faculty
        doc.setFontSize(16);
        doc.setFont(pdfFont, "normal");
        doc.setTextColor(50, 50, 50);
        doc.text(facultyInput.value, centerX, currentY, { align: 'center' });

        // Space for middle content
        currentY += 60;

        // 5. Assignment Title
        const titleSize = parseInt(titleSizeInput.value);
        doc.setFontSize(titleSize); // Use slider value
        doc.setFont(pdfFont, "bold");
        doc.setTextColor(hexColor); // Theme color for Title

        // Handle Case
        let titleText = assignmentTitleInput.value;
        const caseStyle = titleCaseInput.value;
        if (caseStyle === 'uppercase') titleText = titleText.toUpperCase();
        else if (caseStyle === 'lowercase') titleText = titleText.toLowerCase();
        else if (caseStyle === 'capitalize') {
            // Simple capitalize
            titleText = titleText.replace(/\b\w/g, l => l.toUpperCase());
        }

        doc.text(titleText, centerX, currentY, { align: 'center' });

        // Underline or Decoration
        doc.setLineWidth(0.5);
        doc.setDrawColor(hexColor);
        const titleWidth = doc.getTextWidth(titleText);
        doc.line(centerX - (titleWidth / 2) - 5, currentY + 2, centerX + (titleWidth / 2) + 5, currentY + 2);

        currentY += 20;

        // 6. Module
        doc.setFontSize(18);
        doc.setTextColor(0, 0, 0); // Black for Module
        doc.text(moduleInput.value, centerX, currentY, { align: 'center' });

        // Push details to bottom
        currentY = pageHeight - 80;

        // 7. Student Details (Left Aligned but centered block roughly)
        doc.setFontSize(14);
        doc.setFont(pdfFont, "normal");

        const detailsX = margin + 20;
        const valueX = margin + 60;
        const lineHeight = 10;

        doc.text("Name:", detailsX, currentY);
        doc.text(studentNameInput.value, valueX, currentY);

        currentY += lineHeight;

        doc.text("Index No:", detailsX, currentY);
        doc.text(studentIdInput.value, valueX, currentY);

        currentY += lineHeight;

        if (submissionDateInput.value) {
            doc.text("Date:", detailsX, currentY);
            doc.text(submissionDateInput.value, valueX, currentY);
        }

        // Footer Text
        doc.setFontSize(10);
        doc.setTextColor(150);
        // Watermark removed as requested

        // Save
        const fileName = `CoverSheet_${studentIdInput.value || 'Assignment'}.pdf`;
        doc.save(fileName);
    });

    // --- DOCX Generation ---
    document.getElementById('btnDocx').addEventListener('click', async () => {
        const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, WidthType, BorderStyle } = window.docx;

        // Font Mapping
        let docxFont = "Times New Roman";
        const selectedFont = fontFamilyInput.value;
        if (selectedFont === 'helvetica') docxFont = "Arial";
        if (selectedFont === 'courier') docxFont = "Courier New";
        if (selectedFont === 'georgia') docxFont = "Georgia";
        if (selectedFont === 'verdana') docxFont = "Verdana";
        if (selectedFont === 'sinhala') docxFont = "Iskoola Pota";

        // Note: DOCX generation in browser has limitations compared to PDF canvas.
        // We will do a best effort structured layout.

        const children = [];

        // Logo
        if (logoBase64) {
            // Need to convert dataURL to buffer/uint8array for docx?
            // docx.js handles base64 string directly usually needs clean base64
            const cleanBase64 = logoBase64.split(',')[1];
            children.push(
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: cleanBase64,
                            transformation: { width: 100, height: 100 },
                            type: "jpg" // Assuming jpg/png
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 }
                })
            );
        }

        // University
        children.push(
            new Paragraph({
                children: [new TextRun({
                    text: universityInput.value.toUpperCase(),
                    bold: true,
                    size: 44, // Half-points 
                    color: themeColorInput.value.replace('#', ''),
                    font: docxFont
                })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            })
        );

        // Faculty
        children.push(
            new Paragraph({
                children: [new TextRun({
                    text: facultyInput.value,
                    size: 32,
                    color: "333333",
                    font: docxFont
                })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 1200 } // Gap
            })
        );

        // Assignment Title
        let titleText = assignmentTitleInput.value;
        const caseStyle = titleCaseInput.value;
        if (caseStyle === 'uppercase') titleText = titleText.toUpperCase();
        else if (caseStyle === 'lowercase') titleText = titleText.toLowerCase();
        else if (caseStyle === 'capitalize') {
            // Simple capitalize
            titleText = titleText.replace(/\b\w/g, l => l.toUpperCase());
        }

        children.push(
            new Paragraph({
                children: [new TextRun({
                    text: titleText,
                    bold: true,
                    size: parseInt(titleSizeInput.value) * 2, // docx uses half-pts 
                    underline: {},
                    color: themeColorInput.value.replace('#', ''),
                    font: docxFont
                })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );

        // Module
        children.push(
            new Paragraph({
                children: [new TextRun({
                    text: moduleInput.value,
                    bold: true,
                    size: 36,
                    font: docxFont
                })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 2000 } // Big Gap to bottom
            })
        );

        // Details
        const createDetailLine = (label, value) => {
            return new Paragraph({
                children: [
                    new TextRun({ text: label + "\t", bold: true, size: 28, font: docxFont }),
                    new TextRun({ text: value, size: 28, font: docxFont })
                ],
                tabStops: [{ position: 2000, type: "left" }],
                indent: { left: 1440 } // 1 inch indent
            });
        };

        children.push(createDetailLine("Name:", studentNameInput.value));
        children.push(createDetailLine("Index No:", studentIdInput.value));
        children.push(createDetailLine("Date:", submissionDateInput.value));

        const doc = new Document({
            sections: [{
                properties: {},
                children: children
            }]
        });

        Packer.toBlob(doc).then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            document.body.appendChild(a);
            a.style = "display: none";
            a.href = url;
            a.download = `CoverSheet_${studentIdInput.value || 'Assignment'}.docx`;
            a.click();
            window.URL.revokeObjectURL(url);
        });
    });

    // Set Default Date to Today
    submissionDateInput.valueAsDate = new Date();
    updatePreview();
});
