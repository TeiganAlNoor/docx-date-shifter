// Simple DOCX sample creator for testing
// This creates the minimal structure for a DOCX file with sample date ranges

const createSampleDocx = () => {
  // DOCX structure files
  const files = {
    // Content Types
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,

    // Main relationships
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,

    // Word relationships
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`,

    // Main document with sample date ranges
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>جدول الأسابيع - Sample Weekly Schedule</w:t>
            </w:r>
        </w:p>
        <w:tbl>
            <w:tblPr>
                <w:tblStyle w:val="TableGrid"/>
                <w:tblW w:w="0" w:type="auto"/>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الأول</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>6/9-12/9</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة المرسلات</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الثاني</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>13/9-19/9</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة المجادلة الآية (8-1)</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الثالث</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>20/9-26/9</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة الحشر (21-10)</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الرابع</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>27/9-3/10</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة الممتحنة (13-6)</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
        <w:p>
            <w:r>
                <w:t>Additional test ranges: 15/9-21/9 and 22/9-28/9</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>`
  };

  return files;
};

// Function to create the sample ZIP when called from the main app
if (typeof window !== 'undefined' && window.createSampleZip) {
  window.createSampleDocxFiles = createSampleDocx;
}