// Insert table into slide XML when no original table exists
function insertTableIntoSlide(slideXml, tableData, element) {
  if (!tableData || !Array.isArray(tableData) || tableData.length === 0) {
    tableData = Array(5).fill().map((_, i) => 
      Array(5).fill().map((_, j) => `Cell ${i+1}-${j+1}`)
    );
  }
  
  const x = Math.round((element.x || 1) * 914400);
  const y = Math.round((element.y || 1) * 914400);
  const w = Math.round((element.width || 6) * 914400);
  const h = Math.round((element.height || 3) * 914400);
  const colW = Math.round(w / tableData[0].length);
  
  const tableXml = `
    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="${Date.now()}" name="Table"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="${x}" y="${y}"/>
        <a:ext cx="${w}" cy="${h}"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid>
              ${tableData[0].map(() => `<a:gridCol w="${colW}"/>`).join('')}
            </a:tblGrid>
            ${tableData.map((row, rowIndex) => {
              const bgColor = rowIndex % 2 === 0 ? 'FFFFFF' : 'FF0000';
              return `
            <a:tr h="${Math.round(h/tableData.length)}">
              ${row.map(cell => `
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:r>
                      <a:t>${cell || ''}</a:t>
                    </a:r>
                  </a:p>
                </a:txBody>
                <a:tcPr>
                  <a:solidFill>
                    <a:srgbClr val="${bgColor}"/>
                  </a:solidFill>
                </a:tcPr>
              </a:tc>`).join('')}
            </a:tr>`;
            }).join('')}
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>`;
  
  return slideXml.replace('</p:spTree>', tableXml + '</p:spTree>');
}

// Replace table content in XML
function replaceTableInXml(slideXml, newTableData, originalElement) {
  if (!newTableData || !Array.isArray(newTableData) || newTableData.length === 0) {
    newTableData = Array(5).fill().map((_, i) => 
      Array(5).fill().map((_, j) => `Cell ${i+1}-${j+1}`)
    );
  }

  const tableRegex = /<a:tbl[\s\S]*?<\/a:tbl>/g;
  const colW = 1200000; // Adjust column width if needed
  const rowH = 370840;

  return slideXml.replace(tableRegex, () => {
    const newTableXml = `<a:tbl>
      <a:tblPr firstRow="1" bandRow="1">
        <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
      </a:tblPr>
      <a:tblGrid>
        ${newTableData[0].map(() => `<a:gridCol w="${colW}"/>`).join('')}
      </a:tblGrid>
      ${newTableData.map((row, rowIndex) => {
        const bgColor = rowIndex % 2 === 0 ? 'FFFFFF' : 'FF0000';
        return `
        <a:tr h="${rowH}">
          ${row.map(cell => `
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:t>${cell || ''}</a:t>
                </a:r>
              </a:p>
            </a:txBody>
            <a:tcPr>
              <a:solidFill>
                <a:srgbClr val="${bgColor}"/>
              </a:solidFill>
            </a:tcPr>
          </a:tc>`).join('')}
        </a:tr>`;
      }).join('')}
    </a:tbl>`;

    return newTableXml;
  });
}



module.exports = {
  insertTableIntoSlide,
  replaceTableInXml
};