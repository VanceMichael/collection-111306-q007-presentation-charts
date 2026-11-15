import {deflateRawSync} from 'node:zlib';
import {crc32} from '../../src/zip.mjs';

/**
 * Minimal deterministic ZIP writer used only by tests to build realistic
 * OOXML packages (presentation or workbook) from parts.
 */
export function buildZip(files) {
  const chunks = [];
  const directory = [];
  let offset = 0;
  const encoder = new TextEncoder();
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = Buffer.from(name, 'utf8');
    const raw = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const compressed = deflateRawSync(raw);
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28);
    chunks.push(local, nameBytes, compressed);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBytes.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    directory.push(central, nameBytes);
    offset += local.length + nameBytes.length + compressed.length;
  }
  const directoryBuffer = Buffer.concat(directory);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  const count = Object.keys(files).length;
  end.writeUInt16LE(count, 8);
  end.writeUInt16LE(count, 10);
  end.writeUInt32LE(directoryBuffer.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, directoryBuffer, end]);
}

export function rels(entries) {
  return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map(entry => `<Relationship Id="${entry.id}" Type="${entry.type}" Target="${entry.target}"${entry.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
}

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function slideXml({charts = [], show = null} = {}) {
  const frames = charts.map((chart, index) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${chart.name ?? `图表${index + 1}`}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></p:xfrm><p:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${chart.id}"/></a:graphicData></p:graphic></p:graphicFrame>`).join('');
  return `<p:sld ${NS}${show === false ? ' show="0"' : ''}><p:cSld><p:spTree>${frames}</p:spTree></p:cSld></p:sld>`;
}

export function chartXml({series = [], externalData = 'rIdBook'} = {}) {
  const body = series.map((entry, index) => {
    const name = entry.name === undefined ? '' : `<c:tx><c:strRef><c:f>${entry.name.formula}</c:f><c:strCache><c:ptCount val="${entry.name.ptCount ?? entry.name.points.length}"/>${entry.name.points.map(point => `<c:pt idx="${point.idx}"><c:v>${point.value}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:tx>`;
    const cat = entry.cat ? `<c:cat>${entry.cat.literal ? `<c:strLit><c:ptCount val="${entry.cat.ptCount ?? entry.cat.points.length}"/>${entry.cat.points.map(point => `<c:pt idx="${point.idx}"><c:v>${point.value}</c:v></c:pt>`).join('')}</c:strLit>` : `<c:strRef><c:f>${entry.cat.formula}</c:f><c:strCache><c:ptCount val="${entry.cat.ptCount ?? entry.cat.points.length}"/>${entry.cat.points.map(point => `<c:pt idx="${point.idx}"><c:v>${point.value}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`}</c:cat>` : '';
    const val = `<c:val>${entry.val.literal ? `<c:numLit><c:formatCode>0</c:formatCode><c:ptCount val="${entry.val.ptCount ?? entry.val.points.length}"/>${entry.val.points.map(point => `<c:pt idx="${point.idx}"><c:v>${point.value}</c:v></c:pt>`).join('')}</c:numLit>` : `<c:numRef><c:f>${entry.val.formula}</c:f><c:numCache><c:formatCode>0</c:formatCode><c:ptCount val="${entry.val.ptCount ?? entry.val.points.length}"/>${entry.val.points.map(point => `<c:pt idx="${point.idx}"><c:v>${point.value}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`}</c:val>`;
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${cat}${val}</c:ser>`;
  }).join('');
  return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:barChart>${body}</c:barChart></c:plotArea></c:chart>${externalData ? `<c:externalData r:id="${externalData}"/>` : ''}</c:chartSpace>`;
}

export function workbookParts({sheets, sharedStrings = null, workbookRelsExtra = ''}) {
  const files = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
    '_rels/.rels': rels([{id: 'rId1', type: `${REL}/officeDocument`, target: 'xl/workbook.xml'}]),
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"${sheet.state ? ` state="${sheet.state}"` : ''}/>`).join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([
      ...sheets.map((sheet, index) => ({id: `rId${index + 1}`, type: `${REL}/worksheet`, target: `worksheets/sheet${index + 1}.xml`})),
      ...(sharedStrings ? [{id: 'rIdShared', type: `${REL}/sharedStrings`, target: 'sharedStrings.xml'}] : []),
      ...workbookRelsExtra,
    ]),
  };
  sheets.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = sheet.xml;
  });
  if (sharedStrings) files['xl/sharedStrings.xml'] = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStrings.map(value => `<si><t>${value}</t></si>`).join('')}</sst>`;
  return files;
}

export function presentationPackage({slides, charts, workbooks = {}, extraParts = {}}) {
  const files = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>',
    '_rels/.rels': rels([{id: 'rIdOffice', type: `${REL}/officeDocument`, target: 'ppt/presentation.xml'}]),
    'ppt/presentation.xml': `<p:presentation ${NS}><p:sldIdLst>${slides.map((slide, index) => `<p:sldId id="${256 + index}" r:id="${slide.id ?? `rIdSlide${index + 1}`}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels(slides.map((slide, index) => ({id: slide.id ?? `rIdSlide${index + 1}`, type: `${REL}/slide`, target: slide.target ?? `slides/slide${index + 1}.xml`}))),
  };
  slides.forEach((slide, index) => {
    const target = slide.target ?? `slides/slide${index + 1}.xml`;
    files[`ppt/${target}`] = slide.xml ?? slideXml({charts: slide.charts ?? [], show: slide.show});
    files[`ppt/${target.replace(/([^/]+)$/, '_rels/$1.rels')}`] = rels(slide.rels ?? []);
  });
  for (const [name, xml] of Object.entries(charts)) {
    files[`ppt/charts/${name}`] = xml;
  }
  for (const [name, parts] of Object.entries(workbooks)) {
    files[name] = parts instanceof Buffer ? parts : buildZip(parts);
  }
  for (const [name, content] of Object.entries(extraParts)) files[name] = content;
  return buildZip(files);
}
