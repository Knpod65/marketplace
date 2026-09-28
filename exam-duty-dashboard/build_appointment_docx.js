// Per-person invigilator appointment letters (Word), in the layout of the faculty order
// that appoints exam proctors: order text, the person's own duty table, signature block.
//
// Usage:
//   node build_appointment_docx.js data.json out.docx            one page per person, filled in
//   node build_appointment_docx.js --template out_template.docx  blank template with {{placeholders}}
//
// Needs the `docx` npm package (npm install docx).
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, VerticalAlign, HeightRule,
} = require('docx');

// ---- order details (edit here for a new round) ----
const ORDER = {
  faculty: 'คณะรัฐศาสตร์และรัฐประศาสนศาสตร์',
  no: '122/2569',
  subject: 'แต่งตั้งกรรมการคุมสอบปลายภาค ประจำภาคการศึกษาที่ 1 ปีการศึกษา 2569',
  term: 'ภาคการศึกษาที่ 1 ปีการศึกษา 2569',
  examDates: 'ระหว่างวันที่ 19 – 22 ตุลาคม และวันที่ 24 ตุลาคม – 2 พฤศจิกายน พ.ศ. 2569',
  orderedOn: '21 สิงหาคม พ.ศ. 2569',
  signer: 'รองศาสตราจารย์ ดร.ไพลิน ภู่จีนาพันธุ์',
  signerTitle: 'คณบดีคณะรัฐศาสตร์และรัฐประศาสนศาสตร์',
  contact: 'งานบริการการศึกษา คณะรัฐศาสตร์และรัฐประศาสนศาสตร์ โทร. ..............................',
};

const FONT = 'TH SarabunPSK';
// PREVIEW_SCALE only for rendering previews with a wider stand-in font; leave unset for real output.
const SCALE = Number(process.env.PREVIEW_SCALE || 1);
const PT = 16 * SCALE;    // body size (pt)
const PT_TABLE = 15 * SCALE; // table size (pt)
const PAGE_W = 11906, M_LEFT = 1418, M_RIGHT = 1134;   // A4, 2.5 cm / 2 cm
const CONTENT_W = PAGE_W - M_LEFT - M_RIGHT;           // 9354
const COLS = [800, 2050, 1350, 1850, 1804, 1500];      // sums to CONTENT_W
const INDENT = 1134;                                   // 2 cm first-line indent

const DOW = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์'];
const MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const ROLE = { proctor: 'คุมสอบ', own: 'คุมสอบ (อาจารย์ผู้สอน)', dist: 'จ่ายข้อสอบ', opener: 'เปิดห้องสอบ' };

// ---- small builders ----
const run = (text, o = {}) => new TextRun({ text, font: FONT, size: Math.round((o.pt || PT) * 2), bold: o.bold, highlight: o.hl });
const para = (children, o = {}) => new Paragraph({
  children: (Array.isArray(children) ? children : [children]).map(c => typeof c === 'string' ? run(c, o) : c),
  alignment: o.align || AlignmentType.LEFT,
  indent: o.indent,
  spacing: { before: o.before || 0, after: o.after || 0, line: 240 },
  keepNext: o.keepNext,
  pageBreakBefore: o.pageBreakBefore,
});
const blank = (pt = PT) => para('', { pt });
const justified = (text, o = {}) => para(text, { align: AlignmentType.THAI_DISTRIBUTE, ...o });

const border = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(lines, width, o = {}) {
  lines = Array.isArray(lines) ? lines : [lines];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 20, bottom: 20, left: 80, right: 80 },
    children: lines.map(l => para(l, { pt: PT_TABLE, bold: o.bold, align: o.align, hl: o.hl })),
  });
}
const HEAD = ['ลำดับ', 'วัน', 'เวลา', 'ห้องสอบ', 'กระบวนวิชา/ตอน', 'หน้าที่'];
function dutyTable(rows, hl) {
  const C = AlignmentType.CENTER;
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: COLS,
    rows: [
      new TableRow({ tableHeader: true, children: HEAD.map((h, i) => cell(h, COLS[i], { bold: true, align: C })) }),
      ...rows.map((r, n) => new TableRow({
        cantSplit: true,
        height: { value: 0, rule: HeightRule.AUTO },
        children: [
          cell(String(r.no ?? n + 1), COLS[0], { align: C, hl }),
          cell(r.day, COLS[1], { hl }),
          cell(r.time, COLS[2], { align: C, hl }),
          cell(r.rooms, COLS[3], { hl }),
          cell(r.courses, COLS[4], { hl }),
          cell(r.role, COLS[5], { hl }),
        ],
      })),
    ],
  });
}

// ---- one letter ----
function letter(name, rows, { hl } = {}) {
  const body = [
    para('- - สำเนา - -', { align: AlignmentType.CENTER }),
    para(`คำสั่ง${ORDER.faculty}`, { align: AlignmentType.CENTER, bold: true }),
    para(`ที่ ${ORDER.no}`, { align: AlignmentType.CENTER, bold: true }),
    blank(),
    para(`เรื่อง ${ORDER.subject}`, { bold: true }),
    blank(10),
    justified(`ตามที่มหาวิทยาลัยเชียงใหม่ได้กำหนดวันสอบปลายภาค ประจำ${ORDER.term} ${ORDER.examDates} `
      + `เพื่อให้การสอบปลายภาค ประจำ${ORDER.term} ดำเนินไปด้วยความเรียบร้อย อาศัยอำนาจตามความในมาตรา 40 `
      + 'แห่งพระราชบัญญัติมหาวิทยาลัยเชียงใหม่ พ.ศ.2551 ประกอบกับข้อ 6 และข้อ 23 แห่งข้อบังคับมหาวิทยาลัยเชียงใหม่'
      + 'ว่าด้วยชื่อตำแหน่ง คุณสมบัติเฉพาะตำแหน่ง วาระการดำรงตำแหน่ง การพ้นจากตำแหน่ง และอำนาจและหน้าที่ของรองอธิการบดี '
      + 'ผู้ช่วยอธิการบดี หัวหน้าส่วนงาน รองหัวหน้าส่วนงาน และตำแหน่งบริหารอื่นในส่วนงานวิชาการและส่วนงานอื่น พ.ศ.2566 '
      + `${ORDER.faculty} จึงแต่งตั้งบุคคลผู้มีรายชื่อต่อไปนี้ เป็นกรรมการคุมสอบตามวัน – เวลา และสถานที่ที่ปรากฏในตาราง `
      + 'และให้นักศึกษาทุกคนปฏิบัติ ดังนี้.-', { indent: { firstLine: INDENT } }),
    justified('1. ให้นักศึกษาแต่งกายเข้าห้องสอบตามข้อบังคับมหาวิทยาลัยเชียงใหม่ ว่าด้วยการสอบ พ.ศ. 2564', { indent: { firstLine: INDENT } }),
    justified('2. ห้ามนักศึกษาเข้าห้องสอบหลังจากเริ่มสอบไปแล้ว 20 นาที และห้ามนักศึกษาออกจากห้องสอบภายใน 30 นาที หลังจากเริ่มสอบ', { indent: { firstLine: INDENT } }),
    justified('3. ในกรณีที่มีการทุจริตในการสอบ จะมีการลงโทษตามข้อบังคับมหาวิทยาลัยเชียงใหม่ ว่าด้วยวินัยและการดำเนินการทางวินัยนักศึกษา พ.ศ. 2564 ดังนี้.-', { indent: { firstLine: INDENT } }),
    justified('ก. พักการศึกษา มีกำหนดตั้งแต่หนึ่งภาคการศึกษาถึงสี่ภาคการศึกษา หรือระงับการเสนอชื่อให้สำเร็จการศึกษา มีกำหนดตั้งแต่หนึ่งภาคการศึกษาถึงสี่ภาคการศึกษา ในกรณีที่ไม่สามารถพักการศึกษาได้ และให้ลำดับขั้น F ในกระบวนวิชานั้นพร้อมการทำงานบริการสังคมหรือสาธารณประโยชน์ด้วย', { indent: { firstLine: INDENT + 284 } }),
    justified('ข. ลบชื่อออกจากการเป็นนักศึกษา', { indent: { firstLine: INDENT + 284 } }),
    justified('4. ให้นักศึกษาปิดเครื่องมือสื่อสารทุกชนิด', { indent: { firstLine: INDENT } }),
    blank(),
    para([run('กรรมการคุมสอบ '), run(name, { bold: true, hl })], { keepNext: true }),
    para([run('สังกัดคณะ/ภาควิชา '), run(ORDER.faculty, { bold: true })], { keepNext: true, after: 60 }),
    dutyTable(rows, hl),
    blank(),
    justified('ทั้งนี้ ให้กรรมการคุมสอบปฏิบัติหน้าที่ตามข้อบังคับมหาวิทยาลัยเชียงใหม่ ว่าด้วยการสอบ พ.ศ. 2564 '
      + `และแนวปฏิบัติในการจัดสอบของ${ORDER.faculty} โดยเคร่งครัด`, { indent: { firstLine: INDENT } }),
    para(`สั่ง ณ วันที่ ${ORDER.orderedOn}`, { align: AlignmentType.CENTER }),
    blank(),
    ...[`(ลงนาม) ${ORDER.signer}`, `(${ORDER.signer})`, ORDER.signerTitle]
      .map(t => para(t, { align: AlignmentType.CENTER, indent: { left: 3600 }, keepNext: true })),
    para('', { keepNext: true }),
    para('กรณีมีปัญหาในการคุมสอบ กรุณาติดต่อ', { keepNext: true }),
    para([run(ORDER.contact.split(' โทร.')[0], { bold: true }), run(' โทร.' + ORDER.contact.split(' โทร.')[1])]),
  ];
  return body;
}

function makeDoc(pages) {
  const children = [];
  pages.forEach((p, i) => {
    if (i) p[0] = para('- - สำเนา - -', { align: AlignmentType.CENTER, pageBreakBefore: true });
    children.push(...p);
  });
  return new Document({
    styles: { default: { document: { run: { font: FONT, size: Math.round(PT * 2) } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1134, bottom: 851, left: M_LEFT, right: M_RIGHT } } },
      children,
    }],
  });
}

// ---- data -> rows per person ----
function thaiDate(s) {
  const d = new Date(s + 'T00:00:00');
  return `${DOW[(d.getDay() + 6) % 7]} ${d.getDate()} ${MON[d.getMonth()]} ${d.getFullYear() + 543}`;
}
const fmtTime = t => t.replace('-', ' - ');
const fmtRoom = r => r.startsWith('ECON') ? 'คณะเศรษฐศาสตร์ (แจ้งเลขห้องภายหลัง)' : r.replace(/^([A-Z]+)(\d)/, '$1 $2');
function secLabel(e) {   // CMU style: course lec lab
  const k = String(e.key).split('|');
  return k.length === 4 ? `${k[1]} ${k[2]} ${k[3]}` : `${e.code} ${e.sec}`;
}
// "AUD 50, PSB 1101, 1201, 1204" - building prefix written once per building
function compactRooms(rooms) {
  let last = '';
  return rooms.map(fmtRoom).map(r => {
    const [b, n] = r.split(' ');
    const out = b === last && n ? n : r;
    last = b;
    return out;
  }).join(', ');
}
const cleanName = n => n.replace(/\s*\(.*\)\s*$/, '').replace(/^(นางสาว|นาง|นาย)/, '');

function rowsFor(data, pid) {
  const bldg = data.roomBldg || {};
  const slotRooms = (date, time) => Object.keys(data.rooms).filter(k => k.startsWith(`${date}|${time}|`)).map(k => k.split('|')[2]);
  const out = [];
  const mine = data.duties.filter(d => d.pid === pid).sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
  for (const d of mine) {
    if (d.role === 'proctor' || d.role === 'own') {
      const secs = (data.rooms[`${d.date}|${d.time}|${d.room}`] || []).map(secLabel);
      out.push({ day: thaiDate(d.date), time: fmtTime(d.time), rooms: fmtRoom(d.room), courses: secs.length ? secs : ['–'], role: ROLE[d.role] });
    } else if (d.role === 'dist') {
      const name = data.people.find(p => p.id === pid).name;
      const b = (data.distBldg || {})[`${d.date}|${d.time}|${name}`] || 'PS';
      const rooms = slotRooms(d.date, d.time).filter(r => (bldg[r] || 'PS') === b);
      out.push({ day: thaiDate(d.date), time: fmtTime(d.time),
        rooms: b === 'EC' ? ['คณะเศรษฐศาสตร์ (แจ้งเลขห้องภายหลัง)'] : [compactRooms(rooms.sort())],
        courses: ['ทุกกระบวนวิชาที่สอบในช่วงเวลานี้'], role: ROLE.dist });
    } else if (d.role === 'opener') {  // one duty per day: open every PS room used that day
      const keys = Object.keys(data.rooms).filter(k => k.startsWith(d.date + '|') && (bldg[k.split('|')[2]] || 'PS') === 'PS');
      const times = [...new Set(keys.map(k => k.split('|')[1]))].sort();
      const rooms = [...new Set(keys.map(k => k.split('|')[2]))].sort();
      const start = times[0].split('-')[0], end = times.map(t => t.split('-')[1]).sort().pop();
      out.push({ day: thaiDate(d.date), time: `${start} - ${end}`, rooms: [compactRooms(rooms)], courses: ['–'], role: ROLE.opener });
    }
  }
  return out;
}

async function main() {
  const args = process.argv.slice(2);
  let doc;
  if (args[0] === '--template') {
    const ph = { day: '{{วัน}}', time: '{{เวลา}}', rooms: '{{ห้องสอบ}}', courses: '{{กระบวนวิชา/ตอน}}', role: '{{หน้าที่}}' };
    doc = makeDoc([letter('{{ชื่อ-สกุล}}', [1, 2, 3].map(no => ({ ...ph, no })), { hl: 'yellow' })]);
    fs.writeFileSync(args[1], await Packer.toBuffer(doc));
    console.log('wrote template', args[1]);
    return;
  }
  const data = JSON.parse(fs.readFileSync(args[0], 'utf8'));
  const people = data.people.filter(p => data.duties.some(d => d.pid === p.id));
  doc = makeDoc(people.map(p => letter(cleanName(p.name), rowsFor(data, p.id))));
  fs.writeFileSync(args[1], await Packer.toBuffer(doc));
  console.log(`wrote ${args[1]}: ${people.length} letters`);
}
main();
