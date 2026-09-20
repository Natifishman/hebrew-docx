// An invoice in Hebrew: RTL table, numbers that stay left-to-right, totals.
const { HebrewDoc, heading, p, table, run, ltr } = require('../src');

const items = [
  ['ייעוץ אפיון מערכת', '12', '₪450', '₪5,400'],
  ['פיתוח מודול דוחות', '30', '₪380', '₪11,400'],
  ['בדיקות קבלה', '8', '₪320', '₪2,560'],
];
const total = '₪19,360';

const doc = new HebrewDoc({ font: 'David', size: 12 });
doc.footer('חשבונית 2026-0042');

doc.add(
  heading('חשבונית מס', 1),
  p([
    run('מספר חשבונית: '), run(ltr('2026-0042'), { bold: true }),
    run('   |   תאריך: '), run(ltr('20/09/2026')),
  ]),
  p('לכבוד: חברת אלפא בע״מ, רחוב הברזל 12, תל אביב'),
  p(''),
  table(
    ['תיאור השירות', 'שעות', 'תעריף', 'סה״כ'],
    items,
    { widths: [55, 12, 15, 18], headerFill: 'D9D9D9' },
  ),
  p(''),
  p([run('סה״כ לתשלום: '), run(total, { bold: true, ltr: true })], { align: 'left' }),
  p('תנאי תשלום: שוטף + 30. המחירים כוללים מע״מ.', { size: 11 }),
);

doc.save('invoice.docx').then((f) => console.log('wrote', f));
