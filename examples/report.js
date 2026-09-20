// A report: headings, lists, a landscape section for a wide table,
// and mixed Hebrew/Latin text that keeps its direction.
const { HebrewDoc, heading, p, bullet, numbered, table, run, ltr, isolate } = require('../src');

const doc = new HebrewDoc({ font: 'David', size: 12, margin: 2 });
doc.footer('דוח רבעוני Q3');

doc.add(
  heading('דוח פעילות רבעוני', 1),
  p('הדוח מסכם את פעילות המערכת ברבעון השלישי ומציג את ממצאי הבדיקות.', { align: 'justify' }),

  heading('ממצאים עיקריים', 2),
  bullet('זמן התגובה הממוצע ירד ב-18% לעומת הרבעון הקודם'),
  bullet(isolate('שיעור השגיאות בשירות ה-API עומד על 0.4% (יעד: מתחת ל-1%)')),
  bullet('שלוש תקלות קריטיות נסגרו בתוך פחות מארבע שעות'),

  heading('משימות לרבעון הבא', 2),
  numbered('השלמת המעבר לסביבת הענן החדשה'),
  numbered('הטמעת ניטור אוטומטי לכל השירותים'),
  numbered(['שדרוג גרסת ', run(ltr('Node.js 18 → 22'))]),
);

// Wide tables read better in landscape.
doc.section({ landscape: true });
doc.add(
  heading('פירוט לפי צוות', 2),
  table(
    ['צוות', 'משימות שנסגרו', 'זמן ממוצע (שעות)', 'שביעות רצון', 'הערות'],
    [
      ['תשתיות', '42', '3.1', '4.6', 'עומס בחודש אוגוסט'],
      ['אפליקציה', '67', '2.4', '4.8', '-'],
      ['נתונים', '19', '5.7', '4.2', 'שתי משימות נדחו לרבעון הבא'],
    ],
    { widths: [16, 18, 20, 16, 30], width: 25, headerFill: 'D9D9D9' },
  ),
);

doc.save('report.docx').then((f) => console.log('wrote', f));
