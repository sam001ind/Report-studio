// Official Kannur University affiliated colleges & departments default email directory
export const DEFAULT_COLLEGE_EMAILS = [
  { code: '200', name: 'Govt. College Mananthavady', email: 'gcmananthavady@gmail.com' },
  { code: '201', name: 'Govt. Brennen College, Thalassery', email: 'brennencollege@gmail.com' },
  { code: '205', name: 'Govt. College Kasaragod', email: 'gckasaragod@gmail.com' },
  { code: '207', name: "K.M.M. Govt. Women's College, Kannur", email: 'kmmgwc@gmail.com' },
  { code: '209', name: 'Payyannur College, Payyannur', email: 'payyannurcollege@gmail.com' },
  { code: '211', name: 'Sir Syed College, Taliparamba', email: 'sirsyedcollege@gmail.com' },
  { code: '212', name: 'Sree Narayana College, Kannur', email: 'sncollegekannur@gmail.com' },
  { code: '214', name: 'Nirmalagiri College, Kuthuparamba', email: 'nirmalagiricollege@gmail.com' },
  { code: '215', name: 'Mahatma Gandhi College, Iritty', email: 'mgciritty@gmail.com' },
  { code: '216', name: 'N.A.M. College, Kallikandy', email: 'namcollege@gmail.com' },
  { code: '241', name: 'Nehru Arts and Science College, Kanhangad', email: 'nehrucollege@gmail.com' },
  { code: '251', name: 'St. Pius X College, Rajapuram', email: 'stpiusrpm@gmail.com' },
  { code: '265', name: 'SES College, Sreekandapuram', email: 'sescollege.skp@gmail.com' },
  { code: '266', name: 'Mary Matha Arts and Science College, Mananthavady', email: 'marymathacollege@gmail.com' },
  { code: '430', name: 'Department of Information Technology, Mangattuparamba', email: 'itdept@kannuruniv.ac.in' },
  { code: '432', name: 'School of Behavioural Sciences, Mangattuparamba', email: 'behavioural@kannuruniv.ac.in' },
  { code: '433', name: 'Department of Physics, Swami Anandatheertha Campus', email: 'physics@kannuruniv.ac.in' },
  { code: '434', name: 'Department of Commerce and Business Studies', email: 'commerce@kannuruniv.ac.in' },
  { code: '435', name: 'Department of Anthropology', email: 'anthropology@kannuruniv.ac.in' }
];

// Strict DD/MM/YYYY date normalizer
export const formatExamDate = (raw) => {
  if (raw === null || raw === undefined || raw === '') return '';
  const s = String(raw).trim();
  if (!s) return '';

  // Excel serial number (e.g. 45600)
  if (/^\d{5}$/.test(s)) {
    const serial = parseInt(s, 10);
    if (serial > 30000 && serial < 60000) {
      const utcDays = serial - 25569;
      const d = new Date(utcDays * 86400 * 1000);
      d.setHours(12, 0, 0, 0); // avoid UTC offset shifting date
      const dd = String(d.getDate()).padStart(2, '0');
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const yyyy = d.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }
  }

  // Already DD/MM/YYYY
  const ddmmyyyy = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (ddmmyyyy) {
    const [, d, m, y] = ddmmyyyy;
    return `${d.padStart(2, '0')}/${m.padStart(2, '0')}/${y}`;
  }

  // YYYY-MM-DD
  const yyyymmdd = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (yyyymmdd) {
    const [, y, m, d] = yyyymmdd;
    return `${d.padStart(2, '0')}/${m.padStart(2, '0')}/${y}`;
  }

  const dt = new Date(s);
  if (!isNaN(dt.getTime())) {
    const dd = String(dt.getDate()).padStart(2, '0');
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const yyyy = dt.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  return s;
};

export const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');
